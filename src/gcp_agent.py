"""GCP Pub/Sub 에이전트 — Cloud Scheduler 메시지를 받아 src.main 을 실행한다.

실행 방법:
  python -m src.gcp_agent              # 정식 실행 (스트리밍 풀 루프)
  python -m src.gcp_agent --test daily # 테스트: Pub/Sub 없이 daily 직접 실행

GCP Cloud Scheduler → Pub/Sub 토픽(ai-rpa-daily / ai-rpa-monthly) →
  이 에이전트 구독 수신 → subprocess: python -m src.main daily/monthly

설계 원칙:
- 메시지 본문이 "daily" 또는 "monthly" 인지 확인 후 RPA 실행
- 이미 실행 중이면 ACK 후 건너뜀 (lockfile 기반)
- RPA 종료 후 항상 ACK (NACK → 즉시 재전달 → 무한 루프 방지)
- 재시도는 Cloud Scheduler 가 30분마다 다음 메시지를 발행하는 방식으로 처리
"""

from __future__ import annotations

import argparse
import logging
import subprocess
import sys
import time
from pathlib import Path

log = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOCKFILE = PROJECT_ROOT / "state" / "agent.lock"
LOCK_MAX_AGE_SECONDS = 3 * 3600  # 3시간 이상 된 락은 스테일로 간주


# ── 락파일 헬퍼 ───────────────────────────────────────────────────────

def _is_rpa_running() -> bool:
    """현재 RPA 프로세스가 실행 중인지 확인 (lockfile 기반)."""
    if not LOCKFILE.exists():
        return False
    age = time.time() - LOCKFILE.stat().st_mtime
    if age > LOCK_MAX_AGE_SECONDS:
        log.warning(f"[AGENT] 스테일 락파일 제거 (age={age/3600:.1f}h)")
        LOCKFILE.unlink(missing_ok=True)
        return False
    return True


def _acquire_lock() -> None:
    LOCKFILE.parent.mkdir(parents=True, exist_ok=True)
    LOCKFILE.touch()


def _release_lock() -> None:
    LOCKFILE.unlink(missing_ok=True)


# ── RPA 실행 ─────────────────────────────────────────────────────────

def run_rpa(mode: str) -> bool:
    """src.main daily 또는 monthly 를 subprocess 로 실행. 성공 시 True."""
    if mode not in ("daily", "monthly"):
        log.error(f"[AGENT] 잘못된 mode: {mode}")
        return False

    log.info(f"[AGENT] RPA 실행: mode={mode}")
    _acquire_lock()
    try:
        result = subprocess.run(
            [sys.executable, "-m", "src.main", mode],
            cwd=str(PROJECT_ROOT),
        )
        success = result.returncode == 0
        if success:
            log.info(f"[AGENT] RPA 완료: mode={mode}")
        else:
            log.warning(f"[AGENT] RPA 비정상 종료: returncode={result.returncode}")
        return success
    finally:
        _release_lock()


# ── Pub/Sub 콜백 ──────────────────────────────────────────────────────

def _make_callback():
    def callback(message) -> None:
        try:
            mode = message.data.decode("utf-8").strip().lower()
        except Exception as exc:
            log.error(f"[AGENT] 메시지 디코딩 실패: {exc}")
            message.ack()
            return

        log.info(f"[AGENT] 메시지 수신: mode={mode}, id={message.message_id}")

        if mode not in ("daily", "monthly"):
            log.warning(f"[AGENT] 알 수 없는 mode '{mode}' — 무시")
            message.ack()
            return

        if _is_rpa_running():
            log.info("[AGENT] 이미 실행 중 — 건너뜀")
            message.ack()
            return

        run_rpa(mode)
        message.ack()

    return callback


# ── 메인 루프 ─────────────────────────────────────────────────────────

def run_agent(config) -> None:
    """두 Pub/Sub 구독을 스트리밍 풀로 수신한다."""
    from google.cloud import pubsub_v1

    if not config.gcp.project_id:
        raise RuntimeError(
            "GCP_PROJECT_ID 가 설정되지 않았습니다. .env 파일을 확인하세요."
        )

    log.info("[AGENT] Pub/Sub 에이전트 시작")
    log.info(f"  project            : {config.gcp.project_id}")
    log.info(f"  daily subscription : {config.gcp.daily_subscription}")
    log.info(f"  monthly subscription: {config.gcp.monthly_subscription}")

    subscriber = pubsub_v1.SubscriberClient()
    daily_path = subscriber.subscription_path(
        config.gcp.project_id, config.gcp.daily_subscription
    )
    monthly_path = subscriber.subscription_path(
        config.gcp.project_id, config.gcp.monthly_subscription
    )

    callback = _make_callback()
    futures = []
    with subscriber:
        futures.append(subscriber.subscribe(daily_path, callback=callback))
        futures.append(subscriber.subscribe(monthly_path, callback=callback))
        log.info("[AGENT] 대기 중 (Ctrl+C 또는 서비스 중지로 종료)")
        try:
            while True:
                time.sleep(30)
        except KeyboardInterrupt:
            log.info("[AGENT] 중단됨")
        finally:
            for f in futures:
                f.cancel()
                try:
                    f.result(timeout=5)
                except Exception:
                    pass


# ── CLI 진입점 ────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="B2C RPA GCP Pub/Sub 에이전트")
    parser.add_argument(
        "--test",
        metavar="MODE",
        help="Pub/Sub 없이 직접 RPA 실행 (daily 또는 monthly)",
    )
    args = parser.parse_args()

    from .config import load_config
    from .logger import setup_logging

    try:
        config = load_config()
    except RuntimeError as exc:
        print(f"[ERROR] 설정 오류: {exc}", file=sys.stderr)
        sys.exit(2)

    setup_logging(config.runtime.logs_dir)

    if args.test:
        log.info(f"[AGENT] 테스트 모드: mode={args.test}")
        success = run_rpa(args.test)
        sys.exit(0 if success else 1)
    else:
        run_agent(config)
