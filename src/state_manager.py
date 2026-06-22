"""일별/월별 실행 상태를 JSON 파일로 관리."""

from __future__ import annotations

import json
import logging
from datetime import date, datetime, timezone, timedelta
from pathlib import Path

KST = timezone(timedelta(hours=9))
log = logging.getLogger(__name__)

STATUS_PENDING = "pending"
STATUS_SUCCESS = "success"
STATUS_FAILED = "failed"

# 일별: 최대 재시도 시간 (08:30 기준 3시간 → 11:30 이후 포기)
DAILY_CUTOFF_HOURS = 3
# 월별: 1일 기준 최대 N일 후까지 재시도
MONTHLY_MAX_DAYS = 4


def _now_kst() -> str:
    return datetime.now(KST).isoformat()


class DailyState:
    """하루 단위 실행 상태 (daily_YYYY-MM-DD.json)."""

    def __init__(self, state_dir: Path, target_date: date):
        state_dir.mkdir(parents=True, exist_ok=True)
        self.path = state_dir / f"daily_{target_date.isoformat()}.json"
        self.target_date = target_date
        self._d = self._load()

    def _load(self) -> dict:
        if self.path.exists():
            return json.loads(self.path.read_text(encoding="utf-8"))
        return {
            "status": STATUS_PENDING,
            "attempts": 0,
            "sources_done": [],
            "sheets_written": [],
            "first_attempt_at": None,
            "last_attempt_at": None,
            "completed_at": None,
            "give_up_notified": False,
            "errors": [],
        }

    def _save(self) -> None:
        self.path.write_text(
            json.dumps(self._d, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )

    @property
    def is_done(self) -> bool:
        return self._d["status"] == STATUS_SUCCESS

    @property
    def attempts(self) -> int:
        return self._d["attempts"]

    @property
    def give_up_notified(self) -> bool:
        return self._d.get("give_up_notified", False)

    def mark_give_up_notified(self) -> None:
        self._d["give_up_notified"] = True
        self._save()

    def should_give_up(self) -> bool:
        """첫 시도 후 DAILY_CUTOFF_HOURS 경과 시 포기."""
        first = self._d.get("first_attempt_at")
        if not first:
            return False
        elapsed = datetime.now(KST) - datetime.fromisoformat(first)
        return elapsed.total_seconds() > DAILY_CUTOFF_HOURS * 3600

    def record_attempt(self, error: str | None = None) -> None:
        now = _now_kst()
        if self._d["first_attempt_at"] is None:
            self._d["first_attempt_at"] = now
        self._d["attempts"] += 1
        self._d["last_attempt_at"] = now
        if error:
            self._d["errors"].append({"at": now, "msg": error})
        self._save()

    def mark_source_done(self, source: str) -> None:
        if source not in self._d["sources_done"]:
            self._d["sources_done"].append(source)
        self._save()

    def mark_sheet_written(self, sheet: str) -> None:
        if sheet not in self._d["sheets_written"]:
            self._d["sheets_written"].append(sheet)
        self._save()

    def mark_success(self) -> None:
        self._d["status"] = STATUS_SUCCESS
        self._d["completed_at"] = _now_kst()
        self._save()
        log.info(f"  [STATE] {self.target_date} 성공 완료 (총 {self._d['attempts']}회 시도)")

    def mark_failed(self) -> None:
        self._d["status"] = STATUS_FAILED
        self._save()
        log.warning(f"  [STATE] {self.target_date} 최종 실패 ({self._d['attempts']}회 시도)")


class MonthlyState:
    """월 단위 실행 상태 (monthly_YYYY-MM.json)."""

    def __init__(self, state_dir: Path, year: int, month: int):
        state_dir.mkdir(parents=True, exist_ok=True)
        self.run_id = f"{year:04d}-{month:02d}"
        self.path = state_dir / f"monthly_{self.run_id}.json"
        self._d = self._load()

    def _load(self) -> dict:
        if self.path.exists():
            return json.loads(self.path.read_text(encoding="utf-8"))
        return {
            "status": STATUS_PENDING,
            "attempts": 0,
            "sources_done": [],
            "sheets_written": [],
            "first_attempt_at": None,
            "last_attempt_at": None,
            "completed_at": None,
            "give_up_notified": False,
            "errors": [],
        }

    def _save(self) -> None:
        self.path.write_text(
            json.dumps(self._d, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )

    @property
    def is_done(self) -> bool:
        return self._d["status"] == STATUS_SUCCESS

    @property
    def attempts(self) -> int:
        return self._d["attempts"]

    @property
    def give_up_notified(self) -> bool:
        return self._d.get("give_up_notified", False)

    def mark_give_up_notified(self) -> None:
        self._d["give_up_notified"] = True
        self._save()

    def should_give_up(self) -> bool:
        """1일 첫 시도 후 MONTHLY_MAX_DAYS일 경과 시 포기."""
        first = self._d.get("first_attempt_at")
        if not first:
            return False
        elapsed = datetime.now(KST) - datetime.fromisoformat(first)
        return elapsed.total_seconds() > MONTHLY_MAX_DAYS * 86400

    def record_attempt(self, error: str | None = None) -> None:
        now = _now_kst()
        if self._d["first_attempt_at"] is None:
            self._d["first_attempt_at"] = now
        self._d["attempts"] += 1
        self._d["last_attempt_at"] = now
        if error:
            self._d["errors"].append({"at": now, "msg": error})
        self._save()

    def mark_source_done(self, source: str) -> None:
        if source not in self._d["sources_done"]:
            self._d["sources_done"].append(source)
        self._save()

    def mark_sheet_written(self, sheet: str) -> None:
        if sheet not in self._d["sheets_written"]:
            self._d["sheets_written"].append(sheet)
        self._save()

    def mark_success(self) -> None:
        self._d["status"] = STATUS_SUCCESS
        self._d["completed_at"] = _now_kst()
        self._save()
        log.info(f"  [STATE] {self.run_id} 월별 성공 완료 (총 {self._d['attempts']}회 시도)")

    def mark_failed(self) -> None:
        self._d["status"] = STATUS_FAILED
        self._save()
        log.warning(f"  [STATE] {self.run_id} 월별 최종 실패 ({self._d['attempts']}회 시도)")
