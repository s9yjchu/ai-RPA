#!/usr/bin/env bash
# ============================================================
#  B2C RPA — GCP 인프라 최초 설정 스크립트
#  실행 전: gcloud auth login && gcloud config set project $PROJECT
# ============================================================
set -euo pipefail

PROJECT="${GCP_PROJECT_ID:?GCP_PROJECT_ID 환경변수를 설정하세요}"
REGION="asia-northeast3"  # 서울 리전

echo "=== 프로젝트: $PROJECT ==="

# ── API 활성화 ────────────────────────────────────────────────────────
gcloud services enable pubsub.googleapis.com cloudscheduler.googleapis.com \
  --project="$PROJECT"

# ── Pub/Sub 토픽 ──────────────────────────────────────────────────────
gcloud pubsub topics create ai-rpa-daily   --project="$PROJECT" || true
gcloud pubsub topics create ai-rpa-monthly --project="$PROJECT" || true

# ── Pub/Sub 구독 (Windows 에이전트가 pull) ────────────────────────────
gcloud pubsub subscriptions create ai-rpa-daily-sub \
  --topic=ai-rpa-daily \
  --ack-deadline=600 \
  --message-retention-duration=10m \
  --project="$PROJECT" || true

gcloud pubsub subscriptions create ai-rpa-monthly-sub \
  --topic=ai-rpa-monthly \
  --ack-deadline=600 \
  --message-retention-duration=10m \
  --project="$PROJECT" || true

# ── 서비스 계정 (Pub/Sub 구독 전용) ──────────────────────────────────
SA_NAME="b2c-rpa-pubsub-agent"
SA_EMAIL="${SA_NAME}@${PROJECT}.iam.gserviceaccount.com"

gcloud iam service-accounts create "$SA_NAME" \
  --display-name="B2C RPA PubSub Agent" \
  --project="$PROJECT" || true

gcloud projects add-iam-policy-binding "$PROJECT" \
  --member="serviceAccount:${SA_EMAIL}" \
  --role="roles/pubsub.subscriber"

echo "=== service_account.json 키 파일 생성 ==="
gcloud iam service-accounts keys create service_account.json \
  --iam-account="$SA_EMAIL"
echo "  → service_account.json 을 프로젝트 루트에 복사하세요"

# ── Cloud Scheduler — 일별 (평일 08:30~11:30 KST = 23:30~02:30 UTC) ──
# KST 08:30 = UTC 23:30 (전날)  ·  9시간 차이
DAILY_TIMES=("30 23" "0 0" "30 0" "0 1" "30 1" "0 2" "30 2")
for T in "${DAILY_TIMES[@]}"; do
  MIN=$(echo "$T" | cut -d' ' -f1)
  HOUR=$(echo "$T" | cut -d' ' -f2)
  JOB="daily-rpa-${HOUR}h${MIN}m"
  gcloud scheduler jobs create pubsub "$JOB" \
    --schedule="${MIN} ${HOUR} * * 1-5" \
    --time-zone="UTC" \
    --location="$REGION" \
    --topic="ai-rpa-daily" \
    --message-body="daily" \
    --project="$PROJECT" 2>/dev/null || \
  gcloud scheduler jobs update pubsub "$JOB" \
    --schedule="${MIN} ${HOUR} * * 1-5" \
    --time-zone="UTC" \
    --location="$REGION" \
    --topic="ai-rpa-daily" \
    --message-body="daily" \
    --project="$PROJECT"
  echo "  daily job created: $JOB  (cron: ${MIN} ${HOUR} * * 1-5 UTC)"
done

# ── Cloud Scheduler — 월별 (매월 1~4일 08:30~11:30 KST) ──────────────
for DAY in 1 2 3 4; do
  for T in "${DAILY_TIMES[@]}"; do
    MIN=$(echo "$T" | cut -d' ' -f1)
    HOUR=$(echo "$T" | cut -d' ' -f2)
    # 월별 1일 23:30 UTC = 2일 08:30 KST — day-of-month offset 처리
    # 1일 08:30 KST: cron day = 1-1 = 31 (전달 마지막 날) → 간단하게 UTC date 사용
    # 실제 UTC 날짜 계산:
    #   KST day D, time T → UTC day = D (T >= 09:00) or D-1 (T < 09:00)
    #   모든 시간대 08:30~11:30 KST → UTC 23:30(D-1)~02:30(D)
    if [ "$HOUR" -eq 23 ]; then
      UTCDAY=$((DAY - 1))
      [ "$UTCDAY" -eq 0 ] && UTCDAY=28  # 말일 근사 (월 시작 전날)
    else
      UTCDAY=$DAY
    fi
    JOB="monthly-rpa-d${DAY}-${HOUR}h${MIN}m"
    gcloud scheduler jobs create pubsub "$JOB" \
      --schedule="${MIN} ${HOUR} ${UTCDAY} * *" \
      --time-zone="UTC" \
      --location="$REGION" \
      --topic="ai-rpa-monthly" \
      --message-body="monthly" \
      --project="$PROJECT" 2>/dev/null || \
    gcloud scheduler jobs update pubsub "$JOB" \
      --schedule="${MIN} ${HOUR} ${UTCDAY} * *" \
      --time-zone="UTC" \
      --location="$REGION" \
      --topic="ai-rpa-monthly" \
      --message-body="monthly" \
      --project="$PROJECT"
    echo "  monthly job created: $JOB  (cron: ${MIN} ${HOUR} ${UTCDAY} * * UTC = day ${DAY} KST)"
  done
done

echo ""
echo "=== 완료 ==="
echo "  다음 단계:"
echo "  1. service_account.json 을 ai-RPA 프로젝트 루트에 복사"
echo "  2. .env 에 GCP_PROJECT_ID=$PROJECT 추가"
echo "  3. Windows 에서: run_agent.bat --test daily (동작 확인)"
echo "  4. Windows 에서: install_agent_service.bat (서비스 등록)"
