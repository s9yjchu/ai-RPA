"""파일 + 콘솔 로깅 설정."""

from __future__ import annotations

import logging
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

KST = timezone(timedelta(hours=9))


def setup_logging(logs_dir: Path) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_file = logs_dir / f"{datetime.now(KST).strftime('%Y%m%d_%H%M%S')}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return log_file
