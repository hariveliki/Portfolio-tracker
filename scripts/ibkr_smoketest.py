#!/usr/bin/env python3
"""Minimal IBKR connectivity smoke test scaffold."""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path


class JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        payload = {
            "ts": datetime.now(timezone.utc).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }
        return json.dumps(payload)


def configure_logging(log_name: str) -> logging.Logger:
    log_dir = Path(__file__).resolve().parents[1] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger(log_name)
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    handler = logging.FileHandler(log_dir / f"{log_name}.log", encoding="utf-8")
    handler.setFormatter(JsonFormatter())
    logger.addHandler(handler)

    return logger


def main() -> int:
    logger = configure_logging("ibkr_smoketest")
    logger.info("Starting IBKR smoke test")
    logger.info("Smoke test scaffold complete; implement broker ping here")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
