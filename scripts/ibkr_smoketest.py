#!/usr/bin/env python3
"""Smoke test for IBKR Client Portal Gateway connectivity."""

from __future__ import annotations

import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

import yaml

from ibkr_client import GatewayNotAuthenticated, GatewayRequestFailed, IbkrClient


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


def load_gateway_url() -> str:
    config_path = Path(__file__).resolve().parents[1] / "config" / "portfolio.yaml"
    if not config_path.exists():
        return "https://localhost:5001/v1/api"
    with config_path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}
    return str(raw.get("cpapi", {}).get("gateway_url", "https://localhost:5001/v1/api"))


def main() -> int:
    logger = configure_logging("ibkr_smoketest")
    logger.info("Starting IBKR Client Portal Gateway smoke test")

    gateway_url = load_gateway_url()
    print(f"Gateway URL: {gateway_url}")

    try:
        with IbkrClient(base_url=gateway_url) as client:
            auth = client.check_auth()
            logger.info("Authentication check passed")
            print(f"Authenticated: {auth.get('authenticated')}")
            print(f"Connected: {auth.get('connected')}")
            print(f"Competing: {auth.get('competing')}")

            tickle = client.tickle()
            sso_expires = tickle.get("ssoExpires", 0)
            logger.info(f"Tickle OK, SSO expires in {sso_expires}ms")
            print(f"SSO expires in: {sso_expires}ms")

            accounts = client.get_accounts()
            account_ids = [a.get("id", "unknown") for a in accounts]
            logger.info(f"Accounts: {account_ids}")
            print(f"Accounts: {account_ids}")

    except GatewayNotAuthenticated as exc:
        logger.error(f"Not authenticated: {exc}")
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    except GatewayRequestFailed as exc:
        logger.error(f"Request failed: {exc}")
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        logger.error(f"Connection error: {exc}")
        print(f"FAIL: Cannot reach gateway at {gateway_url}: {exc}", file=sys.stderr)
        return 1

    logger.info("Smoke test passed")
    print("Smoke test passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
