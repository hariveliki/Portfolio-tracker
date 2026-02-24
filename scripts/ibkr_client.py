#!/usr/bin/env python3
"""Thin wrapper around IBKR Client Portal REST API (httpx-based)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import httpx


class GatewayNotAuthenticated(Exception):
    pass


class GatewayRequestFailed(Exception):
    pass


@dataclass
class IbkrClient:
    base_url: str = "https://localhost:5001/v1/api"
    account_id: str = ""
    timeout: float = 30.0
    _http: httpx.Client = field(init=False, repr=False)

    def __post_init__(self) -> None:
        self._http = httpx.Client(
            base_url=self.base_url,
            verify=False,
            timeout=self.timeout,
        )

    def close(self) -> None:
        self._http.close()

    def __enter__(self) -> IbkrClient:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def _get(self, path: str, params: dict[str, Any] | None = None) -> Any:
        resp = self._http.get(path, params=params)
        if resp.status_code >= 400:
            raise GatewayRequestFailed(
                f"GET {path} returned {resp.status_code}: {resp.text}"
            )
        return resp.json()

    def _post(self, path: str, json_body: dict[str, Any] | None = None) -> Any:
        resp = self._http.post(path, json=json_body or {})
        if resp.status_code >= 400:
            raise GatewayRequestFailed(
                f"POST {path} returned {resp.status_code}: {resp.text}"
            )
        return resp.json()

    def check_auth(self) -> dict[str, Any]:
        data = self._post("/iserver/auth/status")
        if not data.get("authenticated"):
            raise GatewayNotAuthenticated(
                "Client Portal Gateway session is not authenticated. "
                "Log in at your gateway URL in a browser first."
            )
        return data

    def tickle(self) -> dict[str, Any]:
        return self._post("/tickle")

    def get_accounts(self) -> list[dict[str, Any]]:
        return self._get("/portfolio/accounts")

    def resolve_account_id(self) -> str:
        if self.account_id:
            return self.account_id
        accounts = self.get_accounts()
        if not accounts:
            raise GatewayRequestFailed("No accounts returned by /portfolio/accounts")
        self.account_id = accounts[0]["id"]
        return self.account_id

    def get_trades(self, days: int = 7) -> list[dict[str, Any]]:
        return self._get("/iserver/account/trades", params={"days": str(days)})

    def search_contract(self, symbol: str, sec_type: str = "STK") -> list[dict[str, Any]]:
        return self._get(
            "/iserver/secdef/search",
            params={"symbol": symbol, "secType": sec_type},
        )

    def resolve_conid(self, symbol: str, sec_type: str = "STK") -> int:
        results = self.search_contract(symbol, sec_type)
        if not results:
            raise GatewayRequestFailed(f"No contract found for symbol {symbol!r}")
        conid = results[0].get("conid")
        if conid is None:
            raise GatewayRequestFailed(
                f"Contract search for {symbol!r} returned no conid"
            )
        return int(conid)

    def get_market_history(
        self,
        conid: int,
        period: str = "1y",
        bar: str = "1d",
        outside_rth: bool = False,
    ) -> dict[str, Any]:
        return self._get(
            "/iserver/marketdata/history",
            params={
                "conid": str(conid),
                "period": period,
                "bar": bar,
                "outsideRth": str(outside_rth).lower(),
            },
        )

    def get_positions(self, account_id: str, page: int = 0) -> list[dict[str, Any]]:
        return self._get(f"/portfolio/{account_id}/positions/{page}")

    def get_account_ledger(self, account_id: str) -> dict[str, Any]:
        return self._get(f"/portfolio/{account_id}/ledger")
