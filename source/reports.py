# source/reports.py
from __future__ import annotations

import requests
from datetime import datetime, timedelta, timezone

from source.auth import get_token   # já devolve (token, expires_at)

# -------------- CONSTANTES DA API -----------------------------------------
_BASE_URL   = "https://backend-dashboards.prd.franqueado.grupoboticario.digital"
_ENDPOINT   = f"{_BASE_URL}/store-indicators"

_COMMON_HDR = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json",
    "x-api-key": "8V5pUI1u1y3dFASezqZYY6iZvkUXDHZO6Ol66ja5",
    "cp-code": "10140",
    "revenue-type": "gmv-revenue",
    "dash-view-name": "store-view/general",
    "Origin": "https://extranet.grupoboticario.com.br",
    "Referer": "https://extranet.grupoboticario.com.br/",
}

_LEEWAY = timedelta(seconds=30)   # renova antes de faltar 30 s
# --------------------------------------------------------------------------


class ReportClient:
    """Cliente de relatórios com renovação automática de token."""

    def __init__(self) -> None:
        self.session = requests.Session()
        self.token, self.expires_at = get_token()
        self._update_auth_header()

    # ---------- utilidades privadas ---------------------------------------

    def _update_auth_header(self) -> None:
        """Atualiza o header Authorization da sessão."""
        self.session.headers.update(
            _COMMON_HDR | {"authorization": f"Bearer {self.token}"}
        )

    def _token_near_expiry(self) -> bool:
        return self.expires_at - datetime.now(timezone.utc) < _LEEWAY

    def _refresh_token(self) -> None:
        """Força renovação e atualiza headers."""
        self.token, self.expires_at = get_token()
        self._update_auth_header()

    # ---------- método público --------------------------------------------

    def get_report(self, start_date: str, end_date: str) -> bytes:
        """Baixa e devolve o XLSX (bytes). Renova token quando preciso."""

        # 1) renova preventivamente
        if self._token_near_expiry():
            self._refresh_token()

        params = {
            "orderBy": "revenueVariationPercentage",
            "aggregation": "pdv",
            "order": "DESC",
            "years": start_date[:4],
            "pillars": "Todos",
            "startCurrentDate": start_date,
            "endCurrentDate": end_date,
            "startPreviousDate": start_date,
            "endPreviousDate": end_date,
            "calendarType": "calendar",
            "previousPeriodCycleType": "retail-year",
            "previousPeriodCalendarType": "retail-year",
            "hour": "00:00 - 23:00",
            "separationType": "businessDays",
            "download": "true",
            "channels": "LOJ",
        }

        # 2) faz a requisição
        resp = self.session.get(_ENDPOINT, params=params, timeout=30)

        # 3) se expirou no meio, renova e tenta 1×
        if resp.status_code == 401:
            self._refresh_token()
            resp = self.session.get(_ENDPOINT, params=params, timeout=30)

        resp.raise_for_status()

        download_url = resp.json()["data"]
        file_resp = requests.get(download_url, timeout=60)
        file_resp.raise_for_status()
        return file_resp.content   # bytes do Excel
