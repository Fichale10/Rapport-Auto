"""Client HTTP pour l'API ticketing (authentification JWT + export)."""
from __future__ import annotations

import requests
from typing import Any

ALL_EXPORT_COLUMNS = [
    "noc_engineer", "ticket_number", "incident_nature", "alarm_time",
    "site_parent", "site_name", "site_id", "region", "impact_equipement",
    "impact_service", "plateforme", "technologies", "alarm_text", "cause",
    "escalade", "informed_technician", "duration_escalade", "action",
    "maintenance_technician", "root_cause", "observation", "point_bloquant",
    "cancel_time", "duration", "status",
]

FILTER_RESOURCES = {
    "plateformes": "/api/plateformes/",
}


class TicketingApiClient:
    def __init__(self, base_url: str, timeout: int = 120) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.token: str | None = None

    def _url(self, path: str) -> str:
        return f"{self.base_url}{path}"

    def _headers(self) -> dict[str, str]:
        headers = {"Accept": "application/json"}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        return headers

    def login(self, username: str, password: str) -> dict[str, Any]:
        response = self.session.post(
            self._url("/api/auth/login"),
            json={"username": username, "password": password},
            headers=self._headers(),
            timeout=self.timeout,
        )
        response.raise_for_status()
        payload = response.json()
        self.token = payload.get("token")
        if not self.token:
            raise RuntimeError("Token JWT absent dans la réponse de login.")
        return payload

    def get_all(self, resource_path: str) -> list[dict[str, Any]]:
        response = self.session.get(
            self._url(resource_path),
            headers=self._headers(),
            timeout=self.timeout,
        )
        response.raise_for_status()
        data = response.json()
        if isinstance(data, list):
            return data
        if isinstance(data, dict) and isinstance(data.get("data"), list):
            items = list(data["data"])
            total_page = int(data.get("totalPage") or 1)
            current_page = int(data.get("currentPage") or 1)
            while current_page < total_page:
                current_page += 1
                page_r = self.session.get(
                    self._url(resource_path),
                    params={"page": current_page},
                    headers=self._headers(),
                    timeout=self.timeout,
                )
                page_r.raise_for_status()
                page_data = page_r.json()
                items.extend(page_data.get("data", []))
            return items
        raise RuntimeError(f"Réponse inattendue pour {resource_path}: {type(data)}")

    def export_data(
        self,
        date_debut: str,
        date_fin: str,
        plateformes_id: str | None = None,
        extra_params: dict[str, Any] | None = None,
    ) -> list[dict[str, Any]]:
        if not plateformes_id:
            plateformes = self.get_all("/api/plateformes/")
            plateformes_id = ",".join(str(p["id"]) for p in plateformes if p.get("id"))

        params: dict[str, Any] = {
            "date_debut": date_debut,
            "date_fin": date_fin,
            "plateformes_id": plateformes_id,
            "selected_columns": ",".join(ALL_EXPORT_COLUMNS),
        }
        if extra_params:
            params.update(extra_params)

        response = self.session.get(
            self._url("/api/exports/data"),
            params=params,
            headers=self._headers(),
            timeout=self.timeout,
        )
        response.raise_for_status()
        data = response.json()
        if isinstance(data, list):
            return data
        raise RuntimeError(f"Réponse export inattendue: {type(data)}")
