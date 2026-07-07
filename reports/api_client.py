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

# domaine_id par réseau (tel que retourné par /api/plateformes/)
NETWORK_DOMAINE_IDS: dict[str, int] = {
    "mobile":       42,
    "transmission": 43,
    "fixe":         44,
    "core":         45,
    "energie":      8,
}

# Filtre additionnel sur le libellé de la plateforme (après filtre domaine_id).
# Liste exacte des plateformes couvertes par l'export manuel netXcare mobile.
NETWORK_PLATFORM_NAMES: dict[str, list[str]] = {
    "mobile": [
        "Outils de supervision", "Cellule", "LNBTS", "BTS", "MRBTS",
        "NRCELL", "WBTS", "NRBTS", "WCEL", "SBTS", "LNCEL", "BCF",
    ],
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

    def get_plateformes_for_network(self, network: str | None = None) -> list[dict[str, Any]]:
        """Retourne la liste des plateformes du réseau donné.
        Filtre par domaine_id côté serveur si possible, sinon côté client.
        Applique ensuite un filtre sur le libellé si NETWORK_PLATFORM_NAMES le définit."""
        domaine_id    = NETWORK_DOMAINE_IDS.get(network) if (network and network != "all") else None
        name_filter   = NETWORK_PLATFORM_NAMES.get(network) if network else None
        name_filter_u = [n.upper() for n in name_filter] if name_filter else None

        # Essaie le filtre côté serveur d'abord (plus fiable)
        if domaine_id is not None:
            try:
                plateformes = self.get_all(f"/api/plateformes/?domaine_id={domaine_id}")
                if plateformes:
                    if name_filter_u:
                        plateformes = [
                            p for p in plateformes
                            if str(p.get("libelle", "")).upper() in name_filter_u
                        ]
                    return plateformes
            except Exception:
                pass  # Fallback : filtre client ci-dessous

        # Filtre client : récupère tout et filtre par domaine_id puis par libellé
        plateformes = self.get_all("/api/plateformes/")
        if domaine_id is not None:
            plateformes = [p for p in plateformes if p.get("domaine_id") == domaine_id]
        if name_filter_u:
            plateformes = [
                p for p in plateformes
                if str(p.get("libelle", "")).upper() in name_filter_u
            ]
        return plateformes

    def get_plateformes_ids_for_network(self, network: str | None = None) -> str:
        """Retourne les IDs des plateformes du réseau donné (comma-separated)."""
        plateformes = self.get_plateformes_for_network(network)
        return ",".join(str(p["id"]) for p in plateformes if p.get("id"))

    def export_data(
        self,
        date_debut: str,
        date_fin: str,
        plateformes_id: str | None = None,
        network: str | None = None,
        extra_params: dict[str, Any] | None = None,
    ) -> list[dict[str, Any]]:
        import logging as _logging
        log = _logging.getLogger(__name__)

        # Filtrer les plateformes par réseau (domaine_id) — toutes les plateformes du domaine incluses.
        if not plateformes_id:
            plateformes_id = self.get_plateformes_ids_for_network(network)

        params: dict[str, Any] = {
            "date_debut": date_debut,
            "date_fin": date_fin,
            "plateformes_id": plateformes_id,
            "selected_columns": ",".join(ALL_EXPORT_COLUMNS),
        }
        if extra_params:
            params.update(extra_params)

        resp = self.session.get(
            self._url("/api/exports/data"),
            params=params,
            headers=self._headers(),
            timeout=self.timeout,
        )
        resp.raise_for_status()
        data = resp.json()

        # Format liste directe — retour immédiat (pas de pagination implicite)
        if isinstance(data, list):
            log.info("export_data: %d lignes (liste) — %s → %s (réseau: %s)",
                     len(data), date_debut, date_fin, network or "all")
            return data

        # Format dict paginé {"data": [...], "totalPage": N, "currentPage": N}
        if isinstance(data, dict) and isinstance(data.get("data"), list):
            all_items: list[dict[str, Any]] = list(data["data"])
            total_page = int(data.get("totalPage") or 1)
            current_page = int(data.get("currentPage") or 1)
            log.info("export_data: page %d/%d (%d items) — réseau: %s",
                     current_page, total_page, len(all_items), network or "all")
            while current_page < total_page:
                current_page += 1
                page_r = self.session.get(
                    self._url("/api/exports/data"),
                    params={**params, "page": current_page},
                    headers=self._headers(),
                    timeout=self.timeout,
                )
                page_r.raise_for_status()
                page_data = page_r.json()
                new_items = page_data.get("data", []) if isinstance(page_data, dict) else page_data
                all_items.extend(new_items)
                log.info("export_data: page %d/%d (+%d items)", current_page, total_page, len(new_items))
            log.info("export_data: total %d lignes", len(all_items))
            return all_items

        raise RuntimeError(f"Réponse export inattendue: {type(data)}")
