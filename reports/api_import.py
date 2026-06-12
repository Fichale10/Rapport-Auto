"""Logique d'import API : JSON → DataFrame → process_file() → UploadedReport."""
from __future__ import annotations

import logging
import os
from datetime import date, timedelta
from typing import Any

import pandas as pd
from django.conf import settings
from django.contrib.auth import get_user_model
from django.core.files import File
from django.utils import timezone

from .api_client import TicketingApiClient
from .models import UploadedReport

logger = logging.getLogger(__name__)

User = get_user_model()

# Mapping colonnes API (noms français retournés) → noms attendus par treatement.py
# Les colonnes critiques (Alarm text, Alarm Time, Cancel Time, Site Name,
# Site Parent, Escalade, Status, Cause) arrivent déjà avec le bon nom.
COLUMN_MAPPING = {
    "Ingénieur NOC":           "NOC Engineer",
    "Numero du ticket":        "Ticket Number",
    "Nature de l'incident":    "Incident Nature",
    "Site ID":                 "Site ID",
    "Région":                  "Région",
    "Impact - Equipement":     "Impact Equipement",
    "Impact - Service":        "Impact Service",
    "Technicien Informé":      "Informed Technician",
    "Durée escalade":          "Duration Escalade",
    "Technicien de maintenance": "Maintenance Technician",
    "Point bloquant":          "Point Bloquant",
}


def json_to_dataframe(rows: list[dict[str, Any]]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    df.rename(columns={k: v for k, v in COLUMN_MAPPING.items() if k in df.columns}, inplace=True)
    return df


def _get_or_create_api_user() -> Any:
    user, _ = User.objects.get_or_create(
        username="api_import",
        defaults={"is_active": False, "first_name": "Import", "last_name": "API"},
    )
    return user


def fetch_and_save_api(
    date_debut: str,
    date_fin: str,
    user: Any = None,
) -> "UploadedReport":
    """
    Récupère les données API pour la période donnée, les sauvegarde en Excel
    et crée un UploadedReport **non traité** (processed=False).

    Retourne le rapport créé (à rediriger vers process_report).
    """
    import datetime as _dt

    api_url  = settings.TICKETING_API_URL
    api_user = settings.TICKETING_API_USERNAME
    api_pass = settings.TICKETING_API_PASSWORD

    if not api_url or not api_user or not api_pass:
        raise RuntimeError(
            "Identifiants API manquants. Renseignez TICKETING_API_URL, "
            "TICKETING_API_USERNAME et TICKETING_API_PASSWORD dans .env"
        )

    client = TicketingApiClient(api_url)
    client.login(api_user, api_pass)
    logger.info("API login OK — fetch %s → %s", date_debut, date_fin)

    rows = client.export_data(date_debut, date_fin)
    if not rows:
        raise ValueError(f"Aucune donnée retournée par l'API pour {date_debut} → {date_fin}")

    df = json_to_dataframe(rows)

    # Sauvegarde en Excel dans media/uploads/
    label      = f"API_{date_debut[:10]}_{date_fin[:10]}"
    filename   = f"{label}.xlsx"
    uploads_dir = os.path.join(settings.MEDIA_ROOT, "uploads")
    os.makedirs(uploads_dir, exist_ok=True)
    file_path  = os.path.join(uploads_dir, filename)
    df.to_excel(file_path, index=False)
    logger.info("Fichier API sauvegardé : %s (%d lignes)", file_path, len(df))

    # Dates
    d_start = _dt.date.fromisoformat(date_debut[:10])
    d_end   = _dt.date.fromisoformat(date_fin[:10])

    # Crée le rapport non traité
    report = UploadedReport()
    report.user              = user or _get_or_create_api_user()
    report.original_filename = filename
    report.date_rapport      = d_start
    report.date_fin          = d_end if d_end != d_start else None
    report.processed         = False
    report.source            = "api"

    with open(file_path, "rb") as fh:
        report.file.save(filename, File(fh), save=False)

    report.save()
    logger.info("UploadedReport créé (non traité) : id=%s", report.pk)
    return report


def run_import(
    date_debut: str,
    date_fin: str,
    triggered_by: Any = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """
    Importe les données de l'API pour la période date_debut → date_fin.

    Retourne un dict { 'created': int, 'skipped': int, 'errors': list[str] }.
    """
    from treatement import process_file  # import local pour éviter les circulaires

    api_url = settings.TICKETING_API_URL
    api_user = settings.TICKETING_API_USERNAME
    api_pass = settings.TICKETING_API_PASSWORD

    if not api_url or not api_user or not api_pass:
        raise RuntimeError(
            "Identifiants API manquants. Renseignez TICKETING_API_URL, "
            "TICKETING_API_USERNAME et TICKETING_API_PASSWORD dans .env"
        )

    result = {"created": 0, "skipped": 0, "errors": []}
    user = triggered_by or _get_or_create_api_user()

    client = TicketingApiClient(api_url)
    client.login(api_user, api_pass)
    logger.info("API login OK — import %s → %s", date_debut, date_fin)

    rows = client.export_data(date_debut, date_fin)
    if not rows:
        logger.info("Aucune donnée retournée par l'API pour %s → %s", date_debut, date_fin)
        return result

    df = json_to_dataframe(rows)

    # Détermine le label de fichier
    label = f"API_{date_debut[:10]}_{date_fin[:10]}"

    # Vérifie si un rapport pour cette période existe déjà
    existing = UploadedReport.objects.filter(
        original_filename__startswith="API_",
        date_rapport=date_debut[:10],
        date_fin=date_fin[:10] if date_fin[:10] != date_debut[:10] else None,
    ).first()

    if existing and not overwrite:
        logger.info("Rapport déjà existant pour %s (id=%s), ignoré.", label, existing.pk)
        result["skipped"] += 1
        return result

    # Garde une copie avec Duration_Sec AVANT que process_file la supprime
    # (process_file retourne df_detail/df_synth sans Duration_Sec)
    try:
        df_detail, df_synth, df_synthesis = process_file(df, date_debut[:10], date_fin[:10])
    except Exception as exc:
        msg = f"Erreur process_file pour {label}: {exc}"
        logger.exception(msg)
        result["errors"].append(msg)
        return result

    # Recompute Duration_Sec depuis df_detail (qui a les colonnes Alarm/Cancel Time)
    import pandas as _pd
    from collections import defaultdict

    def _compute_duration_sec(df_src):
        """Recalcule Duration_Sec depuis Alarm Time et Cancel Time."""
        try:
            at = _pd.to_datetime(df_src["Alarm Time"],  dayfirst=True, format='mixed', errors='coerce')
            ct = _pd.to_datetime(df_src["Cancel Time"], dayfirst=True, format='mixed', errors='coerce')
            fin_jour = _pd.to_datetime(f"{date_fin[:10]} 23:59:00")
            ct = ct.fillna(fin_jour)
            dur = (ct - at).dt.total_seconds().clip(lower=0)
            return dur
        except Exception:
            return _pd.Series(dtype=float)

    dur_detail = _compute_duration_sec(df_detail)
    dur_synth  = _compute_duration_sec(df_synth)

    # Crée ou met à jour l'UploadedReport
    report = existing if (existing and overwrite) else UploadedReport()
    report.user = user
    report.original_filename = f"{label}.xlsx"
    report.date_rapport = date_debut[:10]
    report.date_fin = date_fin[:10] if date_fin[:10] != date_debut[:10] else None
    report.processed = True
    report.source = "api"

    total_incidents = len(df_synth)
    report.total_incidents = total_incidents

    unresolved = 0
    if "Status" in df_synth.columns:
        unresolved = int((df_synth["Status"].astype(str).str.upper() == "OUVERT").sum())
    report.unresolved_count = unresolved

    # total_duration_sec = outage total (avec doublons) depuis df_detail
    report.total_duration_sec = int(dur_detail.sum()) if len(dur_detail) else 0

    report.synthesis_json = df_synthesis.to_dict(orient="records")
    report.uploaded_at = timezone.now()

    # Top sites (depuis df_synth dédupliqué)
    if "Site Name" in df_synth.columns:
        site_counts = df_synth["Site Name"].value_counts().head(10)
        report.top_sites_json = [
            {"name": k, "count": int(v)} for k, v in site_counts.items()
        ]

    # Top causes (avec durée recalculée)
    if "Cause" in df_synth.columns:
        cause_data = defaultdict(float)
        for i, row in df_synth.iterrows():
            cause = str(row.get("Cause", "")).strip()
            if cause and cause not in ("nan", ""):
                cause_data[cause] += float(dur_synth.get(i, 0) or 0)
        report.top_causes_json = [
            {"name": k, "duration_sec": v}
            for k, v in sorted(cause_data.items(), key=lambda x: x[1], reverse=True)[:10]
        ]

    # Region sites
    region_col = next(
        (c for c in ("Région", "Region", "REGION", "region") if c in df_synth.columns), None
    )
    site_col = next(
        (c for c in ("Site Name", "site_name") if c in df_synth.columns), None
    )
    if region_col and site_col:
        region_sites: dict[str, list[str]] = {}
        for region, grp in df_synth.groupby(region_col):
            region_sites[str(region).strip()] = (
                grp[site_col].dropna().astype(str).unique().tolist()
            )
        report.region_sites_json = region_sites
    else:
        report.region_sites_json = {}

    # outage_journalier_json : outage par escalade et par jour (pour la section Disponibilité)
    esc_col = "Escalade" if "Escalade" in df_detail.columns else None
    if esc_col:
        outage_jour: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
        alarm_times  = _pd.to_datetime(df_detail["Alarm Time"],  dayfirst=True, format='mixed', errors='coerce')
        cancel_times = _pd.to_datetime(df_detail["Cancel Time"], dayfirst=True, format='mixed', errors='coerce')
        fin_dt = _pd.Timestamp(f"{date_fin[:10]} 23:59:00")
        cancel_times = cancel_times.fillna(fin_dt)

        for i, esc in df_detail[esc_col].items():
            esc = str(esc).strip()
            if not esc or esc in ("nan", ""):
                continue
            t_start = alarm_times.get(i)
            t_end   = cancel_times.get(i)
            if _pd.isna(t_start) or _pd.isna(t_end) or t_end <= t_start:
                continue
            cur = t_start.normalize()
            while cur <= t_end:
                day_str   = cur.strftime("%Y-%m-%d")
                seg_start = max(t_start, cur)
                seg_end   = min(t_end, cur + _pd.Timedelta(days=1) - _pd.Timedelta(seconds=1))
                sec = max(0.0, (seg_end - seg_start).total_seconds())
                outage_jour[esc][day_str] += sec
                cur += _pd.Timedelta(days=1)

        report.outage_journalier_json = {k: dict(v) for k, v in outage_jour.items()}
    else:
        report.outage_journalier_json = {}

    report.save()

    result["created"] += 1
    logger.info("Rapport créé : %s (id=%s, %d incidents)", label, report.pk, total_incidents)
    return result


def run_import_months(
    date_debut: str,
    date_fin: str,
    triggered_by: Any = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """
    Si la période > 1 mois, découpe mois par mois.
    Retourne le cumul des résultats.
    """
    from datetime import date as _date

    d_start = _date.fromisoformat(date_debut[:10])
    d_end   = _date.fromisoformat(date_fin[:10])
    delta   = (d_end - d_start).days

    cumul = {"created": 0, "skipped": 0, "errors": []}

    if delta <= 31:
        res = run_import(date_debut, date_fin, triggered_by, overwrite)
        _merge_results(cumul, res)
        return cumul

    # Découpe mois par mois
    current = d_start
    while current <= d_end:
        # Dernier jour du mois courant
        if current.month == 12:
            last_of_month = _date(current.year + 1, 1, 1) - timedelta(days=1)
        else:
            last_of_month = _date(current.year, current.month + 1, 1) - timedelta(days=1)

        chunk_end = min(last_of_month, d_end)
        res = run_import(
            current.isoformat(),
            chunk_end.isoformat(),
            triggered_by,
            overwrite,
        )
        _merge_results(cumul, res)
        current = chunk_end + timedelta(days=1)

    return cumul


def _merge_results(cumul: dict, res: dict) -> None:
    cumul["created"]  += res.get("created", 0)
    cumul["skipped"]  += res.get("skipped", 0)
    cumul["errors"].extend(res.get("errors", []))
