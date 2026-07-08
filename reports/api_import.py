"""Logique d'import API : JSON → DataFrame → traitement indépendant → UploadedReport."""
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
    network: str = "mobile",
) -> "UploadedReport":
    """
    Récupère les données API pour la période et le réseau donnés, sauvegarde en Excel
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
    logger.info("API login OK — fetch %s → %s (réseau: %s)", date_debut, date_fin, network)

    # Élargir la plage d'un jour de chaque côté pour capturer les tickets
    # créés ET fermés pendant la journée (l'API semble filtrer les tickets
    # "ouverts avant date_debut", pas les tickets actifs pendant la journée).
    _d_start = _dt.date.fromisoformat(date_debut[:10])
    _d_end   = _dt.date.fromisoformat(date_fin[:10])
    api_date_debut = (_d_start - _dt.timedelta(days=1)).isoformat()
    api_date_fin   = (_d_end   + _dt.timedelta(days=1)).isoformat()
    logger.info("Plage API élargie : %s → %s (pour capter tickets intraday)", api_date_debut, api_date_fin)

    rows = client.export_data(api_date_debut, api_date_fin, network=network)
    if not rows:
        raise ValueError(
            f"Aucune donnée retournée par l'API pour {date_debut} → {date_fin} "
            f"(réseau: {network})"
        )

    df = json_to_dataframe(rows)
    logger.info("API retourne %d lignes brutes (plage élargie %s → %s)", len(df), api_date_debut, api_date_fin)

    # Pré-filtrer avant sauvegarde : ne garder que les tickets qui chevauchent
    # la période cible [date_debut 00:00 … date_fin 23:59].
    # Cela élimine les tickets fermés AVANT date_debut (Jan 8-9 révolus) et ceux
    # ouverts APRÈS date_fin — qui seraient de toute façon supprimés par treatement.py
    # mais alourdissent inutilement le fichier (surtout pour un import mensuel).
    debut_dt = pd.Timestamp(f"{date_debut[:10]} 00:00:00")
    fin_dt   = pd.Timestamp(f"{date_fin[:10]}  23:59:59")
    if "Alarm Time" in df.columns:
        at = pd.to_datetime(df["Alarm Time"], dayfirst=True, format="mixed", errors="coerce")
        ct = pd.to_datetime(df.get("Cancel Time", pd.Series(dtype="object")),
                            dayfirst=True, format="mixed", errors="coerce")
        mask = (at <= fin_dt) & (ct.isna() | (ct >= debut_dt))
        avant = len(df)
        df = df[mask].copy()
        logger.info("Pré-filtre date : %d → %d lignes (-%d tickets hors période)",
                    avant, len(df), avant - len(df))

    # Sauvegarde en Excel dans media/uploads/
    net_label   = network.upper() if network and network != "all" else "ALL"
    label       = f"API_{net_label}_{date_debut[:10]}_{date_fin[:10]}"
    filename    = f"{label}.xlsx"
    uploads_dir = os.path.join(settings.MEDIA_ROOT, "uploads")
    os.makedirs(uploads_dir, exist_ok=True)
    file_path   = os.path.join(uploads_dir, filename)

    # Supprimer l'ancien fichier s'il existe pour éviter les erreurs de permission
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
        except PermissionError:
            # Fichier ouvert ailleurs (ex: Excel) — utiliser un nom unique avec timestamp
            import time as _time
            ts       = int(_time.time())
            filename  = f"{label}_{ts}.xlsx"
            file_path = os.path.join(uploads_dir, filename)
            logger.warning("Fichier verrouillé, utilisation du nom alternatif : %s", filename)

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


def _process_api_dataframe(
    df: "pd.DataFrame",
    date_debut: str,
    date_fin: str,
) -> "tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]":
    """
    Traitement des données API brutes : bornage dates, durée, déduplication, synthèse.

    Équivalent à treatement.process_file() SANS le filtre Alarm text
    (ce filtre est propre aux exports Excel netXcare manuels).

    Retourne (df_detail, df_synth, df_synthesis) — même structure que process_file().
    """
    debut_jour = pd.Timestamp(f"{date_debut} 00:00:00")
    fin_jour   = pd.Timestamp(f"{date_fin}   23:59:00")

    df = df.copy()
    df['_at'] = pd.to_datetime(df.get('Alarm Time',  pd.Series(dtype='object')),
                               dayfirst=True, format='mixed', errors='coerce')
    df['_ct'] = pd.to_datetime(df.get('Cancel Time', pd.Series(dtype='object')),
                               dayfirst=True, format='mixed', errors='coerce')

    # Borner les incidents qui chevauchent la période
    cond_avant = (df['_at'] < debut_jour) & ((df['_ct'] >= debut_jour) | df['_ct'].isna())
    df.loc[cond_avant, '_at'] = debut_jour

    cond_apres = (df['_at'] <= fin_jour) & ((df['_ct'] > fin_jour) | df['_ct'].isna())
    df.loc[cond_apres, '_ct'] = fin_jour

    # Supprimer les incidents hors période
    cond_ok = (
        (df['_at'] >= debut_jour) & (df['_at'] <= fin_jour) &
        (df['_ct'] >= debut_jour) & (df['_ct'] <= fin_jour)
    )
    df = df[cond_ok].copy()

    # Calcul durée
    delta = df['_ct'] - df['_at']
    df['Duration_Sec'] = delta.dt.total_seconds().clip(lower=0)
    df['Duration'] = delta.apply(
        lambda x: (f"{int(x.total_seconds()//3600):02d}:"
                   f"{int((x.total_seconds()%3600)//60):02d}:"
                   f"{int(x.total_seconds()%60):02d}")
        if pd.notnull(x) else ""
    ).astype(str)

    # Réécrire les colonnes affichées avec les valeurs bornées
    df['Alarm Time']  = df['_at']
    df['Cancel Time'] = df['_ct']
    df = df.drop(columns=['_at', '_ct'])

    df_complet = df.sort_values('Alarm Time').copy()

    # Déduplication pour la synthèse (identique à treatement.py)
    df_pour_synthese = df.copy()
    if 'Site Parent' in df_pour_synthese.columns and 'Site Name' in df_pour_synthese.columns:
        df_pour_synthese['_racine'] = (
            df_pour_synthese['Site Parent']
            .replace(['', 'N/A', 'nan', 'NaN'], pd.NA)
            .fillna(df_pour_synthese['Site Name'])
        )
        df_pour_synthese = df_pour_synthese.drop_duplicates(
            subset=['_racine', 'Alarm Time'], keep='first'
        ).drop(columns=['_racine'])
    elif 'Site Name' in df_pour_synthese.columns:
        df_pour_synthese = df_pour_synthese.drop_duplicates(
            subset=['Site Name', 'Alarm Time'], keep='first'
        )

    # Tableau de synthèse par escalade (identique à treatement.py)
    ESCALADES_ORDRE = [
        "ENERGIE", "TRANS FH-FIELD O", "RAN-FIELD O", "ENERGIE / TRANS / RAN",
        "TRANS / RAN", "INFRA", "PROJET", "TRANS FO",
        "TRANS FTTM", "TRANS IP", "ENVIRONNEMENT", "BSS",
    ]

    def _fmt(secs: float) -> str:
        s = int(secs or 0)
        return f"{s//3600}:{(s%3600)//60:02d}:{s%60:02d}"

    lignes = []
    for esc in ESCALADES_ORDRE:
        s_esc = (df_pour_synthese[df_pour_synthese['Escalade'] == esc]
                 if 'Escalade' in df_pour_synthese.columns else pd.DataFrame())
        c_esc = (df_complet[df_complet['Escalade'] == esc]
                 if 'Escalade' in df_complet.columns else pd.DataFrame())
        count = len(s_esc)
        if count > 0:
            duree_s  = float(s_esc['Duration_Sec'].sum())
            outage_s = float(c_esc['Duration_Sec'].sum())
            mttr_s   = duree_s / count
            if 'Status' in s_esc.columns:
                non_res = int((s_esc['Status'].astype(str).str.upper() == 'OUVERT').sum())
            else:
                non_res = 0
            statut = f"{non_res} Non resolu" if non_res > 0 else "Résolu"
        else:
            duree_s = outage_s = mttr_s = 0.0
            statut = "N/A"
        lignes.append({
            "Escalade": esc, "Inc count": count,
            "DUREE":  _fmt(duree_s), "MTTR": _fmt(mttr_s),
            "OUTAGE": _fmt(outage_s), "Status": statut,
        })

    df_synthesis = pd.DataFrame(lignes)
    total_count  = len(df_pour_synthese)
    total_duree  = float(df_pour_synthese['Duration_Sec'].sum()) if len(df_pour_synthese) else 0.0
    total_outage = float(df_complet['Duration_Sec'].sum()) if len(df_complet) else 0.0
    total_mttr   = total_duree / total_count if total_count else 0.0
    df_synthesis = pd.concat([df_synthesis, pd.DataFrame([{
        "Escalade": "TOTAL", "Inc count": total_count,
        "DUREE":  _fmt(total_duree), "MTTR": _fmt(total_mttr),
        "OUTAGE": _fmt(total_outage), "Status": "",
    }])], ignore_index=True)

    df_detail = df_complet.drop(columns=['Duration_Sec'])
    df_synth  = df_pour_synthese.drop(columns=['Duration_Sec'], errors='ignore')
    return df_detail, df_synth, df_synthesis


def run_import(
    date_debut: str,
    date_fin: str,
    triggered_by: Any = None,
    overwrite: bool = False,
    network: str = "mobile",
) -> dict[str, Any]:
    """
    Importe les données de l'API pour la période date_debut → date_fin.

    Retourne un dict { 'created': int, 'skipped': int, 'errors': list[str] }.
    """

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
    logger.info("API login OK — import %s → %s (réseau: %s)", date_debut, date_fin, network)

    # Élargir la plage d'un jour de chaque côté pour capturer les tickets
    # créés ET fermés pendant la journée (l'API semble filtrer les tickets
    # "ouverts avant date_debut", pas les tickets actifs pendant la journée).
    d_start = date.fromisoformat(date_debut[:10])
    d_end   = date.fromisoformat(date_fin[:10])
    api_date_debut = (d_start - timedelta(days=1)).isoformat()
    api_date_fin   = (d_end   + timedelta(days=1)).isoformat()
    logger.info("Plage API élargie : %s → %s (rapport: %s → %s)", api_date_debut, api_date_fin, date_debut, date_fin)

    rows = client.export_data(api_date_debut, api_date_fin, network=network)
    if not rows:
        logger.info("Aucune donnée retournée par l'API pour %s → %s (réseau: %s)", date_debut, date_fin, network)
        return result

    df = json_to_dataframe(rows)
    logger.info("API retourne %d lignes brutes (plage élargie %s → %s)", len(df), api_date_debut, api_date_fin)

    # Pré-filtrer : ne garder que les tickets qui chevauchent [date_debut … date_fin]
    debut_dt = pd.Timestamp(f"{date_debut[:10]} 00:00:00")
    fin_dt   = pd.Timestamp(f"{date_fin[:10]}  23:59:59")
    if "Alarm Time" in df.columns:
        at   = pd.to_datetime(df["Alarm Time"], dayfirst=True, format="mixed", errors="coerce")
        ct   = pd.to_datetime(df.get("Cancel Time", pd.Series(dtype="object")),
                              dayfirst=True, format="mixed", errors="coerce")
        mask = (at <= fin_dt) & (ct.isna() | (ct >= debut_dt))
        avant = len(df)
        df = df[mask].copy()
        logger.info("Pré-filtre date : %d → %d lignes (-%d tickets hors période)",
                    avant, len(df), avant - len(df))

    # Détermine le label de fichier (inclut le réseau)
    net_label = network.upper() if network and network != "all" else "ALL"
    label = f"API_{net_label}_{date_debut[:10]}_{date_fin[:10]}"

    # Vérifie si un rapport pour cette période et ce réseau existe déjà
    existing = UploadedReport.objects.filter(
        original_filename__startswith=f"API_{net_label}_",
        date_rapport=date_debut[:10],
        date_fin=date_fin[:10] if date_fin[:10] != date_debut[:10] else None,
    ).first()

    if existing and not overwrite:
        logger.info("Rapport déjà existant pour %s (id=%s), ignoré.", label, existing.pk)
        result["skipped"] += 1
        return result

    # ── Traitement selon le réseau ───────────────────────────────────────────
    if network == "fixe":
        return _run_import_fixe(df, date_debut, date_fin, label, user, result, existing, overwrite)

    if network == "transmission":
        return _run_import_transmission(df, date_debut, date_fin, label, user, result, existing, overwrite)

    try:
        df_detail, df_synth, df_synthesis = _process_api_dataframe(df, date_debut[:10], date_fin[:10])
    except Exception as exc:
        msg = f"Erreur _process_api_dataframe pour {label}: {exc}"
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

    # Mapping escalade brut → clé normalisée (identique à ESC_MAPPING dans views.py)
    _ESC_MAPPING = {
        'ENERGIE':          'ENERGIE',
        'RAN-FIELD O':      'RAN',
        'RAN':              'RAN',
        'TRANS FH-FIELD O': 'TRANS FH',
        'TRANS FH':         'TRANS FH',
        'TRANS IP':         'TRANS IP',
    }

    def _parse_hms(s):
        """Convertit HH:MM:SS en secondes (retourne 0 si invalide)."""
        try:
            parts = str(s).split(':')
            if len(parts) == 3:
                return int(float(parts[0])) * 3600 + int(float(parts[1])) * 60 + int(float(parts[2]))
        except (ValueError, AttributeError):
            pass
        return 0

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
    site_col_synth = next((c for c in ("Site Name", "site_name") if c in df_synth.columns), None)
    site_col_detail = next((c for c in ("Site Name", "site_name") if c in df_detail.columns), None)
    if site_col_synth:
        site_counts = df_synth[site_col_synth].value_counts().head(10)
        report.top_sites_json = [
            {"name": k, "count": int(v)} for k, v in site_counts.items()
        ]

    # Top causes par durée — depuis df_detail (tous incidents, cohérent avec la voie upload)
    cause_col_detail = next((c for c in ("Cause", "Root Cause") if c in df_detail.columns), None)
    if cause_col_detail and "Duration" in df_detail.columns:
        cause_dur: dict[str, float] = defaultdict(float)
        for _, row in df_detail.iterrows():
            cause = str(row.get(cause_col_detail, "")).strip()
            dur_s = _parse_hms(str(row.get("Duration", "")))
            if cause and cause not in ("nan", "") and dur_s > 0:
                cause_dur[cause] += dur_s
        report.top_causes_json = [
            {"name": k, "duration_sec": v}
            for k, v in sorted(cause_dur.items(), key=lambda x: x[1], reverse=True)[:10]
        ]

    # site_duration_json : durée cumulée par site (depuis df_detail, tous incidents)
    if site_col_detail and "Duration" in df_detail.columns:
        site_dur: dict[str, float] = {}
        for _, row in df_detail.iterrows():
            s = str(row.get(site_col_detail, "")).strip()
            d = _parse_hms(str(row.get("Duration", "")))
            if s and s != "nan" and d > 0:
                site_dur[s] = site_dur.get(s, 0) + d
        report.site_duration_json = site_dur
    else:
        report.site_duration_json = {}

    # site_top_cause_json : cause principale par site (depuis df_synth dédupliqué)
    cause_col_synth = next((c for c in ("Cause", "Root Cause") if c in df_synth.columns), None)
    if site_col_synth and cause_col_synth:
        _sc: dict = {}
        for _, row in df_synth.iterrows():
            s = str(row.get(site_col_synth, "")).strip()
            c = str(row.get(cause_col_synth, "")).strip()
            if s and s != "nan" and c and c != "nan":
                if s not in _sc:
                    _sc[s] = {}
                _sc[s][c] = _sc[s].get(c, 0) + 1
        report.site_top_cause_json = {
            s: max(causes, key=causes.get)
            for s, causes in _sc.items() if causes
        }
    else:
        report.site_top_cause_json = {}

    # Region sites
    region_col = next(
        (c for c in ("Région", "Region", "REGION", "region") if c in df_synth.columns), None
    )
    if region_col and site_col_synth:
        region_sites: dict[str, list[str]] = {}
        for region, grp in df_synth.groupby(region_col):
            region_sites[str(region).strip()] = (
                grp[site_col_synth].dropna().astype(str).unique().tolist()
            )
        report.region_sites_json = region_sites
    else:
        report.region_sites_json = {}

    # outage_journalier_json : outage par escalade normalisée et par jour (pour Disponibilité)
    # Les clés sont normalisées via _ESC_MAPPING pour correspondre à NB_SITES dans views.py
    esc_col = "Escalade" if "Escalade" in df_detail.columns else None
    if esc_col:
        outage_jour: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
        alarm_times  = _pd.to_datetime(df_detail["Alarm Time"],  dayfirst=True, format='mixed', errors='coerce')
        cancel_times = _pd.to_datetime(df_detail["Cancel Time"], dayfirst=True, format='mixed', errors='coerce')
        fin_dt = _pd.Timestamp(f"{date_fin[:10]} 23:59:00")
        cancel_times = cancel_times.fillna(fin_dt)

        for i, esc_raw in df_detail[esc_col].items():
            esc_key = _ESC_MAPPING.get(str(esc_raw).strip())
            if not esc_key:
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
                outage_jour[esc_key][day_str] += sec
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
    network: str = "mobile",
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
        res = run_import(date_debut, date_fin, triggered_by, overwrite, network=network)
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
            network=network,
        )
        _merge_results(cumul, res)
        current = chunk_end + timedelta(days=1)

    return cumul


def _merge_results(cumul: dict, res: dict) -> None:
    cumul["created"]  += res.get("created", 0)
    cumul["skipped"]  += res.get("skipped", 0)
    cumul["errors"].extend(res.get("errors", []))


def _run_import_fixe(df, date_debut, date_fin, label, user, result, existing, overwrite):
    """Traitement spécifique réseau fixe (sans filtrage alarmes mobile)."""
    import pandas as _pd
    from collections import defaultdict

    d_debut = date_debut[:10]
    d_fin   = date_fin[:10]

    def _fmt(secs):
        secs = int(secs or 0)
        return f"{secs // 3600}:{(secs % 3600) // 60:02d}:{secs % 60:02d}"

    # Calcul des durées
    debut_ts = _pd.Timestamp(f"{d_debut} 00:00:00")
    fin_ts   = _pd.Timestamp(f"{d_fin} 23:59:59")

    df = df.copy()
    df['_at'] = _pd.to_datetime(df.get('Alarm Time',  _pd.Series(dtype='object')), dayfirst=True, format='mixed', errors='coerce')
    df['_ct'] = _pd.to_datetime(df.get('Cancel Time', _pd.Series(dtype='object')), dayfirst=True, format='mixed', errors='coerce')
    df['_ct'] = df['_ct'].fillna(fin_ts).clip(upper=fin_ts)
    df['_at'] = df['_at'].clip(lower=debut_ts)
    df['_dur'] = (df['_ct'] - df['_at']).dt.total_seconds().clip(lower=0)

    total_incidents = len(df)
    # Détection insensible à la casse pour Status
    status_col = next((c for c in df.columns if c.strip().lower() == 'status'), None)
    unresolved  = int((df[status_col].astype(str).str.upper().isin(['OUVERT', 'OPEN'])).sum()) if status_col else 0

    # ── METIER (Escalade) : count + MTTR ─────────────────────────────────────
    metier_stats = []
    if 'Escalade' in df.columns:
        for esc, grp in df.groupby('Escalade'):
            esc = str(esc).strip()
            if not esc or esc == 'nan':
                continue
            n       = len(grp)
            tot_dur = float(grp['_dur'].sum())
            mttr    = tot_dur / n if n > 0 else 0
            metier_stats.append({
                'name': esc, 'count': n,
                'total_dur_sec': tot_dur, 'mttr_sec': mttr,
                'mttr': _fmt(mttr),
            })
        metier_stats.sort(key=lambda x: x['count'], reverse=True)

    # ── REGION : count + MTTR ────────────────────────────────────────────────
    region_stats = []
    reg_col = next((c for c in ('Région', 'Region', 'REGION') if c in df.columns), None)
    if reg_col:
        for region, grp in df.groupby(reg_col):
            region = str(region).strip()
            if not region or region == 'nan':
                continue
            n       = len(grp)
            tot_dur = float(grp['_dur'].sum())
            mttr    = tot_dur / n if n > 0 else 0
            region_stats.append({
                'name': region, 'count': n,
                'total_dur_sec': tot_dur, 'mttr_sec': mttr,
                'mttr': _fmt(mttr),
            })
        region_stats.sort(key=lambda x: x['count'], reverse=True)

    # ── CAUSES (incidents ouverts, fallback sur tous) ─────────────────────────
    open_causes = []
    cause_col = next((c for c in df.columns if c.strip().lower() == 'cause'), None)
    if cause_col:
        if status_col:
            df_open = df[df[status_col].astype(str).str.upper().isin(['OUVERT', 'OPEN'])]
        else:
            df_open = df
        # Fallback : si aucun incident ouvert, prendre tous les incidents
        if df_open.empty:
            df_open = df
        cause_counts = df_open[cause_col].dropna().astype(str).value_counts()
        for cause, cnt in cause_counts.items():
            cause = str(cause).strip()
            if cause and cause.lower() != 'nan':
                open_causes.append({'name': cause, 'count': int(cnt)})

    # ── TYPES D'INCIDENTS — recherche insensible à la casse ──────────────────
    incident_types = []
    nat_col = next(
        (c for c in df.columns
         if c.strip().lower() in ("incident nature", "nature de l'incident",
                                   "nature de l’incident", "incident_nature",
                                   "type incident", "type d'incident")),
        None
    )
    if nat_col:
        type_counts = df[nat_col].dropna().astype(str).value_counts().head(20)
        for itype, cnt in type_counts.items():
            itype = str(itype).strip()
            if itype and itype.lower() != 'nan':
                incident_types.append({'name': itype, 'count': int(cnt)})

    # ── Top sites ─────────────────────────────────────────────────────────────
    top_sites = []
    site_col = next((c for c in ('Site Name', 'site_name') if c in df.columns), None)
    if site_col:
        for nm, cnt in df[site_col].value_counts().head(10).items():
            top_sites.append({'name': str(nm), 'count': int(cnt)})

    # ── Top causes par durée (toutes) ─────────────────────────────────────────
    top_causes = []
    if 'Cause' in df.columns:
        cause_dur = defaultdict(float)
        for _, row in df.iterrows():
            c = str(row.get('Cause', '')).strip()
            if c and c != 'nan':
                cause_dur[c] += float(row.get('_dur', 0) or 0)
        top_causes = [
            {'name': k, 'duration_sec': v}
            for k, v in sorted(cause_dur.items(), key=lambda x: x[1], reverse=True)[:10]
        ]

    # ── Sauvegarder ──────────────────────────────────────────────────────────
    report = existing if (existing and overwrite) else UploadedReport()
    report.user              = user
    report.original_filename = f"{label}.xlsx"
    report.date_rapport      = d_debut
    report.date_fin          = d_fin if d_fin != d_debut else None
    report.processed         = True
    report.source            = "api"
    report.total_incidents   = total_incidents
    report.unresolved_count  = unresolved
    report.total_duration_sec = int(df['_dur'].sum())
    report.synthesis_json    = []   # non utilisé pour fixe
    report.top_sites_json    = top_sites
    report.top_causes_json   = top_causes
    report.region_sites_json = {}
    report.outage_journalier_json = {}
    report.fixe_stats_json   = {
        'metier':         metier_stats,
        'region':         region_stats,
        'open_causes':    open_causes,
        'incident_types': incident_types,
    }
    report.uploaded_at = timezone.now()
    report.save()

    result["created"] += 1
    logger.info("Rapport FIXE créé : %s (id=%s, %d incidents)", label, report.pk, total_incidents)
    return result


def _run_import_transmission(df, date_debut, date_fin, label, user, result, existing, overwrite):
    """Traitement spécifique réseau transmission (sans filtrage alarmes mobile)."""
    import pandas as _pd
    from collections import defaultdict

    d_debut = date_debut[:10]
    d_fin   = date_fin[:10]

    def _fmt(secs):
        secs = int(secs or 0)
        return f"{secs // 3600}:{(secs % 3600) // 60:02d}:{secs % 60:02d}"

    # Calcul des durées
    debut_ts = _pd.Timestamp(f"{d_debut} 00:00:00")
    fin_ts   = _pd.Timestamp(f"{d_fin} 23:59:59")
    period_secs = (fin_ts - debut_ts).total_seconds()

    df = df.copy()
    df['_at'] = _pd.to_datetime(df.get('Alarm Time',  _pd.Series(dtype='object')), dayfirst=True, format='mixed', errors='coerce')
    df['_ct'] = _pd.to_datetime(df.get('Cancel Time', _pd.Series(dtype='object')), dayfirst=True, format='mixed', errors='coerce')
    df['_ct'] = df['_ct'].fillna(fin_ts).clip(upper=fin_ts)
    df['_at'] = df['_at'].clip(lower=debut_ts)
    df['_dur'] = (df['_ct'] - df['_at']).dt.total_seconds().clip(lower=0)

    total_incidents = len(df)
    cols_lower = {c.strip().lower(): c for c in df.columns}
    status_col = next((c for c in df.columns if c.strip().lower() == 'status'), None)
    unresolved = int((df[status_col].astype(str).str.upper().isin(['OUVERT', 'OPEN'])).sum()) if status_col else 0
    esc_col = cols_lower.get('escalade')
    reg_col = cols_lower.get('région') or cols_lower.get('region')

    # ── CATEGORIES (Backhaul vs BackBone) ────────────────────────────────────
    cat_col = None
    for col in df.columns:
        vals_lower = df[col].dropna().astype(str).str.strip().str.lower()
        if vals_lower.isin(['backhaul', 'backbone', 'back haul', 'back bone']).any():
            cat_col = col
            break
    if not cat_col and esc_col:
        esc_vals = df[esc_col].dropna().astype(str).str.lower()
        if esc_vals.str.contains('backhaul|backbone', na=False).any():
            cat_col = esc_col

    def _normalize_cat(v):
        v_lower = str(v).strip().lower()
        if 'backhaul' in v_lower or 'back haul' in v_lower:
            return 'Backhaul'
        if 'backbone' in v_lower or 'back bone' in v_lower:
            return 'BackBone'
        return str(v).strip()

    categories = []
    if cat_col:
        df['_cat'] = df[cat_col].apply(_normalize_cat)
        for cat, grp in df.groupby('_cat'):
            n = len(grp); td = float(grp['_dur'].sum()); mttr = td / n if n else 0
            categories.append({'name': cat, 'count': n, 'total_dur_sec': td,
                                'mttr_sec': mttr, 'mttr': _fmt(mttr)})
        categories.sort(key=lambda x: x['count'], reverse=True)
    elif esc_col:
        for val, grp in df.groupby(esc_col):
            val = str(val).strip()
            if not val or val.lower() == 'nan': continue
            n = len(grp); td = float(grp['_dur'].sum()); mttr = td / n if n else 0
            categories.append({'name': val, 'count': n, 'total_dur_sec': td,
                                'mttr_sec': mttr, 'mttr': _fmt(mttr)})
        categories.sort(key=lambda x: x['count'], reverse=True)

    # ── REGION × METIER ──────────────────────────────────────────────────────
    region_metier = {}
    if reg_col and esc_col:
        for region, rgrp in df.groupby(reg_col):
            region = str(region).strip()
            if not region or region.lower() == 'nan': continue
            metier_list = []
            for metier, mgrp in rgrp.groupby(esc_col):
                metier = str(metier).strip()
                if not metier or metier.lower() == 'nan': continue
                n = len(mgrp); td = float(mgrp['_dur'].sum()); mttr = td / n if n else 0
                metier_list.append({'name': metier, 'count': n, 'total_dur_sec': td,
                                    'mttr_sec': mttr, 'mttr': _fmt(mttr)})
            metier_list.sort(key=lambda x: x['count'], reverse=True)
            region_metier[region] = metier_list

    # ── PARTENAIRES / Clients IPT & IPLC ─────────────────────────────────────
    partenaires = []
    part_col = None
    for keyword in ('partenaire', 'client', 'liens', 'lien', 'service client'):
        if keyword in cols_lower:
            part_col = cols_lower[keyword]
            break

    if part_col:
        for name, grp in df.groupby(part_col):
            name = str(name).strip()
            if not name or name.lower() == 'nan': continue
            n = len(grp); td = float(grp['_dur'].sum())
            taux = max(0.0, 100.0 - (td / period_secs * 100)) if period_secs > 0 else 100.0
            partenaires.append({'name': name, 'nbre_inc': n, 'total_dur_sec': td,
                                'duree': _fmt(td), 'taux_dispo': round(taux, 2)})
        partenaires.sort(key=lambda x: x['taux_dispo'])

    # ── Top sites / causes ────────────────────────────────────────────────────
    top_sites = []
    site_col = next((c for c in ('Site Name', 'site_name') if c in df.columns), None)
    if site_col:
        for nm, cnt in df[site_col].value_counts().head(10).items():
            top_sites.append({'name': str(nm), 'count': int(cnt)})

    top_causes = []
    cause_col = next((c for c in df.columns if c.strip().lower() == 'cause'), None)
    if cause_col:
        cause_dur = defaultdict(float)
        for _, row in df.iterrows():
            c = str(row.get(cause_col, '')).strip()
            if c and c != 'nan': cause_dur[c] += float(row.get('_dur', 0) or 0)
        top_causes = [{'name': k, 'duration_sec': v}
                      for k, v in sorted(cause_dur.items(), key=lambda x: x[1], reverse=True)[:10]]

    # ── Sauvegarder ──────────────────────────────────────────────────────────
    report = existing if (existing and overwrite) else UploadedReport()
    report.user              = user
    report.original_filename = f"{label}.xlsx"
    report.date_rapport      = d_debut
    report.date_fin          = d_fin if d_fin != d_debut else None
    report.processed         = True
    report.source            = "api"
    report.total_incidents   = total_incidents
    report.unresolved_count  = unresolved
    report.total_duration_sec = int(df['_dur'].sum())
    report.synthesis_json    = []
    report.top_sites_json    = top_sites
    report.top_causes_json   = top_causes
    report.region_sites_json = {}
    report.outage_journalier_json = {}
    report.fixe_stats_json   = {}
    report.transmission_stats_json = {
        'total': total_incidents,
        'categories': categories,
        'region_metier': region_metier,
        'partenaires': partenaires,
    }
    report.uploaded_at = timezone.now()
    report.save()

    result["created"] += 1
    logger.info("Rapport TRANSMISSION créé : %s (id=%s, %d incidents)", label, report.pk, total_incidents)
    return result
