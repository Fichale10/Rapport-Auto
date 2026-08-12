"""
Commande de backfill : recalcule unresolved_details_json / resolved_details_json
(détail incident par incident : site, région, alarm time, durée, cause) pour les
UploadedReport réseau MOBILE qui ont des incidents mais aucun détail enregistré
(champ ajouté après coup — anciens rapports non retraités depuis).

C'est ce détail qui alimente le modal "Synthèse par Escalade" de l'accueil ;
sans lui, le modal affiche "Détail indisponible... Relancez avec MAJ forcée"
— message qui persiste tant que le rapport concerné n'a pas été retraité,
car "MAJ forcée" ne retraite que la période/le réseau actuellement affichés,
pas les anciens rapports Excel importés manuellement.

Stratégie par rapport :
  1. Si le fichier Excel original (`file`) est encore sur disque → retraité
     localement via treatement.process_file (pas d'appel réseau).
  2. Sinon, si source == 'api' → réimport via l'API ticketing (run_import,
     overwrite=True) sur exactement la période du rapport.
  3. Sinon → ignoré (aucune source disponible).

Usage :
    python manage.py backfill_incident_details
    python manage.py backfill_incident_details --dry-run
    python manage.py backfill_incident_details --limit 50
"""
import os

from django.core.management.base import BaseCommand

from reports.models import UploadedReport


def _report_network(fname):
    name = (fname or '').upper()
    if name.startswith('API_') and len(name.split('_')) > 1:
        return name.split('_')[1]
    return 'MOBILE'


def _open_cell(row, col):
    v = row.get(col, '') if col else ''
    v = '' if v is None else str(v).strip()
    return '' if v.lower() in ('nan', 'nat') else v


def _details_from_dedup(df_dedup):
    """Reproduit exactement la logique de process_report (views.py)."""
    unresolved_details, resolved_details = [], []
    site_col = next((c for c in ('Site Name', 'Site name', 'SITE NAME') if c in df_dedup.columns), None)
    if not site_col or len(df_dedup) == 0:
        return unresolved_details, resolved_details

    if 'Status' in df_dedup.columns:
        open_mask = df_dedup['Status'].astype(str).str.upper() == 'OUVERT'
    elif 'Cancel Time' in df_dedup.columns:
        open_mask = df_dedup['Cancel Time'].isna()
    else:
        return unresolved_details, resolved_details

    reg_col   = next((c for c in ('Région', 'Region', 'REGION') if c in df_dedup.columns), None)
    cause_col = next((c for c in ('Cause', 'Root Cause') if c in df_dedup.columns), None)

    def _row(row):
        return {
            'site':       _open_cell(row, site_col),
            'escalade':   _open_cell(row, 'Escalade'),
            'region':     _open_cell(row, reg_col),
            'alarm_time': _open_cell(row, 'Alarm Time'),
            'duration':   _open_cell(row, 'Duration'),
            'cause':      _open_cell(row, cause_col),
        }

    for _, r in df_dedup.loc[open_mask].iterrows():
        unresolved_details.append(_row(r))
    for _, r in df_dedup.loc[~open_mask].iterrows():
        resolved_details.append(_row(r))
    return unresolved_details, resolved_details


class Command(BaseCommand):
    help = "Backfill unresolved_details_json / resolved_details_json (réseau mobile)"

    def add_arguments(self, parser):
        parser.add_argument('--dry-run', action='store_true', help="N'écrit rien, affiche seulement")
        parser.add_argument('--limit', type=int, default=0, help="Nombre max de rapports à traiter (0 = tous)")
        parser.add_argument('--skip-api', action='store_true', help="Ne pas tenter de réimport API (fichier local uniquement)")

    def handle(self, *args, **options):
        dry_run  = options['dry_run']
        limit    = options['limit']
        skip_api = options['skip_api']

        candidates = [
            r for r in UploadedReport.objects.filter(processed=True, total_incidents__gt=0)
            if _report_network(r.original_filename) == 'MOBILE'
            and not (r.unresolved_details_json or []) and not (r.resolved_details_json or [])
        ]
        if limit:
            candidates = candidates[:limit]

        total = len(candidates)
        self.stdout.write(f"Rapports à backfiller : {total}")

        ok_file = ok_api = skipped = errors = 0

        for i, report in enumerate(candidates, 1):
            tag = f"[{i}/{total}] {report.original_filename} ({report.date_rapport} → {report.date_fin or report.date_rapport})"

            if report.file and os.path.exists(report.file.path):
                try:
                    from treatement import process_file
                    _df_export, df_dedup, _df_synth = process_file(
                        report.file.path,
                        report.date_rapport.isoformat(),
                        date_fin=report.date_fin.isoformat() if report.date_fin else None,
                    )
                    unres, res = _details_from_dedup(df_dedup)
                except Exception as e:
                    self.stdout.write(self.style.ERROR(f"  {tag} — retraitement fichier échoué : {e}"))
                    errors += 1
                    continue

                self.stdout.write(f"  {tag} — fichier local : {len(unres)} non résolus, {len(res)} résolus")
                if not dry_run and (unres or res):
                    report.unresolved_details_json = unres
                    report.resolved_details_json = res
                    report.save(update_fields=['unresolved_details_json', 'resolved_details_json'])
                ok_file += 1
                continue

            if report.source == 'api' and not skip_api:
                try:
                    from reports.api_import import run_import
                    d_deb = report.date_rapport.isoformat()
                    d_fin = (report.date_fin or report.date_rapport).isoformat()
                    if dry_run:
                        self.stdout.write(f"  {tag} — [dry-run] réimport API prévu")
                        ok_api += 1
                        continue
                    res = run_import(d_deb, d_fin, overwrite=True, network='mobile')
                    if res.get('errors'):
                        self.stdout.write(self.style.ERROR(f"  {tag} — API : {res['errors']}"))
                        errors += 1
                        continue
                    self.stdout.write(f"  {tag} — réimporté via API")
                    ok_api += 1
                except Exception as e:
                    self.stdout.write(self.style.ERROR(f"  {tag} — réimport API échoué : {e}"))
                    errors += 1
                continue

            self.stdout.write(f"  {tag} — aucune source disponible, ignoré")
            skipped += 1

        self.stdout.write("")
        self.stdout.write(self.style.SUCCESS(
            f"Terminé — Fichier local: {ok_file}  |  API: {ok_api}  |  Ignorés: {skipped}  |  Erreurs: {errors}"
            + (" [DRY-RUN, rien sauvegardé]" if dry_run else "")
        ))
