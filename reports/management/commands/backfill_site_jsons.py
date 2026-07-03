"""
Commande de backfill : recalcule site_duration_json et site_top_cause_json
pour tous les UploadedReport qui ont un fichier détaillé Excel mais dont
ces champs sont encore vides (dict vide = {}).

Usage :
    python manage.py backfill_site_jsons
    python manage.py backfill_site_jsons --dry-run
    python manage.py backfill_site_jsons --limit 50
"""

from django.core.management.base import BaseCommand
from django.conf import settings

from reports.models import UploadedReport


def _parse_hms(s):
    try:
        parts = str(s).split(':')
        if len(parts) == 3:
            return int(float(parts[0])) * 3600 + int(float(parts[1])) * 60 + int(float(parts[2]))
    except Exception:
        pass
    return 0


class Command(BaseCommand):
    help = 'Backfill site_duration_json et site_top_cause_json depuis les fichiers Excel détaillés'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run', action='store_true',
            help="Affiche les résultats sans sauvegarder"
        )
        parser.add_argument(
            '--limit', type=int, default=0,
            help="Nombre maximum de rapports à traiter (0 = tous)"
        )
        parser.add_argument(
            '--force', action='store_true',
            help="Retraiter même les rapports déjà peuplés"
        )

    def handle(self, *args, **options):
        import pandas as pd

        dry_run = options['dry_run']
        limit   = options['limit']
        force   = options['force']

        qs = UploadedReport.objects.filter(processed=True).exclude(detailed_file='').order_by('-date_rapport')
        if not force:
            # Seulement ceux dont au moins un champ est vide
            qs = qs.filter(site_top_cause_json={})
        if limit:
            qs = qs[:limit]

        total = qs.count()
        self.stdout.write(f"Rapports à traiter : {total}")

        ok = skipped = errors = 0

        for i, report in enumerate(qs, 1):
            file_name = report.detailed_file.name or ''
            if not (('results/' in file_name or 'results\\' in file_name)
                    and file_name.endswith('_detailed.xlsx')):
                self.stdout.write(f"  [{i}/{total}] {report.original_filename} — fichier non détaillé, ignoré")
                skipped += 1
                continue

            try:
                path = report.detailed_file.path
                df = pd.read_excel(path)
            except Exception as e:
                self.stdout.write(self.style.ERROR(f"  [{i}/{total}] {report.original_filename} — lecture Excel échouée : {e}"))
                errors += 1
                continue

            site_col  = next((c for c in df.columns if c.strip().lower() == 'site name'), None)
            cause_col = next((c for c in ('Root Cause', 'Cause') if c in df.columns), None)
            dur_col   = 'Duration' if 'Duration' in df.columns else None

            # Sites connus de top_sites_json — on filtre pour éviter les doublons
            # (df_export contient les sites enfants ; top_sites_json utilise df_dedup dédoublonné)
            valid_sites = {s['name'] for s in report.top_sites_json} if report.top_sites_json else None

            # ── site_duration_json ──
            sd: dict = {}
            if site_col and dur_col:
                for _, row in df.iterrows():
                    s = str(row.get(site_col, '')).strip()
                    d = _parse_hms(row.get(dur_col, ''))
                    if s and s != 'nan' and d > 0:
                        sd[s] = sd.get(s, 0) + d

            # ── site_top_cause_json — filtré sur les sites de top_sites_json ──
            sc: dict = {}
            if site_col and cause_col:
                for _, row in df.iterrows():
                    s = str(row.get(site_col, '')).strip()
                    # Ignorer les sites non présents dans top_sites_json (sites enfants/doublons)
                    if valid_sites and s not in valid_sites:
                        continue
                    c = str(row.get(cause_col, '')).strip()
                    if s and s != 'nan' and c and c != 'nan':
                        if s not in sc:
                            sc[s] = {}
                        sc[s][c] = sc[s].get(c, 0) + 1

            top_cause = {
                s: max(causes, key=causes.get)
                for s, causes in sc.items() if causes
            }

            self.stdout.write(
                f"  [{i}/{total}] {report.original_filename} "
                f"— {len(sd)} sites durée, {len(top_cause)} sites cause"
            )

            if not dry_run:
                if sd:
                    report.site_duration_json = sd
                if top_cause:
                    report.site_top_cause_json = top_cause
                report.save(update_fields=['site_duration_json', 'site_top_cause_json'])

            ok += 1

        self.stdout.write("")
        self.stdout.write(self.style.SUCCESS(
            f"Terminé — OK: {ok}  |  Ignorés: {skipped}  |  Erreurs: {errors}"
            + (" [DRY-RUN, rien sauvegardé]" if dry_run else "")
        ))
