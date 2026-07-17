"""
Commande de backfill : recalcule cause_par_escalade_json pour tous les
UploadedReport qui ont un fichier Excel détaillé (colonnes Escalade + Cause
+ Duration) mais dont ce champ est encore vide (dict vide = {}).

Utilisé par les filtres « par escalade » des graphes Incident par Cause
(Nombre / Durée) de la page Statistiques.

Usage :
    python manage.py backfill_cause_escalade
    python manage.py backfill_cause_escalade --dry-run
    python manage.py backfill_cause_escalade --limit 50
    python manage.py backfill_cause_escalade --force
    python manage.py backfill_cause_escalade --refetch-api   # ré-importe via l'API
                                                             # les rapports API sans fichier
"""

from django.core.management.base import BaseCommand

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
    help = 'Backfill cause_par_escalade_json depuis les fichiers Excel détaillés'

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
        parser.add_argument(
            '--refetch-api', action='store_true',
            help="Ré-importer via l'API netXcare les rapports API_MOBILE_* sans "
                 "fichier détaillé (nécessite l'accès à l'API)"
        )

    def handle(self, *args, **options):
        import pandas as pd

        dry_run = options['dry_run']
        limit   = options['limit']
        force   = options['force']

        qs = UploadedReport.objects.filter(processed=True).exclude(detailed_file='').order_by('-date_rapport')
        if not force:
            qs = qs.filter(cause_par_escalade_json={})
        if limit:
            qs = qs[:limit]

        total = qs.count()
        self.stdout.write(f"Rapports à traiter : {total}")

        done = skipped = errors = 0
        for report in qs:
            try:
                path = report.detailed_file.path
            except (ValueError, AttributeError):
                skipped += 1
                continue
            try:
                df = pd.read_excel(path)
            except Exception as exc:
                self.stderr.write(f"  ✗ {report.pk} : lecture impossible ({exc})")
                errors += 1
                continue

            cause_col = next((c for c in ('Cause', 'Root Cause') if c in df.columns), None)
            if not cause_col or 'Escalade' not in df.columns or 'Duration' not in df.columns:
                skipped += 1
                continue

            cause_par_esc: dict = {}
            for _, row in df.iterrows():
                cause = str(row.get(cause_col, '')).strip()
                esc   = str(row.get('Escalade', '')).strip()
                if not cause or cause == 'nan' or not esc or esc.lower() == 'nan':
                    continue
                dur = _parse_hms(str(row.get('Duration', '')))
                _e = cause_par_esc.setdefault(esc, {})
                _c = _e.setdefault(cause, {'count': 0, 'duration_sec': 0.0})
                _c['count'] += 1
                _c['duration_sec'] += dur

            if not cause_par_esc:
                skipped += 1
                continue

            if dry_run:
                self.stdout.write(
                    f"  [dry-run] {report.pk} ({report.original_filename}) : "
                    f"{len(cause_par_esc)} escalades"
                )
            else:
                report.cause_par_escalade_json = cause_par_esc
                report.save(update_fields=['cause_par_escalade_json'])
            done += 1

        self.stdout.write(self.style.SUCCESS(
            f"Terminé : {done} peuplés, {skipped} ignorés, {errors} erreurs"
        ))

        if options['refetch_api']:
            self._refetch_api(dry_run)

    def _refetch_api(self, dry_run):
        """Ré-importe via l'API les rapports API_MOBILE_* sans ventilation ni fichier."""
        from reports.api_import import run_import

        qs = UploadedReport.objects.filter(
            processed=True,
            cause_par_escalade_json={},
            original_filename__startswith='API_MOBILE_',
        ).filter(detailed_file='').order_by('-date_rapport')

        total = qs.count()
        self.stdout.write(f"\nRapports API à ré-importer : {total}")
        ok = ko = 0
        for report in qs:
            d_from = report.date_rapport.isoformat()
            d_to   = (report.date_fin or report.date_rapport).isoformat()
            if dry_run:
                self.stdout.write(f"  [dry-run] {report.original_filename} ({d_from} → {d_to})")
                continue
            try:
                res = run_import(d_from, d_to, overwrite=True, network='mobile')
                if res.get('errors'):
                    self.stderr.write(f"  ✗ {report.original_filename} : {res['errors'][0]}")
                    ko += 1
                else:
                    self.stdout.write(f"  ✓ {report.original_filename}")
                    ok += 1
            except Exception as exc:
                self.stderr.write(f"  ✗ {report.original_filename} : {exc}")
                ko += 1
        if not dry_run:
            self.stdout.write(self.style.SUCCESS(
                f"Ré-import API terminé : {ok} OK, {ko} échecs"
            ))
