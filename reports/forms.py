from django import forms
from .models import UploadedReport


class UploadForm(forms.ModelForm):
    class Meta:
        model = UploadedReport
        fields = ['file', 'period_type', 'date_rapport', 'date_fin']
        widgets = {
            'period_type': forms.Select(attrs={
                'class': 'form-control',
                'id': 'periodType',
            }),
            'date_rapport': forms.DateInput(attrs={
                'type': 'date',
                'class': 'form-control',
                'id': 'id_date_rapport',
            }),
            'date_fin': forms.DateInput(attrs={
                'type': 'date',
                'class': 'form-control',
                'id': 'id_date_fin',
            }),
            'file': forms.ClearableFileInput(attrs={
                'accept': '.xlsx,.xls',
                'class': 'form-control',
                'id': 'fileInput',
            }),
        }
        labels = {
            'file': 'Fichier Excel',
            'period_type': 'Type de période',
            'date_rapport': 'Date de début',
            'date_fin': 'Date de fin',
        }

    def clean(self):
        cleaned = super().clean()
        period = cleaned.get('period_type')
        date_debut = cleaned.get('date_rapport')
        date_fin = cleaned.get('date_fin')
        if period and period != UploadedReport.PERIOD_DAY:
            if not date_fin:
                self.add_error('date_fin', 'Veuillez saisir la date de fin.')
            elif date_debut and date_fin < date_debut:
                self.add_error('date_fin', 'La date de fin doit être postérieure ou égale à la date de début.')
        return cleaned
