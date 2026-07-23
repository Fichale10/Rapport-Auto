"""
Configuration centrale des plateformes et de leurs rapports.
Chaque plateforme déclare ses rapports Excel et PowerPoint.
"""

PLATFORMS = {
    'mobile-dr2': {
        'label':   'Réseau Mobile & DR2',
        'icon':    '📡',
        'color':   '#003087',
        'color2':  '#0047cc',
        'domains': ['mobile', 'dr2'],
        'excel_reports': [
            {
                'num':      1,
                'title':    'Rapport 1',
                'subtitle': 'Bases des Incidents',
                'desc':     'Onglets Réseau Mobile + DR2 générés depuis le fichier source',
                'url_name': 'platform_bases_incidents',
                'url_kwargs': {'platform': 'mobile-dr2'},
                'icon':     '📋',
                'badge':    'Excel',
            },
        ],
        'tool_reports': [
            {
                'num':      1,
                'title':    'Outil 1',
                'subtitle': 'DR2 Report',
                'desc':     'Indicateur ARCEP — rapport DR2 par région et escalade généré depuis fichier Excel',
                'url_name': 'dr2_daily',
                'url_kwargs': {},
                'icon':     '📱',
                'badge':    'Interactif',
            },
            {
                'num':      2,
                'title':    'Outil 2',
                'subtitle': 'Synthèse Mensuelle Mobile',
                'desc':     'Analyse mensuelle des incidents réseau mobile et violations DR2 par région',
                'url_name': 'reporting_network',
                'url_kwargs': {'platform': 'mobile-dr2'},
                'icon':     '📊',
                'badge':    'Interactif',
            },
            {
                'num':      3,
                'title':    'Outil 3',
                'subtitle': 'Rapport Mensuel Mobile',
                'desc':     'Synthèse mensuelle réseau mobile : DR1/DR2, efficacité par métier/région, points bloquants — depuis fichier Excel',
                'url_name': 'mobile_cgi',
                'url_kwargs': {},
                'icon':     '📡',
                'badge':    'Interactif',
            },
            {
                'num':      4,
                'title':    'Outil 4',
                'subtitle': 'Site Down — Micro-coupures',
                'desc':     'Consolidation mensuelle des alarmes NetAct sites down : Nb & durée par site/jour, Cause/Escalade automatiques',
                'url_name': 'site_down',
                'url_kwargs': {},
                'icon':     '📉',
                'badge':    'Interactif',
            },
        ],
        'import': {
            'domains': ['mobile', 'dr2'],
            'hint':    'Fichier brut ticketing mobile (RESEAU_MOBILE_*.xlsx) ou fichier Bases des Incidents',
            'accept':  '.xlsx,.xls',
        },
    },

    'fixe': {
        'label':   'Réseau Fixe',
        'icon':    '☎️',
        'color':   '#059669',
        'color2':  '#10b981',
        'domains': ['fixe'],
        'tool_reports': [
            {
                'num':      1,
                'title':    'Outil 1',
                'subtitle': 'Rapport Hebdo DCO — FTTH',
                'desc':     'Carte choroplèthe du Togo avec incidents FTTH par région et export PPTX',
                'url_name': 'fixe_rapport_ftth',
                'url_kwargs': {},
                'icon':     '🗺️',
                'badge':    'Interactif',
            },
        ],
        'import': {
            'domains': ['fixe'],
            'hint':    'Fichier brut ticketing réseau fixe',
            'accept':  '.xlsx,.xls',
        },
    },

    'transmission': {
        'label':   'Transport',
        'icon':    '🔗',
        'color':   '#d97706',
        'color2':  '#f59e0b',
        'domains': ['transport'],
        'tool_reports': [
            {
                'num':      1,
                'title':    'Outil 1',
                'subtitle': 'Rapport NOC Transport',
                'desc':     'Génération du rapport hebdomadaire NOC Transport depuis TRANSMISSION_*.xlsx',
                'url_name': 'transport_rapport_noc',
                'url_kwargs': {},
                'icon':     '📋',
                'badge':    'Interactif',
            },
            {
                'num':      2,
                'title':    'Outil 2',
                'subtitle': 'Rapport DCO Liens FO',
                'desc':     'Présentation interactive des coupures fibre optique avec carte et export PPTX',
                'url_name': 'transport_rapport_fo',
                'url_kwargs': {},
                'icon':     '🗺️',
                'badge':    'Interactif',
            },
        ],
        'import': {
            'domains': ['transport'],
            'hint':    'Fichier brut ticketing transport',
            'accept':  '.xlsx,.xls',
        },
    },

    'igw': {
        'label':   'IGW',
        'icon':    '🔌',
        'color':   '#7c3aed',
        'color2':  '#8b5cf6',
        'domains': ['igw'],
        'tool_reports': [
            {
                'num':      1,
                'title':    'Outil 1',
                'subtitle': 'Trafic International',
                'desc':     'Visualisation du trafic international par lien depuis les fichiers CSV',
                'url_name': 'igw_trafic_international',
                'url_kwargs': {},
                'icon':     '🌐',
                'badge':    'Interactif',
            },
        ],
        'import': {
            'domains': ['igw'],
            'hint':    'Fichier brut ticketing IGW',
            'accept':  '.xlsx,.xls',
        },
    },

    'core': {
        'label':   'Core Network',
        'icon':    '🌐',
        'color':   '#dc2626',
        'color2':  '#ef4444',
        'domains': ['core'],
        'tool_reports': [],
        'import': {
            'domains': ['core'],
            'hint':    'Fichier brut ticketing core',
            'accept':  '.xlsx,.xls',
        },
    },
}
