from django import template

register = template.Library()


@register.filter
def access(value, key):
    """Access a dictionary key that contains spaces or special characters."""
    if isinstance(value, dict):
        return value.get(key, '')
    return ''


_ESC_COLORS = {
    'ENERGIE':              '#f97316',
    'TRANS FH-FIELD O':     '#3b82f6',
    'TRANS FO':             '#6366f1',
    'TRANS IP':             '#8b5cf6',
    'TRANS FTTM':           '#a855f7',
    'RAN-FIELD O':          '#10b981',
    'PROJET':               '#64748b',
    'INFRA':                '#14b8a6',
    'BSS':                  '#ec4899',
    'ENVIRONNEMENT':        '#06b6d4',
    'ENERGIE / TRANS / RAN':'#ef4444',
    'TRANS / RAN':          '#f43f5e',
    'CORE SWITCH':          '#7c3aed',
    'GOOGLE':               '#22c55e',
}

@register.filter
def esc_color(escalade):
    """Retourne une couleur hex pour une escalade donnée."""
    return _ESC_COLORS.get(str(escalade).upper().strip(),
           _ESC_COLORS.get(str(escalade).strip(), '#94a3b8'))
