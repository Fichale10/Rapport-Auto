from django import template

register = template.Library()


@register.filter
def access(value, key):
    """Access a dictionary key that contains spaces or special characters."""
    if isinstance(value, dict):
        return value.get(key, '')
    return ''
