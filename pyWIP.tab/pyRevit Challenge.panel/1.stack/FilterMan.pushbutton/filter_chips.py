# -*- coding: utf-8 -*-
# FilterMan — filter_chips.py
# Data model and evaluation logic for Advanced Search filter chips.
# No Revit API, no WPF — pure IronPython 2.7 logic.

# ── Fields ────────────────────────────────────────────────────────────────────

FIELDS = [
    (u'family',   u'Family Name'),
    (u'type',     u'Type Name'),
    (u'category', u'Category'),
    (u'level',    u'Level'),
    (u'param',    u'Parameter...'),
]

# ── Operators ─────────────────────────────────────────────────────────────────

TEXT_OPERATORS = [
    (u'equals',          u'equals'),
    (u'not_equals',      u'does not equal'),
    (u'contains',        u'contains'),
    (u'not_contains',    u'does not contain'),
    (u'begins_with',     u'begins with'),
    (u'not_begins_with', u'does not begin with'),
    (u'ends_with',       u'ends with'),
    (u'not_ends_with',   u'does not end with'),
    (u'has_value',       u'has a value'),
    (u'no_value',        u'has no value'),
]

# Level uses exact-match operators only — value comes from model dropdown.
LEVEL_OPERATORS = [
    (u'equals',     u'equals'),
    (u'not_equals', u'does not equal'),
    (u'has_value',  u'has a value'),
    (u'no_value',   u'has no value'),
]

# Numeric operators (v2 — implemented in eval_chip, not surfaced in v1 UI).
NUMERIC_OPERATORS = [
    (u'eq_num',  u'equals'),
    (u'neq_num', u'does not equal'),
    (u'gt',      u'is greater than'),
    (u'gte',     u'is greater than or equal to'),
    (u'lt',      u'is less than'),
    (u'lte',     u'is less than or equal to'),
    (u'has_value', u'has a value'),
    (u'no_value',  u'has no value'),
]

# Value-less operators — no text input needed in the UI.
NO_VALUE_OPS = frozenset([u'has_value', u'no_value'])


def get_operators_for_field(field):
    """Return [(key, label), ...] for the given field."""
    if field == u'level':
        return LEVEL_OPERATORS
    return TEXT_OPERATORS


# ── FilterChip ────────────────────────────────────────────────────────────────

class FilterChip(object):
    __slots__ = (u'field', u'operator', u'value', u'display',
                 u'field_type', u'param_name')

    def __init__(self, field, operator, value, display,
                 field_type=u'text', param_name=None):
        self.field      = field       # str: 'family'|'type'|'category'|'level'|'param'
        self.operator   = operator    # str: see TEXT_OPERATORS keys
        self.value      = value       # str | float | None  (None for has/no value)
        self.display    = display     # str: human-readable label shown on chip
        self.field_type = field_type  # 'text' | 'numeric'
        self.param_name = param_name  # str | None  (set when field='param')

    def __repr__(self):
        return u'<FilterChip {}>'.format(self.display)


def make_display(field_label, op_label, value, param_name=None):
    """Build the human-readable string shown on the chip."""
    label = param_name if param_name else field_label
    if op_label in (u'has a value', u'has no value'):
        return u'{} {}'.format(label, op_label)
    return u'{} {} "{}"'.format(label, op_label, value)


# ── Evaluation ────────────────────────────────────────────────────────────────

def eval_chip(chip, text_value):
    """Evaluate one chip against a string value.

    text_value: str (the field's value for this element) or None (absent).
    Returns bool.
    """
    op = chip.operator

    if op == u'has_value':
        return text_value is not None and text_value != u''
    if op == u'no_value':
        return text_value is None or text_value == u''

    if text_value is None:
        return False

    chip_val = chip.value.lower() if chip.value else u''
    text     = text_value.lower()

    if op == u'equals':          return text == chip_val
    if op == u'not_equals':      return text != chip_val
    if op == u'contains':        return chip_val in text
    if op == u'not_contains':    return chip_val not in text
    if op == u'begins_with':     return text.startswith(chip_val)
    if op == u'not_begins_with': return not text.startswith(chip_val)
    if op == u'ends_with':       return text.endswith(chip_val)
    if op == u'not_ends_with':   return not text.endswith(chip_val)

    # Numeric operators (v2) — only reached when field_type='numeric'.
    if chip.field_type == u'numeric':
        try:
            num_text = float(text_value)
            num_val  = float(chip.value) if chip.value is not None else 0.0
            if op == u'eq_num':  return num_text == num_val
            if op == u'neq_num': return num_text != num_val
            if op == u'gt':      return num_text >  num_val
            if op == u'gte':     return num_text >= num_val
            if op == u'lt':      return num_text <  num_val
            if op == u'lte':     return num_text <= num_val
        except (ValueError, TypeError):
            return False

    return False


def eval_chips(chips, field_values, match_mode=u'ALL'):
    """Evaluate a list of chips against a dict of field values.

    field_values = {
        'family':   str | None,
        'type':     str | None,
        'category': str | None,
        'level':    str | None,
        'params':   [(param_name, param_val), ...],
    }
    match_mode: 'ALL' (every chip must pass) | 'ANY' (at least one must pass).
    Returns bool.  Empty chips list → True (no filter active).
    """
    if not chips:
        return True

    results = []
    for chip in chips:
        if chip.field == u'param':
            params = field_values.get(u'params') or []
            if chip.param_name:
                # Specific parameter by name
                text_val = _find_param_value(chip.param_name, params)
                results.append(eval_chip(chip, text_val))
            else:
                # No name specified: match if ANY parameter value passes
                results.append(any(eval_chip(chip, val) for _n, val in params))
        else:
            text_val = field_values.get(chip.field)
            results.append(eval_chip(chip, text_val))

    if match_mode == u'ALL':
        return all(results)
    return any(results)


def _find_param_value(param_name, params):
    """Return the value of the first param whose name matches (case-insensitive)."""
    if not param_name:
        return None
    needle = param_name.lower()
    for name, val in params:
        if name.lower() == needle:
            return val
    return None
