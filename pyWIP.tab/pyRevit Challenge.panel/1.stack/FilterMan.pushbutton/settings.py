# -*- coding: utf-8 -*-
"""FilterMan persistent settings (dark_mode, etc.) stored as JSON."""
from __future__ import unicode_literals
import os
import json

_SETTINGS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), u'filterman_settings.json'
)


def load_settings():
    """Return settings dict; empty dict on any error."""
    try:
        with open(_SETTINGS_FILE, u'r') as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    """Merge *data* into persisted settings and write file."""
    current = load_settings()
    current.update(data)
    try:
        with open(_SETTINGS_FILE, u'w') as f:
            json.dump(current, f, indent=2)
    except Exception:
        pass
