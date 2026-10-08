# -*- coding: utf-8 -*-
"""FilterMan 2.0 — Entry point."""
from __future__ import unicode_literals
__title__            = u"FilterMan"
__doc__              = u"""Version = 2.0
Date    = 24.06.2026
FilterMan — advanced element filter and selection browser.
Author: Marcin Marek"""
__persistentengine__ = True

import os
import sys
import clr
clr.AddReference('PresentationFramework')

THIS_DIR          = os.path.dirname(os.path.abspath(__file__))
SPLASH_PATH_DARK  = os.path.join(THIS_DIR, u'resources', u'splash-screen-dark.png')
SPLASH_PATH_LIGHT = os.path.join(THIS_DIR, u'resources', u'splash-screen-light.png')

# ── Singleton guard ────────────────────────────────────────────────────
# _window persists between button clicks via __persistentengine__.
# Second click focuses the existing window instead of rebuilding.
try:
    _window
except NameError:
    _window = None

if _window is not None:
    try:
        # IsLoaded becomes False after Window.Close() — Focus() on a closed
        # WPF window returns False silently without raising, so we must check
        # IsLoaded explicitly to detect that the user closed the window.
        if not _window.IsLoaded:
            _window = None
        else:
            _window.Focus()
            _window.Activate()
    except Exception:
        _window = None

if _window is None:
    # ── Purge stale module cache ───────────────────────────────────────
    for _m in [u'splash', u'ui', u'view_actions', u'filter_chips',
               u'advanced_panel', u'excel_handler', u'filter_storage',
               u'unit_utils', u'settings']:
        sys.modules.pop(_m, None)

    import time
    from pyrevit import HOST_APP

    # ── Load persisted dark_mode before splash so colours match ───────
    _dark_mode = False
    try:
        from settings import load_settings as _load_s
        _dark_mode = bool(_load_s().get(u'dark_mode', False))
    except Exception:
        pass

    # ── Splash ────────────────────────────────────────────────────────
    _splash       = None
    _splash_mod   = None
    _splash_start = time.time()

    _splash_path = SPLASH_PATH_DARK if _dark_mode else SPLASH_PATH_LIGHT
    if os.path.isfile(_splash_path):
        import splash as _splash_mod
        _splash = _splash_mod.show_splash(_splash_path, dark_mode=_dark_mode)

    from ui import BrowserWindow

    uidoc   = HOST_APP.uidoc
    doc     = uidoc.Document
    # Pass splash + start time; BrowserWindow closes it after tree build completes.
    _window = BrowserWindow(uidoc, doc, splash=_splash, splash_start=_splash_start)
    _window.Show()
