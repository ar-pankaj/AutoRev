# -*- coding: utf-8 -*-
"""FilterMan v2 - view_actions.py

Selection (read-only): direct API calls - no Transaction, allowed from WPF handler.
View mutations (Isolate, Hide, Reset, Color): ExternalEvent pattern - Transaction
requires Revit API context; ExternalEvent.Raise() queues execution for next Revit
idle tick when that context is available.

Usage:
    Call init_events() once from BrowserWindow.__init__ (Revit API context).
    Then call apply_* from any WPF handler.
"""
from __future__ import unicode_literals
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, TaskDialog
from Autodesk.Revit.DB import (
    Transaction, TemporaryViewMode, ElementId,
    OverrideGraphicSettings, FilteredElementCollector
)
from System.Collections.Generic import List

_handlers = {}
_events   = {}


def _to_net_list(ids):
    net = List[ElementId]()
    for eid in ids:
        net.Add(eid)
    return net


def _expand_with_subcomponents(ids, doc):
    """Expand ElementId list to include nested shared family sub-instances.

    When a host family contains only nested Shared families, isolating/hiding
    the host ID alone shows nothing. GetSubComponentIds() returns the IDs of
    all nested shared instances so the caller can act on the full set.
    """
    from Autodesk.Revit.DB import FamilyInstance
    try:
        int64 = long
    except NameError:
        int64 = int
    seen = set()
    for eid in ids:
        try:
            seen.add(eid.Value)
        except AttributeError:
            seen.add(int64(eid.IntegerValue))
    result = list(ids)
    for eid in ids:
        el = doc.GetElement(eid)
        if not isinstance(el, FamilyInstance):
            continue
        try:
            for sub_id in el.GetSubComponentIds():
                try:
                    k = sub_id.Value
                except AttributeError:
                    k = int64(sub_id.IntegerValue)
                if k not in seen:
                    result.append(sub_id)
                    seen.add(k)
        except Exception:
            pass
    return result


# ── ExternalEvent handler base ─────────────────────────────────────────────

class _BaseHandler(IExternalEventHandler):
    def __init__(self, name):
        self._name = name
        self.ids   = []
        self.extra = None

    def GetName(self):
        return u'FilterMan: ' + self._name

    def Execute(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            return
        doc  = uidoc.Document
        view = uidoc.ActiveView
        try:
            self._run(doc, view, self.ids, self.extra)
        except Exception as ex:
            TaskDialog.Show(u'FilterMan - ' + self._name, unicode(ex))

    def _run(self, doc, view, ids, extra):
        raise NotImplementedError


# ── Concrete handlers ──────────────────────────────────────────────────────

class _IsolateHandler(_BaseHandler):
    def __init__(self):
        _BaseHandler.__init__(self, u'Isolate')

    def _run(self, doc, view, ids, extra):
        if not ids:
            return
        ids = _expand_with_subcomponents(ids, doc)
        t = Transaction(doc, u'FilterMan: Isolate')
        t.Start()
        try:
            view.IsolateElementsTemporary(_to_net_list(ids))
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise


class _TempHideHandler(_BaseHandler):
    def __init__(self):
        _BaseHandler.__init__(self, u'Temp Hide')

    def _run(self, doc, view, ids, extra):
        if not ids:
            return
        ids = _expand_with_subcomponents(ids, doc)
        t = Transaction(doc, u'FilterMan: Temp Hide')
        t.Start()
        try:
            view.HideElementsTemporary(_to_net_list(ids))
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise


class _ResetHandler(_BaseHandler):
    def __init__(self):
        _BaseHandler.__init__(self, u'Reset View Mode')

    def _run(self, doc, view, ids, extra):
        from Autodesk.Revit.DB import View3D
        has_isolate = view.IsTemporaryHideIsolateActive()
        is_3d       = isinstance(view, View3D)
        has_selbox  = is_3d and view.IsSectionBoxActive
        if not has_isolate and not has_selbox:
            return
        t = Transaction(doc, u'FilterMan: Reset View Mode')
        t.Start()
        try:
            if has_isolate:
                view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
            if has_selbox:
                view.IsSectionBoxActive = False
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise


class _ColorHandler(_BaseHandler):
    def __init__(self):
        _BaseHandler.__init__(self, u'Color Override')

    def _run(self, doc, view, ids, extra):
        # extra = (wpf_color, reset)
        wpf_color, reset = extra if extra is not None else (None, True)
        ogs = OverrideGraphicSettings()
        if not reset and wpf_color is not None:
            from Autodesk.Revit.DB import Color as RvColor, FillPatternElement
            rv_col = RvColor(int(wpf_color.R), int(wpf_color.G), int(wpf_color.B))
            ogs.SetProjectionLineColor(rv_col)
            try:
                ogs.SetSurfaceForegroundPatternColor(rv_col)
                ogs.SetSurfaceForegroundPatternVisible(True)
                # Find the Solid Fill pattern and apply it
                solid_id = None
                for fp in FilteredElementCollector(doc)\
                              .OfClass(FillPatternElement).ToElements():
                    if fp.GetFillPattern().IsSolidFill:
                        solid_id = fp.Id
                        break
                if solid_id is not None:
                    ogs.SetSurfaceForegroundPatternId(solid_id)
            except Exception:
                pass  # surface pattern API unavailable in some Revit versions
        ids = _expand_with_subcomponents(ids, doc)
        t = Transaction(doc, u'FilterMan: Color Override')
        t.Start()
        try:
            for eid in ids:
                view.SetElementOverrides(eid, ogs)
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise


class _ExcelPreviewHandler(_BaseHandler):
    """Dry-run: read element params from Revit, compute green/red flags, no Transaction."""
    def __init__(self):
        _BaseHandler.__init__(self, u'Excel Preview')
        self.import_data = []
        self.window      = None

    def _run(self, doc, view, ids, extra):
        from excel_handler import build_preview_data
        preview_data = build_preview_data(doc, self.import_data)
        win = self.window
        if win is not None:
            import System
            win.Dispatcher.Invoke(
                System.Action(lambda: win._show_import_preview(preview_data))
            )


class _ExcelImportHandler(_BaseHandler):
    def __init__(self):
        _BaseHandler.__init__(self, u'Excel Import')
        self.import_data = []
        self.window      = None   # BrowserWindow reference for Dispatcher callback

    def _run(self, doc, view, ids, extra):
        from excel_handler import apply_excel_import as _xlsx_apply
        result = _xlsx_apply(doc, self.import_data)
        win    = self.window
        if win is not None:
            import System
            win.Dispatcher.Invoke(
                System.Action(lambda: win._after_import(result))
            )
        else:
            TaskDialog.Show(
                u'FilterMan - Import',
                u'Updated: {}, Skipped: {}'.format(
                    result[u'updated'], result[u'skipped'])
            )


class _SelectionBoxHandler(_BaseHandler):
    _PAD = 1.0  # feet (~30 cm) padding around computed bbox

    def __init__(self):
        _BaseHandler.__init__(self, u'Selection Box')

    def Execute(self, uiapp):
        from Autodesk.Revit.DB import BoundingBoxXYZ, XYZ, View3D

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            return
        doc  = uidoc.Document
        view = uidoc.ActiveView
        ids  = self.ids
        if not ids:
            return

        try:
            # --- compute union bounding box ---
            INF = 1e10
            mn  = [INF,  INF,  INF]
            mx  = [-INF, -INF, -INF]
            for eid in ids:
                try:
                    el = doc.GetElement(eid)
                    if el is None:
                        continue
                    bb = el.get_BoundingBox(None)
                    if bb is None:
                        continue
                    mn[0] = min(mn[0], bb.Min.X)
                    mn[1] = min(mn[1], bb.Min.Y)
                    mn[2] = min(mn[2], bb.Min.Z)
                    mx[0] = max(mx[0], bb.Max.X)
                    mx[1] = max(mx[1], bb.Max.Y)
                    mx[2] = max(mx[2], bb.Max.Z)
                except Exception:
                    continue

            if mn[0] >= INF:
                TaskDialog.Show(
                    u'FilterMan',
                    u'No bounding box could be computed for the checked elements.')
                return

            p = self._PAD
            box = BoundingBoxXYZ()
            box.Min = XYZ(mn[0] - p, mn[1] - p, mn[2] - p)
            box.Max = XYZ(mx[0] + p, mx[1] + p, mx[2] + p)

            # --- find target 3D view ---
            # Priority: 1) active view if orthographic 3D
            #           2) any open 3D view (already in a UI tab)
            #           3) first non-template orthographic 3D in model
            def _is_ok_3d(v):
                if v is None or v.IsTemplate:
                    return False
                if not isinstance(v, View3D):
                    return False
                return not getattr(v, u'IsPerspective', False)

            view3d = None
            if _is_ok_3d(view):
                view3d = view

            if view3d is None:
                for uiv in uidoc.GetOpenUIViews():
                    v = doc.GetElement(uiv.ViewId)
                    if _is_ok_3d(v):
                        view3d = v
                        break

            if view3d is None:
                for v in (FilteredElementCollector(doc)
                          .OfClass(View3D)
                          .ToElements()):
                    if _is_ok_3d(v):
                        view3d = v
                        break

            if view3d is None:
                TaskDialog.Show(
                    u'FilterMan',
                    u'No orthographic 3D view found in the model. '
                    u'Please create a 3D view first.')
                return

            # --- apply section box ---
            t = Transaction(doc, u'FilterMan: Selection Box')
            t.Start()
            try:
                view3d.SetSectionBox(box)
                view3d.IsSectionBoxActive = True
                t.Commit()
            except Exception:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                raise

            # --- switch to view3d if it is not the currently active view ---
            try:
                if view.Id != view3d.Id:
                    uidoc.RequestViewChange(view3d)
            except Exception:
                pass

        except Exception as ex:
            TaskDialog.Show(u'FilterMan - Selection Box', unicode(ex))

    def _run(self, doc, view, ids, extra):
        pass  # logic lives in Execute override


# ── Initialisation (call once from Revit API context) ─────────────────────

def init_events():
    """Create IExternalEventHandler + ExternalEvent pairs.
    Must be called while in Revit API context (e.g. from BrowserWindow.__init__).
    Safe to call multiple times - subsequent calls are no-ops.
    """
    global _handlers, _events
    if _events:
        return
    pairs = [
        (u'isolate',        _IsolateHandler()),
        (u'temp_hide',      _TempHideHandler()),
        (u'reset',          _ResetHandler()),
        (u'color',          _ColorHandler()),
        (u'selbox',         _SelectionBoxHandler()),
        (u'excel_preview',  _ExcelPreviewHandler()),
        (u'excel_import',   _ExcelImportHandler()),
    ]
    for key, h in pairs:
        _handlers[key] = h
        _events[key]   = ExternalEvent.Create(h)


def _raise(key, ids, extra=None):
    if key not in _events:
        TaskDialog.Show(u'FilterMan',
                        u'View actions not initialised - please restart FilterMan.')
        return
    _handlers[key].ids   = list(ids)
    _handlers[key].extra = extra
    _events[key].Raise()


# ── Public API ─────────────────────────────────────────────────────────────

def apply_selection(ids, mode=u'replace'):
    """Select elements by ElementId. No Transaction required."""
    from pyrevit import HOST_APP
    try:
        HOST_APP.uidoc.Selection.SetElementIds(_to_net_list(ids))
    except Exception as ex:
        TaskDialog.Show(u'FilterMan', u'Selection error: ' + unicode(ex))


def apply_select_all(scope=u'view'):
    """Select all non-type elements in view or project. No Transaction required."""
    from pyrevit import HOST_APP
    uidoc = HOST_APP.uidoc
    doc   = uidoc.Document
    try:
        if scope == u'view':
            coll = (FilteredElementCollector(doc, uidoc.ActiveView.Id)
                    .WhereElementIsNotElementType()
                    .ToElementIds())
        else:
            coll = (FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .ToElementIds())
        uidoc.Selection.SetElementIds(coll)
    except Exception as ex:
        TaskDialog.Show(u'FilterMan', u'Select all error: ' + unicode(ex))


def apply_isolate(ids):
    _raise(u'isolate', ids)


def apply_temp_hide(ids):
    _raise(u'temp_hide', ids)


def apply_reset():
    _raise(u'reset', [])


def apply_color(ids, wpf_color=None, reset=False):
    _raise(u'color', ids, extra=(wpf_color, reset))


def apply_selection_box(ids):
    _raise(u'selbox', ids)


def apply_excel_preview(import_data, window=None):
    _handlers[u'excel_preview'].import_data = import_data
    _handlers[u'excel_preview'].window      = window
    _events[u'excel_preview'].Raise()


def apply_excel_import(import_data, window=None):
    _handlers[u'excel_import'].import_data = import_data
    _handlers[u'excel_import'].window      = window
    _events[u'excel_import'].Raise()
