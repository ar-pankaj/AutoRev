# -*- coding: utf-8 -*-

__title__ = "GridEnds\nName"

from __future__ import print_function

from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView

# -------------------------
# User parameters
# -------------------------
OFFSET_FT = 0.60          # offset distance from grid end, in feet
TEXT_TYPE_NAME = None     # e.g. "3mm Arial"; None = show UI and/or use last choice
DUP_TOLERANCE_FT = 0.25   # proximity tolerance for duplicate detection
# -------------------------

# -------------------------
# Guards: disallow 3D and Schedules
# -------------------------
if isinstance(view, DB.View3D) or isinstance(view, DB.ViewSchedule):
    forms.alert("Active view cannot host text notes (3D/Schedule).\nOpen a Plan/Section/Elevation.", exitscript=True)

# -------------------------
# XYZ helpers (avoid operator overloads)
# -------------------------
def xyz_sub(a, b):
    try:
        return a.Subtract(b)
    except Exception:
        return DB.XYZ(a.X - b.X, a.Y - b.Y, a.Z - b.Z)

def xyz_add(a, b):
    try:
        return a.Add(b)
    except Exception:
        return DB.XYZ(a.X + b.X, a.Y + b.Y, a.Z + b.Z)

def xyz_mul(a, s):
    try:
        return a.Multiply(s)
    except Exception:
        return DB.XYZ(a.X * s, a.Y * s, a.Z * s)

# -------------------------
# Geometry helpers
# -------------------------
def get_primary_curve_for_view(grid, view):
    """Prefer view-specific curve, else model curve, else grid.Curve."""
    try:
        cvs = grid.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view)
        cvs = list(cvs) if cvs else []
        if cvs:
            return cvs[0]
    except Exception:
        pass
    try:
        cvm = grid.GetCurvesInView(DB.DatumExtentType.Model, view)
        cvm = list(cvm) if cvm else []
        if cvm:
            return cvm[0]
    except Exception:
        pass
    return grid.Curve

def safe_norm(vec):
    try:
        v = vec.Normalize()
        if v.GetLength() > 0.0:
            return v
    except Exception:
        pass
    return DB.XYZ.BasisX

def perp_in_view(dir_vec, view):
    """Return a unit vector perpendicular to dir_vec in the view plane."""
    vdir = view.ViewDirection  # normal of the view plane
    perp = dir_vec.CrossProduct(vdir)
    if perp.GetLength() < 1e-8:
        # Degenerate case: fall back to cross with global Z
        perp = dir_vec.CrossProduct(DB.XYZ.BasisZ)
        if perp.GetLength() < 1e-8:
            perp = DB.XYZ.BasisY
    try:
        return perp.Normalize()
    except Exception:
        return DB.XYZ.BasisY

def textnote_at(doc, view, point, text, typeId):
    opts = DB.TextNoteOptions(typeId)
    opts.HorizontalAlignment = DB.HorizontalTextAlignment.Center
    return DB.TextNote.Create(doc, view.Id, point, text, opts)

def note_exists_near(doc, view, point, text, tol_ft):
    """Case-insensitive check for existing text with same content within tol_ft."""
    notes = DB.FilteredElementCollector(doc, view.Id).OfClass(DB.TextNote).ToElements()
    t_upper = text.strip().upper()
    for n in notes:
        try:
            if n.Text and n.Text.strip().upper() == t_upper:
                p = n.Coord
                if p and p.DistanceTo(point) <= tol_ft:
                    return True
        except Exception:
            continue
    return False

# -------------------------
# UI: Select TextNote Type (patched to avoid AttributeError: Name)
# -------------------------
def _get_text_type_name_safe(txt_type):
    """Return a safe display name for a TextNoteType."""
    name = None
    try:
        p = txt_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            name = p.AsString()
    except Exception:
        pass
    if not name:
        try:
            name = txt_type.Name
        except Exception:
            name = None
    if not name:
        name = u"<Unnamed Type {}>".format(txt_type.Id.IntegerValue)
    return name

def pick_text_type_id_ui(doc, preselect_name=None):
    cfg = script.get_config()
    last_type_id_int = getattr(cfg, 'last_text_type_id', None)
    last_type_name = getattr(cfg, 'last_text_type_name', None)

    types = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
    if not types:
        forms.alert("No TextNoteType found in the project.", exitscript=True)

    try:
        default_id = doc.GetDefaultElementTypeId(DB.ElementTypeGroup.TextNoteType)
        if not default_id or default_id == DB.ElementId.InvalidElementId:
            default_id = None
    except Exception:
        default_id = None

    type_rows = []
    for t in types:
        tname = _get_text_type_name_safe(t)
        tid = t.Id
        tid_int = tid.IntegerValue
        is_default = (default_id is not None and tid_int == default_id.IntegerValue)
        type_rows.append((tname, tid, tid_int, is_default))

    type_rows.sort(key=lambda row: (row[0] or u"").lower())

    seen_names = {}
    for tname, tid, tid_int, is_default in type_rows:
        seen_names[tname] = seen_names.get(tname, 0) + 1

    labels = []
    id_by_label = {}
    label_by_id = {}
    for tname, tid, tid_int, is_default in type_rows:
        base_label = tname
        if is_default:
            base_label = u'{}  (Default)'.format(base_label)
        if seen_names.get(tname, 0) > 1:
            label = u'{}  [id:{}]'.format(base_label, tid_int)
        else:
            label = base_label
        labels.append(label)
        id_by_label[label] = tid
        label_by_id[tid_int] = label

    project_default_label = None
    if default_id is not None:
        default_label_for_id = label_by_id.get(default_id.IntegerValue, None)
        if default_label_for_id:
            clean_default_name = default_label_for_id.replace("  (Default)", "")
        else:
            clean_default_name = u"(unknown)"
        project_default_label = u'Project Default: "{}"'.format(clean_default_name)
        labels = [project_default_label] + labels

    preselected_label = None
    if preselect_name:
        for lb in labels:
            if lb.startswith(u'Project Default:'):
                continue
            lb_clean = lb.replace("  (Default)", "")
            if "  [id:" in lb_clean:
                lb_clean = lb_clean.split("  [id:")[0]
            if lb_clean == preselect_name:
                preselected_label = lb
                break
    if preselected_label is None and last_type_id_int is not None:
        preselected_label = label_by_id.get(int(last_type_id_int), None)
    if preselected_label is None and project_default_label:
        preselected_label = project_default_label
    if preselected_label is None and labels:
        preselected_label = labels[0]

    selected = forms.SelectFromList.show(
        labels,
        title="Select Text Type for END Labels",
        button_name="Use Type",
        multiselect=False,
        default=preselected_label
    )

    if not selected:
        forms.alert("Cancelled by user.", exitscript=True)

    if project_default_label and selected == project_default_label:
        chosen_id = default_id
        if default_id is not None:
            chosen_label = label_by_id.get(default_id.IntegerValue)
            if chosen_label:
                chosen_name = chosen_label.replace("  (Default)", "")
                if "  [id:" in chosen_name:
                    chosen_name = chosen_name.split("  [id:")[0]
            else:
                chosen_name = "Project Default"
        else:
            chosen_name = "Project Default"
    else:
        chosen_id = id_by_label.get(selected)
        chosen_name = selected.replace("  (Default)", "")
        if "  [id:" in chosen_name:
            chosen_name = chosen_name.split("  [id:")[0]

    if not chosen_id or chosen_id == DB.ElementId.InvalidElementId:
        forms.alert("Could not resolve the selected TextNote type.", exitscript=True)

    try:
        cfg.last_text_type_id = int(chosen_id.IntegerValue)
        cfg.last_text_type_name = chosen_name
        script.save_config()
    except Exception:
        pass

    return chosen_id

# -------------------------
# Collect grids visible in the active view
# -------------------------
grids = list(DB.FilteredElementCollector(doc, view.Id).OfClass(DB.Grid).ToElements())
if not grids:
    forms.alert("No grids found in the active view.", warn_icon=True, exitscript=True)

# -------------------------
# Resolve text type (with UI)
# -------------------------
textTypeId = pick_text_type_id_ui(doc, preselect_name=TEXT_TYPE_NAME)

# -------------------------
# Create notes
# -------------------------
created = []
skipped = []
errors = []

with revit.Transaction('Place END 1 / END 2 at Grid Ends'):
    for g in grids:
        try:
            c = get_primary_curve_for_view(g, view)
            if not c:
                skipped.append((g.Id.IntegerValue, "no curve in view"))
                continue

            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)

            # Direction along the grid and a perpendicular in the view plane
            d = safe_norm(xyz_sub(p1, p0))
            perp = perp_in_view(d, view)

            # Offset both labels away from the grid line
            q0 = xyz_add(p0, xyz_mul(perp, OFFSET_FT))
            q1 = xyz_add(p1, xyz_mul(perp, OFFSET_FT))

            # "END 1" at endpoint 0
            if not note_exists_near(doc, view, q0, "END 1", DUP_TOLERANCE_FT):
                n0 = textnote_at(doc, view, q0, "END 1", textTypeId)
                created.append(n0.Id.IntegerValue)
            else:
                skipped.append(("dup", "END 1", g.Id.IntegerValue))

            # "END 2" at endpoint 1
            if not note_exists_near(doc, view, q1, "END 2", DUP_TOLERANCE_FT):
                n1 = textnote_at(doc, view, q1, "END 2", textTypeId)
                created.append(n1.Id.IntegerValue)
            else:
                skipped.append(("dup", "END 2", g.Id.IntegerValue))

        except Exception as ge:
            errors.append("Grid {}: {}".format(g.Id.IntegerValue, str(ge)))
            continue

# -------------------------
# Report
# -------------------------
msg_lines = []
msg_lines.append("Created TextNotes: {}".format(len(created)))
if created:
    msg_lines.append("  IDs: {}".format(", ".join([str(x) for x in created])))
if skipped:
    msg_lines.append("Skipped: {}".format(len(skipped)))
if errors:
    msg_lines.append("Errors: {}".format(len(errors)))
    for e in errors[:10]:
        msg_lines.append("  - {}".format(e))

forms.alert("\n".join(msg_lines), title="END 1 / END 2 at Grid Ends", warn_icon=bool(errors))
