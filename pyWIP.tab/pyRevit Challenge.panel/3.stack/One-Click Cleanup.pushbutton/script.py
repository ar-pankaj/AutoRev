# -*- coding: UTF-8 -*-
"""One-Click Cleanup Assistant — model hygiene scan and safe cleanup."""
# pylint: disable=import-error,invalid-name,broad-except,superfluous-parens
from collections import defaultdict

from pyrevit import revit, DB, UI, HOST_APP
from pyrevit import forms
from pyrevit import script
from pyrevit.compat import get_elementid_value_func


logger = script.get_logger()
output = script.get_output()
doc = revit.doc
get_elementid_value = get_elementid_value_func()

PURGE_UNUSED_PASSES = 3
WARNING_RISK_THRESHOLD = 50


# --- Brand ---

BRAND = {
    'gold': '#F1C40F',
    'goldBright': '#FFCF10',
    'goldHighlight': '#FFD633',
    'slate': '#2C3E50',
    'slateMid': '#34495E',
    'ink': '#0B1014',
}

SEVERITY_OK = 'ok'
SEVERITY_WARN = 'warn'
SEVERITY_RISK = 'risk'


def make_header_html(title, subtitle):
    """Build branded report header."""
    b = BRAND
    return (
        '<div style="background:{slate};color:{goldBright};padding:12px 16px;'
        'margin-bottom:8px;">'
        '<h2 style="margin:0;font-size:1.4em;">{title}</h2>'
        '<p style="margin:4px 0 0;color:#ccc;font-size:0.9em;">{subtitle}</p>'
        '</div>'
    ).format(title=title, subtitle=subtitle, **b)


def make_tile_html(count, label, severity):
    """Build a KPI tile for the dashboard."""
    b = BRAND
    if severity == SEVERITY_OK:
        style = 'background:{gold};color:{ink};'.format(**b)
    elif severity == SEVERITY_RISK:
        style = 'background:{slate};color:{goldBright};'.format(**b)
    else:
        style = 'background:{goldHighlight};color:{slate};'.format(**b)
    return (
        '<p style="{0}font-size:2.2em;margin:4px;padding:10px;width:130px;'
        'height:90px;display:inline-block;font-weight:bold;'
        'box-shadow:1px 1px 3px {ink};vertical-align:top;">'
        '{1}<span style="font-size:0.35em;display:block;text-transform:'
        'uppercase;">{2}</span></p>'
    ).format(style, count, label, ink=b['ink'])


def make_section_title(text):
    """Build a section heading."""
    b = BRAND
    return (
        '<p style="border-left:4px solid {goldBright};padding-left:10px;'
        'margin:14px 0 6px;color:{slate};"><strong>{text}</strong></p>'
    ).format(text=text, **b)


def print_html_block(html):
    """Send HTML to the pyRevit output window."""
    output.print_html(html)


# --- View helpers (from wipeactions) ---

READONLY_VIEWS = [
    DB.ViewType.ProjectBrowser,
    DB.ViewType.SystemBrowser,
    DB.ViewType.Undefined,
    DB.ViewType.DrawingSheet,
    DB.ViewType.Internal,
]

VIEWREF_PREFIX = {
    DB.ViewType.CeilingPlan: 'Reflected Ceiling Plan: ',
    DB.ViewType.FloorPlan: 'Floor Plan: ',
    DB.ViewType.EngineeringPlan: 'Structural Plan: ',
    DB.ViewType.DraftingView: 'Drafting View: ',
    DB.ViewType.Section: 'Section: ',
    DB.ViewType.ThreeD: '3D View: ',
}


def get_referenced_view_names(document):
    """Return names of views referenced by reference viewers."""
    view_refs = DB.FilteredElementCollector(document)\
        .OfCategory(DB.BuiltInCategory.OST_ReferenceViewer)\
        .WhereElementIsNotElementType()\
        .ToElements()
    names = set()
    for view_ref in view_refs:
        ref_param = view_ref.Parameter[
            DB.BuiltInParameter.REFERENCE_VIEWER_TARGET_VIEW
        ]
        if ref_param and ref_param.HasValue:
            names.add(ref_param.AsValueString())
    return names


def get_sheeted_view_ids(document):
    """Return sheeted view ids as integers."""
    sheeted = set()
    viewports = DB.FilteredElementCollector(document)\
        .OfClass(DB.Viewport)\
        .WhereElementIsNotElementType()\
        .ToElements()
    for viewport in viewports:
        sheeted.add(get_elementid_value(viewport.ViewId))
    return sheeted


def is_view_referenced(view, view_ref_names):
    """True when another view depends on this view."""
    refsheet = view.Parameter[DB.BuiltInParameter.VIEW_REFERENCING_SHEET]
    refviewport = view.Parameter[DB.BuiltInParameter.VIEW_REFERENCING_DETAIL]
    refprefix = VIEWREF_PREFIX.get(view.ViewType, '')
    view_label = refprefix + revit.query.get_name(view)
    if refsheet and refviewport:
        if refsheet.HasValue and refviewport.HasValue:
            if refsheet.AsString() != '' and refviewport.AsString() != '':
                return True
    return view_label in view_ref_names


def is_view_sheeted(view, sheeted_view_ids):
    """True when view or a dependent is on a sheet."""
    related_ids = [view.Id] + list(view.GetDependentViewIds())
    for v_id in related_ids:
        if get_elementid_value(v_id) in sheeted_view_ids:
            return True
    return False


def get_open_view_ids():
    """Return ids of views currently open in the UI."""
    if not revit.uidoc:
        return set()
    open_ids = set()
    for uiview in revit.uidoc.GetOpenUIViews():
        open_ids.add(get_elementid_value(uiview.ViewId))
    return open_ids


def is_removable_view(view, view_ref_names, sheeted_view_ids, open_view_ids,
                      keep_sheeted=True, keep_referenced=True):
    """Mirror wipeactions view purge safety checks."""
    if view.ViewType in READONLY_VIEWS:
        return False
    if view.IsTemplate:
        return False
    if view.ViewType == DB.ViewType.ThreeD \
            and revit.query.get_name(view) == '{3D}':
        return False
    if '<' in revit.query.get_name(view):
        return False
    if get_elementid_value(view.Id) in open_view_ids:
        return False
    if keep_referenced and is_view_referenced(view, view_ref_names):
        return False
    if keep_sheeted and is_view_sheeted(view, sheeted_view_ids):
        return False
    return True


# --- Scan ---

def scan_warnings(document):
    """Collect model warnings grouped by description type."""
    warnings = list(document.GetWarnings())
    type_counts = defaultdict(int)
    for warn in warnings:
        desc = warn.GetDescriptionText().replace('\n', ' ').strip()
        type_counts[desc] += 1
    types = []
    for desc, type_count in sorted(
            type_counts.items(), key=lambda x: (-x[1], x[0])):
        types.append({'description': desc, 'count': type_count})
    count = len(warnings)
    if count > WARNING_RISK_THRESHOLD:
        severity = SEVERITY_RISK
    elif count > 0:
        severity = SEVERITY_WARN
    else:
        severity = SEVERITY_OK
    return {
        'items': warnings,
        'types': types,
        'count': count,
        'severity': severity,
    }


def scan_unused_views(document):
    """Find views safe to delete (unsheeted and unreferenced)."""
    view_ref_names = get_referenced_view_names(document)
    sheeted_view_ids = get_sheeted_view_ids(document)
    open_view_ids = get_open_view_ids()
    views = DB.FilteredElementCollector(document)\
        .OfClass(DB.View)\
        .WhereElementIsNotElementType()\
        .ToElements()
    unsheeted = []
    unreferenced = []
    for view in views:
        if not is_removable_view(
                view, view_ref_names, sheeted_view_ids, open_view_ids,
                keep_sheeted=True, keep_referenced=True):
            continue
        unsheeted.append(view)
        if not is_view_referenced(view, view_ref_names):
            unreferenced.append(view)
    items = unsheeted
    count = len(items)
    severity = SEVERITY_WARN if count else SEVERITY_OK
    return {
        'items': items,
        'unsheeted': unsheeted,
        'unreferenced': unreferenced,
        'count': count,
        'severity': severity,
    }


def scan_cad_imports(document):
    """List linked and imported DWG instances."""
    dwgs = DB.FilteredElementCollector(document)\
        .OfClass(DB.ImportInstance)\
        .WhereElementIsNotElementType()\
        .ToElements()
    linked = []
    imported = []
    for dwg in dwgs:
        if dwg.IsLinked:
            linked.append(dwg)
        else:
            imported.append(dwg)
    count = len(imported)
    severity = SEVERITY_WARN if count else SEVERITY_OK
    return {
        'linked': linked,
        'imported': imported,
        'items': imported,
        'count': count,
        'severity': severity,
    }


def scan_unused_families(document):
    """Families with no placed instances."""
    used_family_ids = set()
    instances = DB.FilteredElementCollector(document)\
        .OfClass(DB.FamilyInstance)\
        .WhereElementIsNotElementType()\
        .ToElements()
    for inst in instances:
        fam = inst.Symbol.Family
        used_family_ids.add(get_elementid_value(fam.Id))

    unused = []
    families = revit.query.get_families(document, only_editable=False)
    for fam in families:
        if fam.IsInPlace:
            continue
        if get_elementid_value(fam.Id) not in used_family_ids:
            unused.append(fam)
    count = len(unused)
    severity = SEVERITY_WARN if count else SEVERITY_OK
    return {'items': unused, 'count': count, 'severity': severity}


def scan_worksets(document):
    """Flag worksets with no model elements."""
    empty = []
    if not document.IsWorkshared:
        return {'empty': empty, 'count': 0, 'severity': SEVERITY_OK}
    counts = defaultdict(int)
    elements = DB.FilteredElementCollector(document)\
        .WhereElementIsNotElementType()\
        .ToElements()
    for el in elements:
        ws = revit.query.get_element_workset(el)
        if ws:
            counts[ws.Name] += 1
    worksets = DB.FilteredWorksetCollector(document).ToWorksets()
    for ws in worksets:
        if ws.Kind != DB.WorksetKind.UserWorkset:
            continue
        if counts.get(ws.Name, 0) == 0:
            empty.append(ws.Name)
    count = len(empty)
    severity = SEVERITY_OK
    return {'empty': empty, 'count': count, 'severity': severity}


def scan_unused_filters(document):
    """Parameter filters not applied to any view."""
    views = DB.FilteredElementCollector(document)\
        .OfClass(DB.View)\
        .WhereElementIsNotElementType()\
        .ToElements()
    filters = DB.FilteredElementCollector(document)\
        .OfClass(DB.ParameterFilterElement)\
        .ToElements()
    all_filters = set()
    used_filters = set()
    for flt in filters:
        all_filters.add(get_elementid_value(flt.Id))
    for view in views:
        if view.AreGraphicsOverridesAllowed():
            for filter_id in view.GetFilters():
                used_filters.add(get_elementid_value(filter_id))
    unused_ids = all_filters - used_filters
    items = [document.GetElement(DB.ElementId(x)) for x in unused_ids]
    count = len(items)
    severity = SEVERITY_WARN if count else SEVERITY_OK
    return {'items': items, 'count': count, 'severity': severity}


def scan_unused_view_templates(document):
    """View templates not assigned to any view."""
    viewlist = DB.FilteredElementCollector(document)\
        .OfClass(DB.View)\
        .WhereElementIsNotElementType()\
        .ToElements()
    templates = set()
    used_templates = set()
    regular_views = []
    for view in viewlist:
        if view.IsTemplate and 'master' not in revit.query.get_name(view).lower():
            templates.add(get_elementid_value(view.Id))
        else:
            regular_views.append(view)
    for view in regular_views:
        tid = get_elementid_value(view.ViewTemplateId)
        if tid > 0:
            used_templates.add(tid)
    unused_ids = templates - used_templates
    items = [document.GetElement(DB.ElementId(x)) for x in unused_ids]
    count = len(items)
    severity = SEVERITY_WARN if count else SEVERITY_OK
    return {'items': items, 'count': count, 'severity': severity}


def scan_empty_elevation_markers(document):
    """Elevation markers with no views."""
    markers = DB.FilteredElementCollector(document)\
        .OfClass(DB.ElevationMarker)\
        .WhereElementIsNotElementType()\
        .ToElements()
    items = [m for m in markers if m.CurrentViewCount == 0]
    count = len(items)
    severity = SEVERITY_WARN if count else SEVERITY_OK
    return {'items': items, 'count': count, 'severity': severity}


def scan_empty_tags(document):
    """Independent tags with empty text across the model."""
    tags = DB.FilteredElementCollector(document)\
        .OfClass(DB.IndependentTag)\
        .WhereElementIsNotElementType()\
        .ToElements()
    items = []
    for tag in tags:
        tag_text = tag.TagText
        if tag_text == '' or tag_text is None:
            items.append(tag)
    count = len(items)
    severity = SEVERITY_WARN if count else SEVERITY_OK
    return {'items': items, 'count': count, 'severity': severity}


def build_summary_counts(scan_data):
    """Flatten scan results into count dict for before/after comparison."""
    return {
        'warnings': scan_data['warnings']['count'],
        'unused_views': scan_data['unused_views']['count'],
        'cad_imports': scan_data['cad_imports']['count'],
        'unused_families': scan_data['unused_families']['count'],
        'empty_worksets': scan_data['worksets']['count'],
        'unused_filters': scan_data['unused_filters']['count'],
        'unused_templates': scan_data['unused_templates']['count'],
        'empty_markers': scan_data['empty_markers']['count'],
        'empty_tags': scan_data['empty_tags']['count'],
    }


def count_severity_totals(scan_data):
    """Count OK / warn / risk categories for KPI tiles."""
    categories = [
        scan_data['warnings'],
        scan_data['unused_views'],
        scan_data['cad_imports'],
        scan_data['unused_families'],
        scan_data['worksets'],
        scan_data['unused_filters'],
        scan_data['unused_templates'],
        scan_data['empty_markers'],
        scan_data['empty_tags'],
    ]
    totals = {SEVERITY_OK: 0, SEVERITY_WARN: 0, SEVERITY_RISK: 0}
    for cat in categories:
        sev = cat.get('severity', SEVERITY_OK)
        if cat.get('count', 0) == 0 and sev == SEVERITY_WARN:
            totals[SEVERITY_OK] += 1
        else:
            totals[sev] = totals.get(sev, 0) + 1
    return totals


def scan_model(document):
    """Run all hygiene scans and return structured results."""
    output.set_title('One-Click Cleanup - Scanning...')
    data = {
        'warnings': scan_warnings(document),
        'unused_views': scan_unused_views(document),
        'cad_imports': scan_cad_imports(document),
        'unused_families': scan_unused_families(document),
        'worksets': scan_worksets(document),
        'unused_filters': scan_unused_filters(document),
        'unused_templates': scan_unused_view_templates(document),
        'empty_markers': scan_empty_elevation_markers(document),
        'empty_tags': scan_empty_tags(document),
    }
    data['summary_counts'] = build_summary_counts(data)
    data['severity_totals'] = count_severity_totals(data)
    return data


def scan_counts(document):
    """Lightweight re-scan for before/after summary."""
    return build_summary_counts(scan_model(document))


# --- Report ---

def render_dashboard(scan_data, doc_title):
    """Print branded hygiene report to the output window."""
    print_html_block(make_header_html('One-Click Cleanup', doc_title))

    totals = scan_data['severity_totals']
    print_html_block(make_tile_html(
        totals.get(SEVERITY_OK, 0), 'OK', SEVERITY_OK))
    print_html_block(make_tile_html(
        totals.get(SEVERITY_WARN, 0), 'Optional', SEVERITY_WARN))
    print_html_block(make_tile_html(
        totals.get(SEVERITY_RISK, 0), 'Risky', SEVERITY_RISK))
    output.print_md('')

    _print_warnings_section(scan_data['warnings'])
    _print_views_section(scan_data['unused_views'])
    _print_cad_section(scan_data['cad_imports'])
    _print_families_usage_section(scan_data['unused_families'])
    _print_worksets_section(scan_data['worksets'])
    _print_filters_section(scan_data['unused_filters'])
    _print_templates_section(scan_data['unused_templates'])
    _print_markers_section(scan_data['empty_markers'])
    _print_tags_section(scan_data['empty_tags'])

    b = BRAND
    print_html_block(
        '<p style="color:{slateMid};font-size:0.85em;margin-top:16px;">'
        'Powered by pyRevit</p>'.format(**b))


def _print_warnings_section(data):
    print_html_block(make_section_title('Warnings ({0})'.format(data['count'])))
    if not data['count']:
        output.print_md('No warnings.')
        return
    rows = []
    for warn_type in data.get('types', []):
        desc = warn_type['description']
        if len(desc) > 200:
            desc = desc[:197] + '...'
        rows.append([desc, warn_type['count']])
    output.print_table(
        table_data=rows, columns=['Warning Type', 'Count'])


def _print_views_section(data):
    print_html_block(make_section_title('Unused Views ({0})'.format(data['count'])))
    if not data['count']:
        output.print_md('No unused views detected.')
        return
    rows = []
    for view in data['items'][:40]:
        rows.append([
            revit.query.get_name(view),
            str(view.ViewType),
            output.linkify(view.Id),
        ])
    output.print_table(table_data=rows, columns=['Name', 'Type', 'Id'])


def _print_cad_section(data):
    print_html_block(make_section_title(
        'CAD Imports ({0}) / Linked ({1})'.format(
            len(data['imported']), len(data['linked']))))
    rows = []
    for dwg in data['imported'][:30]:
        name = dwg.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()
        rows.append(['IMPORT', name or '?', output.linkify(dwg.Id)])
    for dwg in data['linked'][:10]:
        name = dwg.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()
        rows.append(['LINK', name or '?', output.linkify(dwg.Id)])
    if rows:
        output.print_table(table_data=rows, columns=['Mode', 'Name', 'Id'])
    else:
        output.print_md('No CAD imports found.')


def _print_families_usage_section(data):
    print_html_block(make_section_title(
        'Unused Families ({0})'.format(data['count'])))
    if not data['count']:
        output.print_md('All families appear to have instances.')
        return
    rows = [[revit.query.get_name(f)] for f in data['items'][:30]]
    output.print_table(table_data=rows, columns=['Family'])


def _print_worksets_section(data):
    print_html_block(make_section_title(
        'Empty Worksets ({0})'.format(data['count'])))
    if data['count']:
        output.print_md(', '.join(data['empty']))
    else:
        output.print_md('No empty user worksets.')


def _print_filters_section(data):
    print_html_block(make_section_title(
        'Unused Filters ({0})'.format(data['count'])))
    if data['count']:
        rows = [[f.Name] for f in data['items'][:30]]
        output.print_table(table_data=rows, columns=['Filter'])


def _print_templates_section(data):
    print_html_block(make_section_title(
        'Unused View Templates ({0})'.format(data['count'])))
    if data['count']:
        rows = [[revit.query.get_name(v)] for v in data['items'][:30]]
        output.print_table(table_data=rows, columns=['Template'])


def _print_markers_section(data):
    print_html_block(make_section_title(
        'Empty Elevation Markers ({0})'.format(data['count'])))


def _print_tags_section(data):
    print_html_block(make_section_title(
        'Empty Tags ({0})'.format(data['count'])))


def render_summary(before, after):
    """Print before/after count comparison."""
    b = BRAND
    html = (
        '<div style="background:{slate};color:{goldBright};padding:12px;'
        'margin-top:12px;"><strong>Cleanup Summary</strong><br/>'
    ).format(**b)
    labels = [
        ('unused_views', 'Unused views'),
        ('cad_imports', 'CAD imports'),
        ('unused_filters', 'Unused filters'),
        ('unused_templates', 'Unused templates'),
        ('empty_markers', 'Empty elevation markers'),
        ('empty_tags', 'Empty tags'),
        ('warnings', 'Warnings'),
    ]
    for key, label in labels:
        html += '{0}: {1} &rarr; {2}<br/>'.format(
            label, before.get(key, 0), after.get(key, 0))
    html += '</div>'
    print_html_block(html)


# --- UI ---

class CleanupAction(forms.TemplateListItem):
    """Selectable cleanup action for multiselect UI."""

    def __init__(self, key, label, needs_pick=False, destructive=False,
                 checked=False):
        super(CleanupAction, self).__init__(key, checked=checked)
        self.key = key
        self.label = label
        self.needs_pick = needs_pick
        self.destructive = destructive

    @property
    def name(self):
        return self.label

    def unwrap(self):
        """Return the action object, not the key string stored in self.item."""
        return self


class NamedElementItem(forms.TemplateListItem):
    """List item wrapping a Revit element."""

    @property
    def name(self):
        if hasattr(self.item, 'Name') and self.item.Name:
            return self.item.Name
        return revit.query.get_name(self.item)


ACTION_DEFS = [
    ('report_only', 'Report only (no changes)', False, False, True),
    ('delete_views', 'Delete selected unused views', True, True, False),
    ('remove_cad', 'Remove selected CAD imports', True, True, False),
    ('purge_filters', 'Purge unused filters', True, True, False),
    ('purge_templates', 'Purge unused view templates', True, True, False),
    ('remove_markers', 'Remove empty elevation markers', False, True, False),
    ('remove_tags', 'Remove empty tags', False, True, False),
    ('purge_unused', 'Launch Revit Purge Unused (3 passes)', False, False, False),
]


def build_action_list():
    """Build cleanup action options for SelectFromList."""
    return [
        CleanupAction(key, label, needs_pick=needs_pick,
                      destructive=destructive, checked=checked)
        for key, label, needs_pick, destructive, checked in ACTION_DEFS
    ]


def pick_items(title, items, button_name):
    """Let user pick elements from scan results."""
    if not items:
        return []
    wrapped = [NamedElementItem(x) for x in items]
    return forms.SelectFromList.show(
        wrapped,
        title=title,
        button_name=button_name,
        multiselect=True,
        width=500,
    ) or []


def confirm_destructive(actions, scan_data, picks):
    """Confirm destructive cleanup with item counts."""
    lines = []
    for action in actions:
        if not action.destructive:
            continue
        key = action.key
        if key == 'delete_views':
            lines.append('Views: {0}'.format(len(picks.get('views', []))))
        elif key == 'remove_cad':
            lines.append('CAD: {0}'.format(len(picks.get('cad', []))))
        elif key == 'purge_filters':
            lines.append('Filters: {0}'.format(len(picks.get('filters', []))))
        elif key == 'purge_templates':
            lines.append('Templates: {0}'.format(
                len(picks.get('templates', []))))
        elif key == 'remove_markers':
            lines.append('Markers: {0}'.format(
                scan_data['empty_markers']['count']))
        elif key == 'remove_tags':
            lines.append('Tags: {0}'.format(scan_data['empty_tags']['count']))
    if not lines:
        return True
    msg = 'Proceed with cleanup?\n\n' + '\n'.join(lines)
    return forms.alert(msg, yes=True, no=True)


# --- Execute ---

def delete_elements_safe(elements, transaction_name):
    """Delete elements inside a transaction, logging failures."""
    if not elements:
        return 0
    removed = 0
    with revit.Transaction(transaction_name):
        for el in elements:
            try:
                doc.Delete(el.Id)
                removed += 1
            except Exception as ex:
                logger.error('Delete failed: %s | %s', el.Id, ex)
    return removed


def call_purge_unused():
    """Queue Revit Purge Unused three times (multiple passes needed)."""
    cid = UI.RevitCommandId.LookupPostableCommandId(
        UI.PostableCommand.PurgeUnused)
    for _ in range(PURGE_UNUSED_PASSES):
        HOST_APP.uiapp.PostCommand(cid)


def run_cleanup(selected_actions, scan_data, picks):
    """Execute selected cleanup actions inside a transaction group."""
    action_keys = set(a.key for a in selected_actions)
    if action_keys == set(['report_only']) or not action_keys:
        return

    with revit.TransactionGroup('One-Click Cleanup', doc):
        if 'delete_views' in action_keys:
            views = picks.get('views', [])
            delete_elements_safe(views, 'Delete Unused Views')
        if 'remove_cad' in action_keys:
            cad = picks.get('cad', [])
            delete_elements_safe(cad, 'Remove CAD Imports')
        if 'purge_filters' in action_keys:
            flts = picks.get('filters', [])
            delete_elements_safe(flts, 'Purge Unused Filters')
        if 'purge_templates' in action_keys:
            tmpls = picks.get('templates', [])
            delete_elements_safe(tmpls, 'Purge Unused View Templates')
        if 'remove_markers' in action_keys:
            delete_elements_safe(
                scan_data['empty_markers']['items'],
                'Remove Empty Elevation Markers')
        if 'remove_tags' in action_keys:
            delete_elements_safe(
                scan_data['empty_tags']['items'],
                'Remove Empty Tags')

    if 'purge_unused' in action_keys:
        call_purge_unused()


def collect_item_picks(selected_actions, scan_data):
    """Prompt for per-category item selection when needed."""
    picks = {}
    keys = set(a.key for a in selected_actions)
    if 'delete_views' in keys and scan_data['unused_views']['count']:
        picks['views'] = pick_items(
            'Select Views to Delete',
            scan_data['unused_views']['items'],
            'Delete Views')
    if 'remove_cad' in keys and scan_data['cad_imports']['count']:
        picks['cad'] = pick_items(
            'Select CAD Imports to Remove',
            scan_data['cad_imports']['imported'],
            'Remove CAD')
    if 'purge_filters' in keys and scan_data['unused_filters']['count']:
        picks['filters'] = pick_items(
            'Select Filters to Purge',
            scan_data['unused_filters']['items'],
            'Purge Filters')
    if 'purge_templates' in keys and scan_data['unused_templates']['count']:
        picks['templates'] = pick_items(
            'Select View Templates to Purge',
            scan_data['unused_templates']['items'],
            'Purge Templates')
    return picks


# --- Main ---

def main():
    """Scan, report, and optionally execute safe cleanup."""
    try:
        doc_title = doc.Title or 'Untitled'
        if doc.PathName:
            doc_title = '{0} - {1}'.format(doc_title, doc.PathName)

        scan_data = scan_model(doc)
        before_counts = scan_data['summary_counts']

        selected = forms.SelectFromList.show(
            build_action_list(),
            title='Cleanup Actions',
            button_name='Run Cleanup',
            multiselect=True,
            width=520,
        )
        if not selected:
            script.exit()

        render_dashboard(scan_data, doc_title)

        action_keys = set(a.key for a in selected)
        if action_keys == set(['report_only']):
            logger.info('Report only - no changes made.')
            script.exit()

        destructive = [a for a in selected if a.destructive]
        picks = {}
        if destructive:
            picks = collect_item_picks(selected, scan_data)

        if not confirm_destructive(selected, scan_data, picks):
            script.exit()

        run_cleanup(selected, scan_data, picks)
        after_counts = scan_counts(doc)
        render_summary(before_counts, after_counts)
        logger.info('One-Click Cleanup finished.')
    except Exception as ex:
        logger.error('One-Click Cleanup failed: %s', ex)
        output.print_md('## Error')
        output.print_md('`{0}`'.format(ex))
        raise


main()
