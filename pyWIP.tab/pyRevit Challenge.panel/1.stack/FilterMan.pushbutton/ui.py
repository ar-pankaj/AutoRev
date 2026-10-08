# -*- coding: utf-8 -*-
"""FilterMan 2.0 - BrowserWindow and UI logic."""
from __future__ import unicode_literals
import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, Visibility, GridLength, GridUnitType, Thickness,
    VerticalAlignment, HorizontalAlignment, FontWeights,
    TextWrapping, TextAlignment
)
from System.Windows.Controls import (
    CheckBox, StackPanel, TextBlock, Border,
    ComboBox, ComboBoxItem, TextBox, Button, DockPanel,
    DataGridTextColumn, Orientation, ScrollViewer
)
from System.Windows.Media import SolidColorBrush, ColorConverter, Brushes, FontFamily
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind
from System.Windows.Threading import DispatcherTimer, DispatcherPriority
from System import TimeSpan

from pyrevit.framework import wpf
from Autodesk.Revit.DB import FilteredElementCollector

THIS_DIR  = os.path.dirname(os.path.abspath(__file__))
XAML_PATH = os.path.join(THIS_DIR, u'browser.xaml')
ICON_PATH = os.path.join(THIS_DIR, u'resources', u'icon.png')

_SEGOE = FontFamily(u'Segoe UI')


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  DATA MODELS                                                         ║
# ╚══════════════════════════════════════════════════════════════════════╝

def eid_key(element_id):
    """Canonical ElementId -> int. Revit 2024+ (.Value) and older (.IntegerValue)."""
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


class TreeNode(object):
    __slots__ = [
        u'element_id', u'level', u'label', u'is_checked',
        u'is_indeterminate', u'children', u'parent',
        u'category_name', u'family_name', u'type_name', u'wpf_item',
        u'label_lower', u'_header_border', u'_checkbox',
        u'_expand_icon', u'_count_label', u'extra_indent',
    ]

    def __init__(self, level, label, element_id=None):
        self.element_id       = element_id
        self.level            = level        # 'category'|'family'|'type'|'instance'
        self.label            = label
        self.label_lower      = label.lower()
        self.is_checked       = False
        self.is_indeterminate = False
        self.children         = []
        self.parent           = None
        self.category_name    = u''
        self.family_name      = u''
        self.type_name        = u''
        self.wpf_item         = None    # flat row Border in tree_items StackPanel
        self._header_border   = None    # alias for wpf_item (bg/border colour updates)
        self._checkbox        = None    # direct CheckBox reference
        self._expand_icon     = None    # expand/collapse TextBlock (▼/►)
        self._count_label     = None    # count TextBlock for non-instance nodes
        self.extra_indent     = 0       # additional left padding (px) for nested groups


class FilterCondition(object):
    __slots__ = [u'field', u'param_name', u'operator', u'value']

    FIELDS = [
        (u'family_name', u'Family Name'),
        (u'type_name',   u'Type Name'),
        (u'category',    u'Category'),
        (u'level',       u'Level'),
        (u'parameter',   u'Parameter...'),
    ]

    TEXT_OPS = [
        (u'contains',    u'contains'),
        (u'equals',      u'equals'),
        (u'not_equals',  u'not equals'),
        (u'starts_with', u'starts with'),
        (u'ends_with',   u'ends with'),
        (u'has_value',   u'has value'),
        (u'no_value',    u'no value'),
    ]

    NUM_OPS = [
        (u'equals', u'equals'),
        (u'gt',     u'>'),
        (u'gte',    u'>='),
        (u'lt',     u'<'),
        (u'lte',    u'<='),
    ]

    def __init__(self, field=u'family_name', operator=u'contains',
                 value=u'', param_name=None):
        self.field      = field
        self.param_name = param_name
        self.operator   = operator
        self.value      = value


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  BROWSER WINDOW                                                      ║
# ╚══════════════════════════════════════════════════════════════════════╝

class BrowserWindow(Window):

    def __init__(self, uidoc, doc, splash=None, splash_start=None):
        self.uidoc = uidoc
        self.doc   = doc

        wpf.LoadComponent(self, XAML_PATH)

        # While splash is visible let it stay on top - suppress our Topmost
        # until the tree is ready (restored in _deferred_first_build).
        if splash is not None:
            self.Topmost = False

        # ── Icon ──────────────────────────────────────────────────────
        if os.path.isfile(ICON_PATH):
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource        = Uri(ICON_PATH, UriKind.Absolute)
            bmp.DecodePixelWidth = 22
            bmp.EndInit()
            self.img_logo.Source = bmp

        # ── State ─────────────────────────────────────────────────────
        self._search_query   = u''
        self._search_scope   = {
            u'family':      True,
            u'type':        True,
            u'param_name':  False,
            u'param_value': False,
        }
        self._root_nodes        = []
        self._all_nodes         = []
        self._eid_map           = {}
        self._checked_ids       = {}
        self._key_to_nodes      = {}  # {int_key: [TreeNode, ...]} for cross-tree sync
        self._show_checked_only = False
        self._last_revit_ids    = frozenset()
        self._updating          = False
        self._dark_mode         = False
        self._colored_ids       = {}
        self._adv_open          = False
        self._adv_scope         = u'checked'
        self._adv_conditions    = []
        self._import_failures    = {}        # {id_str: [param_key, ...]} from last import
        self._last_import_data   = []        # raw import rows (used by Overwrite)
        self._import_preview_mode = False   # True while preview table is shown
        self._adv_columns       = [         # list of (header, row_key) tuples
            (u'Family Name', u'family_name'),
            (u'Type Name',   u'type_name'),
        ]
        self._adv_param_cache   = {}   # {param_name: 'I'|'T'}
        self._adv_rows          = None      # System.Data.DataTable
        self._model_levels      = []
        self._param_cache       = {}
        self._poll_timer        = None
        self._debounce_timer    = None
        self._color_popup       = None
        self._adv_anim_timer    = None
        self._splash_ref        = splash
        self._splash_start      = splash_start
        self._defer_splash_timer = None
        # Set of id(node) for non-leaf nodes whose children are hidden (collapsed)
        self._collapsed         = set()

        # ── Load persisted settings (dark mode) ───────────────────────
        try:
            from settings import load_settings
            _s = load_settings()
            if _s.get(u'dark_mode', False):
                self._dark_mode = True
        except Exception:
            pass

        self._wire_events()
        self._update_search_placeholder()

        # ── Apply dark palette if restored from settings ───────────────
        if self._dark_mode:
            for _key, _hex in _DARK_PALETTE.iteritems():
                self.Resources[_key] = SolidColorBrush(
                    ColorConverter.ConvertFromString(_hex)
                )

        # Create ExternalEvent handlers while still in Revit API context.
        # These are needed for Transaction-requiring view actions (Isolate etc.).
        try:
            from view_actions import init_events
            init_events()
        except Exception:
            pass

        self._start_polling()
        self._start_shimmer()

        # Defer tree build so window shows before the blocking Revit API scan.
        # Splash is closed AFTER tree build completes (not on a racing timer).
        self._defer_build_timer = DispatcherTimer()
        self._defer_build_timer.Interval = TimeSpan.FromMilliseconds(50)

        def _deferred_first_build(s, a):
            self._defer_build_timer.Stop()
            self._build_tree()
            self._stop_shimmer()
            # Restore Topmost now that the tree is ready
            try:
                self.Topmost = bool(self.cb_always_on_top.IsChecked)
            except Exception:
                self.Topmost = True
            # Close splash now that the tree is ready
            _ref = self._splash_ref
            self._splash_ref = None
            if _ref is not None:
                SPLASH_MIN_MS = 800.0
                import time as _t
                remaining = 0.0
                if self._splash_start is not None:
                    elapsed = (_t.time() - self._splash_start) * 1000.0
                    remaining = max(0.0, SPLASH_MIN_MS - elapsed)
                if remaining <= 0:
                    try:
                        _ref.Close()
                    except Exception:
                        pass
                else:
                    _timer = DispatcherTimer()
                    _timer.Interval = TimeSpan.FromMilliseconds(remaining)
                    def _close(ss, aa, r=_ref, t=_timer):
                        t.Stop()
                        try:
                            r.Close()
                        except Exception:
                            pass
                    _timer.Tick += _close
                    _timer.Start()
                    self._defer_splash_timer = _timer

        self._defer_build_timer.Tick += _deferred_first_build
        self._defer_build_timer.Start()

    # ──────────────────────────────────────────────────────────────────
    # EVENT WIRING
    # ──────────────────────────────────────────────────────────────────

    def _wire_events(self):
        # Header
        self.cb_always_on_top.Checked   += self._always_on_top_changed
        self.cb_always_on_top.Unchecked += self._always_on_top_changed
        self.btn_theme.Click            += self._toggle_theme
        self.btn_refresh.Click          += self._refresh_click
        self.cb_active_view.Checked     += self._active_view_changed
        self.cb_active_view.Unchecked   += self._active_view_changed

        # Search
        self.txt_search.TextChanged   += self._search_changed
        self.btn_search_clear.Click   += self._clear_search

        # Search scope
        for _name in (u'cb_scope_family', u'cb_scope_type',
                      u'cb_scope_paramname', u'cb_scope_paramval'):
            _cb = getattr(self, _name)
            _cb.Checked   += self._scope_changed
            _cb.Unchecked += self._scope_changed

        # Tree toolbar
        self.btn_expand_all.Click       += self._expand_all
        self.btn_collapse_all.Click     += self._collapse_all
        self.btn_select_all.Click       += self._select_all
        self.btn_deselect_all.Click     += self._deselect_all
        self.btn_checked_only.Checked   += self._checked_only_on
        self.btn_checked_only.Unchecked += self._checked_only_off

        # Action row
        self.btn_isolate.Click        += self._isolate_click
        self.btn_temp_hide.Click      += self._temp_hide_click
        self.btn_color.Click          += self._color_click
        self.btn_selbox.Click         += self._selbox_click
        self.btn_reset_all.Click      += self._reset_click
        self.btn_sel_inview.Click     += self._sel_inview_click
        self.btn_sel_inproj.Click     += self._sel_inproj_click

        # Status bar
        self.btn_check_from_revit.Click += self._check_from_revit
        self.btn_select_in_revit.Click  += self._select_in_revit

        # Bottom bar
        self.btn_save_json.Click    += self._save_json_click
        self.btn_load_json.Click    += self._load_json_click
        self.btn_toggle_adv.Click   += self._toggle_advanced

        # Advanced panel
        self.btn_scope_checked.Click += self._scope_checked_click
        self.btn_scope_all.Click     += self._scope_all_click
        self.btn_qb_apply.Click         += self._qb_apply
        self.btn_qb_clear.Click         += self._qb_clear
        self.btn_qb_add_condition.Click += self._add_condition_row
        self.btn_add_column.Click       += self._add_column_click
        self.btn_import_excel.Click     += self._import_excel_click
        self.btn_overwrite_excel.Click  += self._overwrite_click
        self.btn_export_excel.Click     += self._export_excel_click

        # Window lifecycle
        self.Closing           += self._on_closing
        self.SourceInitialized += self._on_source_initialized

    # ──────────────────────────────────────────────────────────────────
    # TREE BUILD
    # ──────────────────────────────────────────────────────────────────

    def _build_tree(self):
        """Collect elements and build Category > Family > Type > Instance hierarchy."""
        try:
            self._build_tree_inner()
        except Exception as ex:
            import traceback
            tb_str = traceback.format_exc()
            # Show error in tree panel
            try:
                self.tree_items.Children.Clear()
                err = TextBlock()
                err.Text        = u'Tree build error:\n{}\n\n{}'.format(
                    unicode(ex), tb_str)
                err.Foreground  = SolidColorBrush(
                    ColorConverter.ConvertFromString(u'#C0392B'))
                err.TextWrapping = TextWrapping.Wrap
                err.Margin       = Thickness(10)
                err.FontSize     = 11.0
                self.tree_items.Children.Add(err)
            except Exception:
                pass
            # Always also show via TaskDialog so it is visible regardless
            try:
                from Autodesk.Revit.UI import TaskDialog
                td = TaskDialog(u'FilterMan - Tree Error')
                td.MainContent = u'{}\n\n{}'.format(unicode(ex), tb_str[:1200])
                td.Show()
            except Exception:
                pass

    def _build_tree_inner(self):
        if self.cb_active_view.IsChecked:
            try:
                col = FilteredElementCollector(
                    self.doc, self.uidoc.ActiveView.Id
                )
            except Exception:
                col = FilteredElementCollector(self.doc)
        else:
            col = FilteredElementCollector(self.doc)

        elements  = col.WhereElementIsNotElementType().ToElements()
        hierarchy = {}   # {cat: {fam: {type: [eid_key]}}}
        eid_map   = {}
        skip_nocat = 0
        skip_nosym = 0
        skip_ex    = 0

        from Autodesk.Revit.DB import (Group as _Group, BuiltInCategory as _BIC)
        # ID-based exclusion (locale-independent) — long() avoids IronPython int/long issues
        _excl_cat_ids = set()
        for _bic_val in [_BIC.OST_Cameras, _BIC.OST_Viewports,
                         _BIC.OST_ProjectBasePoint, _BIC.OST_SharedBasePoint]:
            try:
                _excl_cat_ids.add(long(_bic_val))
            except Exception:
                pass
        for _bic_name in [u'OST_SectionBox', u'OST_VolumeOfInterest']:
            try:
                _excl_cat_ids.add(long(getattr(_BIC, _bic_name)))
            except Exception:
                pass
        # Name-based fallback (English Revit) — catches cases where ID conversion fails
        _EXCL_CAT_NAMES = frozenset([
            u'Cameras', u'Viewports', u'Section Boxes', u'Scope Boxes',
            u'Project Base Point', u'Survey Point',
        ])
        for el in elements:
            try:
                if isinstance(el, _Group):
                    continue  # Model/Detail Groups shown in group subtree
                cat = el.Category
                if cat is None:
                    skip_nocat += 1
                    continue
                # Name check FIRST — always reliable, evaluated independently
                try:
                    if cat.Name in _EXCL_CAT_NAMES:
                        continue
                except Exception:
                    pass
                # ID check SECOND — independent try/except so a failure here
                # does NOT block the name check above
                try:
                    try:
                        _cid = cat.Id.IntegerValue
                    except AttributeError:
                        _cid = int(cat.Id.Value)
                    if long(_cid) in _excl_cat_ids:
                        continue
                except Exception:
                    pass
                cat_name  = cat.Name
                sym       = self.doc.GetElement(el.GetTypeId())
                if sym is None:
                    skip_nosym += 1
                    continue
                fam       = getattr(sym, u'Family', None)
                fam_name  = fam.Name if fam else cat_name
                type_name = _safe_sym_name(sym)
                k         = eid_key(el.Id)
                eid_map[k] = el.Id
                hierarchy \
                    .setdefault(cat_name, {}) \
                    .setdefault(fam_name, {}) \
                    .setdefault(type_name, []) \
                    .append(k)
            except Exception:
                skip_ex += 1
                continue

        skipped = skip_nocat + skip_nosym + skip_ex

        self._eid_map      = eid_map
        self._key_to_nodes = {}
        self._root_nodes   = []
        self._all_nodes    = []

        for cat_name in sorted(hierarchy):
            cat_node = TreeNode(u'category', cat_name)
            cat_node.category_name = cat_name

            for fam_name in sorted(hierarchy[cat_name]):
                fam_node = TreeNode(u'family', fam_name)
                fam_node.category_name = cat_name
                fam_node.family_name   = fam_name
                fam_node.parent        = cat_node

                for type_name in sorted(hierarchy[cat_name][fam_name]):
                    type_node = TreeNode(u'type', type_name)
                    type_node.category_name = cat_name
                    type_node.family_name   = fam_name
                    type_node.type_name     = type_name
                    type_node.parent        = fam_node

                    for k in hierarchy[cat_name][fam_name][type_name]:
                        inst = TreeNode(
                            u'instance',
                            u'{} | id:{}'.format(type_name, k),
                            element_id=eid_map[k]
                        )
                        inst.category_name = cat_name
                        inst.family_name   = fam_name
                        inst.type_name     = type_name
                        inst.parent        = type_node
                        self._key_to_nodes.setdefault(k, []).append(inst)
                        type_node.children.append(inst)
                        self._all_nodes.append(inst)

                    fam_node.children.append(type_node)
                    self._all_nodes.append(type_node)

                cat_node.children.append(fam_node)
                self._all_nodes.append(fam_node)

            self._root_nodes.append(cat_node)
            self._all_nodes.append(cat_node)

        # Remove stale checked_ids
        self._checked_ids = {
            k: v for k, v in self._checked_ids.iteritems() if k in eid_map
        }
        self._param_cache  = {}

        # Append group hierarchy BEFORE building _collapsed so group nodes are included.
        self._add_group_nodes()
        # Re-sort root nodes so "Model Groups" appears in its correct alpha position.
        self._root_nodes.sort(key=lambda n: n.label.lower())

        # Default: families+types collapsed → cat+fam rows visible, type+inst hidden.
        # Expanding a family reveals type rows; expanding a type reveals instances.
        self._collapsed = set(
            id(n) for n in self._all_nodes
            if n.level in (u'category', u'family', u'type') and n.children
        )
        self._model_levels = self._collect_level_names()
        self._render_tree(skipped)
        self.Title = u'FilterMan'
        self._update_status()

    def _collect_level_names(self):
        from Autodesk.Revit.DB import Level
        try:
            levels = FilteredElementCollector(self.doc) \
                .OfClass(Level).ToElements()
            return sorted([lv.Name for lv in levels if lv.Name])
        except Exception:
            return []

    # ──────────────────────────────────────────────────────────────────
    # GROUP TREE
    # ──────────────────────────────────────────────────────────────────

    def _add_group_nodes(self):
        """Collect Model Group instances and append a hierarchical subtree."""
        from Autodesk.Revit.DB import Group as _Group
        try:
            if self.cb_active_view.IsChecked:
                try:
                    col = FilteredElementCollector(
                        self.doc, self.uidoc.ActiveView.Id
                    )
                except Exception:
                    col = FilteredElementCollector(self.doc)
            else:
                col = FilteredElementCollector(self.doc)
            all_groups = (col.OfClass(_Group)
                           .WhereElementIsNotElementType()
                           .ToElements())
        except Exception:
            return

        model_groups = []
        for g in all_groups:
            try:
                cat = g.Category
                if cat is not None and u'Model' in cat.Name:
                    model_groups.append(g)
            except Exception:
                continue

        if not model_groups:
            return

        by_type = {}
        for g in model_groups:
            try:
                gt = g.GroupType
                type_name = _safe_sym_name(gt) if gt is not None else u'(unnamed type)'
            except Exception:
                type_name = u'(unnamed type)'
            by_type.setdefault(type_name, []).append(g)

        grp_cat = TreeNode(u'category', u'Model Groups')
        grp_cat.category_name = u'Model Groups'

        for type_name in sorted(by_type.keys()):
            grp_type = TreeNode(u'family', type_name)
            grp_type.category_name = u'Model Groups'
            grp_type.family_name   = type_name
            grp_type.parent        = grp_cat

            for g in sorted(by_type[type_name],
                            key=lambda x: eid_key(x.Id)):
                try:
                    members = list(g.GetMemberIds())
                except Exception:
                    members = []

                k = eid_key(g.Id)
                grp_inst = TreeNode(
                    u'type',
                    u'id:{} ({} members)'.format(k, len(members))
                )
                grp_inst.category_name = u'Model Groups'
                grp_inst.family_name   = type_name
                grp_inst.parent        = grp_type

                self._add_group_children(g, grp_inst, depth=0)

                grp_type.children.append(grp_inst)
                self._all_nodes.append(grp_inst)

            grp_cat.children.append(grp_type)
            self._all_nodes.append(grp_type)

        self._root_nodes.append(grp_cat)
        self._all_nodes.append(grp_cat)

    def _add_group_children(self, group_el, parent_node, depth):
        """Recursively add member elements of group_el under parent_node.

        depth=0 means direct children of a top-level group instance.
        extra_indent = depth * 28 keeps nested groups properly indented
        relative to the parent group instance (28px per nesting level).
        """
        from Autodesk.Revit.DB import Group as _Group
        extra = depth * 28

        try:
            member_ids = group_el.GetMemberIds()
        except Exception:
            return

        for mid in member_ids:
            try:
                mem = self.doc.GetElement(mid)
                if mem is None:
                    continue
                k = eid_key(mem.Id)

                if isinstance(mem, _Group):
                    # Nested group - show Group Type → Group Instance → children.
                    try:
                        sub_gt = mem.GroupType
                        sub_type_name = (
                            _safe_sym_name(sub_gt) if sub_gt is not None else u'(unnamed)'
                        )
                    except Exception:
                        sub_type_name = u'(unnamed)'
                    try:
                        sub_members = list(mem.GetMemberIds())
                    except Exception:
                        sub_members = []

                    sub_type = TreeNode(u'family', sub_type_name)
                    sub_type.category_name = u'Model Groups'
                    sub_type.family_name   = sub_type_name
                    sub_type.parent        = parent_node
                    sub_type.extra_indent  = (depth + 1) * 28

                    sub_inst = TreeNode(
                        u'type',
                        u'id:{} ({} members)'.format(k, len(sub_members))
                    )
                    sub_inst.category_name = u'Model Groups'
                    sub_inst.family_name   = sub_type_name
                    sub_inst.parent        = sub_type
                    sub_inst.extra_indent  = (depth + 1) * 28

                    self._add_group_children(mem, sub_inst, depth + 1)

                    sub_type.children.append(sub_inst)
                    self._all_nodes.append(sub_inst)

                    parent_node.children.append(sub_type)
                    self._all_nodes.append(sub_type)

                else:
                    # Regular element - flat leaf with category label.
                    try:
                        cat = mem.Category
                        cat_name = cat.Name if cat is not None else u'?'
                    except Exception:
                        cat_name = u'?'

                    child = TreeNode(
                        u'instance',
                        u'{} | id:{}'.format(cat_name, k),
                        element_id=mem.Id
                    )
                    child.category_name = cat_name
                    child.parent        = parent_node
                    child.extra_indent  = extra

                    # Inherit family/type names from the matching main-tree node
                    # (used by _node_matches for search).
                    for existing in self._key_to_nodes.get(k, []):
                        child.family_name = existing.family_name
                        child.type_name   = existing.type_name
                        break

                    # Register for cross-tree checkbox sync.
                    self._key_to_nodes.setdefault(k, []).append(child)

                    # Restore persisted checked state.
                    if k in self._checked_ids:
                        child.is_checked = True

                    parent_node.children.append(child)
                    self._all_nodes.append(child)

            except Exception:
                continue

    # ──────────────────────────────────────────────────────────────────
    # TREE RENDER - flat StackPanel rows
    # ──────────────────────────────────────────────────────────────────

    def _render_tree(self, skipped=0):
        """Rebuild the flat row list from scratch, then apply visibility."""
        self.tree_items.Children.Clear()

        if not self._root_nodes:
            msg = TextBlock()
            if skipped > 0:
                msg.Text = (
                    u'No elements to display.\n'
                    u'{} elements were skipped (null category or unresolvable type).\n\n'
                    u'Possible causes: model has only non-categorised elements, '
                    u'or the active document has no placed instances.'
                ).format(skipped)
            else:
                msg.Text = (
                    u'No elements found in the model '
                    u'(FilteredElementCollector returned 0 non-type elements).'
                )
            msg.Foreground   = SolidColorBrush(ColorConverter.ConvertFromString(u'#888888'))
            msg.Margin       = Thickness(10)
            msg.FontSize     = 11.0
            msg.FontFamily   = _SEGOE
            msg.TextWrapping = TextWrapping.Wrap
            self.tree_items.Children.Add(msg)
            return

        # Add ALL nodes in display order; initial visibility set in _build_row.
        # type/instance rows start Collapsed; cat/family rows start Visible.
        # _filter_tree() is called later via user actions (search, expand, etc.).
        for node in self._root_nodes:
            self._add_node_and_children(node)

    def _add_node_and_children(self, node):
        """Add node row and recursively all its descendants to tree_items."""
        row = self._build_row(node)
        self.tree_items.Children.Add(row)
        for child in node.children:
            self._add_node_and_children(child)

    def _build_row(self, node):
        """Build a single flat row Border for the given TreeNode."""
        row = Border()
        row.MinHeight   = 24.0
        node.wpf_item   = row
        node._header_border = row
        # Only category rows are visible on first render.
        # family / type / instance start hidden; clicking ▶ on a category
        # calls _filter_tree() which reveals families, and so on.
        if node.level != u'category':
            row.Visibility = Visibility.Collapsed

        _xi = node.extra_indent
        if node.level == u'category':
            row.SetResourceReference(Border.BackgroundProperty,   u'CatBgBrush')
            row.SetResourceReference(Border.BorderBrushProperty,  u'AccentBrush')
            row.BorderThickness = Thickness(3, 0, 0, 0)
            row.Padding         = Thickness(_xi, 2, 6, 2)
        elif node.level == u'family':
            row.SetResourceReference(Border.BackgroundProperty,   u'FamBgBrush')
            row.BorderThickness = Thickness(0)
            row.Padding         = Thickness(14 + _xi, 2, 6, 2)
        elif node.level == u'type':
            row.Background      = Brushes.Transparent
            row.BorderThickness = Thickness(0)
            row.Padding         = Thickness(28 + _xi, 2, 6, 2)
        else:  # instance
            row.Background      = Brushes.Transparent
            row.BorderThickness = Thickness(0)
            row.Padding         = Thickness(42 + _xi, 2, 6, 2)

        sp             = StackPanel()
        sp.Orientation = Orientation.Horizontal

        # Expand / collapse icon (▼ / ►)
        tex = TextBlock()
        tex.Width             = 14.0
        tex.FontSize          = 8.0
        tex.TextAlignment     = TextAlignment.Center
        tex.VerticalAlignment = VerticalAlignment.Center
        tex.Margin            = Thickness(0, 0, 4, 0)
        tex.SetResourceReference(TextBlock.ForegroundProperty, u'Text2Brush')

        if node.level == u'instance' or not node.children:
            tex.Text          = u''
            node._expand_icon = None
        elif id(node) in self._collapsed:
            tex.Text              = u'▶'
            tex.Tag               = node
            tex.MouseLeftButtonUp += self._toggle_expand_node
            tex.Cursor            = _hand_cursor()
            node._expand_icon     = tex
        else:
            tex.Text              = u'▼'
            tex.Tag               = node
            tex.MouseLeftButtonUp += self._toggle_expand_node
            tex.Cursor            = _hand_cursor()
            node._expand_icon     = tex
        sp.Children.Add(tex)

        # CheckBox
        cb = CheckBox()
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Width             = 13.0
        cb.Margin            = Thickness(0, 0, 6, 0)
        if node.is_indeterminate:
            cb.IsChecked = None
        else:
            cb.IsChecked = node.is_checked
        cb.Tag            = node
        cb.Checked       += self._on_node_checked
        cb.Unchecked     += self._on_node_unchecked
        cb.Indeterminate += self._on_node_indeterminate
        node._checkbox    = cb
        sp.Children.Add(cb)

        # Label
        lbl = TextBlock()
        lbl.Text             = node.label
        lbl.VerticalAlignment = VerticalAlignment.Center
        lbl.FontFamily       = _SEGOE
        lbl.FontSize         = 11.0
        if node.level == u'category':
            lbl.FontWeight = FontWeights.SemiBold
        lbl.SetResourceReference(TextBlock.ForegroundProperty, u'TextBrush')
        sp.Children.Add(lbl)

        # Count badge (non-instance)
        if node.level != u'instance':
            total   = self._count_instances(node)
            checked = self._count_checked(node)
            cnt = TextBlock()
            cnt.Text              = u' ({}/{})'.format(checked, total)
            cnt.FontSize          = 10.0
            cnt.SetResourceReference(TextBlock.ForegroundProperty, u'Text3Brush')
            cnt.VerticalAlignment = VerticalAlignment.Center
            node._count_label     = cnt
            sp.Children.Add(cnt)
        else:
            node._count_label = None

        row.Child = sp

        # Instance background (checked / Revit selected)
        if node.level == u'instance':
            self._update_instance_bg(node)

        return row

    # ──────────────────────────────────────────────────────────────────
    # EXPAND / COLLAPSE
    # ──────────────────────────────────────────────────────────────────

    def _toggle_expand_node(self, sender, args):
        node = sender.Tag
        if node is None or node.level == u'instance':
            return
        if id(node) in self._collapsed:
            self._collapsed.discard(id(node))
            if node._expand_icon is not None:
                node._expand_icon.Text = u'▼'
        else:
            self._collapsed.add(id(node))
            if node._expand_icon is not None:
                node._expand_icon.Text = u'▶'
        self._filter_tree()

    def _is_ancestor_expanded(self, node):
        """Return True when every ancestor of node is NOT in self._collapsed."""
        parent = node.parent
        while parent is not None:
            if id(parent) in self._collapsed:
                return False
            parent = parent.parent
        return True

    # ──────────────────────────────────────────────────────────────────
    # CHECKBOX HANDLERS
    # ──────────────────────────────────────────────────────────────────

    def _on_node_checked(self, sender, args):
        if self._updating:
            return
        node = sender.Tag
        if node is None:
            return
        self._updating = True
        try:
            self._set_subtree_checked(node, True)
            self._update_parent_states()
            self._update_status()
        finally:
            self._updating = False

    def _on_node_unchecked(self, sender, args):
        if self._updating:
            return
        node = sender.Tag
        if node is None:
            return
        self._updating = True
        try:
            self._set_subtree_checked(node, False)
            self._update_parent_states()
            self._update_status()
        finally:
            self._updating = False

    def _on_node_indeterminate(self, sender, args):
        pass  # user cannot set indeterminate directly

    def _set_subtree_checked(self, node, state):
        node.is_checked       = state
        node.is_indeterminate = False

        if node.level == u'instance' and node.element_id:
            k = eid_key(node.element_id)
            if state:
                self._checked_ids[k] = node.element_id
            else:
                self._checked_ids.pop(k, None)
            self._update_instance_bg(node)
            # Cross-sync: mirror state to all sibling nodes (same element in
            # main tree AND group subtree can both point to the same ElementId).
            for sibling in self._key_to_nodes.get(k, []):
                if sibling is node:
                    continue
                sibling.is_checked       = state
                sibling.is_indeterminate = False
                if sibling._checkbox is not None:
                    sibling._checkbox.IsChecked = state
                self._update_instance_bg(sibling)

        if node._checkbox is not None:
            node._checkbox.IsChecked = state

        for child in node.children:
            self._set_subtree_checked(child, state)

    def _update_instance_bg(self, node):
        if node._header_border is None:
            return
        k = eid_key(node.element_id) if node.element_id else None
        try:
            if node.is_checked:
                node._header_border.Background = self.Resources[u'CkBgBrush']
            elif k is not None and k in self._last_revit_ids:
                node._header_border.Background = self.Resources[u'RvBgBrush']
            else:
                node._header_border.Background = Brushes.Transparent
        except Exception:
            node._header_border.Background = Brushes.Transparent

    # ──────────────────────────────────────────────────────────────────
    # INSTANCE COUNTS
    # ──────────────────────────────────────────────────────────────────

    def _count_instances(self, node):
        if node.level == u'instance':
            return 1
        total = 0
        for child in node.children:
            total += self._count_instances(child)
        return total

    def _count_checked(self, node):
        if node.level == u'instance':
            return 1 if node.is_checked else 0
        total = 0
        for child in node.children:
            total += self._count_checked(child)
        return total

    # ──────────────────────────────────────────────────────────────────
    # POLLING & STATUS
    # ──────────────────────────────────────────────────────────────────

    def _start_polling(self):
        self._poll_timer          = DispatcherTimer()
        self._poll_timer.Interval = TimeSpan.FromMilliseconds(500)
        self._poll_timer.Tick    += self._poll_tick
        self._poll_timer.Start()

    def _collect_group_member_keys(self, group_el, result_set):
        """Recursively add eid_keys of all non-Group members to result_set."""
        from Autodesk.Revit.DB import Group as _Group
        try:
            for mid in group_el.GetMemberIds():
                mem = self.doc.GetElement(mid)
                if mem is None:
                    continue
                if isinstance(mem, _Group):
                    self._collect_group_member_keys(mem, result_set)
                else:
                    result_set.add(eid_key(mem.Id))
        except Exception:
            pass

    def _expand_selection_keys(self):
        """Return frozenset of eid_keys from current Revit selection.
        Group elements are replaced by their member eid_keys (recursively).
        Uses _eid_map as a fast membership test: keys absent from _eid_map
        may be Groups and are resolved via GetElement only when needed.
        """
        from Autodesk.Revit.DB import Group as _Group
        result = set()
        try:
            for eid in self.uidoc.Selection.GetElementIds():
                k = eid_key(eid)
                if k in self._eid_map:
                    result.add(k)
                else:
                    el = self.doc.GetElement(eid)
                    if el is not None and isinstance(el, _Group):
                        self._collect_group_member_keys(el, result)
                    else:
                        result.add(k)
        except Exception:
            pass
        return frozenset(result)

    def _poll_tick(self, sender, args):
        try:
            ids = self._expand_selection_keys()
            if ids != self._last_revit_ids:
                self._last_revit_ids = ids
                self._sync_revit_visual()
                self._update_status()
        except Exception:
            pass

    def _update_status(self):
        self.txt_checked_count.Text = unicode(len(self._checked_ids))
        self.txt_revit_count.Text   = unicode(len(self._last_revit_ids))
        self._update_adv_scope_text()

    def _sync_revit_visual(self):
        for node in self._all_nodes:
            if node.level == u'instance' and node._header_border is not None:
                self._update_instance_bg(node)

    # ──────────────────────────────────────────────────────────────────
    # SEARCH
    # ──────────────────────────────────────────────────────────────────

    def _search_changed(self, sender, args):
        q = self.txt_search.Text or u''
        self.txt_placeholder.Visibility = (
            Visibility.Collapsed if q else Visibility.Visible
        )
        self.btn_search_clear.Visibility = (
            Visibility.Visible if q else Visibility.Collapsed
        )
        self._search_query = q.lower()
        self._start_debounce()

    def _clear_search(self, sender, args):
        self.txt_search.Text = u''

    def _start_debounce(self):
        if self._debounce_timer is not None:
            self._debounce_timer.Stop()
        self._debounce_timer          = DispatcherTimer()
        self._debounce_timer.Interval = TimeSpan.FromMilliseconds(300)
        self._debounce_timer.Tick    += self._debounce_tick
        self._debounce_timer.Start()

    def _debounce_tick(self, sender, args):
        self._debounce_timer.Stop()
        if self._search_query:
            self._auto_expand_for_search()
        self._filter_tree()

    def _auto_expand_for_search(self):
        """Expand ancestor nodes of all instances matching the current query.
        Called once per query change - manual collapse during search is preserved."""
        q            = self._search_query
        scope        = self._search_scope
        checked_only = self._show_checked_only
        for node in self._all_nodes:
            if node.level != u'instance':
                continue
            if checked_only and not node.is_checked:
                continue
            if not self._node_matches(node, q, scope):
                continue
            parent = node.parent
            while parent is not None:
                if id(parent) in self._collapsed:
                    self._collapsed.discard(id(parent))
                    if parent._expand_icon is not None:
                        parent._expand_icon.Text = u'▼'
                parent = parent.parent

    def _filter_tree(self):
        q            = self._search_query
        scope        = self._search_scope
        checked_only = self._show_checked_only
        has_filter   = bool(q) or checked_only

        if not has_filter:
            # No active filter - show/hide based on expand/collapse state only.
            for node in self._all_nodes:
                if node.wpf_item is not None:
                    node.wpf_item.Visibility = (
                        Visibility.Visible
                        if self._is_ancestor_expanded(node)
                        else Visibility.Collapsed
                    )
            return

        # Filter active: compute which nodes match, ignore collapse state.
        vis = {}

        # Pass 1: instance visibility
        for node in self._all_nodes:
            if node.level != u'instance':
                vis[id(node)] = False
                continue
            show = True
            if checked_only and not node.is_checked:
                show = False
            if show and q:
                show = self._node_matches(node, q, scope)
            vis[id(node)] = show

        # Pass 2: propagate up - forward = bottom-up because leaves are
        # appended before their parents in _all_nodes.
        for node in self._all_nodes:
            if node.level == u'instance':
                continue
            vis[id(node)] = any(vis.get(id(c), False) for c in node.children)

        # Apply: all nodes respect both vis (matching) and expand/collapse state.
        # _auto_expand_for_search() has already opened paths of matching results.
        for node in self._all_nodes:
            if node.wpf_item is not None:
                show = vis.get(id(node), False) and self._is_ancestor_expanded(node)
                node.wpf_item.Visibility = (
                    Visibility.Visible if show else Visibility.Collapsed
                )

    def _node_matches(self, node, query, scope):
        if scope.get(u'family') and query in node.family_name.lower():
            return True
        if scope.get(u'type') and query in node.type_name.lower():
            return True
        if scope.get(u'param_name') or scope.get(u'param_value'):
            for name, val in self._get_params(node):
                if scope.get(u'param_name') and query in name.lower():
                    return True
                if scope.get(u'param_value') and query in val.lower():
                    return True
        return False

    def _get_params(self, node):
        if not node.element_id:
            return []
        k = eid_key(node.element_id)
        if k not in self._param_cache:
            result = []
            try:
                el = self.doc.GetElement(node.element_id)
                if el is not None:
                    for p in el.Parameters:
                        try:
                            defn = p.Definition
                            if defn is None:
                                continue
                            name = defn.Name
                            val  = p.AsString()
                            if not val:
                                val = p.AsValueString() or u''
                            result.append((name, val))
                        except Exception:
                            pass
            except Exception:
                pass
            self._param_cache[k] = result
        return self._param_cache[k]

    def _update_adv_scope_text(self):
        if self._adv_scope == u'checked':
            n    = len(self._checked_ids)
            noun = u'checked element' if n == 1 else u'checked elements'
            self.txt_adv_scope.Text = (
                u'Scope: {} {} - build conditions '
                u'below, then Apply'.format(n, noun)
            )
        else:
            self.txt_adv_scope.Text = (
                u'Scope: all elements in model - '
                u'build conditions below, then Apply'
            )

    def _update_search_placeholder(self):
        parts = []
        if self.cb_scope_family.IsChecked:
            parts.append(u'families')
        if self.cb_scope_type.IsChecked:
            parts.append(u'types')
        if self.cb_scope_paramname.IsChecked:
            parts.append(u'parameter names')
        if self.cb_scope_paramval.IsChecked:
            parts.append(u'parameter values')
        if parts:
            text = u'Search ' + u', '.join(parts) + u'...'
        else:
            text = u'Search...'
        self.txt_placeholder.Text = text

    def _scope_changed(self, sender, args):
        self._search_scope = {
            u'family':      self.cb_scope_family.IsChecked,
            u'type':        self.cb_scope_type.IsChecked,
            u'param_name':  self.cb_scope_paramname.IsChecked,
            u'param_value': self.cb_scope_paramval.IsChecked,
        }
        self._update_search_placeholder()
        self._filter_tree()

    # ──────────────────────────────────────────────────────────────────
    # TREE TOOLBAR
    # ──────────────────────────────────────────────────────────────────

    def _expand_all(self, sender, args):
        self._collapsed.clear()
        for node in self._all_nodes:
            if node._expand_icon is not None:
                node._expand_icon.Text = u'▼'
        self._filter_tree()

    def _collapse_all(self, sender, args):
        for node in self._all_nodes:
            if node.level != u'instance':
                self._collapsed.add(id(node))
                if node._expand_icon is not None:
                    node._expand_icon.Text = u'▶'
        self._filter_tree()

    def _select_all(self, sender, args):
        self._updating = True
        try:
            for node in self._all_nodes:
                if node.level != u'instance' or not node.element_id:
                    continue
                node.is_checked = True
                self._checked_ids[eid_key(node.element_id)] = node.element_id
                if node._checkbox is not None:
                    node._checkbox.IsChecked = True
                self._update_instance_bg(node)
            self._update_parent_states()
            self._update_status()
        finally:
            self._updating = False

    def _deselect_all(self, sender, args):
        self._updating = True
        try:
            for node in self._all_nodes:
                if node.level != u'instance':
                    continue
                node.is_checked = False
                if node.element_id:
                    self._checked_ids.pop(eid_key(node.element_id), None)
                if node._checkbox is not None:
                    node._checkbox.IsChecked = False
                self._update_instance_bg(node)
            self._update_parent_states()
            self._update_status()
        finally:
            self._updating = False

    def _checked_only_on(self, sender, args):
        self._show_checked_only = True
        self._expand_ancestors_of_checked()
        self._filter_tree()

    def _checked_only_off(self, sender, args):
        self._show_checked_only = False
        self._filter_tree()

    # ──────────────────────────────────────────────────────────────────
    # PARENT STATE PROPAGATION
    # ──────────────────────────────────────────────────────────────────

    def _update_parent_states(self):
        # Forward = bottom-up: leaves are appended before their parents in _all_nodes,
        # so iterating forward ensures children are processed before their parents.
        for node in self._all_nodes:
            if node.level == u'instance' or not node.children:
                continue
            checked = sum(1 for c in node.children if c.is_checked)
            indet   = sum(1 for c in node.children if c.is_indeterminate)
            total   = len(node.children)
            if checked == total:
                node.is_checked       = True
                node.is_indeterminate = False
            elif checked > 0 or indet > 0:
                node.is_checked       = False
                node.is_indeterminate = True
            else:
                node.is_checked       = False
                node.is_indeterminate = False
            if node._checkbox is not None:
                node._checkbox.IsChecked = (
                    None if node.is_indeterminate else node.is_checked
                )
            if node._count_label is not None:
                node._count_label.Text = u' ({}/{})'.format(
                    self._count_checked(node),
                    self._count_instances(node)
                )

    def _expand_ancestors_of_checked(self):
        """Expand family/category ancestors of checked nodes, but keep types
        collapsed (▶). In Checked Only view: types are visible (non-leaf, vis-based)
        but their instances stay hidden until the user clicks ▶ on the type."""
        for node in self._all_nodes:
            if node.level != u'instance' or not node.is_checked:
                continue
            # Skip the type (direct parent) - start expanding from family upward.
            p = node.parent   # type  - leave in _collapsed
            if p is not None:
                p = p.parent  # family - expand from here
            while p is not None:
                if id(p) in self._collapsed:
                    self._collapsed.discard(id(p))
                    if p._expand_icon is not None:
                        p._expand_icon.Text = u'▼'
                p = p.parent

    # ──────────────────────────────────────────────────────────────────
    # HEADER ACTIONS
    # ──────────────────────────────────────────────────────────────────

    def _always_on_top_changed(self, sender, args):
        self.Topmost = bool(self.cb_always_on_top.IsChecked)

    def _active_view_changed(self, sender, args):
        self._build_tree()

    def _refresh_click(self, sender, args):
        self.txt_search.Text        = u''
        self._search_query          = u''
        self._checked_ids           = {}
        self._show_checked_only     = False
        self.btn_checked_only.IsChecked = False
        self._start_shimmer()
        self._build_tree()
        self._stop_shimmer()

    # ──────────────────────────────────────────────────────────────────
    # SYNC BUTTONS
    # ──────────────────────────────────────────────────────────────────

    def _check_from_revit(self, sender, args):
        # Re-query and expand selection at click time so Group selections
        # are resolved to their member eid_keys before checking tree nodes.
        target_keys = self._expand_selection_keys()
        self._updating = True
        try:
            for node in self._all_nodes:
                if node.level != u'instance' or not node.element_id:
                    continue
                k       = eid_key(node.element_id)
                checked = k in target_keys
                node.is_checked = checked
                if checked:
                    self._checked_ids[k] = node.element_id
                else:
                    self._checked_ids.pop(k, None)
                if node._checkbox is not None:
                    node._checkbox.IsChecked = checked
                self._update_instance_bg(node)
            self._update_parent_states()
            self._update_status()
        finally:
            self._updating = False
        # Switch to Checked Only so the user immediately sees what was imported.
        # Expand all nodes so parent rows are visible, but keep group-instance
        # nodes (the id:X containers) collapsed - the user sees which group was
        # matched without the member list being auto-expanded.
        self._show_checked_only         = True
        self.btn_checked_only.IsChecked = True
        self._collapsed.clear()
        for node in self._all_nodes:
            if node._expand_icon is not None:
                node._expand_icon.Text = u'▼'
            if (node.level == u'type'
                    and node.category_name == u'Model Groups'
                    and node.children):
                self._collapsed.add(id(node))
                if node._expand_icon is not None:
                    node._expand_icon.Text = u'▶'
        self._filter_tree()

    def _select_in_revit(self, sender, args):
        from view_actions import apply_selection
        ids = list(self._checked_ids.values())
        self._last_revit_ids = frozenset(self._checked_ids.keys())
        apply_selection(ids, mode=u'replace')
        self._update_status()

    # ──────────────────────────────────────────────────────────────────
    # ACTION ROW
    # ──────────────────────────────────────────────────────────────────

    def _isolate_click(self, sender, args):
        if not self._checked_ids:
            return
        from view_actions import apply_isolate
        apply_isolate(list(self._checked_ids.values()))

    def _temp_hide_click(self, sender, args):
        if not self._checked_ids:
            return
        from view_actions import apply_temp_hide
        apply_temp_hide(list(self._checked_ids.values()))

    def _reset_click(self, sender, args):
        from view_actions import apply_reset, apply_color
        apply_reset()
        if self._colored_ids:
            apply_color(
                [self._eid_map[k] for k in self._colored_ids if k in self._eid_map],
                wpf_color=None, reset=True
            )
            self._colored_ids.clear()

    def _sel_inview_click(self, sender, args):
        from view_actions import apply_select_all
        apply_select_all(scope=u'view')

    def _sel_inproj_click(self, sender, args):
        from view_actions import apply_select_all
        apply_select_all(scope=u'project')

    def _selbox_click(self, sender, args):
        if not self._checked_ids:
            return
        from view_actions import apply_selection_box
        apply_selection_box(list(self._checked_ids.values()))

    def _color_click(self, sender, args):
        from System.Windows.Controls.Primitives import Popup, PlacementMode
        from System.Windows.Media import SolidColorBrush as SCB, ColorConverter as CC

        SWATCHES = [
            (u'#E67E22', u'Orange'),
            (u'#C0392B', u'Red'),
            (u'#27AE60', u'Green'),
            (u'#2980B9', u'Blue'),
            (u'#8E44AD', u'Purple'),
            (u'#F1C40F', u'Yellow'),
        ]

        if self._color_popup is None:
            popup                 = Popup()
            popup.PlacementTarget = self.btn_color
            popup.Placement       = PlacementMode.Top
            popup.StaysOpen       = False

            row             = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Background  = SCB(CC.ConvertFromString(u'#FFFFFF'))

            for hex_c, name in SWATCHES:
                btn            = Button()
                btn.Width      = 22
                btn.Height     = 22
                btn.Margin     = Thickness(3)
                btn.Background = SCB(CC.ConvertFromString(hex_c))
                btn.ToolTip    = name
                btn.Tag        = hex_c
                btn.Click     += self._color_swatch_click
                row.Children.Add(btn)

            popup.Child       = row
            self._color_popup = popup

        self._color_popup.IsOpen = True

    def _color_swatch_click(self, sender, args):
        from System.Windows.Media import ColorConverter
        from view_actions import apply_color
        color = ColorConverter.ConvertFromString(sender.Tag)
        eids  = list(self._checked_ids.values())
        self._colored_ids = dict(self._checked_ids)
        apply_color(eids, color, reset=False)
        self._color_popup.IsOpen = False

    # ──────────────────────────────────────────────────────────────────
    # JSON
    # ──────────────────────────────────────────────────────────────────

    def _save_json_click(self, sender, args):
        try:
            from filter_storage import save_checked_ids
            from System.Windows import MessageBox
            if not self._checked_ids:
                MessageBox.Show(self, u'No elements checked.', u'FilterMan')
                return
            from Microsoft.Win32 import SaveFileDialog
            dlg = SaveFileDialog()
            dlg.Filter     = u'JSON files (*.json)|*.json|All files (*.*)|*.*'
            dlg.DefaultExt = u'json'
            dlg.FileName   = u'filterman_selection'
            ok = dlg.ShowDialog(self)  # owner = FilterMan; dialog appears in front
            if ok == True:
                save_checked_ids(self._checked_ids, dlg.FileName)
        except Exception as ex:
            from System.Windows import MessageBox
            MessageBox.Show(self, unicode(ex), u'FilterMan - Save JSON')

    def _load_json_click(self, sender, args):
        try:
            from filter_storage import load_checked_ids
            from Microsoft.Win32 import OpenFileDialog
            dlg = OpenFileDialog()
            dlg.Filter = u'JSON files (*.json)|*.json|All files (*.*)|*.*'
            ok = dlg.ShowDialog(self)  # owner = FilterMan; dialog appears in front
            if ok != True:
                return
            filepath = dlg.FileName
            try:
                int_ids = set(load_checked_ids(filepath))
            except Exception as ex:
                from System.Windows import MessageBox
                MessageBox.Show(
                    self,
                    u'Could not read file:\n{}'.format(unicode(ex)),
                    u'FilterMan')
                return
            self._updating = True
            try:
                for node in self._all_nodes:
                    if node.level != u'instance' or not node.element_id:
                        continue
                    k       = eid_key(node.element_id)
                    checked = k in int_ids
                    node.is_checked = checked
                    if checked:
                        self._checked_ids[k] = node.element_id
                    else:
                        self._checked_ids.pop(k, None)
                    if node._checkbox is not None:
                        node._checkbox.IsChecked = checked
                    self._update_instance_bg(node)
                self._update_parent_states()
                self._update_status()
            finally:
                self._updating = False
        except Exception as ex:
            from System.Windows import MessageBox
            MessageBox.Show(self, unicode(ex), u'FilterMan - Load JSON')

    # ──────────────────────────────────────────────────────────────────
    # ADVANCED PANEL
    # ──────────────────────────────────────────────────────────────────

    def _toggle_advanced(self, sender, args):
        self._adv_open = not self._adv_open

        if self._adv_open:
            self.adv_panel.Visibility     = Visibility.Visible
            self.grid_splitter.Visibility = Visibility.Visible
            self.btn_toggle_adv.Content   = u'Close Advanced ◄'
            # _load_advanced_data called inside _animate_adv after panel is fully open
        else:
            self.btn_toggle_adv.Content = u'Advanced - Filter Table & Excel ►'
            # Visibility.Collapsed is set inside _animate_adv on completion
            if self._import_preview_mode:
                self._set_preview_locked(False)

        self._animate_adv(self._adv_open)

    def _animate_adv(self, opening):
        """Slide advanced panel in/out over ~112ms with ease-out."""
        ADV_W   = 620
        STEPS   = 8
        STEP_MS = 14   # ~112ms total - snappy but smooth

        if self._adv_anim_timer is not None:
            try:
                self._adv_anim_timer.Stop()
            except Exception:
                pass

        # Pin col_main to current pixel width during animation.
        # On opening completion: col_adv becomes Star so resizing the window
        # grows the Advanced panel (right side). col_main stays fixed.
        # On closing completion: col_main becomes Star; col_adv = 0.
        _mw = self.col_main.ActualWidth
        _mw_px = float(_mw) if _mw > 0 else 460.0
        self.col_main.Width = GridLength(_mw_px)

        # Use ActualWidth (pixels) regardless of whether col_adv is Star or fixed.
        start_col = float(self.col_adv.ActualWidth)
        # If col_adv is currently Star, freeze it at its actual pixel size first.
        if self.col_adv.Width.IsStar:
            self.col_adv.Width = GridLength(start_col)
        end_col   = float(ADV_W) if opening else 0.0
        _aw       = self.ActualWidth
        start_win = float(_aw) if _aw > 0 else 660.0
        end_win   = (start_win + ADV_W + 5) if opening else (start_win - ADV_W - 5)

        _step = [0]

        def _tick(s, a):
            _step[0] += 1
            t    = _step[0] / float(STEPS)
            ease = 1.0 - (1.0 - t) * (1.0 - t)
            w_col = start_col + (end_col - start_col) * ease
            w_win = start_win + (end_win - start_win) * ease
            self.col_adv.Width = GridLength(max(0.0, w_col))
            self.Width         = max(460.0, w_win)
            if _step[0] >= STEPS:
                self._adv_anim_timer.Stop()
                self.Width = max(460.0, end_win)
                if opening:
                    # col_main fixed, col_adv Star - window resize grows Advanced panel
                    self.col_main.Width = GridLength(_mw_px)
                    self.col_adv.Width  = GridLength(1, GridUnitType.Star)
                    self._load_advanced_data()
                else:
                    self.col_adv.Width  = GridLength(0.0)
                    self.col_main.Width = GridLength(1, GridUnitType.Star)
                    self.adv_panel.Visibility     = Visibility.Collapsed
                    self.grid_splitter.Visibility = Visibility.Collapsed

        tmr          = DispatcherTimer()
        tmr.Interval = TimeSpan.FromMilliseconds(STEP_MS)
        tmr.Tick    += _tick
        tmr.Start()
        self._adv_anim_timer = tmr

    def _scope_checked_click(self, sender, args):
        self.btn_scope_checked.IsChecked = True
        self.btn_scope_all.IsChecked     = False
        self._adv_scope = u'checked'
        self._update_adv_scope_text()
        if self._adv_open:
            self._load_advanced_data()

    def _scope_all_click(self, sender, args):
        self.btn_scope_all.IsChecked     = True
        self.btn_scope_checked.IsChecked = False
        self._adv_scope = u'all'
        self._update_adv_scope_text()
        if self._adv_open:
            self._load_advanced_data()

    def _load_advanced_data(self):
        if self._import_preview_mode:
            return  # table locked while preview is active
        try:
            from advanced_panel import (
                get_elements_for_scope, apply_conditions,
                build_table_rows, get_available_params
            )
            view_id  = self.uidoc.ActiveView.Id \
                       if self.cb_active_view.IsChecked == True else None
            elements = get_elements_for_scope(
                self.doc, self._adv_scope, self._checked_ids, view_id
            )
            elements              = apply_conditions(elements, self._adv_conditions)
            self._adv_param_cache = get_available_params(elements)
            _builtin_keys = {u'family_name', u'type_name'}
            param_cols = [key for _, key in self._adv_columns
                          if key not in _builtin_keys]
            self._adv_rows = build_table_rows(elements, param_cols, self.doc)
            self._apply_import_failures(self._adv_rows)
            self._rebuild_datagrid()
            count = self._adv_rows.Rows.Count
            self.txt_table_count.Text = unicode(count)
            scope_label = (
                u'checked elements' if self._adv_scope == u'checked' else u'elements'
            )
            self.txt_adv_scope.Text = (
                u'Scope: {} {} - build conditions below, then Apply'.format(
                    count, scope_label
                )
            )
            self._update_import_btn_state()
        except Exception as ex:
            from System.Windows import MessageBox
            MessageBox.Show(self, unicode(ex), u'FilterMan - Advanced')

    def _rebuild_datagrid(self):
        from System.Windows.Data import Binding
        from System.Windows.Controls import Dock, DataGridLength
        from System.Windows import Style, Setter, TextAlignment, VerticalAlignment
        import System.Windows.Media as SWM

        self.data_grid.Columns.Clear()

        # Right-align style for ID column
        right_style = Style()
        right_style.Setters.Add(
            Setter(TextBlock.TextAlignmentProperty, TextAlignment.Right)
        )

        def _fixed_col(header, key, width=None, align_right=False):
            c                = DataGridTextColumn()
            c.Header         = header
            c.Binding        = Binding(key)
            c.SortMemberPath = key
            c.IsReadOnly     = True
            if width is not None:
                c.Width = DataGridLength(float(width))
            if align_right:
                c.ElementStyle = right_style
            self.data_grid.Columns.Add(c)

        def _removable_col(header, key):
            from System.Windows.Controls.Primitives import DataGridColumnHeader
            from System.Windows import HorizontalAlignment as HA
            c               = DataGridTextColumn()
            c.Binding       = Binding(key)
            c.SortMemberPath = key
            c.IsReadOnly    = True
            c.CanUserResize = True
            c.Width         = DataGridLength(130.0)
            # Inherit ColumnHeaderStyle so dark/light theme applies; add Stretch
            hdr_style = Style(DataGridColumnHeader,
                              self.data_grid.ColumnHeaderStyle)
            hdr_style.Setters.Add(
                Setter(DataGridColumnHeader.HorizontalContentAlignmentProperty,
                       HA.Stretch)
            )
            c.HeaderStyle = hdr_style
            hdr = DockPanel()
            hdr.LastChildFill = True
            btn = Button()
            btn.Content           = u'x'
            btn.Width             = 18
            btn.Height            = 18
            btn.Background        = SWM.Brushes.Transparent
            btn.BorderThickness   = Thickness(0)
            btn.SetResourceReference(Button.ForegroundProperty, u'TextBrush')
            btn.VerticalAlignment = VerticalAlignment.Center
            DockPanel.SetDock(btn, Dock.Right)
            hdr.Children.Add(btn)   # must be added before LastChildFill element
            lbl = TextBlock()
            lbl.Text              = header
            lbl.TextAlignment     = TextAlignment.Center
            lbl.VerticalAlignment = VerticalAlignment.Center
            hdr.Children.Add(lbl)  # LastChildFill - fills space left of button
            # Red cell highlight for import failures
            # Guard: _f_{key} column exists only for custom param columns,
            # not for built-in family_name/type_name (no _f_ counterpart).
            # Binding to a non-existent DataTable column causes ArgumentException
            # during WPF row rendering, which makes the entire DataGrid appear empty.
            _f_col = u'_f_' + key
            if (self._adv_rows is not None
                    and self._adv_rows.Columns.Contains(_f_col)):
                from System.Windows import DataTrigger, Setter as WPFSetter
                from System.Windows.Controls import Control as WPFCtrl
                cell_style = Style()
                # '1' = red (not writable / failed)
                trig_red = DataTrigger()
                trig_red.Binding = Binding(_f_col)
                trig_red.Value   = u'1'
                trig_red.Setters.Add(
                    WPFSetter(WPFCtrl.BackgroundProperty, SWM.Brushes.LightCoral)
                )
                cell_style.Triggers.Add(trig_red)
                # '2' = green (value will change and param is writable)
                trig_green = DataTrigger()
                trig_green.Binding = Binding(_f_col)
                trig_green.Value   = u'2'
                trig_green.Setters.Add(
                    WPFSetter(WPFCtrl.BackgroundProperty, SWM.Brushes.LightGreen)
                )
                cell_style.Triggers.Add(trig_green)
                c.CellStyle = cell_style
            def _remove(s, a, _h=header, _k=key):
                if (_h, _k) in self._adv_columns:
                    self._adv_columns.remove((_h, _k))
                    if self._adv_open:
                        self._load_advanced_data()
            btn.Click += _remove
            c.Header = hdr
            self.data_grid.Columns.Add(c)

        _fixed_col(u'ID', u'id', width=65, align_right=True)
        for header, key in self._adv_columns:
            _removable_col(header, key)

        self.data_grid.ItemsSource = (
            self._adv_rows.DefaultView if self._adv_rows is not None else None
        )

    # ── QB (Phase 13) ─────────────────────────────────────────────────

    def _qb_apply(self, sender, args):
        conditions = []
        for row in self.qb_conditions_panel.Children:
            try:
                refs = row.Tag
                if not isinstance(refs, dict):
                    continue
                field_item = refs[u'cb_field'].SelectedItem
                op_item    = refs[u'cb_op'].SelectedItem
                if field_item is None or op_item is None:
                    continue
                field      = field_item.Tag
                op         = op_item.Tag
                value      = refs[u'txt_val'].Text or u''
                param_name = None
                if field == u'parameter':
                    pi = refs[u'cb_param'].SelectedItem
                    if pi is not None:
                        param_name = pi.Tag
                conditions.append(
                    FilterCondition(field=field, operator=op,
                                    value=value, param_name=param_name)
                )
            except Exception:
                continue
        self._adv_conditions = conditions
        if self._adv_open:
            self._load_advanced_data()

    def _qb_clear(self, sender, args):
        self.qb_conditions_panel.Children.Clear()
        self._adv_conditions = []
        if self._adv_open:
            self._load_advanced_data()

    def _add_condition_row(self, sender, args):
        from System.Windows.Controls import Dock
        row = DockPanel()
        row.Margin        = Thickness(0, 3, 0, 0)
        row.LastChildFill = True

        # Remove button - must be docked Right before LastChildFill child
        btn_rm = Button()
        btn_rm.Content = u'×'
        btn_rm.Width   = 24
        btn_rm.Margin  = Thickness(4, 0, 0, 0)
        btn_rm.SetResourceReference(Button.BackgroundProperty, u'Surface2Brush')
        btn_rm.SetResourceReference(Button.ForegroundProperty, u'TextBrush')
        DockPanel.SetDock(btn_rm, Dock.Right)
        row.Children.Add(btn_rm)

        def _dock_left(ctrl, width, mr=4):
            ctrl.Width  = width
            ctrl.Margin = Thickness(0, 0, mr, 0)
            DockPanel.SetDock(ctrl, Dock.Left)
            row.Children.Add(ctrl)

        # Field ComboBox
        cb_field = ComboBox()
        for key, label in FilterCondition.FIELDS:
            it = ComboBoxItem()
            it.Content = label
            it.Tag     = key
            cb_field.Items.Add(it)
        cb_field.SelectedIndex = 0
        _dock_left(cb_field, 120)

        # Param ComboBox - shown only when field == 'parameter', before Op
        cb_param = ComboBox()
        cb_param.Visibility = Visibility.Collapsed
        _inst_params = sorted(
            [pn for pn, k in self._adv_param_cache.items() if k == u'I'],
            key=lambda s: s.lower()
        )
        _type_params = sorted(
            [pn for pn, k in self._adv_param_cache.items() if k == u'T'],
            key=lambda s: s.lower()
        )
        for _grp_label, _grp_names in (
            (u'Instance parameters', _inst_params),
            (u'Type parameters',     _type_params),
        ):
            if not _grp_names:
                continue
            hdr = ComboBoxItem()
            hdr.Content    = _grp_label
            hdr.IsEnabled  = False
            hdr.FontWeight = FontWeights.Bold
            cb_param.Items.Add(hdr)
            for pn in _grp_names:
                it = ComboBoxItem()
                it.Content = pn
                it.Tag     = pn
                cb_param.Items.Add(it)
        _dock_left(cb_param, 130)

        # Operator ComboBox
        cb_op = ComboBox()
        for key, label in FilterCondition.TEXT_OPS:
            it = ComboBoxItem()
            it.Content = label
            it.Tag     = key
            cb_op.Items.Add(it)
        cb_op.SelectedIndex = 0
        _dock_left(cb_op, 110)

        # Value TextBox - fills remaining width
        txt_val = TextBox()
        txt_val.SetResourceReference(TextBox.BackgroundProperty, u'Surface2Brush')
        txt_val.SetResourceReference(TextBox.ForegroundProperty, u'TextBrush')
        row.Children.Add(txt_val)

        # Tag stores refs so _qb_apply can read them
        row.Tag = {
            u'cb_field': cb_field,
            u'cb_op':    cb_op,
            u'cb_param': cb_param,
            u'txt_val':  txt_val,
        }

        # Field change → rebuild op list + toggle param ComboBox
        def _on_field_changed(s, a):
            sel = s.SelectedItem
            if sel is None:
                return
            is_param = (sel.Tag == u'parameter')
            cb_param.Visibility = (
                Visibility.Visible if is_param else Visibility.Collapsed
            )
            ops = (FilterCondition.TEXT_OPS + FilterCondition.NUM_OPS
                   if is_param else FilterCondition.TEXT_OPS)
            cb_op.Items.Clear()
            for k, lbl in ops:
                it = ComboBoxItem()
                it.Content = lbl
                it.Tag     = k
                cb_op.Items.Add(it)
            cb_op.SelectedIndex = 0

        cb_field.SelectionChanged += _on_field_changed

        def _on_remove(s, a):
            self.qb_conditions_panel.Children.Remove(row)

        btn_rm.Click += _on_remove

        self.qb_conditions_panel.Children.Add(row)

    # ── Column picker (Phase 14) ──────────────────────────────────────

    def _add_column_click(self, sender, args):
        from System.Windows import MessageBox
        import System.Windows as SW
        import System.Windows.Controls as SWC

        already = set(key for _, key in self._adv_columns)
        available = [p for p in self._adv_param_cache if p not in already]
        if not available:
            msg = (u'No parameters available. Run a search first.'
                   if not self._adv_param_cache
                   else u'All available parameters are already shown as columns.')
            MessageBox.Show(self, msg, u'FilterMan')
            return

        win = SW.Window()
        win.Owner  = self   # appears in front of FilterMan; no Topmost toggle needed
        win.Title  = u'Add Column'
        win.Width  = 320
        win.Height = 460
        win.WindowStartupLocation = SW.WindowStartupLocation.CenterOwner
        win.ResizeMode = SW.ResizeMode.NoResize

        outer = SWC.DockPanel()
        outer.Margin = SW.Thickness(10)
        outer.LastChildFill = True

        btn_ok = SWC.Button()
        btn_ok.Content = u'Add selected'
        btn_ok.Height  = 28
        btn_ok.Margin  = SW.Thickness(0, 8, 0, 0)
        SWC.DockPanel.SetDock(btn_ok, SWC.Dock.Bottom)
        outer.Children.Add(btn_ok)

        lbl = SWC.TextBlock()
        lbl.Text   = u'Select parameters to add as columns:'
        lbl.Margin = SW.Thickness(0, 0, 0, 6)
        SWC.DockPanel.SetDock(lbl, SWC.Dock.Top)
        outer.Children.Add(lbl)

        sv = SWC.ScrollViewer()
        sv.VerticalScrollBarVisibility = SWC.ScrollBarVisibility.Auto
        sp = SWC.StackPanel()
        sv.Content = sp
        outer.Children.Add(sv)  # LastChildFill

        checkboxes = []
        for pn in sorted(available):
            cb = SWC.CheckBox()
            tb = SWC.TextBlock()
            kind = self._adv_param_cache.get(pn, u'') if isinstance(self._adv_param_cache, dict) else u''
            tb.Text    = u'{} ({})'.format(pn, kind) if kind else pn
            cb.Content = tb
            cb.Margin  = SW.Thickness(2, 2, 2, 2)
            sp.Children.Add(cb)
            checkboxes.append((cb, pn))

        win.Content = outer

        def _ok(s, a):
            win.Close()

        btn_ok.Click += _ok

        win.ShowDialog()  # blocks until closed; Owner=self keeps it in front

        added = False
        for cb, pn in checkboxes:
            if cb.IsChecked:
                kind = self._adv_param_cache.get(pn, u'') if isinstance(self._adv_param_cache, dict) else u''
                hdr  = u'{} ({})'.format(pn, kind) if kind else pn
                self._adv_columns.append((hdr, pn))
                added = True
        if added and self._adv_open:
            self._load_advanced_data()

    # ── Excel (Phase 15) ─────────────────────────────────────────────

    def _update_import_btn_state(self):
        """Import Excel button: blue (empty table) / red 'Reset Table' (has data)."""
        if self._import_preview_mode:
            return  # Cancel Preview state is managed by _set_preview_locked(True)
        has_data = (self._adv_rows is not None
                    and self._adv_rows.Rows.Count > 0)
        try:
            if has_data:
                self.btn_import_excel.Content = u'Reset Table'
                _red = SolidColorBrush(
                    ColorConverter.ConvertFromString(u'#C0392B'))
                self.btn_import_excel.Background  = _red
                self.btn_import_excel.BorderBrush = _red
            else:
                self.btn_import_excel.Content = u'Import Excel'
                self.btn_import_excel.SetResourceReference(
                    Button.BackgroundProperty, u'AccentBrush')
                self.btn_import_excel.SetResourceReference(
                    Button.BorderBrushProperty, u'AccentBrush')
        except Exception:
            pass

    def _set_preview_locked(self, locked):
        """Enable/disable controls that must be frozen during import preview."""
        self._import_preview_mode = locked
        for ctl in [self.btn_scope_checked, self.btn_scope_all,
                    self.cb_active_view, self.btn_add_column,
                    self.btn_qb_add_condition]:
            try:
                ctl.IsEnabled = not locked
            except Exception:
                pass
        self.btn_overwrite_excel.IsEnabled = locked
        try:
            if locked:
                col = SolidColorBrush(ColorConverter.ConvertFromString(u'#E67E22'))
            else:
                col = SolidColorBrush(ColorConverter.ConvertFromString(u'#9E9E9E'))
            self.btn_overwrite_excel.Background  = col
            self.btn_overwrite_excel.BorderBrush = col
        except Exception:
            pass
        try:
            if locked:
                self.btn_import_excel.Content = u'Cancel Preview'
                _red = SolidColorBrush(
                    ColorConverter.ConvertFromString(u'#C0392B'))
                self.btn_import_excel.Background  = _red
                self.btn_import_excel.BorderBrush = _red
            else:
                # Delegate to tri-state updater (Import Excel / Reset Table)
                self._update_import_btn_state()
        except Exception:
            pass

    def _show_import_preview(self, preview_data):
        """Build and display the import-preview table (Excel values, green/red flags)."""
        from advanced_panel import build_preview_table
        from Autodesk.Revit.DB import ElementId

        # Add custom param columns from the preview data to _adv_columns
        _builtin_hdrs = {u'Family Name', u'Type Name'}
        existing_keys = {key for _, key in self._adv_columns}
        if preview_data:
            for pn in preview_data[0].get(u'params', {}).keys():
                if pn not in _builtin_hdrs and pn not in existing_keys:
                    kind = self._adv_param_cache.get(pn, u'') if isinstance(self._adv_param_cache, dict) else u''
                    hdr  = u'{} ({})'.format(pn, kind) if kind else pn
                    self._adv_columns.append((hdr, pn))
                    existing_keys.add(pn)

        # Build DataTable from preview data (Excel values, not Revit)
        self._adv_rows = build_preview_table(preview_data, self._adv_columns)

        # Switch scope to imported element IDs
        import_ids = {}
        for row in preview_data:
            try:
                eid_int = long(row[u'id'])
                import_ids[eid_int] = ElementId(eid_int)
            except Exception:
                pass
        if import_ids:
            self._adv_scope   = u'checked'
            self._checked_ids = import_ids
            self.btn_scope_checked.IsChecked = True
            self.btn_scope_all.IsChecked     = False

        # Open Advanced panel if closed
        if not self._adv_open:
            self._adv_open = True
            self.adv_panel.Visibility     = Visibility.Visible
            self.grid_splitter.Visibility = Visibility.Visible
            self.btn_toggle_adv.Content   = u'Close Advanced ◄'
            self._animate_adv(True)

        # Render the DataGrid and lock scope/column controls
        self._rebuild_datagrid()
        count = self._adv_rows.Rows.Count
        self.txt_table_count.Text = unicode(count)
        self._set_preview_locked(True)

        # Count green / red cells for status text
        n_green = sum(
            1 for row in preview_data
            for pv in row.get(u'preview', {}).values()
            if pv.get(u'writable', True) and pv.get(u'changed', False)
        )
        n_red = sum(
            1 for row in preview_data
            for pv in row.get(u'preview', {}).values()
            if not pv.get(u'writable', True)
        )
        self.txt_adv_scope.Text = (
            u'Preview: {} rows - {} cells to update (green), {} read-only (red)'
            u' - click Overwrite Data to apply'.format(count, n_green, n_red)
        )

    def _apply_import_failures(self, dt):
        """Mark _f_{key} = '1' in DataTable rows where import could not set param."""
        if not self._import_failures or dt is None:
            return
        for dr in dt.Rows:
            try:
                row_id = unicode(dr[u'id'])
            except Exception:
                continue
            failed = self._import_failures.get(row_id)
            if not failed:
                continue
            for pn in failed:
                try:
                    dr[u'_f_' + pn] = u'1'
                except Exception:
                    pass   # column may not exist if param not in current columns

    def _after_import(self, result):
        """Called on WPF thread via Dispatcher.Invoke after Transaction completes."""
        # Exit preview mode, unlock controls, disable Overwrite button
        self._set_preview_locked(False)
        self._import_failures = result.get(u'failures', {})

        # Reload table from Revit (actual post-transaction values)
        self._load_advanced_data()

        updated = result.get(u'updated', 0)
        skipped = result.get(u'skipped', 0)
        errors  = result.get(u'errors', [])
        fails   = self._import_failures
        fails_n = sum(len(v) for v in fails.itervalues()) if fails else 0

        if updated == 0:
            scope_text = (
                u'Import complete - 0 parameters updated. '
                u'{} skipped (read-only or not found).'.format(skipped)
            )
            msg = (
                u'No parameters were updated.\n'
                u'Skipped: {} (read-only or not found in model).\n\n'
                u'Red cells in the preview mark read-only parameters.'.format(skipped)
            )
        else:
            scope_text = u'Import complete - {} updated, {} skipped.'.format(
                updated, skipped)
            msg = u'Updated: {} parameter value(s).'.format(updated)
            if fails_n:
                msg += u'\nCould not set: {} cell(s) (red, read-only).'.format(fails_n)
            msg += u'\nSkipped: {}'.format(skipped)

        if errors:
            scope_text += u' ({} errors)'.format(len(errors))
            msg += u'\n\nErrors ({}):\n'.format(len(errors)) + u'\n'.join(errors[:5])

        try:
            self.txt_adv_scope.Text = scope_text
        except Exception:
            pass
        from System.Windows import MessageBox
        MessageBox.Show(self, msg, u'FilterMan - Import Complete')

    def _export_excel_click(self, sender, args):
        from System.Windows import MessageBox
        if self._adv_rows is None or self._adv_rows.Rows.Count == 0:
            MessageBox.Show(self,
                            u'No data to export. '
                            u'Open the Advanced panel and load elements first.',
                            u'FilterMan - Export')
            return
        try:
            from Microsoft.Win32 import SaveFileDialog
            dlg            = SaveFileDialog()
            dlg.Title      = u'Export to Excel'
            dlg.Filter     = u'Excel Workbook (*.xlsx)|*.xlsx'
            dlg.DefaultExt = u'.xlsx'
            dlg.FileName   = u'FilterMan_Export'
            ok = dlg.ShowDialog(self)  # owner = FilterMan; dialog appears in front
            if ok != True:
                return
            filepath = dlg.FileName
            from excel_handler import export_to_excel
            export_to_excel(self._adv_rows, self._adv_columns, filepath)
            MessageBox.Show(self,
                            u'Exported {} rows to:\n{}'.format(
                                self._adv_rows.Rows.Count, filepath),
                            u'FilterMan - Export')
        except Exception as ex:
            MessageBox.Show(self, unicode(ex), u'FilterMan - Export Error')

    def _overwrite_click(self, sender, args):
        """Apply the pending import data to the Revit model (Transaction)."""
        import_data = getattr(self, u'_last_import_data', [])
        if not import_data:
            return
        import view_actions
        self._import_failures = {}
        # Keep preview locked until _after_import is called via Dispatcher.Invoke
        view_actions.apply_excel_import(import_data, window=self)

    def _import_excel_click(self, sender, args):
        from System.Windows import MessageBox
        try:
            # State 3: PREVIEW → Cancel Preview
            if self._import_preview_mode:
                self._set_preview_locked(False)  # calls _update_import_btn_state
                self._last_import_data = []
                self._adv_columns = [
                    (u'Family Name', u'family_name'),
                    (u'Type Name',   u'type_name'),
                ]
                self._adv_rows = None
                self._rebuild_datagrid()
                self.txt_table_count.Text = u'0'
                try:
                    self.txt_adv_scope.Text = u'Table cleared - ready for new import'
                except Exception:
                    pass
                self._update_import_btn_state()
                return

            # State 2: HAS DATA → Reset Table
            if (self._adv_rows is not None
                    and self._adv_rows.Rows.Count > 0):
                self._last_import_data = []
                self._adv_columns = [
                    (u'Family Name', u'family_name'),
                    (u'Type Name',   u'type_name'),
                ]
                self._adv_rows = None
                self._adv_conditions = []
                try:
                    self.qb_conditions_panel.Children.Clear()
                except Exception:
                    pass
                self._adv_scope = u'checked'
                try:
                    self.btn_scope_checked.IsChecked = True
                    self.btn_scope_all.IsChecked     = False
                except Exception:
                    pass
                self._rebuild_datagrid()
                self.txt_table_count.Text = u'0'
                try:
                    self.txt_adv_scope.Text = u'Table cleared'
                except Exception:
                    pass
                self._update_import_btn_state()
                return

            # State 1: EMPTY → Import Excel
            from Microsoft.Win32 import OpenFileDialog
            dlg            = OpenFileDialog()
            dlg.Title      = u'Import from Excel'
            dlg.Filter     = u'Excel Workbook (*.xlsx)|*.xlsx'
            dlg.DefaultExt = u'.xlsx'
            ok = dlg.ShowDialog(self)  # owner = FilterMan; dialog appears in front
            if ok != True:
                return
            filepath = dlg.FileName
            from excel_handler import import_from_excel
            import_data = import_from_excel(filepath)
            if not import_data:
                MessageBox.Show(self,
                                u'No data rows found in the file.',
                                u'FilterMan - Import')
                return
            import view_actions
            self._import_failures  = {}
            self._last_import_data = import_data
            view_actions.apply_excel_preview(import_data, window=self)
        except Exception as ex:
            MessageBox.Show(self, unicode(ex), u'FilterMan - Import Error')

    # ──────────────────────────────────────────────────────────────────
    # THEME
    # ──────────────────────────────────────────────────────────────────

    def _toggle_theme(self, sender, args):
        self._dark_mode = not self._dark_mode
        palette = _DARK_PALETTE if self._dark_mode else _LIGHT_PALETTE
        for key, hex_c in palette.iteritems():
            self.Resources[key] = SolidColorBrush(
                ColorConverter.ConvertFromString(hex_c)
            )
        self._apply_tree_theme()
        self._apply_titlebar_theme()
        try:
            from settings import save_settings
            save_settings({u'dark_mode': self._dark_mode})
        except Exception:
            pass

    def _apply_titlebar_theme(self):
        """Set Windows native title bar dark/light via DwmSetWindowAttribute."""
        try:
            import ctypes
            from System.Windows.Interop import WindowInteropHelper
            hwnd = WindowInteropHelper(self).Handle.ToInt32()
            val  = ctypes.c_int(1 if self._dark_mode else 0)
            # DWMWA_USE_IMMERSIVE_DARK_MODE: 20 (Win10 20H1+), 19 (Win10 1903-19H2)
            for attr in (20, 19):
                try:
                    r = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, attr, ctypes.byref(val), ctypes.sizeof(val))
                    if r == 0:
                        break
                except Exception:
                    pass
        except Exception:
            pass

    def _apply_tree_theme(self):
        # Cat/fam borders use SetResourceReference and update automatically.
        # Only instance backgrounds need explicit refresh (state-dependent).
        for node in self._all_nodes:
            if node.level == u'instance':
                self._update_instance_bg(node)

    # ──────────────────────────────────────────────────────────────────
    # SHIMMER (Phase 16 - stub)
    # ──────────────────────────────────────────────────────────────────

    def _start_shimmer(self):
        """Moving accent-coloured shimmer on canvas_shimmer.
        Uses TranslateTransform + BeginAnimation - runs on WPF compositor thread,
        continues animating even while Python (UI thread) is building the tree."""
        try:
            from System.Windows.Shapes import Rectangle
            from System.Windows.Controls import Canvas
            from System.Windows.Media import (
                LinearGradientBrush, GradientStop, Color, TranslateTransform
            )
            from System.Windows.Media.Animation import DoubleAnimation, RepeatBehavior
            from System.Windows import Duration, Point
            from System import TimeSpan

            _SHIMMER_W   = 120
            _DURATION_MS = 1200

            # Pick accent colour from current theme palette
            try:
                accent = self.Resources[u'AccentBrush'].Color
                r, g, b = int(accent.R), int(accent.G), int(accent.B)
            except Exception:
                r, g, b = 0x06, 0x96, 0xD7

            rect        = Rectangle()
            rect.Width  = float(_SHIMMER_W)
            rect.Height = 4.0
            Canvas.SetLeft(rect, 0.0)
            Canvas.SetTop(rect, 0.0)

            brush = LinearGradientBrush()
            brush.StartPoint = Point(0.0, 0.5)
            brush.EndPoint   = Point(1.0, 0.5)
            for alpha, offset in [(0, 0.0), (220, 0.5), (0, 1.0)]:
                stop        = GradientStop()
                stop.Color  = Color.FromArgb(alpha, r, g, b)
                stop.Offset = offset
                brush.GradientStops.Add(stop)
            rect.Fill = brush

            transform = TranslateTransform()
            rect.RenderTransform = transform

            anim = DoubleAnimation()
            anim.From           = float(-_SHIMMER_W)
            anim.To             = 1500.0   # wider than any window; ClipToBounds clips
            anim.Duration       = Duration(TimeSpan.FromMilliseconds(_DURATION_MS))
            anim.RepeatBehavior = RepeatBehavior.Forever
            transform.BeginAnimation(TranslateTransform.XProperty, anim)

            self.canvas_shimmer.Children.Clear()
            self.canvas_shimmer.Children.Add(rect)
            self._shimmer_rect = rect
        except Exception:
            pass

    def _stop_shimmer(self):
        """Stop shimmer animation and clear the canvas."""
        try:
            rect = getattr(self, u'_shimmer_rect', None)
            if rect is not None:
                try:
                    from System.Windows.Media import TranslateTransform
                    t = rect.RenderTransform
                    if t is not None:
                        t.BeginAnimation(TranslateTransform.XProperty, None)
                except Exception:
                    pass
                try:
                    self.canvas_shimmer.Children.Remove(rect)
                except Exception:
                    pass
                self._shimmer_rect = None
        except Exception:
            pass

    # ──────────────────────────────────────────────────────────────────
    # LIFECYCLE
    # ──────────────────────────────────────────────────────────────────

    def _on_source_initialized(self, sender, args):
        """Apply title bar theme immediately when the HWND is first created.
        SourceInitialized fires before Show(), so the title bar is dark
        from the first paint - no visible flash."""
        self._apply_titlebar_theme()

    def _on_closing(self, sender, args):
        self._stop_shimmer()
        for _t in (self._poll_timer, self._debounce_timer,
                   self._adv_anim_timer,
                   getattr(self, u'_defer_build_timer', None),
                   getattr(self, u'_defer_splash_timer', None)):
            if _t is not None:
                try:
                    _t.Stop()
                except Exception:
                    pass


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  HELPERS                                                             ║
# ╚══════════════════════════════════════════════════════════════════════╝

def _hand_cursor():
    from System.Windows.Input import Cursors
    return Cursors.Hand


def _safe_sym_name(sym):
    """Return ElementType name, falling back to BuiltInParameter when
    sym.Name is write-only (raises AttributeError in IronPython 2.7 for
    certain Revit builds where the getter is not exposed)."""
    try:
        n = sym.Name
        return n if n else u'(unnamed type)'
    except AttributeError:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter
        for _bip in (BuiltInParameter.SYMBOL_NAME_PARAM,
                     BuiltInParameter.ALL_MODEL_TYPE_NAME):
            try:
                p = sym.get_Parameter(_bip)
                if p is not None:
                    n = p.AsString()
                    if n:
                        return n
            except Exception:
                pass
    except Exception:
        pass
    return u'(unnamed type)'


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  THEME PALETTES                                                      ║
# ╚══════════════════════════════════════════════════════════════════════╝

_LIGHT_PALETTE = {
    u'BgBrush':          u'#F0F0F0',
    u'SurfaceBrush':     u'#FFFFFF',
    u'Surface2Brush':    u'#F5F7FA',
    u'Surface3Brush':    u'#E8E8E8',
    u'BorderBrush':      u'#D0D0D0',
    u'Border2Brush':     u'#B0B0B0',
    u'TextBrush':        u'#1A1A1A',
    u'Text2Brush':       u'#555555',
    u'Text3Brush':       u'#999999',
    u'AccentBrush':      u'#0696D7',
    u'AccentHoverBrush': u'#047DB8',
    u'AccentLiteBrush':  u'#E8F4FB',
    u'DangerBrush':      u'#C0392B',
    u'OkBrush':          u'#27AE60',
    u'CkBgBrush':        u'#C8F0D8',
    u'RvBgBrush':        u'#D0E8F8',
    u'CatBgBrush':       u'#E8EEF4',
    u'FamBgBrush':       u'#F5F7FA',
}

_DARK_PALETTE = {
    u'BgBrush':          u'#1C1C1C',
    u'SurfaceBrush':     u'#282828',
    u'Surface2Brush':    u'#232323',
    u'Surface3Brush':    u'#202020',
    u'BorderBrush':      u'#3A3A3A',
    u'Border2Brush':     u'#555555',
    u'TextBrush':        u'#E0E0E0',
    u'Text2Brush':       u'#999999',
    u'Text3Brush':       u'#666666',
    u'AccentBrush':      u'#1BA0E8',
    u'AccentHoverBrush': u'#3BB0F5',
    u'AccentLiteBrush':  u'#1A3548',
    u'DangerBrush':      u'#E04040',
    u'OkBrush':          u'#2ECC71',
    u'CkBgBrush':        u'#1A3D28',
    u'RvBgBrush':        u'#1A2E40',
    u'CatBgBrush':       u'#263040',
    u'FamBgBrush':       u'#272C32',
}
