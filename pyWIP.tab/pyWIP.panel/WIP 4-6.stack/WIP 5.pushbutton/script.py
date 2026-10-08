# -*- coding: utf-8 -*-
"""
pyRevit: Copy Sheets from Linked Revit Models
SAFE VERSION - Floor Plan / Section Viewport Copy

Important fix in this version:
- This script DOES NOT copy visible model elements from Floor Plans or Sections.
- The previous broad view-content copy could duplicate model elements such as Doors,
  Windows, Generic Models, Room Separation Lines, Walls, etc.
- For model views, this script creates/matches the view and places the viewport only.
- View-owned annotation copy is restricted and disabled for model views by default.

Supported viewport view types:
- FloorPlan
- CeilingPlan
- EngineeringPlan
- AreaPlan
- Section

Recommended use:
- Run on detached/test model first.
- Keep "Update Existing View Contents" OFF unless you understand the consequences.
"""

from pyrevit.framework import List
from pyrevit import forms
from pyrevit import revit, DB
from pyrevit.revit import query
from pyrevit import script
from Autodesk.Revit.DB import Element as DBElement

logger = script.get_logger()
output = script.get_output()

OPTION_SET = None


# -----------------------------------------------------------------------------
# OPTIONS / UI LIST ITEMS
# -----------------------------------------------------------------------------

class Option(forms.TemplateListItem):
    def __init__(self, option_name, default_state=False):
        super(Option, self).__init__(option_name)
        self.option_name = option_name
        self.state = default_state

    @property
    def name(self):
        return self.option_name


class LinkedDocOption(forms.TemplateListItem):
    def __init__(self, link_doc, link_name):
        # Intentionally wrap the document. Some pyRevit versions return the wrapped
        # object directly, so get_source_link_docs() handles both cases.
        super(LinkedDocOption, self).__init__(link_doc)
        self.link_doc = link_doc
        self.link_name = link_name
        self.state = False

    @property
    def name(self):
        return self.link_name


class SheetOption(forms.TemplateListItem):
    def __init__(self, sheet):
        # Intentionally wrap the sheet. Some pyRevit versions return the wrapped
        # object directly, so get_source_sheets() handles both cases.
        super(SheetOption, self).__init__(sheet)
        self.sheet = sheet
        self.state = False

    @property
    def name(self):
        try:
            return "{} - {}".format(self.sheet.SheetNumber, self.sheet.Name)
        except Exception:
            return safe_name(self.sheet)


class OptionSet(object):
    def __init__(self):
        self.op_copy_vports = Option("Copy Viewports", True)
        self.op_copy_titleblock = Option("Copy Sheet Titleblock", True)
        self.op_copy_sheet_properties = Option("Copy Sheet Properties", True)
        self.op_update_exist_view_contents = Option(
            "Update Existing View Contents - ANNOTATIONS ONLY",
            False
        )
        self.op_preserve_detail_numbers = Option("Preserve Detail Numbers", True)
        self.op_copy_annotations_non_model_views = Option(
            "Copy Annotations for NON-MODEL Views Only",
            False
        )

    def all_options(self):
        return [
            self.op_copy_vports,
            self.op_copy_titleblock,
            self.op_copy_sheet_properties,
            self.op_update_exist_view_contents,
            self.op_preserve_detail_numbers,
            self.op_copy_annotations_non_model_views,
        ]


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


# -----------------------------------------------------------------------------
# BASIC HELPERS
# -----------------------------------------------------------------------------

def make_element_id_list(element_ids):
    """Convert Python list of ElementIds into .NET List[DB.ElementId]."""
    id_list = List
    for eid in element_ids:
        try:
            if isinstance(eid, DB.ElementId):
                id_list.Add(eid)
        except Exception:
            pass
    return id_list


def safe_name(element):
    """Safely get a Revit element name."""
    if element is None:
        return ""
    try:
        return query.get_name(element)
    except Exception:
        try:
            return DBElement.Name.GetValue(element)
        except Exception:
            try:
                return element.Name
            except Exception:
                return ""


def get_param_value(param):
    if not param:
        return None
    try:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
        if param.StorageType == DB.StorageType.Integer:
            return param.AsInteger()
        if param.StorageType == DB.StorageType.Double:
            return param.AsDouble()
        if param.StorageType == DB.StorageType.ElementId:
            return param.AsElementId()
    except Exception:
        return None
    return None


def set_param_value(param, value):
    if not param or param.IsReadOnly or value is None:
        return False
    try:
        if param.StorageType == DB.StorageType.String:
            param.Set(str(value))
            return True
        if param.StorageType == DB.StorageType.Integer:
            param.Set(int(value))
            return True
        if param.StorageType == DB.StorageType.Double:
            param.Set(float(value))
            return True
        if param.StorageType == DB.StorageType.ElementId:
            param.Set(value)
            return True
    except Exception:
        return False
    return False


def copy_builtin_param(src_elem, dst_elem, bip):
    try:
        sp = src_elem.get_Parameter(bip)
        dp = dst_elem.get_Parameter(bip)
        if sp and dp and not dp.IsReadOnly:
            return set_param_value(dp, get_param_value(sp))
    except Exception:
        pass
    return False


def is_invalid_element_id(eid):
    try:
        return eid == DB.ElementId.InvalidElementId
    except Exception:
        return True


def get_all_views(doc, include_templates=False):
    views = []
    try:
        for view in DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements():
            try:
                if include_templates or not view.IsTemplate:
                    views.append(view)
            except Exception:
                pass
    except Exception:
        pass
    return views


def get_all_levels(doc):
    try:
        return list(
            DB.FilteredElementCollector(doc)
            .OfClass(DB.Level)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []


def unique_view_name(doc, desired_name):
    existing = set([safe_name(v) for v in get_all_views(doc, True)])
    if desired_name not in existing:
        return desired_name

    i = 1
    while True:
        new_name = "{} - Copy {}".format(desired_name, i)
        if new_name not in existing:
            return new_name
        i += 1


def unique_sheet_number(doc, desired_number):
    existing = set()
    try:
        for sheet in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements():
            existing.add(sheet.SheetNumber)
    except Exception:
        pass

    if desired_number not in existing:
        return desired_number

    i = 1
    while True:
        new_number = "{}-COPY{}".format(desired_number, i)
        if new_number not in existing:
            return new_number
        i += 1


def is_model_view(view):
    """Model views where visible elements must NOT be copied."""
    try:
        return view.ViewType in [
            DB.ViewType.FloorPlan,
            DB.ViewType.CeilingPlan,
            DB.ViewType.EngineeringPlan,
            DB.ViewType.AreaPlan,
            DB.ViewType.Section,
            DB.ViewType.Elevation,
            DB.ViewType.ThreeD,
        ]
    except Exception:
        return True


def is_supported_viewport_view(view):
    try:
        return view.ViewType in [
            DB.ViewType.FloorPlan,
            DB.ViewType.CeilingPlan,
            DB.ViewType.EngineeringPlan,
            DB.ViewType.AreaPlan,
            DB.ViewType.Section,
        ]
    except Exception:
        return False


def get_bic_int(bic):
    try:
        return int(bic)
    except Exception:
        try:
            return bic.IntegerValue
        except Exception:
            return None


# -----------------------------------------------------------------------------
# SELECTION
# -----------------------------------------------------------------------------

def get_user_options():
    global OPTION_SET

    op_set = OptionSet()
    options = op_set.all_options()

    selected = forms.SelectFromList.show(
        options,
        title="Sheet Copy Options - SAFE VERSION",
        multiselect=True,
        button_name="Run Copy"
    )

    if selected is None:
        forms.alert("Copy cancelled.", exitscript=True)

    selected_names = set()
    for x in selected:
        if hasattr(x, "name"):
            selected_names.add(x.name)
        else:
            selected_names.add(str(x))

    for opt in options:
        opt.state = opt.name in selected_names

    OPTION_SET = op_set
    return op_set


def get_source_link_docs():
    """Select loaded Revit links from active document."""
    dest_doc = revit.doc
    link_options = []
    seen = set()

    try:
        links = (
            DB.FilteredElementCollector(dest_doc)
            .OfClass(DB.RevitLinkInstance)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        links = []

    for link in links:
        try:
            link_doc = link.GetLinkDocument()
            if not link_doc:
                continue

            key = link_doc.PathName if link_doc.PathName else link_doc.Title
            if key in seen:
                continue
            seen.add(key)

            link_name = "{} :: {}".format(safe_name(link), link_doc.Title)
            link_options.append(LinkedDocOption(link_doc, link_name))
        except Exception:
            pass

    if not link_options:
        forms.alert("No loaded Revit links found in the active document.", exitscript=True)

    selected = forms.SelectFromList.show(
        sorted(link_options, key=lambda x: x.name),
        title="Select Source Linked Documents",
        multiselect=True,
        button_name="Select Links"
    )

    if not selected:
        forms.alert("No source links selected.", exitscript=True)

    # pyRevit can return wrapper or wrapped Document. Handle both.
    source_docs = []
    for x in selected:
        if hasattr(x, "link_doc"):
            source_docs.append(x.link_doc)
        else:
            source_docs.append(x)

    return source_docs


def get_source_sheets(source_doc):
    """Select sheets from selected linked source document."""
    sheet_options = []

    try:
        sheets = (
            DB.FilteredElementCollector(source_doc)
            .OfClass(DB.ViewSheet)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        sheets = []

    for sheet in sheets:
        try:
            sheet_options.append(SheetOption(sheet))
        except Exception:
            pass

    if not sheet_options:
        forms.alert("No sheets found in linked document: {}".format(source_doc.Title), exitscript=False)
        return []

    selected = forms.SelectFromList.show(
        sorted(sheet_options, key=lambda x: x.name),
        title="Select Sheets from {}".format(source_doc.Title),
        multiselect=True,
        button_name="Select Sheets"
    )

    if not selected:
        return []

    # pyRevit can return wrapper or wrapped ViewSheet. Handle both.
    selected_sheets = []
    for x in selected:
        if hasattr(x, "sheet"):
            selected_sheets.append(x.sheet)
        else:
            selected_sheets.append(x)

    return selected_sheets


# -----------------------------------------------------------------------------
# VIEW MATCHING / CREATION
# -----------------------------------------------------------------------------

def find_matching_view(dest_doc, source_view):
    """Find matching destination view by exact name/type, then view description/type."""
    source_name = safe_name(source_view)
    source_type = source_view.ViewType

    for view in get_all_views(dest_doc):
        try:
            if safe_name(view) == source_name and view.ViewType == source_type:
                return view
        except Exception:
            pass

    try:
        src_desc_param = source_view.get_Parameter(DB.BuiltInParameter.VIEW_DESCRIPTION)
        src_desc = src_desc_param.AsString() if src_desc_param else None
        if src_desc:
            for view in get_all_views(dest_doc):
                try:
                    dst_desc_param = view.get_Parameter(DB.BuiltInParameter.VIEW_DESCRIPTION)
                    dst_desc = dst_desc_param.AsString() if dst_desc_param else None
                    if dst_desc == src_desc and view.ViewType == source_type:
                        return view
                except Exception:
                    pass
    except Exception:
        pass

    return None


def find_level_by_name_or_elevation(dest_doc, source_level, tolerance=0.01):
    if not source_level:
        return None

    source_name = safe_name(source_level)
    try:
        source_elev = source_level.Elevation
    except Exception:
        source_elev = None

    for lvl in get_all_levels(dest_doc):
        if safe_name(lvl) == source_name:
            return lvl

    if source_elev is not None:
        for lvl in get_all_levels(dest_doc):
            try:
                if abs(lvl.Elevation - source_elev) <= tolerance:
                    return lvl
            except Exception:
                pass

    return None


def find_view_family_type(doc, view_family):
    try:
        vfts = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).ToElements()
        for vft in vfts:
            try:
                if vft.ViewFamily == view_family:
                    return vft
            except Exception:
                pass
    except Exception:
        pass
    return None


def find_view_family_type_by_name_or_family(dest_doc, source_vft):
    if not source_vft:
        return None

    source_name = safe_name(source_vft)
    try:
        source_family = source_vft.ViewFamily
    except Exception:
        source_family = None

    try:
        vfts = DB.FilteredElementCollector(dest_doc).OfClass(DB.ViewFamilyType).ToElements()
    except Exception:
        vfts = []

    for vft in vfts:
        if safe_name(vft) == source_name:
            return vft

    if source_family:
        for vft in vfts:
            try:
                if vft.ViewFamily == source_family:
                    return vft
            except Exception:
                pass

    return None


def copy_view_props(source_view, dest_view):
    """Copy safe view settings only. Does not copy model elements."""
    try:
        dest_view.Scale = source_view.Scale
    except Exception:
        pass

    try:
        dest_view.DetailLevel = source_view.DetailLevel
    except Exception:
        pass

    try:
        dest_view.DisplayStyle = source_view.DisplayStyle
    except Exception:
        pass

    try:
        dest_view.Discipline = source_view.Discipline
    except Exception:
        pass

    try:
        dest_view.CropBoxActive = source_view.CropBoxActive
    except Exception:
        pass

    try:
        dest_view.CropBoxVisible = source_view.CropBoxVisible
    except Exception:
        pass

    try:
        if source_view.CropBox:
            dest_view.CropBox = source_view.CropBox
    except Exception:
        pass

    # Safe-ish built-in parameters. Invalid/not-writable parameters are ignored.
    possible_bips = [
        DB.BuiltInParameter.VIEW_DESCRIPTION,
        DB.BuiltInParameter.VIEWER_CROP_REGION_VISIBLE,
        DB.BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE,
        DB.BuiltInParameter.VIEW_PHASE,
        DB.BuiltInParameter.VIEW_PHASE_FILTER,
    ]

    for bip in possible_bips:
        try:
            copy_builtin_param(source_view, dest_view, bip)
        except Exception:
            pass

    # Apply matching view template by name only if it exists in destination.
    try:
        if source_view.ViewTemplateId and source_view.ViewTemplateId != DB.ElementId.InvalidElementId:
            src_template = source_view.Document.GetElement(source_view.ViewTemplateId)
            src_template_name = safe_name(src_template)
            for view in get_all_views(dest_view.Document, include_templates=True):
                try:
                    if view.IsTemplate and safe_name(view) == src_template_name:
                        dest_view.ViewTemplateId = view.Id
                        break
                except Exception:
                    pass
    except Exception:
        pass


# -----------------------------------------------------------------------------
# SAFE VIEW CONTENT COPY
# -----------------------------------------------------------------------------

def get_safe_annotation_category_ids():
    """
    Allowed categories for view-owned annotation/detail copy.
    BuiltInCategory names vary by Revit version, so all are resolved defensively.
    """
    category_names = [
        "OST_TextNotes",
        "OST_Lines",              # Detail lines if OwnerViewId == view.Id
        "OST_DetailComponents",
        "OST_GenericAnnotation",
        "OST_Dimensions",
        "OST_FilledRegion",
        "OST_RevisionClouds",
        "OST_DoorTags",
        "OST_WindowTags",
        "OST_RoomTags",
        "OST_AreaTags",
        "OST_WallTags",
        "OST_MaterialTags",
        "OST_MultiCategoryTags",
        "OST_KeynoteTags",
    ]

    ids = set()
    for name in category_names:
        try:
            bic = getattr(DB.BuiltInCategory, name)
            value = get_bic_int(bic)
            if value is not None:
                ids.add(value)
        except Exception:
            pass
    return ids


def get_view_contents(doc, view):
    """
    SAFE collector: view-owned annotations/details only.

    Critical protection:
    - Requires elem.OwnerViewId == view.Id.
    - Restricts categories to annotation/detail categories only.
    - This prevents copying visible model elements such as Doors, Windows,
      Walls, Floors, Generic Models, Room Separation Lines, etc.
    """
    element_ids = []
    allowed_cat_ids = get_safe_annotation_category_ids()

    try:
        collector = (
            DB.FilteredElementCollector(doc, view.Id)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return element_ids

    for elem in collector:
        try:
            # The most important safety check.
            try:
                if elem.OwnerViewId != view.Id:
                    continue
            except Exception:
                continue

            if elem.Id == view.Id:
                continue

            if not elem.Category:
                continue

            cat_id = elem.Category.Id.IntegerValue
            if cat_id not in allowed_cat_ids:
                continue

            element_ids.append(elem.Id)
        except Exception:
            pass

    return element_ids


def clear_view_contents(dest_doc, dest_view):
    """Delete only safe view-owned annotations/details from destination view."""
    ids = get_view_contents(dest_doc, dest_view)
    if not ids:
        return

    try:
        dest_doc.Delete(make_element_id_list(ids))
        print("\t\tCleared {} annotation/detail element(s).".format(len(ids)))
    except Exception as ex:
        print("\t\tCould not clear annotation/detail contents: {}".format(ex))


def copy_view_contents(source_doc, source_view, dest_doc, dest_view, clear_contents=False):
    """
    Safe annotation/detail copy.

    For Floor Plans, Sections, Ceiling Plans, Area Plans and Engineering Plans,
    content copy is intentionally skipped to avoid model-element duplication.
    """
    if is_model_view(source_view):
        print("\t\tSAFE MODE: Skipping view content copy for model view '{}'.".format(safe_name(source_view)))
        print("\t\tReason: prevents duplicate Doors/Windows/Generic Models/Walls/etc.")
        return []

    if not OPTION_SET or not OPTION_SET.op_copy_annotations_non_model_views.state:
        print("\t\tAnnotation copy for non-model views is OFF.")
        return []

    if clear_contents:
        clear_view_contents(dest_doc, dest_view)

    ids = get_view_contents(source_doc, source_view)
    if not ids:
        print("\t\tNo safe annotation/detail contents found in view: {}".format(safe_name(source_view)))
        return []

    try:
        options = DB.CopyPasteOptions()
        options.SetDuplicateTypeNamesHandler(CopyUseDestination())

        copied = DB.ElementTransformUtils.CopyElements(
            source_view,
            make_element_id_list(ids),
            dest_view,
            DB.Transform.Identity,
            options
        )

        copied_ids = list(copied)
        print("\t\tCopied {} safe annotation/detail element(s).".format(len(copied_ids)))
        return copied_ids
    except Exception as ex:
        print("\t\tCould not copy safe annotation/detail contents from '{}': {}".format(
            safe_name(source_view),
            ex
        ))
        return []


def create_or_match_floor_plan(source_doc, source_view, dest_doc):
    existing = find_matching_view(dest_doc, source_view)
    if existing:
        print("\t\tUsing existing matching plan: {}".format(safe_name(existing)))
        copy_view_props(source_view, existing)
        # Do NOT copy model view contents.
        if OPTION_SET and OPTION_SET.op_update_exist_view_contents.state:
            copy_view_contents(source_doc, source_view, dest_doc, existing, clear_contents=False)
        return existing

    try:
        source_level = source_doc.GetElement(source_view.GenLevel.Id)
    except Exception:
        source_level = None

    dest_level = find_level_by_name_or_elevation(dest_doc, source_level)
    if not dest_level:
        print("\t\tNo matching level found for plan: {}".format(safe_name(source_view)))
        return None

    try:
        source_vft = source_doc.GetElement(source_view.GetTypeId())
    except Exception:
        source_vft = None

    dest_vft = find_view_family_type_by_name_or_family(dest_doc, source_vft)
    if not dest_vft:
        print("\t\tNo matching ViewFamilyType found for plan: {}".format(safe_name(source_view)))
        return None

    try:
        new_view = DB.ViewPlan.Create(dest_doc, dest_vft.Id, dest_level.Id)
        new_view.Name = unique_view_name(dest_doc, safe_name(source_view))
        copy_view_props(source_view, new_view)
        # Do NOT copy model view contents.
        print("\t\tCreated plan: {}".format(safe_name(new_view)))
        return new_view
    except Exception as ex:
        print("\t\tFailed to create plan '{}': {}".format(safe_name(source_view), ex))
        return None


def create_or_match_section(source_doc, source_view, dest_doc):
    existing = find_matching_view(dest_doc, source_view)
    if existing:
        print("\t\tUsing existing matching section: {}".format(safe_name(existing)))
        copy_view_props(source_view, existing)
        # Do NOT copy model view contents.
        if OPTION_SET and OPTION_SET.op_update_exist_view_contents.state:
            copy_view_contents(source_doc, source_view, dest_doc, existing, clear_contents=False)
        return existing

    try:
        source_vft = source_doc.GetElement(source_view.GetTypeId())
    except Exception:
        source_vft = None

    dest_vft = find_view_family_type_by_name_or_family(dest_doc, source_vft)
    if not dest_vft:
        dest_vft = find_view_family_type(dest_doc, DB.ViewFamily.Section)

    if not dest_vft:
        print("\t\tNo Section ViewFamilyType found in destination.")
        return None

    try:
        section_box = source_view.CropBox
        if not section_box:
            print("\t\tSource section has no crop box: {}".format(safe_name(source_view)))
            return None

        new_section = DB.ViewSection.CreateSection(dest_doc, dest_vft.Id, section_box)
        new_section.Name = unique_view_name(dest_doc, safe_name(source_view))
        copy_view_props(source_view, new_section)
        # Do NOT copy model view contents.
        print("\t\tCreated section: {}".format(safe_name(new_section)))
        return new_section
    except Exception as ex:
        print("\t\tFailed to create section '{}': {}".format(safe_name(source_view), ex))
        return None


def copy_view(source_doc, source_view, dest_doc):
    if not source_view:
        return None

    try:
        if source_view.IsTemplate:
            return None
    except Exception:
        pass

    print("\t\tProcessing view: {} [{}]".format(safe_name(source_view), source_view.ViewType))

    if source_view.ViewType in [
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.EngineeringPlan,
        DB.ViewType.AreaPlan,
    ]:
        return create_or_match_floor_plan(source_doc, source_view, dest_doc)

    if source_view.ViewType == DB.ViewType.Section:
        return create_or_match_section(source_doc, source_view, dest_doc)

    # For other view types, only use existing matching view. Do not create/copy model content.
    existing = find_matching_view(dest_doc, source_view)
    if existing:
        return existing

    print("\t\tSkipping unsupported view type: {} ({})".format(source_view.ViewType, safe_name(source_view)))
    return None


# -----------------------------------------------------------------------------
# VIEWPORT COPY HELPERS
# -----------------------------------------------------------------------------

def get_source_vport_data(doc, vport, sheet_view):
    data = {}
    try:
        data["center"] = vport.GetBoxCenter()
        data["bbox"] = vport.get_BoundingBox(sheet_view)
        if data["bbox"]:
            data["bbox_min"] = data["bbox"].Min
            data["bbox_max"] = data["bbox"].Max

        try:
            data["label_offset"] = vport.LabelOffset
        except Exception:
            data["label_offset"] = None

        try:
            data["label_line_length"] = vport.LabelLineLength
        except Exception:
            data["label_line_length"] = None

        try:
            data["type_id"] = vport.GetTypeId()
            data["type_name"] = safe_name(doc.GetElement(vport.GetTypeId()))
        except Exception:
            data["type_id"] = None
            data["type_name"] = None

        try:
            detail_param = vport.get_Parameter(DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
            data["detail_number"] = detail_param.AsString() if detail_param else None
        except Exception:
            data["detail_number"] = None
    except Exception as ex:
        print("\t\tCould not collect viewport data: {}".format(ex))
    return data


def apply_viewport_type(source_doc, source_vport, dest_doc, new_vport):
    try:
        source_type = source_doc.GetElement(source_vport.GetTypeId())
        source_type_name = safe_name(source_type)

        for type_id in new_vport.GetValidTypes():
            dest_type = dest_doc.GetElement(type_id)
            if safe_name(dest_type) == source_type_name:
                new_vport.ChangeTypeId(type_id)
                print("\t\t\tApplied viewport type: {}".format(source_type_name))
                return True

        print("\t\t\tMatching viewport type not found. Using default: {}".format(source_type_name))
    except Exception as ex:
        print("\t\t\tCould not apply viewport type: {}".format(ex))
    return False


def apply_detail_number(source_vport, new_vport):
    if OPTION_SET and not OPTION_SET.op_preserve_detail_numbers.state:
        return False

    try:
        src = source_vport.get_Parameter(DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        dst = new_vport.get_Parameter(DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        if src and dst and not dst.IsReadOnly:
            val = src.AsString()
            if val:
                dst.Set(val)
                print("\t\t\tPreserved detail number: {}".format(val))
                return True
    except Exception as ex:
        print("\t\t\tCould not preserve detail number: {}".format(ex))
    return False


def apply_vport_label_props(new_vport, source_data):
    try:
        label_offset = source_data.get("label_offset")
        label_line_len = source_data.get("label_line_length")

        if label_offset is not None:
            try:
                new_vport.LabelOffset = label_offset
            except Exception:
                pass

        if label_line_len is not None:
            try:
                new_vport.LabelLineLength = label_line_len
            except Exception:
                pass

        print("\t\t\tApplied viewport label properties.")
        return True
    except Exception as ex:
        print("\t\t\tCould not apply label properties: {}".format(ex))
    return False


def correct_vport_by_bbox(dest_doc, new_vport, source_data, dest_sheet):
    try:
        src_min = source_data.get("bbox_min")
        if not src_min:
            return False

        dst_bbox = new_vport.get_BoundingBox(dest_sheet)
        if not dst_bbox:
            return False

        dst_min = dst_bbox.Min
        move_vec = DB.XYZ(src_min.X - dst_min.X, src_min.Y - dst_min.Y, 0)

        if move_vec.GetLength() > 0.0001:
            DB.ElementTransformUtils.MoveElement(dest_doc, new_vport.Id, move_vec)
            print("\t\t\tCorrected viewport position by bbox.")
        return True
    except Exception as ex:
        print("\t\t\tCould not correct viewport position: {}".format(ex))
    return False


def is_view_already_on_sheet(doc, sheet, view_id):
    try:
        for vp_id in sheet.GetAllViewports():
            vp = doc.GetElement(vp_id)
            if vp and vp.ViewId == view_id:
                return True
    except Exception:
        pass
    return False


def is_view_already_placed_on_any_sheet(doc, view_id):
    try:
        vports = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.Viewport)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        for vp in vports:
            try:
                if vp.ViewId == view_id:
                    return True
            except Exception:
                pass
    except Exception:
        pass
    return False


def copy_sheet_viewports(source_doc, source_sheet, dest_doc, dest_sheet):
    print("\tCopying viewports from sheet: {}".format(safe_name(source_sheet)))

    try:
        source_vport_ids = list(source_sheet.GetAllViewports())
    except Exception as ex:
        print("\tCould not get source viewports: {}".format(ex))
        return []

    created = []

    for source_vport_id in source_vport_ids:
        try:
            source_vport = source_doc.GetElement(source_vport_id)
            source_view = source_doc.GetElement(source_vport.ViewId)

            if not source_view:
                continue

            if not is_supported_viewport_view(source_view):
                print("\t\tSkipping unsupported viewport view: {} [{}]".format(
                    safe_name(source_view),
                    source_view.ViewType
                ))
                continue

            source_data = get_source_vport_data(source_doc, source_vport, source_sheet)
            dest_view = copy_view(source_doc, source_view, dest_doc)

            if not dest_view:
                print("\t\tCould not create/find destination view for: {}".format(safe_name(source_view)))
                continue

            if is_view_already_on_sheet(dest_doc, dest_sheet, dest_view.Id):
                print("\t\tView already placed on this sheet: {}".format(safe_name(dest_view)))
                continue

            if is_view_already_placed_on_any_sheet(dest_doc, dest_view.Id):
                print("\t\tView already placed on another sheet. Skipping: {}".format(safe_name(dest_view)))
                continue

            center = source_data.get("center")
            if not center:
                print("\t\tNo valid viewport center. Skipping.")
                continue

            try:
                if hasattr(DB.Viewport, "CanAddViewToSheet"):
                    if not DB.Viewport.CanAddViewToSheet(dest_doc, dest_sheet.Id, dest_view.Id):
                        print("\t\tCannot add view to sheet: {}".format(safe_name(dest_view)))
                        continue
            except Exception:
                pass

            new_vport = DB.Viewport.Create(dest_doc, dest_sheet.Id, dest_view.Id, center)
            print("\t\tCreated viewport for: {}".format(safe_name(dest_view)))

            apply_viewport_type(source_doc, source_vport, dest_doc, new_vport)
            try:
                dest_doc.Regenerate()
            except Exception:
                pass

            apply_vport_label_props(new_vport, source_data)
            try:
                dest_doc.Regenerate()
            except Exception:
                pass

            apply_detail_number(source_vport, new_vport)
            try:
                dest_doc.Regenerate()
            except Exception:
                pass

            correct_vport_by_bbox(dest_doc, new_vport, source_data, dest_sheet)
            created.append(new_vport.Id)

        except Exception as ex:
            print("\t\tFailed to copy viewport: {}".format(ex))

    print("\tCopied {} viewport(s).".format(len(created)))
    return created


# -----------------------------------------------------------------------------
# SHEET / TITLEBLOCK / PROPERTY HELPERS
# -----------------------------------------------------------------------------

def find_sheet_by_number(doc, sheet_number):
    try:
        sheets = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.ViewSheet)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        for sheet in sheets:
            try:
                if sheet.SheetNumber == sheet_number:
                    return sheet
            except Exception:
                pass
    except Exception:
        pass
    return None


def get_sheet_titleblock_instance(doc, sheet):
    try:
        tbs = (
            DB.FilteredElementCollector(doc, sheet.Id)
            .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        for tb in tbs:
            return tb
    except Exception:
        pass
    return None


def find_titleblock_type(dest_doc, source_symbol):
    if not source_symbol:
        return None

    source_type_name = safe_name(source_symbol)
    try:
        source_family_name = source_symbol.Family.Name
    except Exception:
        source_family_name = ""

    try:
        symbols = (
            DB.FilteredElementCollector(dest_doc)
            .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
            .WhereElementIsElementType()
            .ToElements()
        )
    except Exception:
        symbols = []

    for sym in symbols:
        try:
            if safe_name(sym) == source_type_name and sym.Family.Name == source_family_name:
                return sym
        except Exception:
            pass

    for sym in symbols:
        try:
            if safe_name(sym) == source_type_name:
                return sym
        except Exception:
            pass

    return None


def ensure_titleblock_type(source_doc, source_sheet, dest_doc):
    if not OPTION_SET or not OPTION_SET.op_copy_titleblock.state:
        return DB.ElementId.InvalidElementId

    source_tb = get_sheet_titleblock_instance(source_doc, source_sheet)
    if not source_tb:
        return DB.ElementId.InvalidElementId

    source_symbol = source_doc.GetElement(source_tb.GetTypeId())
    dest_symbol = find_titleblock_type(dest_doc, source_symbol)
    if dest_symbol:
        return dest_symbol.Id

    try:
        options = DB.CopyPasteOptions()
        options.SetDuplicateTypeNamesHandler(CopyUseDestination())

        copied = DB.ElementTransformUtils.CopyElements(
            source_doc,
            make_element_id_list([source_symbol.Id]),
            dest_doc,
            DB.Transform.Identity,
            options
        )

        for cid in copied:
            elem = dest_doc.GetElement(cid)
            if elem:
                print("\tCopied titleblock type: {}".format(safe_name(elem)))
                return elem.Id
    except Exception as ex:
        print("\tCould not copy titleblock type: {}".format(ex))

    return DB.ElementId.InvalidElementId


def copy_sheet_properties(source_sheet, dest_sheet):
    if not OPTION_SET or not OPTION_SET.op_copy_sheet_properties.state:
        return

    skip_names = set([
        "Sheet Number",
        "Sheet Name",
        "Current Revision",
        "Current Revision Date",
        "Current Revision Description",
        "Current Revision Issued",
        "Current Revision Issued By",
        "Current Revision Issued To",
    ])

    count = 0
    try:
        params = source_sheet.Parameters
    except Exception:
        params = []

    for src_param in params:
        try:
            pname = src_param.Definition.Name
            if pname in skip_names:
                continue

            dst_param = dest_sheet.LookupParameter(pname)
            if dst_param and not dst_param.IsReadOnly:
                if set_param_value(dst_param, get_param_value(src_param)):
                    count += 1
        except Exception:
            pass

    print("\tCopied {} sheet parameter(s).".format(count))


def create_or_get_destination_sheet(source_doc, source_sheet, dest_doc):
    source_number = source_sheet.SheetNumber
    source_name = source_sheet.Name

    existing = find_sheet_by_number(dest_doc, source_number)
    if existing:
        print("\tUsing existing destination sheet: {} - {}".format(existing.SheetNumber, existing.Name))
        return existing

    titleblock_type_id = ensure_titleblock_type(source_doc, source_sheet, dest_doc)

    try:
        new_sheet = DB.ViewSheet.Create(dest_doc, titleblock_type_id)

        try:
            new_sheet.SheetNumber = source_number
        except Exception:
            new_sheet.SheetNumber = unique_sheet_number(dest_doc, source_number)

        try:
            new_sheet.Name = source_name
        except Exception:
            pass

        print("\tCreated sheet: {} - {}".format(new_sheet.SheetNumber, new_sheet.Name))
        return new_sheet
    except Exception as ex:
        print("\tFailed to create sheet '{} - {}': {}".format(source_number, source_name, ex))
        return None


def copy_sheet(source_doc, source_sheet, dest_doc):
    print("\nCopying sheet: {} - {}".format(source_sheet.SheetNumber, source_sheet.Name))

    dest_sheet = create_or_get_destination_sheet(source_doc, source_sheet, dest_doc)
    if not dest_sheet:
        return None

    copy_sheet_properties(source_sheet, dest_sheet)

    if OPTION_SET and OPTION_SET.op_copy_vports.state:
        copy_sheet_viewports(source_doc, source_sheet, dest_doc, dest_sheet)

    return dest_sheet


# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------

def main():
    dest_doc = revit.doc

    source_docs = get_source_link_docs()
    source_sheet_sets = []

    for source_doc in source_docs:
        selected_sheets = get_source_sheets(source_doc)
        if selected_sheets:
            source_sheet_sets.append((source_doc, selected_sheets))

    if not source_sheet_sets:
        forms.alert("No source sheets were selected.", exitscript=True)

    get_user_options()

    total_work = sum(len(x[1]) for x in source_sheet_sets)
    copied_count = 0

    output.print_md("**SAFE Sheet Copy Started**")
    output.print_md("**Destination Document:** {}".format(dest_doc.Title))
    output.print_md("**Total selected sheet(s):** {}".format(total_work))
    output.print_md("**Safety:** Model view contents are NOT copied. Only viewport placement/settings are copied.")

    with revit.Transaction("SAFE Copy Linked Sheets - Plans and Sections"):
        for source_doc, source_sheets in source_sheet_sets:
            output.print_md("**Source Linked Document:** {}".format(source_doc.Title))

            for source_sheet in source_sheets:
                try:
                    copied_sheet = copy_sheet(source_doc, source_sheet, dest_doc)
                    if copied_sheet:
                        copied_count += 1
                except Exception as ex:
                    print("Failed to copy sheet {} - {}: {}".format(
                        source_sheet.SheetNumber,
                        source_sheet.Name,
                        ex
                    ))

    output.print_md("**Copied {} of {} selected sheet(s).**".format(copied_count, total_work))
    output.print_md("**Reminder:** Run Manage > Inquiry > Warnings and confirm duplicate-instance warnings did not increase.")


if __name__ == "__main__":
    main()