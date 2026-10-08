# -*- coding: utf-8 -*-

__title__ = "Delete\nGroups"

__doc__ = """pyRevit script: Delete Groups with Reliable Model/Detail Detection, Preview, and Browser-Name Resolution"""

from pyrevit import revit, DB, forms

def get_all_group_types(doc):
    """Collect all GroupType element types (Model + Detail)."""
    return list(DB.FilteredElementCollector(doc)
                .OfClass(DB.GroupType)
                .WhereElementIsElementType()
                .ToElements())


def get_all_group_instances(doc):
    """Collect all placed Group instances (Model + Detail)."""
    return list(DB.FilteredElementCollector(doc)
                .OfClass(DB.Group)
                .WhereElementIsNotElementType()
                .ToElements())


def build_category_refs(doc):
    """Resolve BuiltInCategory -> Category objects for model/detail groups."""
    cat_model = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_IOSModelGroups)
    cat_detail = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_IOSDetailGroups)
    return cat_model, cat_detail


def _get_param_string(elem, bip):
    """Helper: safely read a string parameter by BuiltInParameter."""
    try:
        p = elem.get_Parameter(bip)
    except Exception:
        p = None
    if p and p.HasValue:
        try:
            s = p.AsString()
            if s and s.strip():
                return s.strip()
        except Exception:
            pass
        try:
            s = p.AsValueString()
            if s and s.strip():
                return s.strip()
        except Exception:
            pass
    return None


def _get_string_param_by_name(elem, names):
    """Helper: scan element parameters to find a string parameter by its display name."""
    try:
        it = elem.Parameters
    except Exception:
        it = None
    if not it:
        return None
    wanted = set(n.lower() for n in names)
    for p in it:
        try:
            defn = p.Definition
            pname = defn.Name if defn else None
            if not pname:
                continue
            if pname.lower() in wanted and p.StorageType == DB.StorageType.String:
                s = p.AsString()
                if s and s.strip():
                    return s.strip()
        except Exception:
            continue
    return None


def resolve_browser_name_for_type(gt, instance_map):
    """
    Return the group type name as seen in Project Browser.
    Order of attempts:
      1) ElementType.Name
      2) SYMBOL_NAME_PARAM
      3) ALL_MODEL_TYPE_NAME
      4) Any string parameter named 'Type Name' or 'Name'
      5) From instances: GROUP_NAME, then instance.Name
      6) Fallback to 'GroupType_<Id>'
    """
    # 1) Native type name
    try:
        name = getattr(gt, "Name", None)
        if name and name.strip():
            return name.strip()
    except Exception:
        pass

    # 2) and 3) Common built-in parameters carrying type name
    for bip in (DB.BuiltInParameter.SYMBOL_NAME_PARAM,
                DB.BuiltInParameter.ALL_MODEL_TYPE_NAME):
        s = _get_param_string(gt, bip)
        if s:
            return s

    # 4) Scan by parameter display names
    s = _get_string_param_by_name(gt, ["Type Name", "Name"])
    if s:
        return s

    # 5) Derive from instances
    insts = instance_map.get(gt.Id, [])
    if insts:
        # Try GROUP_NAME param on instances first
        try:
            for inst in insts:
                s = _get_param_string(inst, DB.BuiltInParameter.GROUP_NAME)
                if s:
                    return s
        except Exception:
            pass
        # Fallback to instance.Name
        for inst in insts:
            try:
                s = getattr(inst, "Name", None)
                if s and s.strip():
                    return s.strip()
            except Exception:
                continue

    # 6) Last resort
    return "GroupType_{}".format(gt.Id.IntegerValue)


def main():
    doc = revit.doc

    # Works only in .rvt (project) docs
    if getattr(doc, "IsFamilyDocument", False):
        forms.alert("This tool works only in Project documents (.rvt), not in Family files (.rfa).",
                    title="Wrong Document")
        return

    # Ask user for filter (Model / Detail / Both)
    filter_choice = forms.SelectFromList.show(
        ["Model Groups", "Detail Groups", "Both"],
        title="Filter: Which group types do you want to see?",
        multiselect=False
    )
    if not filter_choice:
        forms.alert("No filter selected.", title="Cancelled")
        return

    show_model = filter_choice in ("Model Groups", "Both")
    show_detail = filter_choice in ("Detail Groups", "Both")
    include_unknown_when_specific = (filter_choice != "Both")

    # Collect
    group_types = get_all_group_types(doc)
    group_instances = get_all_group_instances(doc)

    # Fallback: rebuild types from instances if needed
    if not group_types and group_instances:
        type_ids = {g.GroupType.Id for g in group_instances if g.GroupType and g.GroupType.Id}
        group_types = [doc.GetElement(tid) for tid in type_ids if tid]

    if not group_types:
        forms.alert("No group types found in the active project document.", title="No Groups")
        return

    # Category refs
    cat_model, cat_detail = build_category_refs(doc)
    cat_model_id = cat_model.Id if cat_model else None
    cat_detail_id = cat_detail.Id if cat_detail else None

    # Map: GroupTypeId -> instances
    instance_map = {}
    for inst in group_instances:
        tid = inst.GroupType.Id
        instance_map.setdefault(tid, []).append(inst)

    # Derive kind membership from instances (most reliable)
    detail_type_ids_from_insts = set()
    model_type_ids_from_insts = set()
    for inst in group_instances:
        inst_cat = getattr(inst, "Category", None)
        if not inst_cat:
            continue
        if cat_detail_id and inst_cat.Id == cat_detail_id:
            detail_type_ids_from_insts.add(inst.GroupType.Id)
        elif cat_model_id and inst_cat.Id == cat_model_id:
            model_type_ids_from_insts.add(inst.GroupType.Id)

    def classify(gt):
        """Return 'Model', 'Detail', or 'Unknown' for a GroupType."""
        tid = gt.Id
        # 1) From instances
        if tid in detail_type_ids_from_insts:
            return "Detail"
        if tid in model_type_ids_from_insts:
            return "Model"
        # 2) Fallback: from type.Category
        cat = getattr(gt, "Category", None)
        if cat is not None:
            if cat_detail_id and cat.Id == cat_detail_id:
                return "Detail"
            if cat_model_id and cat.Id == cat_model_id:
                return "Model"
        # 3) Unknown
        return "Unknown"

    # Build UI list according to filter, using browser-accurate names
    items = []
    detail_count, model_count, unknown_count = 0, 0, 0

    for gt in group_types:
        if gt is None:
            continue

        kind = classify(gt)
        if kind == "Detail":
            detail_count += 1
        elif kind == "Model":
            model_count += 1
        else:
            unknown_count += 1

        # Apply filter
        include = False
        if kind == "Detail" and show_detail:
            include = True
        elif kind == "Model" and show_model:
            include = True
        elif kind == "Unknown" and include_unknown_when_specific:
            include = True

        if not include:
            continue

        name = resolve_browser_name_for_type(gt, instance_map)
        count = len(instance_map.get(gt.Id, []))
        tag = kind if kind != "Unknown" else "Unknown?"
        label = "[{0}] {1} (TypeId {2}; {3} instance{4})".format(
            tag, name, gt.Id.IntegerValue, count, "" if count == 1 else "s"
        )
        items.append((label, gt.Id))

    if not items:
        msg = ("No matching group types were found.\n\n"
               "Discovered counts - Detail: {0}, Model: {1}, Unknown: {2}\n"
               "Tip: Try choosing 'Both' or ensure detail groups are in this project (not only in links)."
               ).format(detail_count, model_count, unknown_count)
        forms.alert(msg, title="No Groups")
        return

    # Sort and show selection dialog
    items_sorted = sorted(items, key=lambda x: x[0].lower())
    choice_labels = [lbl for (lbl, _) in items_sorted]

    selected = forms.SelectFromList.show(choice_labels,
                                         title="Select Group Types to Delete",
                                         multiselect=True)

    if not selected:
        forms.alert("No groups selected.", title="Cancelled")
        return

    # Resolve labels -> ids
    label_to_id = {lbl: id_ for (lbl, id_) in items_sorted}
    selected_type_ids = [label_to_id[lbl] for lbl in selected]

    # Preview: select all instances of chosen types in UI
    all_sel_insts = []
    for tid in selected_type_ids:
        all_sel_insts.extend(instance_map.get(tid, []))
    if all_sel_insts:
        try:
            revit.get_selection().set_to(all_sel_insts)
        except Exception:
            pass  # Selection preview can fail in some contexts (e.g. edit mode, no active view)

    # Delete instances first, then types
    deleted_types = 0
    with revit.Transaction("Delete Selected Groups"):
        for tid in selected_type_ids:
            # Delete instances
            for inst in list(instance_map.get(tid, [])):
                try:
                    doc.Delete(inst.Id)
                except Exception as ex:
                    print("Could not delete instance {}: {}".format(inst.Id.IntegerValue, ex))
            # Delete the type
            try:
                doc.Delete(tid)
                deleted_types += 1
            except Exception as ex:
                print("Could not delete group type {}: {}".format(tid.IntegerValue, ex))

    forms.alert("Deleted {0} group type{1} and their instances."
                .format(deleted_types, "" if deleted_types == 1 else "s"),
                title="Success")


if __name__ == "__main__":
    main()
