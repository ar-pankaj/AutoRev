# -*- coding: utf-8 -*-
__title__ = "Family Purify"
__doc__ = """Version = 1.0
Date    = 25.05.2026
________________________________________________________________
Description:
Clean, optimize and standardize open Revit family documents.

________________________________________________________________
How-To:
1. Open one or more Revit families in the Family Editor.
2. Select the families to process.
3. Choose cleanup, advanced cleanup and parameter standardization options.
4. Use Preview cleanup to scan only, or Purify Families to run the cleanup.

________________________________________________________________
Main Features:
- Remove CAD imports, images, unused materials, patterns and appearance assets.
- Purge subcategories and compact-save family files.
- Clean nested families and reload them into the parent family.
- Add selected shared parameters for family standardization.
- Show before/after size, health score and cleanup summary.

________________________________________________________________
To-Do:
[FEATURE] Add exportable HTML/CSV cleanup report.
[FEATURE] Add preset profiles for safe, balanced and aggressive cleanup.

________________________________________________________________
Last Updates:
- [28.06.2026] v1.0 Hackathon polish, nested cleanup and shared parameters.
- [27.06.2026] v0.5 Family cleanup workflow and WPF interface.
________________________________________________________________
Author: Justin Biju"""

import os
import time
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows.Controls import ListBoxItem, CheckBox
from System.Windows.Media import SolidColorBrush, Colors

from pyrevit import forms
from pyrevit.forms import WPFWindow

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    View,
    ViewType,
    Transaction,
    TransactionGroup,
    SaveAsOptions,
    BuiltInParameterGroup,
    ElementId,
    StorageType,
    LinePatternElement,
    FillPatternElement,
    AppearanceAssetElement,
    Material,
    FamilySymbol,
    FamilyInstance,
    ImportInstance,
    ImageType,
    IFamilyLoadOptions,
)

try:
    from Autodesk.Revit.DB import GroupTypeId
except:
    GroupTypeId = None

uiapp = __revit__
app = uiapp.Application


# --------------------------
# Family load options (overwrite when reloading nested)
# --------------------------
class OverwriteFamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True


# --------------------------
# Utility
# --------------------------
def _fmt_bytes(n):
    try:
        n = float(n)
    except:
        return "Unknown"
    for unit in ["B", "KB", "MB", "GB"]:
        if n < 1024.0:
            if unit == "B":
                return "%d %s" % (int(n), unit)
            return "%.2f %s" % (n, unit)
        n /= 1024.0
    return "%.2f TB" % n


def _safe_file_size(path):
    try:
        if path and os.path.exists(path):
            return os.path.getsize(path)
    except:
        pass
    return None


class FamilyPurifyWindow(WPFWindow):
    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), "FamilyPurify.xaml")
        if not os.path.exists(xaml_path):
            forms.alert("FamilyPurify.xaml not found:\n{}".format(xaml_path), title="Family Purify")
            return

        WPFWindow.__init__(self, xaml_path)

        self.family_label_to_doc = {}
        self.param_checkboxes = []
        self.type_checkboxes = []
        self.company_param_checkboxes = []
        self.current_family_doc = None

        self.populate_open_families()
        self.populate_company_parameters()
        self.ShowDialog()

    # --------------------------
    # Families
    # --------------------------
    def populate_open_families(self):
        self.FamiliesList.Items.Clear()
        self.family_label_to_doc = {}

        famdocs = []
        for d in list(app.Documents):
            try:
                if d.IsFamilyDocument:
                    famdocs.append(d)
            except:
                pass

        if not famdocs:
            forms.alert("No family documents are open.", title="Family Purify")
            return

        for d in famdocs:
            label = u"{}".format(d.Title)
            self.family_label_to_doc[label] = d
            self.FamiliesList.Items.Add(label)

    def family_selection_changed(self, sender, args):
        selected = list(self.FamiliesList.SelectedItems)
        if not selected:
            self.current_family_doc = None
            self._populate_params([])
            self._populate_types([])
            self._update_family_summary(None)
            return

        label = str(selected[0])
        famdoc = self.family_label_to_doc.get(label)
        self.current_family_doc = famdoc

        if not famdoc:
            self._populate_params([])
            self._populate_types([])
            self._update_family_summary(None)
            return

        self._populate_params(self._get_family_param_names(famdoc))
        self._populate_types(self._get_family_type_names(famdoc))
        self._update_family_summary(famdoc)

    def _set_summary_text(self, control_name, value):
        try:
            getattr(self, control_name).Text = value
        except:
            pass

    def _update_family_summary(self, famdoc):
        if not famdoc:
            self._set_summary_text("SummaryFileSizeText", "File size: --")
            self._set_summary_text("SummaryTypesText", "Family types: --")
            self._set_summary_text("SummaryNestedText", "Nested families: --")
            self._set_summary_text("SummarySharedParamsText", "Shared parameters: --")
            self._set_summary_text("SummaryStatusText", "Status: Select a family")
            return

        audit = self._collect_audit(famdoc)
        size_value = audit.get("size", None)
        size_text = "Unsaved" if size_value is None else _fmt_bytes(size_value)

        self._set_summary_text("SummaryFileSizeText", "File size: {}".format(size_text))
        self._set_summary_text("SummaryTypesText", "Family types: {}".format(audit.get("types", 0)))
        self._set_summary_text("SummaryNestedText", "Nested families: {}".format(audit.get("nested", 0)))
        self._set_summary_text("SummarySharedParamsText", "Shared parameters: {}".format(audit.get("shared_params", 0)))
        self._set_summary_text("SummaryStatusText", "Status: Ready to Purify")

    # --------------------------
    # Parameters UI
    # --------------------------
    def _get_family_param_names(self, famdoc):
        names = []
        try:
            fm = famdoc.FamilyManager
            for p in fm.Parameters:
                try:
                    names.append(p.Definition.Name)
                except:
                    pass
        except:
            pass
        return sorted(list(set(names)))

    def _populate_params(self, param_names):
        self.ParamsList.Items.Clear()
        self.param_checkboxes = []

        for pname in param_names:
            item = ListBoxItem()
            cb = CheckBox()
            cb.Content = pname
            cb.IsChecked = True
            cb.Foreground = SolidColorBrush(Colors.White)
            item.Content = cb
            self.ParamsList.Items.Add(item)
            self.param_checkboxes.append((pname, cb))

    def select_all_params(self, sender, args):
        for _, cb in self.param_checkboxes:
            cb.IsChecked = True

    def select_none_params(self, sender, args):
        for _, cb in self.param_checkboxes:
            cb.IsChecked = False

    def _get_keep_param_set(self):
        keep = set()
        for pname, cb in self.param_checkboxes:
            try:
                if cb.IsChecked:
                    keep.add(pname)
            except:
                pass
        return keep

    # --------------------------
    # Company shared parameters UI
    # --------------------------
    def browse_company_params(self, sender, args):
        picked = None
        try:
            picked = forms.pick_file(file_ext="txt", title="Pick Company Shared Parameter File")
        except:
            try:
                picked = forms.open_file(file_ext="txt", title="Pick Company Shared Parameter File")
            except:
                picked = None

        if picked:
            try:
                app.SharedParametersFilename = picked
            except:
                forms.alert("Could not set the selected shared parameter file.", title="Family Purify")
                return

        self.populate_company_parameters()

    def populate_company_parameters(self):
        self.CompanyParamsList.Items.Clear()
        self.company_param_checkboxes = []

        sp_path = ""
        try:
            sp_path = app.SharedParametersFilename or ""
        except:
            sp_path = ""

        try:
            self.CompanyParamFileLabel.Text = sp_path if sp_path else "No shared parameter file selected"
        except:
            pass

        if not sp_path:
            return

        try:
            def_file = app.OpenSharedParameterFile()
        except:
            def_file = None

        if not def_file:
            try:
                self.CompanyParamFileLabel.Text = "Could not read shared parameter file"
            except:
                pass
            return

        try:
            groups = list(def_file.Groups)
        except:
            groups = []

        for group in groups:
            try:
                group_name = group.Name
                definitions = list(group.Definitions)
            except:
                continue

            for definition in definitions:
                try:
                    label = u"{} / {}".format(group_name, definition.Name)
                    item = ListBoxItem()
                    cb = CheckBox()
                    cb.Content = label
                    cb.IsChecked = False
                    cb.Foreground = SolidColorBrush(Colors.White)
                    item.Content = cb
                    self.CompanyParamsList.Items.Add(item)
                    self.company_param_checkboxes.append((definition, cb))
                except:
                    pass

    def select_all_company_params(self, sender, args):
        for _, cb in self.company_param_checkboxes:
            cb.IsChecked = True

    def select_none_company_params(self, sender, args):
        for _, cb in self.company_param_checkboxes:
            cb.IsChecked = False

    def _get_selected_company_param_defs(self):
        selected = []
        for definition, cb in self.company_param_checkboxes:
            try:
                if cb.IsChecked:
                    selected.append(definition)
            except:
                pass
        return selected

    # --------------------------
    # Types UI
    # --------------------------
    def _get_family_type_names(self, famdoc):
        names = []

        try:
            fm = famdoc.FamilyManager
            for ftype in fm.Types:
                try:
                    if ftype.Name:
                        names.append(ftype.Name)
                except:
                    pass
        except:
            pass

        return sorted(list(set(names)))

    def _populate_types(self, type_names):
        self.TypesList.Items.Clear()
        self.type_checkboxes = []

        for tname in type_names:
            item = ListBoxItem()
            cb = CheckBox()
            cb.Content = tname
            cb.IsChecked = False
            cb.Foreground = SolidColorBrush(Colors.White)
            item.Content = cb
            self.TypesList.Items.Add(item)
            self.type_checkboxes.append((tname, cb))

    def select_all_types(self, sender, args):
        for _, cb in self.type_checkboxes:
            cb.IsChecked = True

    def select_none_types(self, sender, args):
        for _, cb in self.type_checkboxes:
            cb.IsChecked = False

    def _get_selected_type_names_to_delete(self):
        tset = set()
        for tname, cb in self.type_checkboxes:
            try:
                if cb.IsChecked:
                    tset.add(tname)
            except:
                pass
        return tset

    # --------------------------
    # Audit snapshot (internal only)
    # --------------------------
    def _collect_audit(self, famdoc):
        audit = {
            "title": famdoc.Title,
            "types": 0,
            "params": 0,
            "materials": 0,
            "imports": 0,
            "images": 0,
            "nested": 0,
            "shared_params": 0,
            "path": "",
            "size": None
        }

        try:
            audit["path"] = famdoc.PathName or ""
        except:
            audit["path"] = ""

        audit["size"] = _safe_file_size(audit["path"])

        try:
            type_names = self._get_family_type_names(famdoc)
            audit["types"] = len(type_names)
        except:
            try:
                audit["types"] = len(list(FilteredElementCollector(famdoc).OfClass(FamilySymbol).ToElements()))
            except:
                pass

        try:
            fm = famdoc.FamilyManager
            audit["params"] = len([p for p in fm.Parameters])
            shared_count = 0
            for p in fm.Parameters:
                try:
                    if p.IsShared:
                        shared_count += 1
                except:
                    pass
            audit["shared_params"] = shared_count
        except:
            pass

        try:
            audit["materials"] = len(list(FilteredElementCollector(famdoc).OfClass(Material).ToElements()))
        except:
            pass

        try:
            audit["imports"] = len(list(FilteredElementCollector(famdoc).OfClass(ImportInstance).ToElements()))
        except:
            pass

        try:
            audit["images"] = len(list(FilteredElementCollector(famdoc).OfClass(ImageType).ToElements()))
        except:
            pass

        try:
            nested_fams = set()
            insts = FilteredElementCollector(famdoc).OfClass(FamilyInstance).ToElements()
            for fi in insts:
                try:
                    sym = fi.Symbol
                    fam = sym.Family if sym else None
                    if fam:
                        nested_fams.add(fam.Id.IntegerValue)
                except:
                    pass
            audit["nested"] = len(nested_fams)
        except:
            pass

        return audit

    def _health_score(self, audit, file_size=None):
        score = 100

        score -= min(audit.get("imports", 0) * 10, 25)
        score -= min(audit.get("images", 0) * 5, 15)
        score -= min(audit.get("materials", 0), 20)

        if audit.get("types", 0) > 50:
            score -= 10
        if audit.get("params", 0) > 80:
            score -= 10
        if audit.get("nested", 0) > 10:
            score -= 10

        try:
            if file_size and file_size > (20 * 1024 * 1024):
                score -= 10
        except:
            pass

        if score < 0:
            score = 0
        if score > 100:
            score = 100
        return score

    # --------------------------
    # RUN
    # --------------------------
    def clean_selected(self, sender, args):
        started_at = time.time()
        selected_labels = [str(x) for x in list(self.FamiliesList.SelectedItems)]
        if not selected_labels:
            forms.alert("Select at least one family document.", title="Family Purify")
            return

        keep_params = self._get_keep_param_set()

        keep_core_views = bool(self.KeepCoreViewsChk.IsChecked)
        remove_cad_imports = bool(self.RemoveCADImportsChk.IsChecked)
        remove_images = bool(self.RemoveImagesChk.IsChecked)
        remove_imports_images = remove_cad_imports or remove_images
        purge_patterns_assets = bool(self.PurgePatternsAssetsChk.IsChecked)
        keep_only_used_materials = bool(self.KeepOnlyUsedMaterialsChk.IsChecked)
        purge_subcategories = bool(self.PurgeSubcategoriesChk.IsChecked)
        delete_selected_types = bool(self.DeleteSelectedTypesChk.IsChecked)
        compact_save = bool(self.CompactSaveChk.IsChecked)
        scan_only = bool(self.ScanOnlyChk.IsChecked)

        process_nested = bool(self.ProcessNestedChk.IsChecked)
        nested_mode = "SAFE"
        try:
            nested_mode = "FULL" if int(self.NestedModeCombo.SelectedIndex) == 1 else "SAFE"
        except:
            nested_mode = "SAFE"

        company_param_defs = self._get_selected_company_param_defs()
        add_company_params = bool(self.AddCompanyParamsChk.IsChecked) and bool(company_param_defs)
        add_company_params_to_nested = bool(self.CompanyParamsNestedChk.IsChecked)
        company_params_instance = True
        try:
            company_params_instance = False if int(self.CompanyParamKindCombo.SelectedIndex) == 1 else True
        except:
            company_params_instance = True

        type_names_to_delete = set()
        if delete_selected_types:
            type_names_to_delete = self._get_selected_type_names_to_delete()

        rename_value = ""
        try:
            rename_value = (self.RenameBox.Text or "").strip()
        except:
            rename_value = ""

        visited_nested = set()

        for label in selected_labels:
            famdoc = self.family_label_to_doc.get(label)
            if not famdoc:
                continue

            audit_before = self._collect_audit(famdoc)
            size_before = audit_before["size"]
            health_before = self._health_score(audit_before, size_before)

            if scan_only:
                elapsed = time.time() - started_at
                preview = (
                    "SCAN COMPLETE - NO CHANGES MADE\n\n"
                    "Family: {title}\n"
                    "Current size: {size}\n"
                    "Health score: {health}/100\n\n"
                    "Found:\n"
                    "CAD imports: {imports}\n"
                    "Images: {images}\n"
                    "Materials: {materials}\n"
                    "Family types: {types}\n"
                    "Family parameters: {params}\n"
                    "Nested families: {nested}\n\n"
                    "Processing time: {elapsed:.1f}s"
                ).format(
                    title=audit_before.get("title", ""),
                    size=("Unsaved" if size_before is None else _fmt_bytes(size_before)),
                    health=health_before,
                    imports=audit_before.get("imports", 0),
                    images=audit_before.get("images", 0),
                    materials=audit_before.get("materials", 0),
                    types=audit_before.get("types", 0),
                    params=audit_before.get("params", 0),
                    nested=audit_before.get("nested", 0),
                    elapsed=elapsed
                )
                forms.alert(preview, title="Family Purify")
                continue

            stats = self._clean_one_family(
                famdoc=famdoc,
                keep_params=keep_params,
                keep_core_views=keep_core_views,
                remove_imports_images=remove_imports_images,
                remove_cad_imports=remove_cad_imports,
                remove_images=remove_images,
                purge_patterns_assets=purge_patterns_assets,
                keep_only_used_materials=keep_only_used_materials,
                purge_subcategories=purge_subcategories,
                delete_type_names=type_names_to_delete,
                process_nested=process_nested,
                nested_mode=nested_mode,
                visited_nested=visited_nested,
                company_param_defs=(company_param_defs if add_company_params else []),
                company_params_instance=company_params_instance,
                add_company_params_to_nested=add_company_params_to_nested,
                depth=0,
                max_depth=10
            )

            saved_path = None
            if compact_save:
                new_name = rename_value if (rename_value and len(selected_labels) == 1) else ""
                saved_path = self._save_compact(famdoc, new_name)

            audit_after = self._collect_audit(famdoc)
            size_after = _safe_file_size(saved_path) if saved_path else _safe_file_size(audit_after.get("path", ""))

            delta = None
            if size_before is not None and size_after is not None:
                delta = size_after - size_before

            reduction_pct = None
            try:
                if size_before and size_after is not None:
                    reduction_pct = ((size_before - size_after) / float(size_before)) * 100.0
            except:
                reduction_pct = None

            health_after = self._health_score(audit_after, size_after)
            health = "{}/100 -> {}/100".format(health_before, health_after)
            elapsed = time.time() - started_at
            pct_text = "Unknown" if reduction_pct is None else "{:.1f}%".format(reduction_pct)
            if stats.get("nested_failed_details", ""):
                if len(stats["nested_failed_details"]) > 500:
                    stats["nested_failed_details"] = stats["nested_failed_details"][:497] + "..."
                stats["nested_failed_details"] = "Nested failure details: {}\n".format(
                    stats["nested_failed_details"]
                )

            msg_after = (
                "FAMILY PURIFY COMPLETE\n\n"
                "Health score: {health}\n"
                "Size: {sb} -> {sa}{sd}\n"
                "Removed parameters: {rp}   Failed: {fp}\n"
                "Removed types: {rt}        Failed: {ft}\n"
                "Size reduction: {pct}\n"
                "Processing time: {elapsed:.1f}s\n"
                "CAD imports removed: {rcad}   Failed: {fcad}\n"
                "Images removed: {rimg}        Failed: {fimg}\n"
                "Removed subcategories: {rsc}   Failed: {fsc}\n"
                "Removed materials: {rm}     Failed: {fm}\n"
                "Removed line patterns: {rl} Failed: {fl}\n"
                "Removed fill patterns: {rfi} Failed: {ffi}\n"
                "Removed appearance assets: {ra} Failed: {fa}\n\n"
                "Shared parameters added: {ap}   Skipped/Failed: {fap}\n"
                "Nested families processed: {nested}\n"
                "Nested families failed: {nested_failed}\n"
                "{nested_failed_details}"
                "Compact saved: {saved}"
            ).format(
                health=health,
                sb=("Unsaved" if size_before is None else _fmt_bytes(size_before)),
                sa=("Unsaved" if size_after is None else _fmt_bytes(size_after)),
                sd=("" if delta is None else "  ({:+.2f} MB)".format(delta / (1024.0 * 1024.0))),
                pct=pct_text,
                elapsed=elapsed,
                saved=("Yes" if saved_path else "No"),
                **stats
            )

            forms.alert(msg_after, title="Family Purify")

    # --------------------------
    # CORE CLEAN
    # --------------------------
    def _clean_one_family(self, famdoc,
                          keep_params,
                          keep_core_views,
                          remove_imports_images,
                          remove_cad_imports,
                          remove_images,
                          purge_patterns_assets,
                          keep_only_used_materials,
                          purge_subcategories,
                          delete_type_names,
                          process_nested,
                          nested_mode,
                          visited_nested,
                          company_param_defs,
                          company_params_instance,
                          add_company_params_to_nested,
                          depth,
                          max_depth):

        stats = {
            "rp": 0, "fp": 0,
            "rt": 0, "ft": 0,
            "ri": 0, "fi": 0,
            "rcad": 0, "fcad": 0,
            "rimg": 0, "fimg": 0,
            "rsc": 0, "fsc": 0,
            "rm": 0, "fm": 0,
            "rl": 0, "fl": 0,
            "rfi": 0, "ffi": 0,
            "ra": 0, "fa": 0,
            "ap": 0, "fap": 0,
            "nested": 0,
            "nested_failed": 0,
            "nested_failed_details": ""
        }

        cleanup_ok = False
        tg = TransactionGroup(famdoc, "Family Purify Cleanup")
        tg.Start()
        try:
            if keep_core_views:
                self._delete_non_core_views(famdoc)

            rp, fp = self._remove_family_parameters(famdoc, keep_params)
            stats["rp"] += rp
            stats["fp"] += fp

            if delete_type_names:
                rt, ft = self._delete_family_types_by_name(famdoc, delete_type_names)
                stats["rt"] += rt
                stats["ft"] += ft

            if remove_imports_images:
                ri, fi, rcad, fcad, rimg, fimg = self._delete_imports_and_images(
                    famdoc,
                    remove_cad_imports,
                    remove_images
                )
                stats["ri"] += ri
                stats["fi"] += fi
                stats["rcad"] += rcad
                stats["fcad"] += fcad
                stats["rimg"] += rimg
                stats["fimg"] += fimg

            if purge_subcategories:
                rsc, fsc = self._purge_unused_subcategories(famdoc)
                stats["rsc"] += rsc
                stats["fsc"] += fsc

            if purge_patterns_assets:
                rl, fl, rfi, ffi, ra, fa = self._purge_unused_patterns_and_assets(famdoc)
                stats["rl"] += rl
                stats["fl"] += fl
                stats["rfi"] += rfi
                stats["ffi"] += ffi
                stats["ra"] += ra
                stats["fa"] += fa

            if keep_only_used_materials:
                rm, fm = self._keep_only_used_materials(famdoc)
                stats["rm"] += rm
                stats["fm"] += fm

            if company_param_defs and (depth == 0 or add_company_params_to_nested):
                ap, fap = self._add_company_parameters(
                    famdoc=famdoc,
                    company_param_defs=company_param_defs,
                    is_instance=company_params_instance
                )
                stats["ap"] += ap
                stats["fap"] += fap

            tg.Assimilate()
            cleanup_ok = True
        except:
            try:
                tg.RollBack()
            except:
                pass

        if cleanup_ok and process_nested and depth < max_depth:
            nested_stats = self._process_nested_families(
                parent_famdoc=famdoc,
                nested_mode=nested_mode,
                keep_params=keep_params,
                keep_core_views=keep_core_views,
                remove_imports_images=remove_imports_images,
                remove_cad_imports=remove_cad_imports,
                remove_images=remove_images,
                purge_patterns_assets=purge_patterns_assets,
                keep_only_used_materials=keep_only_used_materials,
                purge_subcategories=purge_subcategories,
                visited_nested=visited_nested,
                company_param_defs=company_param_defs,
                company_params_instance=company_params_instance,
                add_company_params_to_nested=add_company_params_to_nested,
                depth=depth + 1,
                max_depth=max_depth
            )
            for key, value in nested_stats.items():
                if key in stats:
                    stats[key] += value

        return stats

    # --------------------------
    # NESTED
    # --------------------------
    def _process_nested_families(self, parent_famdoc,
                                 nested_mode,
                                 keep_params,
                                 keep_core_views,
                                 remove_imports_images,
                                 remove_cad_imports,
                                 remove_images,
                                 purge_patterns_assets,
                                 keep_only_used_materials,
                                 purge_subcategories,
                                 visited_nested,
                                 company_param_defs,
                                 company_params_instance,
                                 add_company_params_to_nested,
                                 depth,
                                 max_depth):

        stats = {
            "rp": 0, "fp": 0,
            "rt": 0, "ft": 0,
            "ri": 0, "fi": 0,
            "rcad": 0, "fcad": 0,
            "rimg": 0, "fimg": 0,
            "rsc": 0, "fsc": 0,
            "rm": 0, "fm": 0,
            "rl": 0, "fl": 0,
            "rfi": 0, "ffi": 0,
            "ra": 0, "fa": 0,
            "ap": 0, "fap": 0,
            "nested": 0,
            "nested_failed": 0,
            "nested_failed_details": ""
        }
        load_opts = OverwriteFamilyLoadOptions()

        nested_families = []
        nested_keys = set()
        try:
            insts = FilteredElementCollector(parent_famdoc).OfClass(FamilyInstance).ToElements()
            for fi in insts:
                try:
                    sym = fi.Symbol
                    fam = sym.Family if sym else None
                    if fam:
                        fid = fam.Id.IntegerValue
                        key = fam.UniqueId if fam.UniqueId else str(fid)
                        if key not in nested_keys:
                            nested_keys.add(key)
                            nested_families.append({
                                "id": fid,
                                "unique_id": fam.UniqueId,
                                "name": fam.Name
                            })
                except:
                    pass
        except:
            pass

        for fam_info in nested_families:
            fam_key = fam_info.get("unique_id") or fam_info.get("id")
            fam_name = fam_info.get("name") or "Nested family"

            if fam_key is not None and fam_key in visited_nested:
                continue
            if fam_key is not None:
                visited_nested.add(fam_key)

            nested_doc = None
            try:
                fam = self._resolve_nested_family(parent_famdoc, fam_info)
                if not fam:
                    raise Exception("Could not resolve nested family after parent reload.")

                nested_doc = parent_famdoc.EditFamily(fam)
                if not nested_doc:
                    continue

                if nested_mode == "FULL":
                    child_stats = self._clean_one_family(
                        famdoc=nested_doc,
                        keep_params=keep_params,
                        keep_core_views=keep_core_views,
                        remove_imports_images=remove_imports_images,
                        remove_cad_imports=remove_cad_imports,
                        remove_images=remove_images,
                        purge_patterns_assets=purge_patterns_assets,
                        keep_only_used_materials=keep_only_used_materials,
                        purge_subcategories=purge_subcategories,
                        delete_type_names=set(),  # safer
                        process_nested=True,
                        nested_mode="FULL",
                        visited_nested=visited_nested,
                        company_param_defs=company_param_defs,
                        company_params_instance=company_params_instance,
                        add_company_params_to_nested=add_company_params_to_nested,
                        depth=depth,
                        max_depth=max_depth
                    )
                    for key, value in child_stats.items():
                        if key in stats:
                            stats[key] += value
                else:
                    self._safe_clean_nested(
                        famdoc=nested_doc,
                        remove_imports_images=remove_imports_images,
                        remove_cad_imports=remove_cad_imports,
                        remove_images=remove_images,
                        purge_patterns_assets=purge_patterns_assets,
                        keep_only_used_materials=keep_only_used_materials,
                        purge_subcategories=purge_subcategories
                    )
                    if company_param_defs and add_company_params_to_nested:
                        ap, fap = self._add_company_parameters(
                            famdoc=nested_doc,
                            company_param_defs=company_param_defs,
                            is_instance=company_params_instance
                        )
                        stats["ap"] += ap
                        stats["fap"] += fap

                try:
                    nested_doc.Regenerate()
                except:
                    pass

                nested_doc.LoadFamily(parent_famdoc, load_opts)
                try:
                    parent_famdoc.Regenerate()
                except:
                    pass

                stats["nested"] += 1

            except Exception as ex:
                stats["nested_failed"] += 1
                detail = "{}: {}".format(fam_name, ex)
                if len(detail) > 180:
                    detail = detail[:177] + "..."
                if stats["nested_failed_details"]:
                    stats["nested_failed_details"] += "; "
                stats["nested_failed_details"] += detail
            finally:
                if nested_doc:
                    try:
                        nested_doc.Close(False)
                    except:
                        pass

        return stats

    def _resolve_nested_family(self, parent_famdoc, fam_info):
        try:
            unique_id = fam_info.get("unique_id")
            if unique_id:
                fam = parent_famdoc.GetElement(unique_id)
                if fam:
                    return fam
        except:
            pass

        try:
            fam_id = fam_info.get("id")
            if fam_id is not None:
                fam = parent_famdoc.GetElement(ElementId(int(fam_id)))
                if fam:
                    return fam
        except:
            pass

        target_name = fam_info.get("name") or ""
        if not target_name:
            return None

        try:
            insts = FilteredElementCollector(parent_famdoc).OfClass(FamilyInstance).ToElements()
            for fi in insts:
                try:
                    sym = fi.Symbol
                    fam = sym.Family if sym else None
                    if fam and fam.Name == target_name:
                        return fam
                except:
                    pass
        except:
            pass

        return None

    def _safe_clean_nested(self, famdoc,
                          remove_imports_images,
                          remove_cad_imports,
                          remove_images,
                          purge_patterns_assets,
                          keep_only_used_materials,
                          purge_subcategories):
        tg = TransactionGroup(famdoc, "Safe Nested Clean")
        tg.Start()
        try:
            if remove_imports_images:
                self._delete_imports_and_images(famdoc, remove_cad_imports, remove_images)
            if purge_subcategories:
                self._purge_unused_subcategories(famdoc)
            if purge_patterns_assets:
                self._purge_unused_patterns_and_assets(famdoc)
            if keep_only_used_materials:
                self._keep_only_used_materials(famdoc)
            tg.Assimilate()
        except:
            try:
                tg.RollBack()
            except:
                pass

    # --------------------------
    # Company shared parameters
    # --------------------------
    def _add_company_parameters(self, famdoc, company_param_defs, is_instance):
        added = 0
        failed_or_skipped = 0

        if not company_param_defs:
            return added, failed_or_skipped

        existing_names = set()
        try:
            fm = famdoc.FamilyManager
            for p in fm.Parameters:
                try:
                    existing_names.add(p.Definition.Name)
                except:
                    pass
        except:
            return added, len(company_param_defs)

        group_arg = BuiltInParameterGroup.PG_DATA
        if GroupTypeId:
            try:
                group_arg = GroupTypeId.Data
            except:
                group_arg = BuiltInParameterGroup.PG_DATA

        t = Transaction(famdoc, "Add Shared Parameters")
        t.Start()
        try:
            for definition in company_param_defs:
                try:
                    pname = definition.Name
                    if pname in existing_names:
                        failed_or_skipped += 1
                        continue

                    famdoc.FamilyManager.AddParameter(definition, group_arg, is_instance)
                    existing_names.add(pname)
                    added += 1
                except:
                    failed_or_skipped += 1
        finally:
            t.Commit()

        return added, failed_or_skipped

    # --------------------------
    # Helpers
    # --------------------------
    def _delete_non_core_views(self, famdoc):
        keep_types = set([ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation, ViewType.Section, ViewType.ThreeD])
        t = Transaction(famdoc, "Delete Non-Core Views")
        t.Start()
        try:
            active_id = famdoc.ActiveView.Id if famdoc.ActiveView else None
            for v in FilteredElementCollector(famdoc).OfClass(View):
                try:
                    if v.IsTemplate:
                        continue
                    if active_id and v.Id == active_id:
                        continue
                    nm = (v.Name or "").lower()
                    keep = (v.ViewType in keep_types) or ("ref" in nm) or ("reference" in nm) or ("level" in nm)
                    if not keep:
                        try:
                            famdoc.Delete(v.Id)
                        except:
                            pass
                except:
                    pass
        finally:
            t.Commit()

    def _remove_family_parameters(self, famdoc, keep_params):
        removed = 0
        failed = 0
        t = Transaction(famdoc, "Remove Family Parameters")
        t.Start()
        try:
            fm = famdoc.FamilyManager
            params = [p for p in fm.Parameters]
            for p in params:
                try:
                    pname = p.Definition.Name
                    if pname in keep_params:
                        continue
                    try:
                        fm.RemoveParameter(p)
                        removed += 1
                    except:
                        failed += 1
                except:
                    failed += 1
        finally:
            t.Commit()
        return removed, failed

    def _delete_family_types_by_name(self, famdoc, type_names_to_delete):
        removed = 0
        failed = 0
        t = Transaction(famdoc, "Delete Selected Types")
        t.Start()
        try:
            fm = famdoc.FamilyManager
            family_types = [ftype for ftype in fm.Types]
            for ftype in family_types:
                try:
                    if ftype.Name not in type_names_to_delete:
                        continue

                    fm.CurrentType = ftype
                    fm.DeleteCurrentType()
                    removed += 1
                except:
                    failed += 1
        finally:
            t.Commit()
        return removed, failed

    def _delete_imports_and_images(self, famdoc, delete_imports=True, delete_images=True):
        removed = 0
        failed = 0
        removed_imports = 0
        failed_imports = 0
        removed_images = 0
        failed_images = 0

        t = Transaction(famdoc, "Delete Imports + Images")
        t.Start()
        try:
            if delete_imports:
                try:
                    imports = list(FilteredElementCollector(famdoc).OfClass(ImportInstance).ToElements())
                    for imp in imports:
                        try:
                            famdoc.Delete(imp.Id)
                            removed += 1
                            removed_imports += 1
                        except:
                            failed += 1
                            failed_imports += 1
                except:
                    pass

            if delete_images:
                try:
                    imgs = list(FilteredElementCollector(famdoc).OfClass(ImageType).ToElements())
                    for img in imgs:
                        try:
                            famdoc.Delete(img.Id)
                            removed += 1
                            removed_images += 1
                        except:
                            failed += 1
                            failed_images += 1
                except:
                    pass
        finally:
            t.Commit()

        return removed, failed, removed_imports, failed_imports, removed_images, failed_images

    def _keep_only_used_materials(self, famdoc):
        removed = 0
        failed = 0

        mats = list(FilteredElementCollector(famdoc).OfClass(Material).ToElements())
        all_mat_ids = set([m.Id for m in mats])
        used_mat_ids = set()

        def scan_params(elem):
            try:
                for p in elem.Parameters:
                    try:
                        if p.StorageType != StorageType.ElementId:
                            continue
                        eid = p.AsElementId()
                        if eid and eid != ElementId.InvalidElementId and eid in all_mat_ids:
                            used_mat_ids.add(eid)
                    except:
                        pass
            except:
                pass

        try:
            for e in FilteredElementCollector(famdoc).WhereElementIsNotElementType().ToElements():
                scan_params(e)
            for t_ in FilteredElementCollector(famdoc).WhereElementIsElementType().ToElements():
                scan_params(t_)
        except:
            pass

        unused_ids = list(all_mat_ids - used_mat_ids)

        t = Transaction(famdoc, "Keep Only Used Materials")
        t.Start()
        try:
            for mid in unused_ids:
                try:
                    famdoc.Delete(mid)
                    removed += 1
                except:
                    failed += 1
        finally:
            t.Commit()

        return removed, failed

    def _purge_unused_subcategories(self, famdoc):
        removed = 0
        failed = 0

        elems = []
        try:
            elems = list(FilteredElementCollector(famdoc).WhereElementIsNotElementType().ToElements())
        except:
            elems = []

        # used categories (including subcategories) by elements
        used_cat_ints = set()
        for e in elems:
            try:
                c = e.Category
                if c:
                    used_cat_ints.add(c.Id.IntegerValue)
            except:
                pass

        # gather parent categories seen
        parents = {}
        for e in elems:
            try:
                c = e.Category
                if not c:
                    continue
                pid = c.Id.IntegerValue
                if pid not in parents:
                    parents[pid] = c
            except:
                pass

        candidates = []
        for _, pc in parents.items():
            try:
                subs = pc.SubCategories
                if not subs:
                    continue
                it = subs.GetEnumerator()
                while it.MoveNext():
                    sc = it.Current
                    try:
                        nm = (sc.Name or "").lower()
                        if nm in ["<sketch>", "invisible lines"]:
                            continue
                        candidates.append(sc)
                    except:
                        pass
            except:
                pass

        t = Transaction(famdoc, "Purge Unused Subcategories")
        t.Start()
        try:
            for sc in candidates:
                try:
                    scid = sc.Id
                    if scid and scid.IntegerValue not in used_cat_ints:
                        try:
                            famdoc.Delete(scid)
                            removed += 1
                        except:
                            failed += 1
                except:
                    failed += 1
        finally:
            t.Commit()

        return removed, failed

    def _purge_unused_patterns_and_assets(self, famdoc):
        line_patterns = list(FilteredElementCollector(famdoc).OfClass(LinePatternElement).ToElements())
        fill_patterns = list(FilteredElementCollector(famdoc).OfClass(FillPatternElement).ToElements())
        assets = list(FilteredElementCollector(famdoc).OfClass(AppearanceAssetElement).ToElements())
        materials = list(FilteredElementCollector(famdoc).OfClass(Material).ToElements())

        all_line_ids = set([e.Id for e in line_patterns])
        all_fill_ids = set([e.Id for e in fill_patterns])
        all_asset_ids = set([e.Id for e in assets])

        used_line_ids = set()
        used_fill_ids = set()
        used_asset_ids = set()

        for m in materials:
            try:
                aid = m.AppearanceAssetId
                if aid and aid != ElementId.InvalidElementId and aid in all_asset_ids:
                    used_asset_ids.add(aid)
            except:
                pass

        def scan_params(elem):
            try:
                for p in elem.Parameters:
                    try:
                        if p.StorageType != StorageType.ElementId:
                            continue
                        eid = p.AsElementId()
                        if not eid or eid == ElementId.InvalidElementId:
                            continue
                        if eid in all_line_ids:
                            used_line_ids.add(eid)
                        if eid in all_fill_ids:
                            used_fill_ids.add(eid)
                        if eid in all_asset_ids:
                            used_asset_ids.add(eid)
                    except:
                        pass
            except:
                pass

        for e in FilteredElementCollector(famdoc).WhereElementIsNotElementType().ToElements():
            scan_params(e)
        for t_ in FilteredElementCollector(famdoc).WhereElementIsElementType().ToElements():
            scan_params(t_)

        unused_line = list(all_line_ids - used_line_ids)
        unused_fill = list(all_fill_ids - used_fill_ids)
        unused_assets = list(all_asset_ids - used_asset_ids)

        rl = fl = rfi = ffi = ra = fa = 0

        t = Transaction(famdoc, "Purge Patterns + Assets")
        t.Start()
        try:
            for eid in unused_assets:
                try:
                    famdoc.Delete(eid)
                    ra += 1
                except:
                    fa += 1
            for eid in unused_fill:
                try:
                    famdoc.Delete(eid)
                    rfi += 1
                except:
                    ffi += 1
            for eid in unused_line:
                try:
                    famdoc.Delete(eid)
                    rl += 1
                except:
                    fl += 1
        finally:
            t.Commit()

        return rl, fl, rfi, ffi, ra, fa

    def _save_compact(self, famdoc, optional_new_name_no_ext=""):
        try:
            current_path = famdoc.PathName or ""
        except:
            current_path = ""

        if optional_new_name_no_ext:
            name = optional_new_name_no_ext
            if not name.lower().endswith(".rfa"):
                name += ".rfa"

            basepath = ""
            try:
                basepath = os.path.dirname(current_path)
            except:
                basepath = ""

            if basepath:
                save_path = os.path.join(basepath, name)
            else:
                save_path = forms.save_file(file_ext="rfa", title="Save Family As (Compact)")
        else:
            save_path = current_path
            if not save_path:
                save_path = forms.save_file(file_ext="rfa", title="Save Family (Compact)")

        if not save_path:
            return None

        opt = SaveAsOptions()
        opt.OverwriteExistingFile = True
        try:
            opt.Compact = True
        except:
            pass
        try:
            opt.MaximumBackups = 1
        except:
            pass

        try:
            famdoc.SaveAs(save_path, opt)
            return save_path
        except:
            return None

    def close_window(self, sender, args):
        self.Close()


FamilyPurifyWindow()
