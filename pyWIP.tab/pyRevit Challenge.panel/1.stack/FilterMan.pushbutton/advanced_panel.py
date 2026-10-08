# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import clr
clr.AddReference(u'System.Data')
from Autodesk.Revit.DB import FilteredElementCollector
from System.Data import DataTable


_ALL_ELEMENTS_CAP = 2000  # safety limit for 'all' mode to prevent freeze/OOM

def get_elements_for_scope(doc, mode, checked_ids, active_view_id=None):
    """
    mode 'checked': return elements whose ElementId is in checked_ids.values()
    mode 'all':     return up to _ALL_ELEMENTS_CAP non-type elements.
    """
    if mode == u'checked':
        result = []
        for eid in checked_ids.values():
            el = doc.GetElement(eid)
            if el:
                result.append(el)
        return result
    else:
        if active_view_id:
            col = FilteredElementCollector(doc, active_view_id)
        else:
            col = FilteredElementCollector(doc)
        els = list(col.WhereElementIsNotElementType().ToElements())
        return els[:_ALL_ELEMENTS_CAP]


def apply_conditions(elements, conditions):
    """Return elements that match ALL conditions (AND logic)."""
    if not conditions:
        return elements
    return [el for el in elements if _matches_all(el, conditions)]


def _matches_all(el, conditions):
    fv = _field_values(el)
    for cond in conditions:
        if not _eval(cond, fv, el):
            return False
    return True


def _field_values(el):
    vals = {}
    try:
        cat = el.Category
        vals[u'category'] = cat.Name if cat else u''
    except Exception:
        vals[u'category'] = u''
    try:
        sym = el.Document.GetElement(el.GetTypeId())
        if sym:
            try:
                fam = getattr(sym, u'Family', None)
                vals[u'family_name'] = fam.Name if fam else u''
            except Exception:
                vals[u'family_name'] = u''
            if not vals[u'family_name']:
                try:
                    from Autodesk.Revit.DB import BuiltInParameter as _BIP
                    _p = sym.get_Parameter(_BIP.SYMBOL_FAMILY_NAME_PARAM)
                    vals[u'family_name'] = (unicode(_p.AsString()).strip()
                                           if _p and _p.AsString() else u'')
                except Exception:
                    pass
            try:
                _rtn = sym.Name
                vals[u'type_name'] = unicode(_rtn).strip() if _rtn else u''
            except Exception:
                vals[u'type_name'] = u''
            if not vals[u'type_name']:
                try:
                    from Autodesk.Revit.DB import BuiltInParameter as _BIP
                    for _bp in (_BIP.SYMBOL_NAME_PARAM, _BIP.ALL_MODEL_TYPE_NAME):
                        try:
                            _p = sym.get_Parameter(_bp)
                            if _p and _p.AsString():
                                vals[u'type_name'] = unicode(_p.AsString()).strip()
                                break
                        except Exception:
                            pass
                except Exception:
                    pass
        else:
            vals[u'family_name'] = vals[u'type_name'] = u''
    except Exception:
        vals[u'family_name'] = vals[u'type_name'] = u''
    try:
        lp = el.LookupParameter(u'Level')
        vals[u'level'] = lp.AsValueString() if lp else u''
    except Exception:
        vals[u'level'] = u''
    return vals


def _eval(cond, fv, el):
    field = cond.field
    op    = cond.operator
    val   = (cond.value or u'').strip().lower()

    if field == u'parameter':
        try:
            p = el.LookupParameter(cond.param_name or u'')
            target = (p.AsString() or p.AsValueString() or u'').lower() \
                     if p else u''
        except Exception:
            target = u''
    else:
        target = fv.get(field, u'').lower()

    if op == u'has_value':
        return bool(target)
    if op == u'no_value':
        return not bool(target)
    if op == u'contains':
        return val in target
    if op == u'not_contains':
        return val not in target
    if op == u'equals':
        return target == val
    if op == u'not_equals':
        return target != val
    if op == u'starts_with':
        return target.startswith(val)
    if op == u'ends_with':
        return target.endswith(val)
    try:
        t_num = float(target)
        v_num = float(val)
        if op == u'gt':  return t_num >  v_num
        if op == u'gte': return t_num >= v_num
        if op == u'lt':  return t_num <  v_num
        if op == u'lte': return t_num <= v_num
    except (ValueError, TypeError):
        pass
    return False


def build_table_rows(elements, param_names, doc):
    """Build a DataTable for WPF DataGrid binding.

    DataTable.DefaultView as ItemsSource + Binding("[ColumnName]") is the
    reliable WPF pattern for dynamic columns in IronPython. Python dicts and
    Hashtable do not bind correctly via the WPF binding engine.
    Instance param checked first; falls back to type param when absent.
    """
    dt = DataTable()
    dt.Columns.Add(u'id')
    dt.Columns.Add(u'family_name')
    dt.Columns.Add(u'type_name')
    added = {u'id', u'family_name', u'type_name'}
    for pn in param_names:
        if pn not in added:
            dt.Columns.Add(pn)
            dt.Columns.Add(u'_f_' + pn)  # import-failure highlight marker
            added.add(pn)

    first_error = [None]
    for el in elements:
        try:
            try:
                id_val = el.Id.Value
            except AttributeError:
                id_val = el.Id.IntegerValue

            try:
                sym = doc.GetElement(el.GetTypeId())
            except Exception:
                sym = None

            try:
                fam = sym.Family if sym is not None else None
            except Exception:
                fam = None

            try:
                fam_name = fam.Name if fam is not None else u''
            except Exception:
                fam_name = u''
            # Fallback for system families (walls, ceilings) where .Family raises
            if not fam_name and sym is not None:
                try:
                    from Autodesk.Revit.DB import BuiltInParameter as _BIP
                    _p = sym.get_Parameter(_BIP.SYMBOL_FAMILY_NAME_PARAM)
                    if _p is not None:
                        _v = _p.AsString()
                        if _v:
                            fam_name = unicode(_v).strip()
                except Exception:
                    pass

            try:
                _raw_tn = sym.Name if sym is not None else u''
                type_name = unicode(_raw_tn).strip() if _raw_tn else u''
            except Exception:
                type_name = u''
            # Fallback when sym.Name returns empty (observed for system + loadable types)
            if not type_name and sym is not None:
                try:
                    from Autodesk.Revit.DB import BuiltInParameter as _BIP
                    for _bp in (_BIP.SYMBOL_NAME_PARAM, _BIP.ALL_MODEL_TYPE_NAME):
                        try:
                            _p = sym.get_Parameter(_bp)
                            if _p is not None:
                                _v = _p.AsString()
                                if _v:
                                    type_name = unicode(_v).strip()
                                    break
                        except Exception:
                            pass
                except Exception:
                    pass

            dr = dt.NewRow()
            dr[u'id']          = unicode(int(id_val))
            dr[u'family_name'] = fam_name
            dr[u'type_name']   = type_name
            for pn in param_names:
                if pn in added:
                    try:
                        p = el.LookupParameter(pn)
                        if p is None and sym is not None:
                            p = sym.LookupParameter(pn)
                        dr[pn] = (p.AsString() or p.AsValueString() or u'') if p else u''
                    except Exception:
                        dr[pn] = u''
                    dr[u'_f_' + pn] = u'0'  # default: no failure
            dt.Rows.Add(dr)
        except Exception as e:
            if first_error[0] is None:
                first_error[0] = unicode(e)
            continue

    if first_error[0] is not None and dt.Rows.Count == 0 and elements:
        raise Exception(
            u'build_table_rows: all {} elements failed. First error: {}'.format(
                len(elements), first_error[0])
        )
    return dt


def build_preview_table(preview_data, adv_columns):
    """Build DataTable from import-preview data (Excel values, no Revit read).

    _f_{pn} flag values:
      '0' = no change (writable but current value == Excel value)
      '1' = not writable (red)
      '2' = will change (green)
    """
    dt = DataTable()
    dt.Columns.Add(u'id')
    dt.Columns.Add(u'family_name')
    dt.Columns.Add(u'type_name')
    _builtin_keys = {u'family_name', u'type_name'}
    param_cols = [key for _, key in adv_columns if key not in _builtin_keys]
    added = {u'id', u'family_name', u'type_name'}
    for pn in param_cols:
        if pn not in added:
            dt.Columns.Add(pn)
            dt.Columns.Add(u'_f_' + pn)
            added.add(pn)
    for row in preview_data:
        try:
            dr = dt.NewRow()
            dr[u'id']          = unicode(row.get(u'id', u''))
            dr[u'family_name'] = row.get(u'family_name', u'')
            dr[u'type_name']   = row.get(u'type_name', u'')
            preview = row.get(u'preview', {})
            params  = row.get(u'params',  {})
            for pn in param_cols:
                if pn in added:
                    dr[pn] = unicode(params.get(pn, u''))
                    pv = preview.get(pn, {})
                    if not pv.get(u'writable', True):
                        dr[u'_f_' + pn] = u'1'   # red
                    elif pv.get(u'changed', False):
                        dr[u'_f_' + pn] = u'2'   # green
                    else:
                        dr[u'_f_' + pn] = u'0'   # no change
            dt.Rows.Add(dr)
        except Exception:
            continue
    return dt


def get_available_params(elements, max_sample=150):
    """Return {param_name: 'I'|'T'} from instance/type params of sampled elements.
    Instance params take precedence: if a param exists on both, kind = 'I'."""
    result = {}
    for el in elements[:max_sample]:
        try:
            for p in el.Parameters:
                if p.Definition and p.Definition.Name not in result:
                    result[p.Definition.Name] = u'I'
        except Exception:
            pass
        try:
            type_el = el.Document.GetElement(el.GetTypeId())
            if type_el is not None:
                for p in type_el.Parameters:
                    if p.Definition and p.Definition.Name not in result:
                        result[p.Definition.Name] = u'T'
        except Exception:
            pass
    return result
