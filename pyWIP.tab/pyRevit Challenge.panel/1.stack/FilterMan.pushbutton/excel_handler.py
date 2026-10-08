# -*- coding: utf-8 -*-
"""FilterMan v2 - excel_handler.py

Export / import .xlsx using only IronPython 2.7 stdlib (zipfile + ElementTree).
No openpyxl or xlsxwriter - those require CPython.
"""
from __future__ import unicode_literals
import zipfile
import shutil
from xml.etree import ElementTree as ET

_NS      = u'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_HDR_RGB = u'FF0696D7'   # ARGB - Revit blue, fully opaque

# ── Minimal xlsx XML parts ─────────────────────────────────────────────────

_CONTENT_TYPES = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    u'<Default Extension="xml" ContentType="application/xml"/>'
    u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    u'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    u'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    u'<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
    u'</Types>'
)

_RELS = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    u'<Relationship Id="rId1"'
    u' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    u' Target="xl/workbook.xml"/>'
    u'</Relationships>'
)

_WORKBOOK = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    u' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    u'<sheets><sheet name="FilterMan Export" sheetId="1" r:id="rId1"/></sheets>'
    u'</workbook>'
)

_WORKBOOK_RELS = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    u'<Relationship Id="rId1"'
    u' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
    u' Target="worksheets/sheet1.xml"/>'
    u'<Relationship Id="rId2"'
    u' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
    u' Target="styles.xml"/>'
    u'</Relationships>'
)

_STYLES = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    u'<fonts count="2">'
    u'<font><sz val="11"/><name val="Calibri"/></font>'
    u'<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    u'</fonts>'
    u'<fills count="3">'
    u'<fill><patternFill patternType="none"/></fill>'
    u'<fill><patternFill patternType="gray125"/></fill>'
    u'<fill><patternFill patternType="solid"><fgColor rgb="' + _HDR_RGB + u'"/></patternFill></fill>'
    u'</fills>'
    u'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    u'<cellXfs count="2">'
    u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    u'<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1">'
    u'<alignment horizontal="center"/>'
    u'</xf>'
    u'</cellXfs>'
    u'</styleSheet>'
)

# Worksheet relationship: links sheet to its table definition
_SHEET_RELS = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    u'<Relationship Id="rId1"'
    u' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"'
    u' Target="../tables/table1.xml"/>'
    u'</Relationships>'
)


# ── Helpers ────────────────────────────────────────────────────────────────

def _col_letter(n):
    """1-based column index to Excel column letter(s)."""
    s = u''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = unichr(65 + r) + s
    return s


def _esc(v):
    return (v.replace(u'&', u'&amp;')
             .replace(u'<', u'&lt;')
             .replace(u'>', u'&gt;')
             .replace(u'"', u'&quot;'))


def _is_num(s):
    """True if s can be interpreted as a number by Excel."""
    if not s:
        return False
    try:
        float(s)
        return True
    except (ValueError, TypeError):
        return False


def _parse_shared_strings(raw):
    """Parse xl/sharedStrings.xml and return list of unicode strings.

    Each <si> element is one shared string. Handles both plain <t> and
    rich-text <r><t> runs by concatenating all <t> text within the <si>.
    """
    root   = ET.fromstring(raw)
    result = []
    for si_el in root.iter(u'{' + _NS + u'}si'):
        parts = []
        for t_el in si_el.iter(u'{' + _NS + u'}t'):
            if t_el.text:
                parts.append(t_el.text)
        result.append(u''.join(parts))
    return result


def _build_sheet(display_headers, row_keys, data_rows):
    """Return xl/worksheets/sheet1.xml content as unicode string.

    Numeric string values are written as number cells (no t attribute) so
    Excel sorts and filters them correctly. Text values use inlineStr.
    xmlns:r declared on root so <tablePart r:id> is valid without local redecl.
    Ends with <tableParts> to link the sheet to table1.xml.
    """
    _R_NS = u'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    parts = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<worksheet xmlns="' + _NS + u'" xmlns:r="' + _R_NS + u'">',
        u'<sheetViews><sheetView tabSelected="1" workbookViewId="0">',
        u'<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
        u'</sheetView></sheetViews>',
        u'<cols>',
        u'<col min="1" max="1" width="12" customWidth="1"/>',
        u'<col min="2" max="200" width="22" customWidth="1"/>',
        u'</cols>',
        u'<sheetData>',
        u'<row r="1">',
    ]
    for ci, h in enumerate(display_headers, 1):
        ref = _col_letter(ci) + u'1'
        parts.append(u'<c r="' + ref + u'" t="inlineStr" s="1"><is><t>'
                     + _esc(h) + u'</t></is></c>')
    parts.append(u'</row>')

    for ri, row_dict in enumerate(data_rows, 2):
        ri_s = unicode(ri)
        parts.append(u'<row r="' + ri_s + u'">')
        for ci, key in enumerate(row_keys, 1):
            ref     = _col_letter(ci) + ri_s
            val_raw = unicode(row_dict.get(key, u''))
            if _is_num(val_raw):
                parts.append(u'<c r="' + ref + u'"><v>' + val_raw + u'</v></c>')
            else:
                parts.append(u'<c r="' + ref + u'" t="inlineStr"><is><t>'
                             + _esc(val_raw) + u'</t></is></c>')
        parts.append(u'</row>')

    parts += [
        u'</sheetData>',
        u'<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
        u'</worksheet>',
    ]
    return u''.join(parts)


def _build_table(display_headers, row_count):
    """Return xl/tables/table1.xml content - Excel Table with autoFilter."""
    n_cols   = len(display_headers)
    last_col = _col_letter(n_cols)
    last_row = max(1, row_count + 1)   # +1 for header row
    ref      = u'A1:' + last_col + unicode(last_row)

    parts = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<table xmlns="' + _NS + u'"',
        u' id="1" name="FilterManData" displayName="FilterManData"',
        u' ref="' + ref + u'" totalsRowShown="0">',
        u'<autoFilter ref="' + ref + u'"/>',
        u'<tableColumns count="' + unicode(n_cols) + u'">',
    ]
    for i, h in enumerate(display_headers, 1):
        parts.append(u'<tableColumn id="' + unicode(i) + u'" name="' + _esc(h) + u'"/>')
    parts += [
        u'</tableColumns>',
        u'<tableStyleInfo name="TableStyleMedium9"'
        u' showFirstColumn="0" showLastColumn="0"'
        u' showRowStripes="1" showColumnStripes="0"/>',
        u'</table>',
    ]
    return u''.join(parts)


# ── Public API ─────────────────────────────────────────────────────────────

def export_to_excel(data_table, adv_columns, filepath):
    """Write DataTable to a new .xlsx file with an Excel Table (for filter/sort).

    data_table:  System.Data.DataTable (from advanced_panel.build_table_rows)
    adv_columns: list of (header_str, row_key_str) from BrowserWindow._adv_columns
    filepath:    full path including .xlsx extension
    """
    # Built-in columns keep their display label (Family Name / Type Name) because
    # they are in skip_set and never used for LookupParameter.
    # Custom param columns use the raw param name (key) so that import_from_excel
    # can pass it directly to LookupParameter without stripping any suffix.
    _BUILTIN_KEYS   = frozenset([u'family_name', u'type_name'])
    display_headers = [u'ID'] + [
        h if k in _BUILTIN_KEYS else k
        for h, k in adv_columns
    ]
    row_keys        = [u'id'] + [k for _, k in adv_columns]

    rows = []
    for dr in data_table.Rows:
        row = {}
        for k in row_keys:
            try:
                v = dr[k]
                row[k] = u'' if v is None else unicode(v)
            except Exception:
                row[k] = u''
        rows.append(row)

    sheet = _build_sheet(display_headers, row_keys, rows)
    table = _build_table(display_headers, len(rows))
    tmp   = filepath + u'.tmp'
    try:
        with zipfile.ZipFile(tmp, u'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(u'[Content_Types].xml',                _CONTENT_TYPES.encode(u'utf-8'))
            zf.writestr(u'_rels/.rels',                        _RELS.encode(u'utf-8'))
            zf.writestr(u'xl/workbook.xml',                    _WORKBOOK.encode(u'utf-8'))
            zf.writestr(u'xl/_rels/workbook.xml.rels',         _WORKBOOK_RELS.encode(u'utf-8'))
            zf.writestr(u'xl/styles.xml',                      _STYLES.encode(u'utf-8'))
            zf.writestr(u'xl/worksheets/sheet1.xml',           sheet.encode(u'utf-8'))
            zf.writestr(u'xl/worksheets/_rels/sheet1.xml.rels', _SHEET_RELS.encode(u'utf-8'))
            zf.writestr(u'xl/tables/table1.xml',               table.encode(u'utf-8'))
        shutil.move(tmp, filepath)
    except Exception:
        try:
            import os
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        raise


def import_from_excel(filepath):
    """Read .xlsx - handles both FilterMan-exported (inlineStr) and
    Excel-resaved (sharedStrings) formats.

    Returns list of {u'id': int, u'params': {param_name: value_str}}.
    Family Name and Type Name columns are skipped (read-only).
    Raises ValueError on malformed files.
    """
    with zipfile.ZipFile(filepath, u'r') as zf:
        raw = zf.read(u'xl/worksheets/sheet1.xml')
        # sharedStrings.xml is written by Excel when it saves the file
        try:
            shared = _parse_shared_strings(zf.read(u'xl/sharedStrings.xml'))
        except KeyError:
            shared = []

    root = ET.fromstring(raw)

    # Build grid: {row_num_int: {col_letter_str: value_str}}
    grid = {}
    for row_el in root.iter(u'{' + _NS + u'}row'):
        r = int(row_el.get(u'r', u'0'))
        for c_el in row_el:
            ref      = c_el.get(u'r', u'')
            col      = u''.join(ch for ch in ref if ch.isalpha())
            cell_t   = c_el.get(u't', u'')
            val      = u''

            if cell_t == u's':
                # sharedStrings index - format used by Excel on save
                v_el = c_el.find(u'{' + _NS + u'}v')
                if v_el is not None and v_el.text is not None:
                    try:
                        val = shared[int(v_el.text)]
                    except (ValueError, IndexError):
                        val = u''
            elif cell_t == u'inlineStr':
                # inlineStr - format written by FilterMan export
                is_el = c_el.find(u'{' + _NS + u'}is')
                if is_el is not None:
                    t_el = is_el.find(u'{' + _NS + u'}t')
                    if t_el is not None and t_el.text:
                        val = t_el.text
            else:
                # numeric or other - read raw value
                v_el = c_el.find(u'{' + _NS + u'}v')
                if v_el is not None and v_el.text:
                    val = v_el.text

            if col:
                grid.setdefault(r, {})[col] = val

    if 1 not in grid:
        raise ValueError(u'No header row found in sheet1.xml')

    col_to_name = grid[1]
    id_col      = None
    for col, name in col_to_name.iteritems():
        if name == u'ID':
            id_col = col
            break
    if id_col is None:
        found = u', '.join(sorted(col_to_name.values())) or u'(none)'
        raise ValueError(
            u'No "ID" column found. Columns detected: {}. '
            u'Export from FilterMan first, then edit and import.'.format(found)
        )

    skip_set   = {u'ID', u'Family Name', u'Type Name'}
    param_cols = {col: name for col, name in col_to_name.iteritems()
                  if name not in skip_set}

    result = []
    for rn in sorted(grid.iterkeys()):
        if rn == 1:
            continue
        row    = grid[rn]
        raw_id = row.get(id_col, u'').strip()
        if not raw_id:
            continue
        try:
            eid = int(float(raw_id))
        except (ValueError, TypeError):
            continue
        params = {}
        for col, pn in param_cols.iteritems():
            v = row.get(col, u'').strip()
            if v:
                params[pn] = v
        result.append({u'id': eid, u'params': params})
    return result


def apply_excel_import(doc, import_data):
    """Apply parameter values from import_data inside a single Transaction.

    Returns {u'updated': int, u'skipped': int, u'errors': [str, ...]}.
    ElementId uses long() to ensure Int64 compatibility with Revit 2024+.
    """
    from Autodesk.Revit.DB import Transaction, ElementId, StorageType

    updated  = 0
    skipped  = 0
    errors   = []
    failures = {}   # {id_str: [param_name, ...]} - cells that could not be set

    t = Transaction(doc, u'FilterMan: Import from Excel')
    t.Start()
    try:
        for row in import_data:
            row_failures = []
            try:
                eid = ElementId(long(row[u'id']))
            except Exception:
                errors.append(u'Invalid ElementId: {}'.format(row[u'id']))
                skipped += 1
                continue
            el = doc.GetElement(eid)
            if el is None:
                errors.append(u'Not found: {}'.format(row[u'id']))
                skipped += 1
                continue
            for pn, val_str in row[u'params'].iteritems():
                try:
                    param = el.LookupParameter(pn)
                    if param is None or param.IsReadOnly:
                        skipped += 1
                        row_failures.append(pn)
                        continue
                    st = param.StorageType
                    if st == StorageType.String:
                        param.Set(val_str)
                    elif st == StorageType.Double:
                        param.Set(float(val_str))
                    elif st == StorageType.Integer:
                        param.Set(int(float(val_str)))
                    else:
                        skipped += 1
                        row_failures.append(pn)
                        continue
                    updated += 1
                except Exception as ex:
                    errors.append(u'{} [{}]: {}'.format(pn, row[u'id'], unicode(ex)))
                    skipped += 1
                    row_failures.append(pn)
            if row_failures:
                failures[unicode(row[u'id'])] = row_failures
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise

    return {u'updated': updated, u'skipped': skipped,
            u'errors': errors, u'failures': failures}


# ── Import preview (dry-run, no Transaction) ──────────────────────────────────

def _param_current_str(p):
    """Return current parameter value as a plain string (same repr used by export)."""
    from Autodesk.Revit.DB import StorageType
    st = p.StorageType
    if st == StorageType.String:
        return p.AsString() or u''
    return p.AsValueString() or u''


def build_preview_data(doc, import_data):
    """Dry-run: for each row check param writability and value change vs Revit.

    Returns a new list with each row enriched by:
      'family_name', 'type_name'  - read from Revit for display
      'preview': {param_name: {'writable': bool, 'changed': bool}, ...}

    No Transaction is opened.
    """
    from Autodesk.Revit.DB import ElementId
    result = []
    for row in import_data:
        enriched = {k: v for k, v in row.iteritems()}
        try:
            eid = ElementId(long(row[u'id']))
        except Exception:
            enriched[u'family_name'] = u''
            enriched[u'type_name']   = u''
            enriched[u'preview']     = {}
            result.append(enriched)
            continue
        el = doc.GetElement(eid)
        if el is None:
            enriched[u'family_name'] = u''
            enriched[u'type_name']   = u''
            enriched[u'preview']     = {}
            result.append(enriched)
            continue
        # Read family/type names — separate try blocks so system families
        # (walls, floors) that raise on .Family still yield a type_name
        try:
            sym = doc.GetElement(el.GetTypeId())
        except Exception:
            sym = None
        # family_name: .Family for loadable, BuiltInParameter fallback for system
        _fam_name_info = u''
        try:
            fam = sym.Family if sym is not None else None
            enriched[u'family_name'] = fam.Name if fam is not None else u''
            _fam_name_info = u'Family.Name={!r}'.format(enriched[u'family_name'])
        except Exception as _fe:
            enriched[u'family_name'] = u''
            _fam_name_info = u'Family RAISES:{}'.format(type(_fe).__name__)
        if not enriched[u'family_name'] and sym is not None:
            try:
                from Autodesk.Revit.DB import BuiltInParameter as _BIP
                _p = sym.get_Parameter(_BIP.SYMBOL_FAMILY_NAME_PARAM)
                if _p is not None:
                    _v = _p.AsString()
                    if _v:
                        enriched[u'family_name'] = unicode(_v).strip()
                        _fam_name_info += u'->FAM_PARAM'
            except Exception:
                pass
        # type_name: sym.Name, then BuiltInParameter fallbacks
        _sym_name_info = u''
        try:
            _raw = sym.Name if sym is not None else u''
            enriched[u'type_name'] = unicode(_raw).strip() if _raw else u''
            _sym_name_info = u'sym.Name={!r}'.format(enriched[u'type_name'])
        except Exception as _te:
            enriched[u'type_name'] = u''
            _sym_name_info = u'sym.Name RAISES:{}'.format(type(_te).__name__)
        if not enriched[u'type_name'] and sym is not None:
            try:
                from Autodesk.Revit.DB import BuiltInParameter as _BIP
                for _bp in (_BIP.SYMBOL_NAME_PARAM, _BIP.ALL_MODEL_TYPE_NAME):
                    try:
                        _p = sym.get_Parameter(_bp)
                        if _p is not None:
                            _v = _p.AsString()
                            if _v:
                                enriched[u'type_name'] = unicode(_v).strip()
                                _sym_name_info += u'->' + unicode(_bp).split(u'.')[-1]
                                break
                    except Exception:
                        pass
            except Exception:
                pass
        # Check each param
        preview = {}
        for pn, excel_val in row.get(u'params', {}).iteritems():
            try:
                p = el.LookupParameter(pn)
                if p is None or p.IsReadOnly:
                    preview[pn] = {u'writable': False, u'changed': False}
                    continue
                current = _param_current_str(p)
                preview[pn] = {
                    u'writable': True,
                    u'changed':  current != unicode(excel_val),
                }
            except Exception:
                preview[pn] = {u'writable': False, u'changed': False}
        enriched[u'preview'] = preview
        result.append(enriched)
    return result
