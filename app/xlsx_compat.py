"""Normalize and validate the XLSX package emitted by the spreadsheet runtime.

This edits package metadata and equivalent string encoding only. It does not
calculate formulas or author workbook data/styles.
"""
from __future__ import annotations

import os
from pathlib import Path
import posixpath
import tempfile
import xml.etree.ElementTree as ET
import zipfile

MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
REL='http://schemas.openxmlformats.org/package/2006/relationships'
DOCREL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WORKBOOK_TYPE='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
WORKSHEET_TYPE='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
NS={'s':MAIN,'r':DOCREL,'p':REL,'c':CT}
ET.register_namespace('x',MAIN)
ET.register_namespace('r',DOCREL)


def xml_bytes(root, namespace=None):
    if namespace:
        ET.register_namespace('',namespace)
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)


def read_package(path):
    with zipfile.ZipFile(path) as archive:
        names=archive.namelist()
        if len(names)!=len(set(names)) or archive.testzip():
            raise ValueError('Invalid or damaged XLSX ZIP package.')
        return {name:archive.read(name) for name in names}


def relationship_target(rels_name, target):
    if target.startswith('/'):
        return target.lstrip('/')
    if rels_name=='_rels/.rels':
        base=''
    else:
        base=posixpath.dirname(posixpath.dirname(rels_name))
    return posixpath.normpath(posixpath.join(base,target))


def validate_xlsx(path):
    parts=read_package(path)
    required={'[Content_Types].xml','_rels/.rels','xl/workbook.xml','xl/_rels/workbook.xml.rels','xl/styles.xml'}
    if not required.issubset(parts):
        raise ValueError('Missing required Excel package parts.')
    roots={name:ET.fromstring(raw) for name,raw in parts.items() if name.endswith(('.xml','.rels'))}
    types=roots['[Content_Types].xml']
    overrides={n.get('PartName'):n.get('ContentType') for n in types.findall('c:Override',NS)}
    # Excel's file-format detection expects an explicit workbook part type.
    if overrides.get('/xl/workbook.xml')!=WORKBOOK_TYPE:
        raise ValueError('Missing explicit XLSX workbook content-type declaration.')
    defaults={n.get('Extension'):n.get('ContentType') for n in types.findall('c:Default',NS)}
    if defaults.get('xml')!='application/xml':
        raise ValueError('Nonstandard default XML content type.')
    for name,root in roots.items():
        if name.endswith('.rels'):
            seen=set()
            for rel in root:
                if rel.get('Id') in seen:
                    raise ValueError('Duplicate relationship ID.')
                seen.add(rel.get('Id'))
                if rel.get('TargetMode')!='External':
                    if relationship_target(name,rel.get('Target','')) not in parts:
                        raise ValueError('Relationship points to a missing part.')
    root_rels=roots['_rels/.rels']
    if not any(r.get('Type')==DOCREL+'/officeDocument' and
               relationship_target('_rels/.rels',r.get('Target',''))=='xl/workbook.xml' for r in root_rels):
        raise ValueError('Missing root workbook relationship.')
    wb=roots['xl/workbook.xml']
    views=wb.findall('s:bookViews/s:workbookView',NS)
    rels={r.get('Id'):r for r in roots['xl/_rels/workbook.xml.rels']}
    names=[]
    for sheet in wb.findall('s:sheets/s:sheet',NS):
        name=sheet.get('name')
        if name in names:
            raise ValueError('Duplicate worksheet name.')
        names.append(name)
        rel=rels.get(sheet.get('{'+DOCREL+'}id'))
        if rel is None or rel.get('Type')!=DOCREL+'/worksheet':
            raise ValueError('Invalid worksheet relationship.')
        target=relationship_target('xl/_rels/workbook.xml.rels',rel.get('Target',''))
        if overrides.get('/'+target)!=WORKSHEET_TYPE:
            raise ValueError('Missing worksheet content type.')
        sheet_root=roots[target]
        for view in sheet_root.findall('s:sheetViews/s:sheetView',NS):
            if int(view.get('workbookViewId','0'))>=len(views):
                raise ValueError('Worksheet refers to a nonexistent workbook view.')
        for cell in sheet_root.findall('.//s:c',NS):
            if cell.get('t')=='str' and cell.find('s:f',NS) is None:
                raise ValueError('Literal text uses formula-result string encoding.')
            if cell.get('t')=='inlineStr' and cell.find('s:is',NS) is None:
                raise ValueError('Invalid inline string.')
    if not names:
        raise ValueError('Workbook has no worksheets.')
    return names


def normalize_xlsx(path):
    path=Path(path)
    parts=read_package(path)
    types=ET.fromstring(parts['[Content_Types].xml'])
    xml_default=next((n for n in types if n.tag=='{'+CT+'}Default' and n.get('Extension')=='xml'),None)
    if xml_default is None:
        xml_default=ET.SubElement(types,'{'+CT+'}Default',Extension='xml')
    old_default=xml_default.get('ContentType','application/xml')
    existing={n.get('PartName') for n in types if n.tag=='{'+CT+'}Override'}
    # Preserve any parts that previously depended on the old XML default.
    for name in parts:
        if name.endswith('.xml') and name!='[Content_Types].xml' and '/'+name not in existing:
            kind=WORKBOOK_TYPE if name=='xl/workbook.xml' else old_default
            ET.SubElement(types,'{'+CT+'}Override',PartName='/'+name,ContentType=kind)
    xml_default.set('ContentType','application/xml')
    parts['[Content_Types].xml']=xml_bytes(types,CT)

    wb=ET.fromstring(parts['xl/workbook.xml'])
    views=wb.find('s:bookViews',NS)
    if views is None:
        views=ET.Element('{'+MAIN+'}bookViews')
        sheets=wb.find('s:sheets',NS)
        wb.insert(list(wb).index(sheets),views)
    if not list(views):
        ET.SubElement(views,'{'+MAIN+'}workbookView',activeTab='0',firstSheet='0')
    parts['xl/workbook.xml']=xml_bytes(wb)
    for name,raw in list(parts.items()):
        if not name.startswith('xl/worksheets/') or not name.endswith('.xml'):
            continue
        root=ET.fromstring(raw)
        changed=False
        for cell in root.findall('.//s:c',NS):
            if cell.get('t')=='str' and cell.find('s:f',NS) is None:
                value=cell.find('s:v',NS)
                content='' if value is None else (value.text or '')
                if value is not None:
                    cell.remove(value)
                cell.set('t','inlineStr')
                inline=ET.SubElement(cell,'{'+MAIN+'}is')
                text=ET.SubElement(inline,'{'+MAIN+'}t')
                text.text=content
                if content!=content.strip():
                    text.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
                changed=True
        if changed:
            parts[name]=xml_bytes(root)
    handle,temp=tempfile.mkstemp(prefix='xlsx-',suffix='.xlsx',dir=path.parent)
    os.close(handle)
    temp=Path(temp)
    try:
        with zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as archive:
            # Conventional Office package order; original cell/style data stays intact.
            order=['[Content_Types].xml','_rels/.rels']+[n for n in parts if n not in ('[Content_Types].xml','_rels/.rels')]
            for name in order:
                archive.writestr(name,parts[name])
        validate_xlsx(temp)
        os.replace(temp,path)
    finally:
        temp.unlink(missing_ok=True)
