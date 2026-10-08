"""Generate child-order tables from the official ECMA-376 Transitional XSD archive.

python3 tools/ooxml/schema-order.py /path/to/part4.zip public/common/opc-order.js /path/to/part1.zip
Sources: https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip
         https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip
         (Part 1 for the preset shapes' adjustment values: PowerPoint repairs a shape whose list is incomplete)
Only ordering facts are emitted; the upstream archive stays outside the repository.
"""
import hashlib
import io
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

XS = '{http://www.w3.org/2001/XMLSchema}'
SELECT = {
    'w': ['CT_RPr', 'CT_PPr', 'CT_SectPr', 'CT_Settings', 'CT_TblPr', 'CT_TcPr', 'CT_TrPr'],
    'a': ['CT_GroupShapeProperties', 'CT_TextBody', 'CT_ShapeProperties', 'CT_EffectList', 'CT_TextBodyProperties', 'CT_TextParagraphProperties', 'CT_TextCharacterProperties', 'CT_Blip', 'CT_BlipFillProperties', 'CT_NonVisualDrawingProps'],
    'wp': ['CT_Anchor', 'CT_Inline'],
    'p': ['CT_Shape', 'CT_Picture', 'CT_GroupShape', 'CT_GraphicalObjectFrame', 'CT_Connector', 'CT_Presentation', 'CT_PresentationProperties', 'CT_ViewProperties', 'CT_Slide', 'CT_CommonSlideData', 'CT_SlideLayout', 'CT_NotesMaster', 'CT_NotesSlide', 'CT_SlideMaster'],
    's': ['CT_Stylesheet', 'CT_Worksheet', 'CT_Workbook', 'CT_Table', 'CT_TableColumn', 'CT_QueryTableRefresh', 'CT_PivotCacheDefinition'],
}
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}

def main():
    source, output, part1 = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3])
    data = source.read_bytes()
    with zipfile.ZipFile(io.BytesIO(data)) as outer:
        archive = outer.read('OfficeOpenXML-XMLSchema-Transitional.zip')
    types, groups, simple, namespaces = {}, {}, {}, {}
    with zipfile.ZipFile(io.BytesIO(archive)) as z:
        for name in z.namelist():
            if not name.endswith('.xsd'): continue
            text = z.read(name)
            ns = dict(event for _, event in ET.iterparse(io.BytesIO(text), events=['start-ns']))
            root = ET.fromstring(text); uri = root.get('targetNamespace')
            namespaces[uri] = ns
            for el in root:
                if el.tag in (XS + 'complexType', XS + 'group'):
                    (types if el.tag == XS + 'complexType' else groups)[(uri, el.get('name'))] = el
                elif el.tag == XS + 'simpleType': simple[(uri, el.get('name'))] = el
    def qname(name, uri):
        prefix, sep, local = name.partition(':')
        return (namespaces[uri][prefix], local) if sep else (uri, name)
    def label(key):
        uri, name = key
        prefix = next((k for k, v in NAMESPACES.items() if v == uri), None)
        return prefix + ':' + name if prefix else '{' + uri + '}' + name
    def particle(el, uri, seen):
        tag = el.tag.removeprefix(XS)
        if tag == 'element':
            return [[label(qname(el.get('ref'), uri) if el.get('ref') else (uri, el.get('name')))]]
        if tag == 'group' and el.get('ref'):
            key = qname(el.get('ref'), uri)
            if ('group', key) in seen: raise ValueError('Recursive ordering group: ' + str(key))
            return particle(groups[key], key[0], seen | {('group', key)})
        if tag == 'extension':
            key = qname(el.get('base'), uri)
            base = particle(types[key], key[0], seen | {('type', key)}) if key in types else []
            return base + [slot for c in el for slot in particle(c, uri, seen)]
        if tag in ('sequence', 'complexType', 'complexContent', 'group', 'restriction'):
            return [slot for c in el for slot in particle(c, uri, seen)]
        if tag in ('choice', 'all'):
            alternatives = [name for c in el for slot in particle(c, uri, seen) for name in slot]
            return [list(dict.fromkeys(alternatives))] if alternatives else []
        return []
    table = {}
    for prefix, names in SELECT.items():
        for name in names:
            key = (NAMESPACES[prefix], name)
            table[prefix + ':' + name] = particle(types[key], key[0], {('type', key)})
    # Office uses an ordered run-property sequence, including its 2010 effects,
    # where ECMA's base XSD has a repeated choice. Match the Open XML SDK's
    # w:CT_RPr/w:rPr particle in data/schemas/
    # schemas_openxmlformats_org_wordprocessingml_2006_main.json, SDK commit
    # 431ab05cf160248cc3885a4a766026d4f8243792. Otherwise adding
    # bold to a run containing only w14:glow puts bold after the effects, which
    # Word's validator rejects. Keep unknown extensions at their original slot.
    run_base = [name for slot in table['w:CT_RPr'] for name in slot if name != 'w:rPrChange']
    run_effects = ['glow', 'shadow', 'reflection', 'textOutline', 'textFill', 'scene3d',
                   'props3d', 'ligatures', 'numForm', 'numSpacing', 'stylisticSets', 'cntxtAlts']
    table['w:CT_RPr'] = [[name] for name in run_base + ['w14:' + name for name in run_effects] + ['w:rPrChange']]
    limits = {}
    for kind, prefix, name in [('shape', 'a', 'ST_DrawingElementId'), ('docPr', 'a', 'ST_DrawingElementId'), ('timing', 'p', 'ST_TLTimeNodeID'), ('sldId', 'p', 'ST_SlideId'), ('sldMasterId', 'p', 'ST_SlideMasterId'), ('sldLayoutId', 'p', 'ST_SlideLayoutId')]:
        restriction = simple[(NAMESPACES[prefix], name)].find(XS + 'restriction')
        assert restriction.get('base') == 'xsd:unsignedInt'
        bound = {'min': 0, 'max': 4294967295}
        for facet in restriction:
            tag, value = facet.tag.removeprefix(XS), int(facet.get('value'))
            if tag == 'minInclusive': bound['min'] = value
            if tag == 'maxInclusive': bound['max'] = value
            if tag == 'maxExclusive': bound['max'] = value - 1
        limits[kind] = bound
    # Word's tracked-change IDs use the SDK's Int32Value (CT_TrackChange@id).
    limits['revision'] = {'min': 0, 'max': 2147483647}
    # every preset shape's adjustment values, in order, with their defaults (presetShapeDefinitions.xml)
    data1 = part1.read_bytes()
    with zipfile.ZipFile(io.BytesIO(data1)) as outer, zipfile.ZipFile(io.BytesIO(outer.read('OfficeOpenXML-DrawingMLGeometries.zip'))) as g:
        presets = ET.fromstring(g.read('presetShapeDefinitions.xml'))
    adjust = {}
    for shape in presets:
        av = shape.find('{' + NAMESPACES['a'] + '}avLst')
        gds = [] if av is None else [(gd.get('name'), gd.get('fmla')) for gd in av]
        assert all(f.startswith('val ') for _, f in gds), shape.tag
        if gds: adjust[shape.tag] = [[n, int(f[4:])] for n, f in gds]
    result = {'namespaces': NAMESPACES, 'types': table, 'limits': limits, 'presetAdjust': adjust}
    text = '/* Generated by tools/ooxml/schema-order.py from ECMA-376 Part 4, fifth edition.\n'
    text += ' * Source archive SHA-256: ' + hashlib.sha256(data).hexdigest() + '\n'
    text += ' * Preset adjustments from ECMA-376 Part 1, fifth edition, SHA-256: ' + hashlib.sha256(data1).hexdigest() + '\n */\n'
    text += '/* Word run properties include the Open XML SDK Office 2010 sequence; see the generator. */\n'
    text += '(function (root) { root.L.opc.schema = ' + json.dumps(result, separators=(',', ':')) + '; })(typeof window !== "undefined" ? window : globalThis);\n'
    output.write_text(text)
    print(str(len(table)) + ' schema types; ' + str(len(text.encode())) + ' bytes')

if __name__ == '__main__':
    main()
