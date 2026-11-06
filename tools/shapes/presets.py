#!/usr/bin/env python3
"""Writes public/common/presets.js, DrawingML's 186 preset shapes (ECMA-376 presetShapeDefinitions.xml) as data:
adjust values, guide formulas, the text rectangle, adjust handles, connection sites and paths, which
common/geometry.js evaluates the way Office does; and its 40 text warps (presetTextWarpDefinitions.xml), the paths
that text is bent along.

Source: Apache OpenOffice main/oox/source/export/preset{Shape,TextWarp}Definitions.xml (Apache-2.0), in refs/openoffice
(git clone --depth 1 https://github.com/apache/openoffice.git refs/openoffice).
  python3 -I tools/shapes/presets.py
"""
import json, pathlib, re, sys
import xml.etree.ElementTree as ET

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = ROOT / 'refs/openoffice/main/oox/source/export/presetShapeDefinitions.xml'
WARPS = ROOT / 'refs/openoffice/main/oox/source/export/presetTextWarpDefinitions.xml'
OUT = ROOT / 'public/common/presets.js'


def arg(v):
    """a formula argument or a coordinate: a number, or the name of a guide"""
    return int(v) if re.fullmatch(r'-?\d+', v) else v


def strip(tag):
    return tag.split('}')[-1]


def guides(el):
    out = []
    for gd in (el if el is not None else []):
        parts = gd.get('fmla').split()
        out.append([gd.get('name'), parts[0]] + [arg(p) for p in parts[1:]])
    return out


def pos(el):
    p = next(c for c in el if strip(c.tag) == 'pos')
    return [arg(p.get('x')), arg(p.get('y'))]


def shape(el):
    s = {}
    for part in el:
        tag = strip(part.tag)
        if tag == 'avLst':
            s['av'] = guides(part)
        elif tag == 'gdLst':
            s['gd'] = guides(part)
        elif tag == 'rect':
            s['rect'] = [arg(part.get(k)) for k in 'ltrb']
        elif tag == 'ahLst':
            s['ah'] = []
            for h in part:
                g = lambda k: arg(h.get(k)) if h.get(k) is not None else None
                if strip(h.tag) == 'ahXY':
                    s['ah'].append(['xy', g('gdRefX'), g('minX'), g('maxX'), g('gdRefY'), g('minY'), g('maxY')] + pos(h))
                else:
                    s['ah'].append(['polar', g('gdRefR'), g('minR'), g('maxR'), g('gdRefAng'), g('minAng'), g('maxAng')] + pos(h))
        elif tag == 'cxnLst':
            s['cxn'] = [[arg(c.get('ang'))] + pos(c) for c in part]
        elif tag == 'pathLst':
            s['path'] = []
            for p in part:
                o = {}
                for k in ('w', 'h'):
                    if p.get(k): o[k] = int(p.get(k))
                if p.get('fill') and p.get('fill') != 'norm': o['fill'] = p.get('fill')
                if p.get('stroke') in ('false', '0'): o['stroke'] = 0
                if p.get('extrusionOk') in ('false', '0'): o['ext'] = 0
                c = []
                for cmd in p:
                    t = strip(cmd.tag)
                    pts = [v for pt in cmd for v in (arg(pt.get('x')), arg(pt.get('y')))]
                    if t == 'moveTo': c.append(['M'] + pts)
                    elif t == 'lnTo': c.append(['L'] + pts)
                    elif t == 'quadBezTo': c.append(['Q'] + pts)
                    elif t == 'cubicBezTo': c.append(['C'] + pts)
                    elif t == 'arcTo': c.append(['A'] + [arg(cmd.get(k)) for k in ('wR', 'hR', 'stAng', 'swAng')])
                    elif t == 'close': c.append(['Z'])
                    else: sys.exit('unknown path command ' + t)
                o['c'] = c
                s['path'].append(o)
    return s


def upside_down(s):
    """a shape mirrored top to bottom: each y it draws, puts a handle or a site at is measured up from the bottom"""
    s = json.loads(json.dumps(s))
    gd, made = s.setdefault('gd', []), {}

    def flip(y):
        if isinstance(y, int): y = str(y)
        if y not in made:
            made[y] = 'fy' + str(len(made))
            gd.append([made[y], '+-', 'b', 0, arg(y)])
        return made[y]
    for p in s['path']:
        for c in p['c']:
            assert c[0] != 'A', 'arcs would turn the other way'
            for i in range(2, len(c), 2): c[i] = flip(c[i])
    l, t, r, b = s['rect']
    s['rect'] = [l, flip(b), r, flip(t)]
    for h in s.get('ah', []): h[8] = flip(h[8])
    ANG = {'cd4': '3cd4', '3cd4': 'cd4'}
    s['cxn'] = [[ANG.get(a, a), x, flip(y)] for a, x, y in s.get('cxn', [])]
    return s


def main():
    root = ET.parse(SRC).getroot()
    shapes = {strip(el.tag): shape(el) for el in root}
    # the 24- and 32-point stars' handles stop at ssd2 in the definitions, a slip: their drawings pin adj to 0..50000
    # as every other star does, and so do the handles here
    for name in ('star24', 'star32'):
        assert shapes[name]['ah'][0][6] == 'ssd2'
        shapes[name]['ah'][0][6] = 50000
    # the pie's text rectangle names its guides in the wrong places (t="ir" r="it"); PowerPoint 2016 puts the text in
    # il, it, ir, ib (LibreOffice's tdf149918_pie_textarea)
    assert shapes['pie']['rect'] == ['il', 'ir', 'it', 'ib']
    shapes['pie']['rect'] = ['il', 'it', 'ir', 'ib']
    # the file has upDownArrow twice and no upArrow, which is downArrow turned upside down
    assert 'upArrow' not in shapes
    shapes['upArrow'] = upside_down(shapes['downArrow'])
    warps = {strip(el.tag): shape(el) for el in ET.parse(WARPS).getroot()}
    rows = lambda d: ',\n'.join('  %s: %s' % (name, json.dumps(s, separators=(',', ':'))) for name, s in sorted(d.items()))
    OUT.write_text('''/* VibeOffice — DrawingML's preset shapes and text warps as data (ECMA-376 presetShapeDefinitions.xml and
 * presetTextWarpDefinitions.xml), evaluated by common/geometry.js.
 * Derived from Apache OpenOffice main/oox/source/export/preset{Shape,TextWarp}Definitions.xml (Apache-2.0).
 * Generated by tools/shapes/presets.py: don't edit by hand.
 * Each shape: av (adjust values) and gd (guides) as [name, operator, ...arguments], rect (text box l, t, r, b),
 * ah (adjust handles), cxn (connection sites: angle, x, y) and path (w, h, fill, stroke: 0, ext: 0, c: commands).
 * Each warp: av, gd, ah and path, the lines its text is bent along. */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  L.presetShapes = {
%s,
  };
  L.presetWarps = {
%s,
  };
})(typeof window !== 'undefined' ? window : globalThis);
''' % (rows(shapes), rows(warps)))
    print(len(shapes), 'shapes,', len(warps), 'warps,', OUT.stat().st_size, 'bytes:', OUT.relative_to(ROOT))


main()
