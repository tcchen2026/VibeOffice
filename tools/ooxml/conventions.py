"""What do our saves write that Office never writes? Schema-valid markup Office rejects (a partial preset
adjustment list, a master and layout sharing an ID) passed the Open XML SDK and LibreOffice; this finds
candidates by difference instead of by rule.

    python3 tools/ooxml/conventions.py ORIGINALS SAVED [--out report.json] [--min 1]

ORIGINALS: the corpus folder (the saved files' sources, matched by name). Office-authored originals
(docProps/app.xml names a Microsoft application) define the conventions; the rest are ignored there.
SAVED: our saved files. A signature counts against a saved file only if its own original lacks it, so
markup carried over from the source is not blamed on the writer. Signatures:
  attr    element E carries attribute A                    (Office never writes A on E)
  child   element E contains child C                       (never inside E)
  order   in E, child B comes before child C               (Office writes C before B, never B before C)
  value   attribute A of E has value V                     (Office uses a small set of values, V not in it)
  gdset   a preset shape lists these adjustment names      (never this set for this preset)
  ctype   content type T is given by Default or Override   (Office does it the other way)
  rel     a part of kind K has a relationship of type R    (never from that kind of part)
Standard library only.
"""
import argparse
import collections
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ProcessPoolExecutor

PREFIX = {
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w',
    'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
    'http://schemas.openxmlformats.org/presentationml/2006/main': 'p',
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main': 'x',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
    'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing': 'wp',
    'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing': 'xdr',
    'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
    'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
    'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
    'http://schemas.openxmlformats.org/package/2006/content-types': 'ct',
    'http://schemas.openxmlformats.org/package/2006/relationships': 'pr',
    'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties': 'ep',
    'http://schemas.openxmlformats.org/package/2006/metadata/core-properties': 'cp',
    'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties': 'op',
    'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes': 'vt',
    'http://purl.org/dc/elements/1.1/': 'dc', 'http://purl.org/dc/terms/': 'dcterms',
    'http://www.w3.org/2001/XMLSchema-instance': 'xsi', 'http://www.w3.org/XML/1998/namespace': 'xml',
    'urn:schemas-microsoft-com:vml': 'v', 'urn:schemas-microsoft-com:office:office': 'o',
    'urn:schemas-microsoft-com:office:excel': 'xv', 'urn:schemas-microsoft-com:office:word': 'w10',
    'http://schemas.microsoft.com/office/word/2010/wordml': 'w14', 'http://schemas.microsoft.com/office/powerpoint/2010/main': 'p14',
    'http://schemas.microsoft.com/office/drawing/2010/main': 'a14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main': 'x14',
}
VALUE_DOMAIN = 40          # attributes with at most this many distinct Office values are treated as enumerations
SKIP_VALUE = re.compile(r'\d|^\{|^[0-9A-Fa-f-]{8,}$')   # numbers, GUIDs and ids are not enumerations


def q(tag):
    if tag[0] != '{': return tag
    ns, local = tag[1:].split('}')
    return PREFIX.get(ns, '{' + ns + '}') + ':' + local


def kind(name):
    return re.sub(r'\d+', 'N', name)


def office(z):
    try:
        app = z.read('docProps/app.xml').decode('utf8', 'replace')
    except Exception:
        return False
    m = re.search(r'<Application>([^<]*)', app)
    return bool(m and 'Microsoft' in m.group(1))


def signatures(path):
    """the set of signatures of one package (None when it cannot be read)"""
    out = set()
    try:
        z = zipfile.ZipFile(path)
        names = z.namelist()
    except Exception:
        return None, False
    is_office = office(z)
    for name in names:
        if not (name.endswith('.xml') or name.endswith('.rels') or name.endswith('.vml')): continue
        try:
            root = ET.fromstring(z.read(name))
        except Exception:   # damaged parts (fuzzing test files) are left out
            continue
        pk = kind(name)
        if name == '[Content_Types].xml':
            for el in root:
                ext = el.get('Extension') or os.path.splitext(el.get('PartName') or '')[1][1:].lower()
                out.add(('ctype', q(el.tag)[3:], (el.get('ContentType') or '').rsplit('/', 1)[-1] + ' .' + ext))
            continue
        if name.endswith('.rels'):
            src = kind(re.sub(r'_rels/([^/]*)\.rels$', r'\1', name))
            for el in root:
                if el.get('TargetMode') != 'External':
                    out.add(('rel', src, (el.get('Type') or '').rsplit('/', 1)[-1]))
            continue
        for el in root.iter():
            if not isinstance(el.tag, str): continue
            e = q(el.tag)
            for a, v in el.attrib.items():
                an = q(a)
                out.add(('attr', e, an))
                if len(v) <= 40 and not SKIP_VALUE.search(v): out.add(('value', e + '@' + an, v))
            kids = [q(c.tag) for c in el if isinstance(c.tag, str)]
            for c in set(kids): out.add(('child', e, c))
            seen = []
            for c in kids:
                if not seen or seen[-1] != c: seen.append(c)
            for i, b in enumerate(seen):
                for c in seen[i + 1:]:
                    if b != c: out.add(('order', e, b + ' < ' + c))
            if e == 'a:prstGeom':
                av = el.find('{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
                gds = sorted(g.get('name') for g in av) if av is not None else []
                if gds: out.add(('gdset', el.get('prst'), ' '.join(gds)))
        out.add(('part', pk, ''))
    return out, is_office


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('originals'); ap.add_argument('saved')
    ap.add_argument('--out'); ap.add_argument('--min', type=int, default=1, help='report signatures in at least this many saved files')
    args = ap.parse_args()
    saved = sorted(f for f in os.listdir(args.saved) if os.path.isfile(os.path.join(args.saved, f)))
    originals = sorted(os.listdir(args.originals))
    stem = lambda f: os.path.splitext(f)[0]
    by_stem = {stem(f): f for f in originals}
    with ProcessPoolExecutor() as pool:
        orig = dict(zip(originals, pool.map(signatures, [os.path.join(args.originals, f) for f in originals], chunksize=8)))
        ours = dict(zip(saved, pool.map(signatures, [os.path.join(args.saved, f) for f in saved], chunksize=8)))
    office_sigs = collections.Counter()
    n_office = 0
    for sigs, is_office in orig.values():
        if sigs is None or not is_office: continue
        n_office += 1
        office_sigs.update(sigs)
    # value domains: enumerations only
    domains = collections.defaultdict(set)
    for (t, k, v) in office_sigs:
        if t == 'value': domains[k].add(v)
    # enumerations: few values, all identifiers (left, pageBottom…), not colours, fonts or names
    enum = {k for k, vs in domains.items() if len(vs) <= VALUE_DOMAIN and all(re.fullmatch(r'[a-z][A-Za-z]*', v) for v in vs)}
    present = set(office_sigs)
    order_seen = {(k, v) for (t, k, v) in present if t == 'order'}
    found = collections.defaultdict(list)
    unmatched = 0
    for f, (sigs, _) in ours.items():
        if sigs is None: continue
        src = by_stem.get(stem(f))
        own = orig.get(src, (set(), False))[0] or set() if src else set()
        if not src: unmatched += 1
        for s in sigs - own:
            t, k, v = s
            if s in present: continue
            if t == 'value' and k not in enum: continue
            if t == 'order':
                b, c = v.split(' < ')
                if (k, c + ' < ' + b) not in order_seen: continue   # Office never writes both: not an inversion
            if t == 'part': continue
            if t == 'ctype':
                other = ('Override' if k == 'Default' else 'Default', v)
                if ('ctype',) + other not in present: continue      # an extension Office never uses either way
            found[s].append(f)
    rows = sorted(found.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    rows = [r for r in rows if len(r[1]) >= args.min]
    print(f'{n_office} Office-authored originals define the conventions; {len(ours)} saved files ({unmatched} without an original)')
    print(f'{len(rows)} signatures our saves introduce and Office never writes:\n')
    for (t, k, v), files in rows:
        extra = f'  (Office values: {", ".join(sorted(domains[k])[:12])})' if t == 'value' else ''
        print(f'{len(files):5}  {t:6} {k}  {v}{extra}   e.g. {files[0]}')
    if args.out:
        with open(args.out, 'w') as fh:
            json.dump([{'type': t, 'key': k, 'value': v, 'files': files} for (t, k, v), files in rows], fh, indent=1)


if __name__ == '__main__':
    sys.exit(main())
