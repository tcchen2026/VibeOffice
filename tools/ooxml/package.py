"""Independent OOXML package checks (Python XML/ZIP, not the application's readers).

check FILE...; compare ORIGINAL SAVED [--policy POLICY.json]
Policy: {parts: {originalName: {target, mode: opaque|merged|regenerated,
         allow: [expanded-name ElementTree paths]}}, removed: [names], expectedSlides: N}
Allow paths apply to both versions. An allowed attribute uses path/@{namespace}name.
There are no implicit loss allowances. Output is JSONL, and failures exit nonzero.
"""
import argparse
from collections import Counter, defaultdict
import copy
import hashlib
import io
import json
from pathlib import Path
import posixpath
import sys
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET
import zipfile

R = {'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
     'http://purl.oclc.org/ooxml/officeDocument/relationships'}
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
W = {'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main'}
P = {'http://schemas.openxmlformats.org/presentationml/2006/main', 'http://purl.oclc.org/ooxml/presentationml/main'}
A = {'http://schemas.openxmlformats.org/drawingml/2006/main', 'http://purl.oclc.org/ooxml/drawingml/main'}
S = {'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main'}

def split(tag):
    if tag.startswith('{'):
        return tuple(tag[1:].split('}', 1))
    return '', tag

def at(el, name, default=None):
    return next((v for k, v in el.attrib.items() if split(k)[1] == name), default)

def rels_owner(name):
    if name == '_rels/.rels':
        return ''
    folder, base = posixpath.split(name)
    if posixpath.basename(folder) != '_rels' or not base.endswith('.rels'):
        raise ValueError('Invalid relationship part name: ' + name)
    return posixpath.join(posixpath.dirname(folder), base[:-5])

def target_of(owner, target):
    u = urlsplit(target)
    if u.scheme or u.netloc or u.query:
        raise ValueError('Internal relationship target is not a part URI: ' + target)
    path = unquote(u.path)
    name = owner if not path else posixpath.normpath(path.lstrip('/') if path.startswith('/') else posixpath.join(posixpath.dirname(owner), path))
    if name.startswith('../') or name in ('..', '.', '') or '\\' in name:
        raise ValueError('Internal target escapes the package: ' + target)
    return name

def walk(root):
    """Every branch, retaining which alternatives are mutually exclusive."""
    def rec(el, path, alternatives):
        yield el, path, alternatives
        counts = Counter()
        for child in el:
            counts[child.tag] += 1
            cp = f'{path}/{child.tag}[{counts[child.tag]}]'
            branches = alternatives
            if el.tag == '{' + MC + '}AlternateContent' and split(child.tag)[1] in ('Choice', 'Fallback'):
                branches = dict(alternatives, **{path: cp})
            yield from rec(child, cp, branches)
    yield from rec(root, '/' + root.tag + '[1]', {})

def compatible(a, b):
    return all(k not in b or b[k] == v for k, v in a.items())

def parse_xml(data):
    # ElementTree expands element/attribute names, but not QName-valued MC attributes.
    # Expand those values while namespace declarations are still available.
    scopes, pending = [], []
    iterator = ET.iterparse(io.BytesIO(data), events=('start-ns', 'start', 'end'))
    for event, value in iterator:
        if event == 'start-ns':
            pending.append(value)
        elif event == 'start':
            scope = dict(scopes[-1]) if scopes else {'xml': 'http://www.w3.org/XML/1998/namespace'}
            scope.update(pending); pending = []; scopes.append(scope)
            for key, val in list(value.attrib.items()):
                ns, name = split(key)
                if ns == MC or (value.tag == '{' + MC + '}Choice' and name == 'Requires') or key == '{http://www.w3.org/2001/XMLSchema-instance}type':
                    def expanded(token):
                        prefix, colon, local = token.partition(':')
                        uri = scope.get(prefix)
                        if uri is None:
                            raise ET.ParseError('Undeclared prefix in ' + key + ': ' + prefix)
                        return '{' + uri + '}' + (local if colon else '')
                    value.set(key, ' '.join(expanded(t) for t in val.split()))
        else:
            scopes.pop()
    return iterator.root

class Package:
    def __init__(self, file):
        self.file = str(file)
        self.issues = []
        self.data, self.xml, self.rels, self.types = {}, {}, {}, {}
        with zipfile.ZipFile(file) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                n = unquote(info.filename)
                if n in self.data:
                    self.issue('duplicate-part', n)
                if n.startswith('/') or posixpath.normpath(n) != n or n.startswith('../'):
                    self.issue('invalid-part-name', n)
                self.data[n] = z.read(info)
        for n, data in self.data.items():
            if n.endswith(('.xml', '.rels', '.vml')):
                try:
                    self.xml[n] = parse_xml(data)
                except ET.ParseError as e:
                    self.issue('invalid-xml', n, str(e))
        cts = self.xml.get('[Content_Types].xml')
        if cts is None:
            self.issue('missing-content-types', '[Content_Types].xml')
        defaults, overrides = {}, {}
        for c in cts if cts is not None else []:
            if c.tag == '{' + CT + '}Default':
                key = c.get('Extension', '').lower()
                if key in defaults:
                    self.issue('duplicate-content-type', '[Content_Types].xml', key)
                defaults[key] = c.get('ContentType')
            elif c.tag == '{' + CT + '}Override':
                key = unquote(c.get('PartName', '').lstrip('/'))
                if key in overrides:
                    self.issue('duplicate-content-type', '[Content_Types].xml', key)
                overrides[key] = c.get('ContentType')
                if key not in self.data:
                    self.issue('missing-override-target', '[Content_Types].xml', key)
        for n in self.data:
            if n == '[Content_Types].xml':
                continue
            self.types[n] = overrides.get(n, defaults.get(n.rsplit('.', 1)[-1].lower()))
            if not self.types[n]:
                self.issue('missing-content-type', n)
        for n, root in self.xml.items():
            if not n.endswith('.rels'):
                continue
            try:
                owner = rels_owner(n)
            except ValueError as e:
                self.issue('invalid-rels-name', n, str(e)); continue
            if owner and owner not in self.data:
                self.issue('missing-rels-owner', n, owner)
            rels = self.rels.setdefault(owner, {})
            for rel in root:
                id_ = rel.get('Id', '')
                if not id_ or id_ in rels:
                    self.issue('duplicate-or-empty-rid', n, id_)
                external = rel.get('TargetMode') == 'External'
                raw = rel.get('Target', '')
                try:
                    target = raw if external else target_of(owner, raw)
                except ValueError as e:
                    self.issue('invalid-target', n, str(e)); continue
                rels[id_] = dict(type=rel.get('Type'), target=target, external=external)
                if not external and target not in self.data:
                    self.issue('missing-target', n, target)

    def issue(self, code, part, detail='', path=''):
        self.issues.append(dict(code=code, part=part, path=path, detail=detail))

    def check(self):
        definitions = defaultdict(list)
        references = []
        caches = set()
        pivot_refs = []
        for part, root in self.xml.items():
            for el, path, branches in walk(root):
                ns, tag = split(el.tag)
                for key, val in el.attrib.items():
                    ans, name = split(key)
                    if (ans in R or (ans == 'urn:schemas-microsoft-com:office:office' and name == 'relid')) and val and val not in self.rels.get(part, {}):
                        self.issue('unresolved-rid', part, val, path)
                scope = part
                kind = None
                if tag == 'docPr' and 'wordprocessingDrawing' in ns:
                    kind, scope = 'docPr', 'word'
                elif tag == 'cNvPr':
                    kind = 'shape'
                elif tag == 'cTn' and ns in P:
                    kind = 'timing'
                elif tag == 'sldId' and ns in P and '/sldIdLst' in path.replace('{' + ns + '}', ''):
                    kind = 'slide'
                elif tag in ('bookmarkStart', 'permStart') and ns in W:
                    kind = 'bookmark' if tag == 'bookmarkStart' else 'permission'
                elif tag == 'pivotCache' and ns in S:
                    id_ = at(el, 'cacheId')
                    if id_ in caches:
                        self.issue('duplicate-cache-id', part, id_, path)
                    caches.add(id_)
                if kind and at(el, 'id') is not None:
                    key = (scope, kind, at(el, 'id'))
                    for prior_part, prior_path, prior_branches in definitions[key]:
                        # Branch paths include a part identifier to avoid conflating unrelated ACs.
                        both = compatible({part + k: v for k, v in branches.items()}, {prior_part + k: v for k, v in prior_branches.items()})
                        if both:
                            self.issue('duplicate-' + kind + '-id', part, key[2], path)
                    definitions[key].append((part, path, branches))
                # PresentationML spid targets cNvPr. VML o:spid is the shape's
                # own identity, not a reference into DrawingML's ID space.
                if ns in P and at(el, 'spid') is not None:
                    references.append((part, 'shape', at(el, 'spid'), path, branches))
                if tag in ('stCxn', 'endCxn') and ns in A:
                    references.append((part, 'shape', at(el, 'id'), path, branches))
                if tag == 'tn' and ns in P:
                    references.append((part, 'timing', at(el, 'val'), path, branches))
                if tag in ('bookmarkEnd', 'permEnd') and ns in W:
                    references.append((part, 'bookmark' if tag == 'bookmarkEnd' else 'permission', at(el, 'id'), path, branches))
                if tag == 'pivotTableDefinition' and ns in S:
                    pivot_refs.append((part, at(el, 'cacheId'), path))
        for scope, kind, val, path, branches in references:
            if not any(compatible(branches, bs) for _, _, bs in definitions.get((scope, kind, val), [])):
                self.issue('unresolved-' + kind + '-id', scope, str(val), path)
        for part, id_, path in pivot_refs:
            if id_ not in caches:
                self.issue('unresolved-cache-id', part, id_, path)
        return self.issues

    def canonical(self, part, root, name_map=None):
        name_map = name_map or {}
        def rec(el):
            attrs = []
            for k, v in el.attrib.items():
                ns, name = split(k)
                if ns in R and v in self.rels.get(part, {}):
                    r = self.rels[part][v]
                    v = json.dumps([r['type'], r['external'], name_map.get(r['target'], r['target'])], ensure_ascii=False)
                attrs.append((k, v))
            text = el.text or ''
            if len(el) and not text.strip():
                text = ''
            return [el.tag, sorted(attrs), text, [rec(c) + [c.tail if c.tail and c.tail.strip() else ''] for c in el]]
        return rec(root)

def allowed(root, rules):
    root = copy.deepcopy(root)
    for rule in rules:
        if '/@' in rule:
            path, attr = rule.rsplit('/@', 1)
            for el in root.findall(path):
                el.attrib.pop(attr, None)
        else:
            remove = set(root.findall(rule))
            for parent in root.iter():
                for child in list(parent):
                    if child in remove:
                        parent.remove(child)
    return root

def semantic(pkg, part, root, name_map):
    """Ordered paragraph/cell text and properties, independent of run splitting and prefixes."""
    paragraphs, cells, objects = [], [], []
    inventory = Counter()
    for el in root.iter():
        ns, tag = split(el.tag)
        if tag in ('sdt', 'object', 'OLEObject', 'control', 'graphicFrame', 'contentPart', 'model3d', 'videoFile', 'audioFile', 'glow', 'reflection', 'softEdge', 'sp3d', 'dataBar', 'ext'):
            inventory[el.tag] += 1
        if tag == 'p' and ns in W | A:
            text = []
            runs = []
            for child in el:
                if split(child.tag)[1] == 'pPr':
                    continue
                for t in child.iter():
                    tns, tn = split(t.tag)
                    if tns in W | A and tn in ('t', 'delText', 'instrText'):
                        text.append(t.text or '')
                    elif tns in W | A and tn in ('tab', 'br', 'cr'):
                        text.append('\t' if tn == 'tab' else '\n')
                if split(child.tag)[1] == 'r':
                    rp = next((r for r in child if split(r.tag)[1] == 'rPr'), None)
                    prop = pkg.canonical(part, rp, name_map) if rp is not None else None
                    txt = ''.join(t.text or '' for t in child if split(t.tag)[1] in ('t', 'delText', 'instrText'))
                    if runs and runs[-1][0] == prop:
                        runs[-1][1] += txt
                    else:
                        runs.append([prop, txt])
            pp = next((c for c in el if split(c.tag)[1] == 'pPr'), None)
            paragraphs.append(dict(text=''.join(text), properties=pkg.canonical(part, pp, name_map) if pp is not None else None, runs=runs))
        elif tag == 'c' and ns in S:
            cells.append(pkg.canonical(part, el, name_map))
        elif tag in ('spPr', 'bodyPr', 'tblPr', 'sectPr'):
            objects.append(pkg.canonical(part, el, name_map))
    return dict(paragraphs=paragraphs, cells=cells, properties=objects, features=dict(sorted(inventory.items())))

def compare(a, b, policy):
    differences = []
    modes = policy.get('parts', {})
    mapping = {n: spec.get('target', n) for n, spec in modes.items()}
    removed = set(policy.get('removed', []))
    summary = Counter()
    for name, data in a.data.items():
        if name in removed:
            continue
        target = mapping.get(name, name)
        spec = modes.get(name, {})
        mode = spec.get('mode', 'opaque')
        if target not in b.data:
            differences.append(dict(part=name, what='missing-part')); continue
        if data == b.data[target]:
            summary['byteIdentical'] += 1
            # The same XML can reference different content after its .rels was rewritten.
            if name in a.xml and target in b.xml and a.canonical(name, a.xml[name], mapping) != b.canonical(target, b.xml[target]):
                differences.append(dict(part=name, what='relationship-retargeted'))
            continue
        if mode not in ('opaque', 'merged', 'regenerated'):
            raise ValueError('Unknown comparison mode: ' + mode)
        if mode == 'opaque':
            differences.append(dict(part=name, what='changed-opaque', original=hashlib.sha256(data).hexdigest(), saved=hashlib.sha256(b.data[target]).hexdigest()))
        elif name in a.xml and target in b.xml:
            ar = allowed(a.xml[name], spec.get('allow', []))
            br = allowed(b.xml[target], spec.get('allow', []))
            ax = semantic(a, name, ar, mapping) if mode == 'regenerated' else a.canonical(name, ar, mapping)
            bx = semantic(b, target, br, {}) if mode == 'regenerated' else b.canonical(target, br)
            if ax != bx:
                fields = [k for k in ax if ax[k] != bx[k]] if isinstance(ax, dict) else ['xml']
                differences.append(dict(part=name, what='changed-' + mode, fields=fields))
            else:
                summary['equivalent'] += 1
        else:
            differences.append(dict(part=name, what='cannot-compare-' + mode))
    # Relationships are a multiset of edges; changing rIds is allowed only if XML references agree.
    for owner, rels in a.rels.items():
        if owner in removed:
            continue
        dest = mapping.get(owner, owner)
        edges = lambda rs, mp: Counter((r['type'], r['external'], mp.get(r['target'], r['target'])) for r in rs.values() if r['target'] not in removed)
        missing = edges(rels, mapping) - edges(b.rels.get(dest, {}), {})
        if missing:
            differences.append(dict(part=owner, what='missing-relationships', edges=[list(e) + [n] for e, n in missing.items()]))
    if 'expectedSlides' in policy:
        count = sum(1 for x in b.xml.values() if split(x.tag)[0] in P and split(x.tag)[1] == 'sld')
        if count != policy['expectedSlides']:
            differences.append(dict(what='slide-count', expected=policy['expectedSlides'], actual=count))
    ai = Counter(json.dumps(i, sort_keys=True) for i in a.check())
    bi = Counter(json.dumps(i, sort_keys=True) for i in b.check())
    new_issues = [json.loads(i) for i, n in (bi - ai).items() for _ in range(n)]
    return dict(original=a.file, saved=b.file, status='failed' if differences or new_issues else 'ok',
                summary=dict(summary), differences=differences, newIssues=new_issues)

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=['check', 'compare'])
    parser.add_argument('files', nargs='+')
    parser.add_argument('--policy')
    args = parser.parse_args()
    failed = False
    if args.command == 'compare':
        if len(args.files) != 2:
            parser.error('compare requires ORIGINAL SAVED')
        try:
            policy = json.loads(Path(args.policy).read_text()) if args.policy else {}
            result = compare(Package(args.files[0]), Package(args.files[1]), policy)
        except Exception as e:
            result = dict(files=args.files, status='failed', error=str(e))
        print(json.dumps(result, ensure_ascii=False))
        failed = result['status'] != 'ok'
    else:
        for file in args.files:
            try:
                pkg = Package(file)
                issues = pkg.check()
                result = dict(file=file, status='failed' if issues else 'ok', parts=len(pkg.data), issues=issues)
            except Exception as e:
                result = dict(file=file, status='failed', error=str(e))
            failed |= result['status'] != 'ok'
            print(json.dumps(result, ensure_ascii=False))
    return int(failed)

if __name__ == '__main__':
    sys.exit(main())
