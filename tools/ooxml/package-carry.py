"""Independently check package features, their bytes, rIds and intended targets.

python3 tools/ooxml/package-carry.py ORIGINALS SAVED RESULTS.jsonl OUT.jsonl
Only package features are selected here. Content controls, pivot mutation rules,
and other model-owned content have separate semantic/edit checks.
"""
import argparse
from collections import Counter
import json
from pathlib import Path
from package import Package, split

R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships/'
FEATURES = {R + tag: label for tag, label in {
    'customXml': 'custom XML', 'customXmlProps': 'custom XML properties',
    'glossaryDocument': 'glossary', 'attachedTemplate': 'attached template',
    'font': 'embedded fonts', 'webSettings': 'web settings', 'connections': 'connections',
    'externalLink': 'external links', 'handoutMaster': 'handout masters',
    'notesMaster': 'notes masters', 'commentAuthors': 'comment authors',
    'tags': 'tags', 'custom-properties': 'custom properties',
}.items()}
FEATURES.update({
    'http://schemas.microsoft.com/office/2006/relationships/vbaProject': 'VBA',
    'http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels': 'labels',
    'http://schemas.microsoft.com/office/2017/10/relationships/rdRichValue': 'rich values',
    'http://schemas.microsoft.com/office/2017/10/relationships/person': 'persons',
    'http://schemas.microsoft.com/office/2011/relationships/people': 'people',
    'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds': 'comment identities',
    'http://schemas.microsoft.com/office/2011/relationships/webextension': 'web extensions',
    'http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes': 'web extension panes',
    'http://schemas.microsoft.com/office/2011/relationships/model': 'data model',
})
# Backlinks into regenerated or merged parts are checked as edges but do not turn
# those parts into opaque content. Classify by namespace/root, not a fixed path.
OWNED = {
    'wordprocessingml': {'document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments', 'styles', 'numbering', 'settings', 'fonts'},
    'spreadsheetml': {'workbook', 'worksheet', 'chartsheet', 'styleSheet', 'sst', 'table', 'comments', 'metadata'},
    'presentationml': {'presentation', 'sld', 'sldMaster', 'sldLayout', 'notes', 'presentationPr'},
}

def canonical_type(type_):
    if type_.startswith(STRICT):
        tag = type_[len(STRICT):]
        return R + {'customProperties': 'custom-properties', 'extendedProperties': 'extended-properties'}.get(tag, tag)
    return type_

def check(a, b, owner, rid, relation):
    problems, seen = [], set()
    def edge(part, id_, rel):
        other = b.rels.get(part, {}).get(id_)
        if other != rel:
            problems.append(dict(what='changed-reference', part=part, id=id_, original=rel, saved=other))
        if not rel['external']:
            visit(rel['target'])
    def visit(part):
        if part in seen:
            return
        seen.add(part)
        if part not in a.data:
            problems.append(dict(what='missing-original-dependency', part=part)); return
        if part not in b.data:
            problems.append(dict(what='missing-part', part=part)); return
        if a.types.get(part) != b.types.get(part):
            problems.append(dict(what='changed-content-type', part=part))
        root = a.xml.get(part)
        ns, tag = split(root.tag) if root is not None else ('', '')
        if any(family in ns and tag in tags for family, tags in OWNED.items()):
            return
        if canonical_type(relation['type']) == R + 'custom-properties' and part == relation['target']:
            if root is None or part not in b.xml or a.canonical(part, root) != b.canonical(part, b.xml[part]):
                problems.append(dict(what='changed-property-types-or-values', part=part))
        elif a.data[part] != b.data[part]:
            problems.append(dict(what='changed-opaque-bytes', part=part))
        for id_, rel in a.rels.get(part, {}).items():
            edge(part, id_, rel)
    edge(owner, rid, relation)
    return problems, len(seen)

def main():
    p = argparse.ArgumentParser(description=__doc__)
    for name in ['originals', 'saved', 'results', 'out']: p.add_argument(name)
    args = p.parse_args()
    rows = [json.loads(line) for line in Path(args.results).read_text().splitlines() if line.strip()]
    totals = Counter()
    with open(args.out, 'w') as out:
        for row in rows:
            report = {k: row[k] for k in ['file', 'scenario', 'status'] if k in row}
            try:
                if row['status'] != 'ok':
                    report['error'] = row.get('error', 'Driver did not complete')
                else:
                    a = Package(Path(args.originals) / row['file'])
                    b = Package(Path(args.saved) / row['scenario'] / row['file'])
                    checks = []
                    for owner, rels in a.rels.items():
                        if owner and owner not in b.data: continue
                        for rid, rel in rels.items():
                            feature = FEATURES.get(canonical_type(rel['type']))
                            if not feature: continue
                            problems, parts = check(a, b, owner, rid, rel)
                            checks.append(dict(feature=feature, owner=owner, id=rid, parts=parts, problems=problems))
                    report.update(checks=checks, status='failed' if any(c['problems'] for c in checks) else 'ok')
            except Exception as error:
                report.update(status='failed', error=str(error))
            totals[report['status']] += 1
            out.write(json.dumps(report) + '\n'); out.flush()
    print(json.dumps(totals))
    if totals['failed']: raise SystemExit(1)

if __name__ == '__main__': main()
