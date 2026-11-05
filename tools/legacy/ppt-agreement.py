"""Compare converted PPT text/order/size with pinned source and LibreOffice references.

python3 -I tools/legacy/ppt-agreement.py RUN_DIR OUTPUT.json
Does not invoke LibreOffice or inspect its source code. Keeps individual slide
differences so ordering, field caches and actual omissions can be distinguished.
"""
import collections
import json
import pathlib
import posixpath
import re
import sys
import xml.etree.ElementTree as E
import zipfile

BASE = pathlib.Path('tmp/legacy/artifacts')
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


def package(path):
    with zipfile.ZipFile(path) as z:
        pres = E.fromstring(z.read('ppt/presentation.xml'))
        rels = {r.get('Id'): r.get('Target') for r in E.fromstring(z.read('ppt/_rels/presentation.xml.rels'))}
        ids = pres.findall('p:sldIdLst/p:sldId', NS)
        slides = []
        for n in ids:
            name = posixpath.normpath('ppt/' + rels[n.get(RID)])
            slide = E.fromstring(z.read(name))
            paragraphs = [''.join(t.text or '' for t in p.findall('.//a:t', NS))
                          for p in slide.findall('.//a:p', NS)]
            def field_text(p):
                return ''.join(('{' + (c.get('type') or 'unknown') + '}')
                               if c.tag == '{' + NS['a'] + '}fld' else
                               ''.join(t.text or '' for t in c.findall('.//a:t', NS))
                               for c in p)
            field_paragraphs = [field_text(p) for p in slide.findall('.//a:p', NS)]
            slides.append({'name': name, 'paragraphs': paragraphs,
                           'text': ''.join(paragraphs),
                           'fieldText': ''.join(field_paragraphs),
                           'fieldParagraphs': field_paragraphs,
                           'fields': [f.get('type') for f in slide.findall('.//a:fld', NS)]})
        return {'slides': slides, 'ids': [n.get('id') for n in ids],
                'size': pres.find('p:sldSz', NS).attrib}


def norm(text):
    return re.sub(r'\s+', '', text)


def main():
    run = pathlib.Path(sys.argv[1])
    rows = [json.loads(s) for s in (run / 'runs.jsonl').read_text().splitlines()]
    originals = [r for r in rows if not r['mutation']]
    source = {r['sha256']: r for r in json.load(open(BASE / 'D1-source.json'))}
    counts = collections.Counter()
    details = []
    for row in originals:
        if row['status'] != 'converted' or row.get('type') not in ('pptx', 'ppsx', 'potx'):
            continue
        path = row.get('pkg') or row.get('package') or row.get('packagePath') or run / 'packages' / (row['sha256'] + '.' + row['type'])
        out = package(path)
        src = source.get(row['sha256'], {})
        counts['convertedFiles'] += 1
        counts['outputSlides'] += len(out['slides'])
        counts['sourceSlides'] += len(src.get('slides', []))
        counts['sourceCountMatches'] += len(out['slides']) == len(src.get('slides', []))
        counts['sourceOrderMatches'] += out['ids'] == [str(s['id']) for s in src.get('slides', [])]
        counts['sourceDimensionsMatch'] += all(str(v) == out['size'].get(k) for k, v in src.get('dimensions', {}).items())
        lo = json.load(open(BASE / 'references' / row['sha256'] / 'result.json'))['conversions']['pptx']
        detail = {'file': row['file'], 'sha256': row['sha256'], 'package': str(path), 'loStatus': lo['status'], 'differences': []}
        details.append(detail)
        if lo['status'] != 'converted':
            continue
        other = package(BASE / 'references' / lo['file'])
        counts['loFiles'] += 1
        counts['loSlides'] += len(other['slides'])
        counts['loCountMatches'] += len(out['slides']) == len(other['slides'])
        for i, (a, b) in enumerate(zip(out['slides'], other['slides'])):
            counts['comparedSlides'] += 1
            same = norm(a['text']) == norm(b['text'])
            counts['textMatches'] += same
            counts['nonemptyLOSlides'] += bool(norm(b['text']))
            counts['nonemptyTextMatches'] += same and bool(norm(b['text']))
            counts['exactTextMatches'] += a['text'] == b['text']
            counts['fieldCacheIndependentMatches'] += norm(a['fieldText']) == norm(b['fieldText'])
            counts['paragraphMultisetMatches'] += collections.Counter(norm(p) for p in a['paragraphs'] if norm(p)) == collections.Counter(norm(p) for p in b['paragraphs'] if norm(p))
            if not same:
                detail['differences'].append({'slide': i + 1, 'output': a, 'lo': b})
    result = {'counts': dict(counts), 'details': details,
              'statuses': dict(collections.Counter(r['status'] for r in rows)),
              'originals': dict(collections.Counter(r['status'] for r in originals))}
    pathlib.Path(sys.argv[2]).write_text(json.dumps(result, ensure_ascii=False, indent=2))
    print(json.dumps({k: v for k, v in result.items() if k != 'details'}, indent=2))


if __name__ == '__main__':
    main()
