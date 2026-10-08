#!/usr/bin/env python3
"""Isolate batch-2 comment repair without reserializing unrelated XML.

python3 tools/quire/test/comment-office-copies.py BATCH PREVIOUS_BATCH OUTPUT
These are diagnostic copies, not proposed preservation behavior.
"""
import argparse
import hashlib
import json
from pathlib import Path
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET


def edit(parts, unique=False, legacy=False, none=False):
    out = dict(parts)
    doc = out['word/document.xml'].decode('utf-8')
    seen = set()

    def reference(match):
        id_ = re.search(r'<w:commentReference\b[^>]*\bw:id="([^"]+)"', match[0])[1]
        duplicate = id_ in seen
        seen.add(id_)
        return '' if none or (unique and duplicate) else match[0]

    # A reference occupies its own run in this saved fixture. Assert that shape
    # so a new writer cannot silently broaden what the diagnostic removes.
    runs = list(re.finditer(r'<w:r\b[^>]*>(?:(?!</w:r>).)*<w:commentReference\b[^>]*/>(?:(?!</w:r>).)*</w:r>', doc, re.S))
    assert len(runs) == doc.count('<w:commentReference ')
    doc = re.sub(r'<w:r\b[^>]*>(?:(?!</w:r>).)*<w:commentReference\b[^>]*/>(?:(?!</w:r>).)*</w:r>', reference, doc, flags=re.S)
    if none:
        doc = re.sub(r'<w:commentRange(?:Start|End)\b[^>]*/>', '', doc)
    out['word/document.xml'] = doc.encode()
    removed = {'word/commentsExtended.xml'} if legacy or none else set()
    if none:
        removed.add('word/comments.xml')
    if removed:
        for part in removed:
            out.pop(part, None)
        rel = 'word/_rels/document.xml.rels'
        types = ('comments', 'commentsExtended') if none else ('commentsExtended',)
        out[rel] = re.sub(rb'<Relationship\b[^>]*/>', lambda m: b'' if any(('/' + t + '"').encode() in m[0] for t in types) else m[0], out[rel])
        out['[Content_Types].xml'] = re.sub(rb'<Override\b[^>]*/>', lambda m: b'' if any(('PartName="/' + p + '"').encode() in m[0] for p in removed) else m[0], out['[Content_Types].xml'])
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('batch', type=Path)
    ap.add_argument('previous', type=Path)
    ap.add_argument('output', type=Path)
    args = ap.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    if any(args.output.iterdir()):
        ap.error('Output directory must be empty; keep earlier Office evidence.')
    file = '12-point-comments.docx'
    saved = args.batch / file
    original = Path(next(x['source'] for x in json.loads((args.batch / 'manifest.json').read_text()) if x['file'] == file))
    with zipfile.ZipFile(saved) as z:
        parts = {n: z.read(n) for n in z.namelist()}
        infos = {n: z.getinfo(n) for n in z.namelist()}
    rows = []

    def emit(name, feature, values=None, control=None, result='Pending'):
        dest = args.output / name
        if control:
            shutil.copyfile(control, dest)
        else:
            with zipfile.ZipFile(dest, 'w') as z:
                for n, data in values.items():
                    z.writestr(infos[n], data)
        with zipfile.ZipFile(dest) as z:
            assert z.testzip() is None
            current = {n: z.read(n) for n in z.namelist()}
            for n, data in current.items():
                if n.endswith(('.xml', '.rels')):
                    ET.fromstring(data)
        changed = [n for n in sorted(parts.keys() | current.keys()) if parts.get(n) != current.get(n)]
        rows.append(dict(file=name, feature=feature, result=result, sha256=hashlib.sha256(dest.read_bytes()).hexdigest(),
                         changedParts=changed, control=str(control) if control else None))

    emit('01-failing-control.docx', 'Exact batch-2 file 12; already reported to offer repair.', control=saved, result='Repair prompt (reported)')
    emit('02-single-reference.docx', 'Remove only the second reference to comment 0. Keep reply threads.', edit(parts, unique=True))
    emit('03-legacy-comments.docx', 'Remove only reply-thread metadata. Keep both references to comment 0.', edit(parts, legacy=True))
    emit('04-single-reference-legacy.docx', 'Combine changes from 02 and 03.', edit(parts, unique=True, legacy=True))
    emit('05-no-comments-control.docx', 'Remove all comment markers and comment parts; retain all document text and formatting.', edit(parts, none=True))
    emit('06-original-source.docx', 'Unchanged public test source; establishes whether its malformed comment IDs already need repair.', control=original)
    emit('07-previous-writer.docx', 'Same source saved by writer 4c47a1d before the point-comment change.', control=args.previous / file)
    (args.output / 'manifest.json').write_text(json.dumps(rows, indent=2) + '\n')
    table = ['# Comment repair isolation', '', 'Open **02–07** in Word. File 01 is the already-failing control.',
             'Report the numbers that offer repair, and whether repaired files then open. Comment removal in the diagnostic copies is intentional.', '',
             '| File | Change / control | Result |', '|---|---|---|']
    table += [f"| [{r['file']}]({r['file']}) | {r['feature']} | {r['result']} |" for r in rows]
    (args.output / 'CHECKLIST.md').write_text('\n'.join(table) + '\n')
    print(json.dumps({'files': len(rows), 'output': str(args.output), 'changedParts': {r['file']: r['changedParts'] for r in rows}}, indent=2))


if __name__ == '__main__':
    main()
