"""Reference check for Ledger's CSV parser: parse each file with Python's csv module using the delimiter and
encoding Ledger chose, and compare the md5 of the rows with Ledger's digest.
Usage: python3 -I csvref.py results.jsonl corpus_dir"""
import csv, hashlib, io, json, sys, os

csv.field_size_limit(1 << 30)
# WHATWG windows-1252: the five undefined bytes decode to the C1 code points
import codecs
codecs.register_error('c1', lambda e: (''.join(chr(b) for b in e.object[e.start:e.end]), e.end))
res = [json.loads(l) for l in open(sys.argv[1], encoding='utf-8')]
d = sys.argv[2]
ok = bad = 0
bads = []
for r in res:
    if 'digest' not in r:
        continue
    enc = {'windows-1252': 'cp1252', 'utf-16le': 'utf-16-le', 'utf-16be': 'utf-16-be'}.get(r['encoding'], r['encoding'])
    raw = open(os.path.join(d, r['file']), 'rb').read()
    try:
        text = raw.decode(enc, errors='c1' if enc == 'cp1252' else 'replace')
    except LookupError:
        text = raw.decode('latin-1')
    if text.startswith('﻿'):
        text = text[1:]
    rows = list(csv.reader(io.StringIO(text, newline=''), delimiter=r['delim']))
    rows = [[] if (len(x) == 1 and x[0] == '') else x for x in rows]
    dig = hashlib.md5(json.dumps(rows, ensure_ascii=False, separators=(',', ':')).encode('utf-8', 'surrogatepass')).hexdigest()
    if dig == r['digest']:
        ok += 1
    else:
        bad += 1
        bads.append((r['file'], len(rows), r['rows']))
print(json.dumps({'match': ok, 'differ': bad}))
for b in bads[:20]:
    print(b)
