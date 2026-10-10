"""SmartArt made by PowerPoint in a layout Lectern has (public/lectern/js/smartart-io.js, OFFICE): for every corpus
deck that has one, open it in Lectern, save it unedited (every diagram part must stay byte-identical), edit one
such diagram through its items (rename the first, add one) and save (tools/ooxml/package.py must pass), and
reopen that save (the diagram must be editable). Needs Chromium with CDP, as tools/shot.mjs.

    python3 tools/lectern/test/smartart-office.py OUT [name filter]

Writes OUT/<deck>-edited.pptx and OUT/results.json; removes the Recent Files entries it made."""
import base64, collections, glob, hashlib, json, os, re, subprocess, sys, zipfile
OUT = sys.argv[1] if len(sys.argv) > 1 else 'tmp/smartart-office'
os.makedirs(OUT, exist_ok=True)
M = {'default','orgChart1','hierarchy1','hierarchy2','vList2','hList1','process1','chevron1','process2','hProcess11','cycle2','cycle3','radial1','venn1','venn3','target1','funnel1','matrix1','pyramid1','pyramid2','pyramid3'}
def page(js):
    steps = [{'js': js + '\nwindow.__b64 = window.__out; return window.__b64.length;', 'wait': 300}] + [{'js': f'return window.__b64.slice({i*40000},{(i+1)*40000})'} for i in range(200)]
    open(OUT + '/steps.json', 'w').write(json.dumps(steps))
    r = subprocess.run(['node', 'tools/shot.mjs', 'lectern/', OUT + '/steps.json', '--out', OUT, '--timeout', '200000'], capture_output=True, text=True)
    vals = [json.loads(l[6:]) for l in r.stdout.splitlines() if l.startswith('js -> ')]
    errs = [l for l in r.stdout.splitlines() if 'error' in l.lower() and 'no errors' not in l]
    if not vals or not isinstance(vals[0], int): return None, (errs or r.stdout[-400:])
    n = vals[0]; s = ''.join(v for v in vals[1:] if isinstance(v, str))[:n]
    return json.loads(s), errs
OPEN = """const b = Uint8Array.from(atob('%s'), (c) => c.charCodeAt(0));
await L.app.openFile(new File([b], '%s')); await new Promise((r) => setTimeout(r, 1200));
const b64 = async () => { const bb = new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()); let s = ''; for (let i = 0; i < bb.length; i += 0x8000) s += String.fromCharCode.apply(null, bb.subarray(i, i + 0x8000)); return btoa(s); };
const found = []; L.pres.slides.forEach((s, i) => L.model.walk(s.shapes, (x) => { if (x.sa) found.push({ i, id: x.id, layout: x.sa.layout, office: !!x.sa.office, n: L.smartart.outline(x.sa.items).length }); return true; }));
"""
def zipb(b): return zipfile.ZipFile(__import__('io').BytesIO(b))
def dparts(z): return {n: hashlib.md5(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('ppt/diagrams/')}
def issues(path):
    out = subprocess.run(['python3', 'tools/ooxml/package.py', 'check', path], capture_output=True, text=True).stdout.splitlines()
    return collections.Counter((i['code'], str(i.get('detail', ''))[:60]) for i in json.loads(out[-1])['issues']) if out else collections.Counter({('no result', ''): 1})
def new_issues(orig, saved):
    """package.py findings in the save that the original does not have (some corpus decks fail as they came)"""
    return sum((issues(saved) - issues(orig)).values())
rows = []
flt = sys.argv[2] if len(sys.argv) > 2 else ''
corpus = os.path.join(os.environ.get('VO_CORPORA', os.path.expanduser('~/corpora')), 'powerpoint')
for f in sorted(glob.glob(os.path.join(corpus, '*.pptx'))):
    name = os.path.basename(f)
    if flt not in name: continue
    try: z = zipfile.ZipFile(f)
    except Exception: continue
    ids = [m.group(1) for n in z.namelist() if re.match(r'ppt/diagrams/data\d*\.xml$', n) for m in [re.search(r'loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/([^"#]+)', z.read(n).decode('utf8', 'replace'))] if m]
    if not any(i in M for i in ids): continue
    raw = open(f, 'rb').read(); b = base64.b64encode(raw).decode()
    js = OPEN % (b, 'deck.pptx') + """
const plain = await b64();
const d = found.find((x) => x.office); let edited = null, info = '';
if (d) {
  L.ed.idx = d.i; L.ed.render();
  const g = L.model.shapeById(L.pres.slides[d.i], d.id); L.ed.select([g.id]); await new Promise((r) => setTimeout(r, 150));
  const first = L.smartart.outline(g.sa.items)[0].item;
  L.diagram.setText(g, first.id, 'Edited in Lectern'); L.diagram.op('after', g, first.id); L.diagram.setText(g, L.diagram.current().sa.items.length ? L.smartart.outline(g.sa.items)[1].item.id : first.id, 'Added here');
  edited = await b64(); info = g.sa.layout + ' items=' + L.smartart.outline(g.sa.items).length;
}
window.__out = JSON.stringify({ found, plain, edited, info, losses: (L.pres.losses || []).map((l) => l.what).slice(0, 5) });"""
    res, errs = page(js)
    if not res: rows.append((name, 'PAGE FAIL', errs)); print(name, 'PAGE FAIL', errs); continue
    o = zipfile.ZipFile(f); p = zipb(base64.b64decode(res['plain']))
    same = dparts(o) == dparts(p)
    row = {'deck': name, 'mapped': sum(1 for x in res['found'] if x['office']), 'diagrams': len(ids), 'unedited_identical': same, 'edit': res['info']}
    if res['edited']:
        ep = os.path.join(OUT, name.replace('.pptx', '-edited.pptx')); open(ep, 'wb').write(base64.b64decode(res['edited']))
        row['new_issues'] = new_issues(f, ep)
        rb = base64.b64encode(open(ep, 'rb').read()).decode()
        res2, e2 = page(OPEN % (rb, 'edited.pptx') + "window.__out = JSON.stringify(found);")
        row['reopened'] = [(x['layout'], x['n'], x['office']) for x in (res2 or [])]
    rows.append(row); print(json.dumps(row))
json.dump(rows, open(OUT + '/results.json', 'w'), indent=1)
clean = "const e = (await VO.recent.list()).filter((x) => x.app === 'lectern' && ['deck.pptx', 'edited.pptx'].includes(x.name)); for (const x of e) { if (x.draft) await VO.recent.discardDraft(x.key); await VO.recent.remove(x.key); } window.__out = JSON.stringify(e.length);"
page(clean)
ok = [r for r in rows if isinstance(r, dict) and r['unedited_identical'] and not r.get('new_issues') and (not r['edit'] or any(not o for _, _, o in r.get('reopened', [])))]
print(f'{len(ok)}/{len(rows)} decks pass')
