"""What does an open-and-save in Lectern lose? Compare each original deck with Lectern's saved copy and count, per
feature, the decks where the original has it and the saved copy no longer does.
Usage: python3 -I loss-audit.py corpusDir savedDir out.json"""
import json, os, re, sys, zipfile, collections

corpus, saved, outp = sys.argv[1], sys.argv[2], sys.argv[3]

def load(path):
    z = zipfile.ZipFile(path)
    files = {}
    for n in z.namelist():
        if n.endswith('.xml') or n.endswith('.rels'):
            try: files[n] = z.read(n).decode('utf8', 'replace')
            except Exception: pass
    return z.namelist(), files

def rel_targets(files, prefix):
    """relationship type tails from .rels files whose source part matches prefix"""
    c = collections.Counter()
    for n, x in files.items():
        if n.endswith('.rels') and re.match(prefix, n):
            for t in re.findall(r'Type="[^"]*/([^"/]+)"', x): c[t] += 1
    return c

def joined(files, pat):
    return '\n'.join(x for n, x in files.items() if re.match(pat, n))

FEATURES = {
    # package parts
    'embedded OLE object (slide)': lambda names, f: rel_targets(f, r'ppt/slides/_rels/')['oleObject'] + rel_targets(f, r'ppt/slides/_rels/')['package'],
    'video / audio on a slide': lambda names, f: sum(rel_targets(f, r'ppt/slides/_rels/')[k] for k in ('video', 'audio', 'media')),
    'comments': lambda names, f: rel_targets(f, r'ppt/slides/_rels/')['comments'] + len(re.findall(r'<p188:cm\b|<p:cm\b', joined(f, r'ppt/comments/'))),
    'VBA macros': lambda names, f: sum(1 for n in names if n.lower().endswith('vbaproject.bin')),
    'custom XML parts / tags': lambda names, f: sum(1 for n in names if n.startswith('customXml/') or n.startswith('ppt/tags/')),
    'custom document properties': lambda names, f: int('docProps/custom.xml' in names),
    'embedded fonts': lambda names, f: len(re.findall(r'<p:embeddedFont>', f.get('ppt/presentation.xml', ''))),
    'sections': lambda names, f: len(re.findall(r'<p14:section ', f.get('ppt/presentation.xml', ''))),
    'custom shows': lambda names, f: len(re.findall(r'<p:custShow ', f.get('ppt/presentation.xml', ''))),
    'handout master': lambda names, f: sum(1 for n in names if n.startswith('ppt/handoutMasters/') and n.endswith('.xml')),
    'SmartArt (editable diagram data)': lambda names, f: rel_targets(f, r'ppt/slides/_rels/')['diagramData'],
    'ink / content parts': lambda names, f: len(re.findall(r'<p:contentPart\b|<p14:contentPart\b', joined(f, r'ppt/slides/slide\d+\.xml'))),
    'Office 2016+ charts (chartEx)': lambda names, f: sum(1 for n in names if re.match(r'ppt/charts/chartEx\d*\.xml', n)),
    '3-D models': lambda names, f: len(re.findall(r'<am3d:model3d\b', joined(f, r'ppt/slides/slide\d+\.xml'))),
    # slide content
    'animations (effects)': lambda names, f: len(re.findall(r'presetClass="(?!mediacall)', joined(f, r'ppt/slides/slide\d+\.xml'))),
    'media play/pause animations': lambda names, f: len(re.findall(r'presetClass="mediacall"', joined(f, r'ppt/slides/slide\d+\.xml'))),
    'slide transitions': lambda names, f: sum(1 for n, x in f.items() if re.match(r'ppt/slides/slide\d+\.xml$', n) and re.search(r'<p:transition\b[^>]*[^/]>\s*<(?!/)', x)),
    'auto-advance timings': lambda names, f: sum(1 for n, x in f.items() if re.match(r'ppt/slides/slide\d+\.xml$', n) and re.search(r'<p:transition\b[^>]*advTm="', x)),
    'hyperlinks (web, file, slide)': lambda names, f: len(re.findall(r'<a:hlinkClick\b(?![^>]*action="ppaction://(?:ole|media|macro|program))[^>]*r:id="[^"]+"', joined(f, r'ppt/slides/slide\d+\.xml'))) + len(re.findall(r'<a:hlinkClick\b[^>]*action="ppaction://(?:hlinksldjump|hlinkshowjump)', joined(f, r'ppt/slides/slide\d+\.xml'))),
    'action settings (run program, macro, OLE verb)': lambda names, f: len(re.findall(r'action="ppaction://(?:ole|macro|program|media)', joined(f, r'ppt/slides/slide\d+\.xml'))),
    'glow / reflection / soft edge': lambda names, f: len(re.findall(r'<a:(glow|reflection|softEdge)\b', joined(f, r'ppt/slides/slide\d+\.xml'))),
    '3-D shape effects (bevel, extrusion)': lambda names, f: len(re.findall(r'<a:sp3d\b[^/>]*(?:/>|>)(?!</a:sp3d>)', joined(f, r'ppt/slides/slide\d+\.xml'))) and len(re.findall(r'<a:bevelT\b|extrusionH="', joined(f, r'ppt/slides/slide\d+\.xml'))),
    'inner shadow': lambda names, f: len(re.findall(r'<a:innerShdw\b', joined(f, r'ppt/slides/slide\d+\.xml'))),
    'picture artistic effects / recolour': lambda names, f: len(re.findall(r'<a:duotone\b|<a14:imgEffect\b|<a:clrRepl\b', joined(f, r'ppt/slides/slide\d+\.xml'))),
    'text effects on runs (glow, outline gradient)': lambda names, f: len(re.findall(r'<a:rPr\b[^>]*>(?:(?!</a:rPr>).)*<a:(glow|reflection)\b', joined(f, r'ppt/slides/slide\d+\.xml'), re.S)),
    'speaker notes with formatting': lambda names, f: len(re.findall(r'<a:rPr\b[^>]*\b(b="1"|i="1"|u="sng")', joined(f, r'ppt/notesSlides/notesSlide\d+\.xml'))),
    'slide layouts (named)': lambda names, f: len(set(re.findall(r'<p:cSld name="([^"]+)"', joined(f, r'ppt/slideLayouts/slideLayout\d+\.xml')))),
    'slide masters': lambda names, f: sum(1 for n in names if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', n)),
}
# features where the saved copy may legitimately have fewer (a count drop, not just presence)
COUNTED = {'animations (effects)', 'slide transitions', 'auto-advance timings', 'hyperlinks (web, file, slide)', 'slide layouts (named)', 'slide masters', 'embedded OLE object (slide)', 'video / audio on a slide', 'comments'}

def words(names, f, original):
    """words of the slides, their SmartArt and the speaker notes, taken per paragraph (a:p)"""
    pats = [r'ppt/slides/slide\d+\.xml$', r'ppt/notesSlides/notesSlide\d+\.xml$']
    if original:
        # SmartArt text lives in the diagram drawing (or, without one, the data model); Lectern saves it as shapes
        pats.append(r'ppt/diagrams/drawing\d+\.xml$')
        drawn = {re.sub(r'drawing', 'data', n) for n in names if re.match(r'ppt/diagrams/drawing\d+\.xml$', n) and len(f.get(n, '')) > 600}
        pats.append('|'.join(re.escape(n) + '$' for n in names if re.match(r'ppt/diagrams/data\d+\.xml$', n) and n not in drawn) or r'^$')
    c = collections.Counter()
    for n, x in f.items():
        if not any(re.match(p, n) for p in pats): continue
        x = re.sub(r'<mc:Fallback\b.*?</mc:Fallback>', '', x, flags=re.S)
        for para in re.split(r'</a:p>', x):
            t = ''.join(m.group(1) for m in re.finditer(r'<a:t>([^<]*)</a:t>', para))
            c.update(re.findall(r'\w+', t.lower()))
    return c

stats = {k: {'decks_with': 0, 'decks_lost_all': 0, 'decks_lost_some': 0, 'items_orig': 0, 'items_saved': 0, 'examples': []} for k in FEATURES}
pairs = 0
text = {'decks': 0, 'decks_missing_words': 0, 'words_orig': 0, 'words_missing': 0, 'worst': []}
for s in sorted(os.listdir(saved)):
    sp = os.path.join(saved, s)
    with open(sp, 'rb') as fh:
        if fh.read(4) == bytes.fromhex('d0cf11e0'): continue
    base = s.rsplit('.', 1)[0]
    orig = next((os.path.join(corpus, base + e) for e in ('.pptx', '.pptm', '.potx', '.ppsx', '.potm', '.ppsm') if os.path.exists(os.path.join(corpus, base + e))), None)
    if not orig: continue
    try:
        on, of = load(orig); sn, sf = load(sp)
    except Exception:
        continue
    pairs += 1
    for k, fn in FEATURES.items():
        try: a, b = int(fn(on, of) or 0), int(fn(sn, sf) or 0)
        except Exception: continue
        if a <= 0: continue
        st = stats[k]
        st['decks_with'] += 1; st['items_orig'] += a; st['items_saved'] += min(a, b)
        if b == 0: st['decks_lost_all'] += 1
        elif b < a and k in COUNTED: st['decks_lost_some'] += 1
        if b < a and len(st['examples']) < 3: st['examples'].append(base)
    wa, wb = words(on, of, True), words(sn, sf, False)
    miss = sum((wa - wb).values())
    text['decks'] += 1; text['words_orig'] += sum(wa.values()); text['words_missing'] += miss
    if miss: text['decks_missing_words'] += 1; text['worst'].append([base, miss, sum(wa.values())])
text['worst'] = sorted(text['worst'], key=lambda r: -r[1])[:30]
json.dump({'pairs': pairs, 'stats': stats, 'text': text}, open(outp, 'w'), indent=1)
print('deck pairs compared', pairs)
print('words: %d of %d decks keep every word; %d of %d words missing in all' % (text['decks'] - text['decks_missing_words'], text['decks'], text['words_missing'], text['words_orig']))
print('%-44s %6s %6s %6s %10s' % ('feature', 'decks', 'lost', 'part', 'kept items'))
for k, st in sorted(stats.items(), key=lambda kv: -kv[1]['decks_with']):
    if st['decks_with']:
        print('%-44s %6d %6d %6d %5d/%-5d' % (k, st['decks_with'], st['decks_lost_all'], st['decks_lost_some'], st['items_saved'], st['items_orig']))
