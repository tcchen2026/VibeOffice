"""What does an open-and-save in Quire lose? Compare each original .docx with Quire's saved copy and count, per
feature, the documents where the original has it and the saved copy has less (or none). Also compares the words
of the whole document (body, headers, footers, notes, comments, text boxes) as an independent text check.
Usage: python3 -I loss-audit.py corpusDir savedDir out.json"""
import collections, json, os, re, sys, zipfile

corpus, saved, outp = sys.argv[1], sys.argv[2], sys.argv[3]

def load(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    files = {}
    for n in names:
        if n.endswith('.xml') or n.endswith('.rels'):
            try:
                x = z.read(n).decode('utf8', 'replace')
                # the Fallback of mc:AlternateContent repeats the Choice (a VML copy of a text box or shape)
                files[n] = re.sub(r'<mc:Fallback\b.*?</mc:Fallback>', '', x, flags=re.S) if 'mc:Fallback' in x else x
            except Exception: pass
    return names, files

def part(f, pat):
    return '\n'.join(x for n, x in f.items() if re.match(pat, n))

BODY = r'word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$'
def cnt(pat, where=BODY, flags=0):
    return lambda names, f: len(re.findall(pat, part(f, where), flags))

FEATURES = {
    'comments': cnt(r'<w:comment\b', r'word/comments\.xml$'),
    'comment replies / resolved state (commentsExtended)': cnt(r'<w15:commentEx\b', r'word/commentsExtended\.xml$'),
    'footnotes': cnt(r'<w:footnote\b(?![^>]*w:type=)', r'word/footnotes\.xml$'),
    'endnotes': cnt(r'<w:endnote\b(?![^>]*w:type=)', r'word/endnotes\.xml$'),
    'headers / footers': lambda names, f: sum(1 for n in names if re.match(r'word/(header|footer)\d*\.xml$', n)),
    'tracked insertions / deletions': cnt(r'<w:(ins|del)\b'),
    'tracked formatting changes': cnt(r'<w:(rPrChange|pPrChange|sectPrChange|tblPrChange)\b'),
    'fields (TOC, PAGE, REF, …)': cnt(r'<w:fldSimple\b|<w:fldChar\b[^>]*w:fldCharType="begin"'),
    'content controls': cnt(r'<w:sdt>|<w:sdt\b'),
    'legacy form fields': cnt(r'<w:ffData\b'),
    'bookmarks': cnt(r'<w:bookmarkStart\b(?![^>]*w:name="_GoBack")'),
    'hyperlinks': cnt(r'<w:hyperlink\b'),
    'pictures': cnt(r'<pic:pic\b|<v:imagedata\b'),
    'text boxes': cnt(r'<w:txbxContent\b'),
    'drawing shapes (wps)': cnt(r'<wps:wsp\b'),
    'VML shapes': cnt(r'<v:(shape|rect|oval|roundrect|line|polyline)\b'),
    'equations': cnt(r'<m:oMath\b'),
    'tables': cnt(r'<w:tbl>'),
    'sections': cnt(r'<w:sectPr\b', r'word/document\.xml$'),
    'multi-column sections': cnt(r'<w:cols\b[^>]*w:num="[2-9]'),
    'page borders': cnt(r'<w:pgBorders\b', r'word/document\.xml$'),
    'line numbering': cnt(r'<w:lnNumType\b', r'word/document\.xml$'),
    'document background colour': cnt(r'<w:background\b', r'word/document\.xml$'),
    'drop caps / frames': cnt(r'<w:framePr\b'),
    'embedded OLE objects': cnt(r'<o:OLEObject\b|<w:object\b'),
    'charts': lambda names, f: sum(1 for n in names if re.match(r'word/charts/chart\d+\.xml$', n)),
    'SmartArt': lambda names, f: sum(1 for n in names if re.match(r'word/diagrams/data\d+\.xml$', n)),
    'Word 2010+ text effects (glow, outline, ligatures)': cnt(r'<w14:(glow|textOutline|textFill|ligatures|shadow|reflection)\b'),
    'smart tags / custom XML markup': cnt(r'<w:(smartTag|customXml)\b'),
    'permission ranges': cnt(r'<w:permStart\b'),
    'ruby (phonetic guide)': cnt(r'<w:ruby>'),
    'embedded HTML/RTF chunks (altChunk)': cnt(r'<w:altChunk\b', r'word/document\.xml$'),
    'glossary (building blocks, AutoText)': lambda names, f: int(any(n.startswith('word/glossary/') for n in names)),
    'custom XML data parts': lambda names, f: sum(1 for n in names if re.match(r'customXml/item\d+\.xml$', n)),
    'custom document properties': lambda names, f: int('docProps/custom.xml' in names),
    'VBA macros': lambda names, f: sum(1 for n in names if n.lower().endswith('vbaproject.bin')),
    'embedded fonts': lambda names, f: sum(1 for n in names if re.match(r'word/fonts/', n)),
    'document protection': cnt(r'<w:documentProtection\b', r'word/settings\.xml$'),
    'attached template': cnt(r'<w:attachedTemplate\b', r'word/settings\.xml$'),
    'paragraph styles defined': cnt(r'<w:style\b[^>]*w:type="paragraph"', r'word/styles\.xml$'),
    'character styles defined': cnt(r'<w:style\b[^>]*w:type="character"', r'word/styles\.xml$'),
    'table styles defined': cnt(r'<w:style\b[^>]*w:type="table"', r'word/styles\.xml$'),
    'list definitions': cnt(r'<w:abstractNum\b', r'word/numbering\.xml$'),
    'theme': lambda names, f: int(any(re.match(r'word/theme/theme\d*\.xml$', n) for n in names)),
}

WORDS_PARTS = r'word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$'
def words(f):
    """visible words: w:t of the main story and the other stories, ignoring deleted text and field codes"""
    x = part(f, WORDS_PARTS)
    x = re.sub(r'<w:del\b[^>]*/>', '', x)
    x = re.sub(r'<w:del\b[^>]*>.*?</w:del>', ' ', x, flags=re.S)
    # words are taken per paragraph, so text split over several runs counts the same as one run
    toks = []
    for para in re.split(r'</w:p>', x):
        toks.append(''.join(m.group(1) for m in re.finditer(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', para)))
    txt = ' '.join(toks)
    txt = re.sub(r'&lt;', '<', re.sub(r'&gt;', '>', re.sub(r'&amp;', '&', txt)))
    return collections.Counter(w for w in re.findall(r'\w+', txt.lower()))

stats = {k: {'docs_with': 0, 'lost_all': 0, 'lost_some': 0, 'items_orig': 0, 'items_saved': 0, 'examples': []} for k in FEATURES}
text = {'docs': 0, 'docs_missing_words': 0, 'words_orig': 0, 'words_missing': 0, 'worst': []}
pairs = 0
for s in sorted(os.listdir(saved)):
    sp = os.path.join(saved, s)
    op = os.path.join(corpus, s)
    if not os.path.exists(op):
        cand = [c for c in os.listdir(corpus) if c.rsplit('.', 1)[0] == s.rsplit('.', 1)[0]] if False else []
        continue
    try:
        with open(sp, 'rb') as fh:
            if fh.read(4) == bytes.fromhex('d0cf11e0'): continue
        on, of = load(op); sn, sf = load(sp)
    except Exception:
        continue
    pairs += 1
    for k, fn in FEATURES.items():
        try: a, b = int(fn(on, of) or 0), int(fn(sn, sf) or 0)
        except Exception: continue
        if a <= 0: continue
        st = stats[k]
        st['docs_with'] += 1; st['items_orig'] += a; st['items_saved'] += min(a, b)
        if b == 0: st['lost_all'] += 1
        elif b < a: st['lost_some'] += 1
        if b < a and len(st['examples']) < 4: st['examples'].append(s)
    wa, wb = words(of), words(sf)
    miss = sum((wa - wb).values())
    text['docs'] += 1; text['words_orig'] += sum(wa.values()); text['words_missing'] += miss
    if miss:
        text['docs_missing_words'] += 1
        text['worst'].append([s, miss, sum(wa.values())])
text['worst'] = sorted(text['worst'], key=lambda r: -r[1])[:25]
json.dump({'pairs': pairs, 'stats': stats, 'text': text}, open(outp, 'w'), indent=1)
print('document pairs compared', pairs)
print('words: %d of %d documents keep every word; %d of %d words missing in all' % (text['docs'] - text['docs_missing_words'], text['docs'], text['words_missing'], text['words_orig']))
print('%-52s %6s %6s %6s %12s' % ('feature', 'docs', 'lost', 'part', 'kept items'))
for k, st in sorted(stats.items(), key=lambda kv: -kv[1]['docs_with']):
    if st['docs_with']:
        print('%-52s %6d %6d %6d %6d/%-6d' % (k, st['docs_with'], st['lost_all'], st['lost_some'], st['items_saved'], st['items_orig']))
