"""Body lines of a PDF, page by page, to compare with the lines Quire draws (tests/quire/lines.test.mjs).

  python3 tools/wordlines.py PDF DOCX OUT.json [--text]

Each page's body (between the .docx's top and bottom margins) as rows: the words at one height (pdftotext -bbox-layout,
tops within 1pt), left to right, so a table's cells on one line make one row, as they do in Quire. A row is
[hash, chars, top, bottom] in points from the top of the page: hash is FNV-1a (32 bit) of the row's last 16 non-blank
characters after NFKC normalisation, chars how many non-blank characters it has. The reference for a document whose text
can't be published (Word's PDF of the walkthrough in /work/corpora/word-ref) is compared without storing its text; --text
also keeps each row's text, for documents of our own (tests/quire/lines/linecases.docx, made by tools/linecases.py).
"""
import html, json, re, subprocess, sys, unicodedata, zipfile

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def key(text):
    """the non-blank characters a row is matched by (the same in tests/lib/lines.mjs)"""
    return ''.join(c for c in unicodedata.normalize('NFKC', text) if not c.isspace() and c not in '\u200b\ufeff\u00ad')


def fnv(text):
    h = 0x811c9dc5
    for c in text:
        h = ((h ^ ord(c)) * 0x01000193) & 0xffffffff
    return h


def body(docx):
    """the page height and the body's top and bottom (points), from the document's last section"""
    import xml.etree.ElementTree as ET
    root = ET.fromstring(zipfile.ZipFile(docx).read('word/document.xml'))
    sect = root.find(f'{W}body/{W}sectPr')
    size, mar = sect.find(f'{W}pgSz'), sect.find(f'{W}pgMar')
    pt = lambda el, a, d: int(el.get(W + a, d)) / 20 if el is not None else d / 20
    h = pt(size, 'h', 15840)
    return h, pt(mar, 'top', 1440), h - pt(mar, 'bottom', 1440)


def read(pdf, docx, out, text=False):
    page_h, top_y, bottom_y = body(docx)
    xml = subprocess.run(['pdftotext', '-bbox-layout', pdf, '-'], capture_output=True, text=True, check=True).stdout
    pages = []
    for page in re.findall(r'<page\b.*?</page>', xml, re.S):
        words = [(float(y0), float(y1), float(x0), html.unescape(w)) for x0, y0, y1, w in
                 re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="([\d.]+)">([^<]*)</word>', page)]
        rows = []
        for w in sorted(words):
            if w[0] < top_y - 1 or w[1] > bottom_y + 1:
                continue
            if rows and abs(rows[-1][0][0] - w[0]) <= 1:
                rows[-1].append(w)
            else:
                rows.append([w])
        out_rows = []
        for r in rows:
            k = key(''.join(w[3] for w in sorted(r, key=lambda w: w[2])))
            if not k:
                continue
            row = [fnv(k[-16:]), len(k), round(min(w[0] for w in r), 2), round(max(w[1] for w in r), 2)]
            if text:
                row.append(' '.join(w[3] for w in sorted(r, key=lambda w: w[2])))
            out_rows.append(row)
        pages.append(out_rows)
    producer = re.search(r'<meta name="Producer" content="([^"]*)"', xml)
    with open(out, 'w') as f:
        f.write('{"pdf": %s, "pageHeight": %s, "body": %s, "pages": [\n' % (json.dumps(producer.group(1) if producer else None, ensure_ascii=False), page_h, json.dumps([round(top_y, 2), round(bottom_y, 2)])))
        f.write(',\n'.join(json.dumps(p, ensure_ascii=False, separators=(',', ':')) for p in pages))
        f.write('\n]}\n')
    print(f'saved {out}: {len(pages)} pages, {sum(len(p) for p in pages)} rows')


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if a != '--text']
    if len(args) != 3:
        sys.exit(__doc__)
    read(*args, text='--text' in sys.argv)
