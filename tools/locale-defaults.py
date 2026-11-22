"""New documents' defaults per language, as Office writes them (docs/i18n.md, Defaults; public/common/locales.js).

  python3 tools/locale-defaults.py [/work/corpora]

Scans the Office-saved files of the corpora (tools/corpora.sh) for the ones a Japanese, Korean, Chinese, French or German
Office made, and prints the most common value of each default with its count:

  Word        page size and margins, document grid, docDefaults fonts, size and languages, the Normal style's
              paragraph settings, themeFontLang, characterSpacingControl, the theme's Latin and East Asian fonts;
              the built-in styles that name a font or size in English Word (headings, Title, …: STYLES), each
              style's whole pPr and rPr as one value
  Excel       the default font (name, size, charset), the default row height, the theme's fonts
  PowerPoint  the slide size, the theme's Latin and East Asian fonts, the text language

Which language made a file: Word, the East Asian language Office set for its theme fonts (themeFontLang eastAsia)
for Japanese, Korean and Chinese, the documents' language (docDefaults w:lang) for French and German; Excel, the default
font's charset (128 Japanese, 129 Korean, 134 Simplified Chinese, 136 Traditional Chinese) or the first sheet's name (Feuil1, Tabelle1); PowerPoint, the text
language of its default text style.
"""
import collections, os, re, sys, zipfile

CORPORA = os.path.expanduser(sys.argv[1] if len(sys.argv) > 1 else '/work/corpora')
LANGS = ['ja', 'ko', 'zh-CN', 'zh-Hant', 'fr', 'de']
# the built-in styles whose font or size Quire's English set names (public/quire/js/dmodel.js, D.builtinStyles)
STYLES = ['heading %d' % i for i in range(1, 10)] + ['Title', 'Subtitle', 'Body Text 3', 'Plain Text', 'caption', 'footnote text',
          'endnote text', 'annotation text', 'Balloon Text', 'envelope address', 'envelope return', 'index heading',
          'Document Map', 'macro', 'HTML Preformatted', 'annotation reference', 'header', 'footer', 'toc 1']


def chinese(tag):
    tag = tag.lower()
    return ('zh-Hant' if re.match(r'^zh-(?:tw|hk|mo|hant)(?:-|$)', tag) else 'zh-CN') if tag == 'zh' or tag.startswith('zh-') else None


def read(z, name):
    try:
        return z.read(name).decode('utf8', 'replace')
    except KeyError:
        return ''


def attr(xml, tag, name):
    m = re.search(r'<' + tag + r'\b[^>]*\b' + name + r'="([^"]*)"', xml)
    return m.group(1) if m else None


def theme_fonts(xml):
    out = {}
    for part in ('majorFont', 'minorFont'):
        m = re.search(r'<a:' + part + r'>(.*?)</a:' + part + '>', xml, re.S)
        if not m:
            continue
        body = m.group(1)
        out[part + ' latin'] = attr(body, 'a:latin', 'typeface')
        out[part + ' ea'] = attr(body, 'a:ea', 'typeface')
        for script in ('Jpan', 'Hang', 'Hans', 'Hant'):
            f = re.search(r'<a:font script="' + script + r'" typeface="([^"]*)"', body)
            if f:
                out[part + ' ' + script] = f.group(1)
    return out


def app_name(z):
    return attr(read(z, 'docProps/app.xml').replace('<Application>', '<Application v="').replace('</Application>', '"/>'), 'Application', 'v') or ''


def word(path):
    z = zipfile.ZipFile(path)
    if 'Word' not in app_name(z):
        return None
    doc, sty, sett = read(z, 'word/document.xml'), read(z, 'word/styles.xml'), read(z, 'word/settings.xml')
    tfl = attr(sett, 'w:themeFontLang', 'w:eastAsia') or ''
    dd = (re.search(r'<w:docDefaults>.*?</w:docDefaults>', sty, re.S) or [''])[0]
    lang = attr(dd, 'w:lang', 'w:val') or ''
    which = chinese(tfl) or ('ja' if tfl.startswith('ja') else 'ko' if tfl.startswith('ko') else 'fr' if lang.startswith('fr') else 'de' if lang.startswith('de') else None)
    if not which:
        return None
    normal = (re.search(r'<w:style [^>]*w:default="1"[^>]*w:type="paragraph"[^>]*>.*?</w:style>|<w:style [^>]*w:type="paragraph"[^>]*w:default="1"[^>]*>.*?</w:style>', sty, re.S) or [''])[0]
    # Japanese, Korean and Simplified Chinese Normal styles are justified without widow control.
    # Traditional Chinese has no widow control, but need not be justified. This excludes English defaults.
    if which in ('ja', 'ko', 'zh-CN') and not ('w:val="both"' in normal and re.search(r'<w:widowControl w:val="(0|false)"', normal)):
        return None
    if which == 'zh-Hant' and not re.search(r'<w:widowControl w:val="(0|false)"', normal):
        return None
    sect = (re.findall(r'<w:sectPr\b.*?</w:sectPr>', doc, re.S) or [''])[-1]
    v = {
        'page': '%s x %s' % (attr(sect, 'w:pgSz', 'w:w'), attr(sect, 'w:pgSz', 'w:h')),
        'margins t r b l': ' '.join(str(attr(sect, 'w:pgMar', 'w:' + k)) for k in ('top', 'right', 'bottom', 'left')),
        'header footer': '%s %s' % (attr(sect, 'w:pgMar', 'w:header'), attr(sect, 'w:pgMar', 'w:footer')),
        'docGrid': '%s %s %s' % (attr(sect, 'w:docGrid', 'w:type'), attr(sect, 'w:docGrid', 'w:linePitch'), attr(sect, 'w:docGrid', 'w:charSpace')),
        'docDefaults rFonts': re.sub(r'\s+', ' ', (re.search(r'<w:rFonts [^>]*/>', dd) or [''])[0]),
        'docDefaults sz': attr(dd, 'w:sz', 'w:val'),
        'docDefaults lang': re.sub(r'\s+', ' ', (re.search(r'<w:lang [^>]*/>', dd) or [''])[0]),
        'docDefaults spacing': re.sub(r'\s+', ' ', (re.search(r'<w:spacing [^>]*/>', dd) or [''])[0]),
        'Normal pPr': re.sub(r'w:rsid\w*="[^"]*"', '', (re.search(r'<w:pPr>.*?</w:pPr>', normal, re.S) or [''])[0]),
        'Normal rPr': (re.search(r'<w:rPr>.*?</w:rPr>', normal, re.S) or [''])[0],
        'themeFontLang': re.sub(r'\s+', ' ', (re.search(r'<w:themeFontLang [^>]*/>', sett) or [''])[0]),
        'characterSpacingControl': attr(sett, 'w:characterSpacingControl', 'w:val'),
        'defaultTabStop': attr(sett, 'w:defaultTabStop', 'w:val'),
        'compat': ' '.join(sorted(set(re.findall(r'<w:(\w+)/>', (re.search(r'<w:compat>.*?</w:compat>', sett, re.S) or [''])[0])))),
    }
    for m in re.finditer(r'<w:style\b[^>]*>.*?</w:style>', sty, re.S):
        name = attr(m.group(0), 'w:name', 'w:val')
        if name in STYLES:
            body = re.sub(r'\s+w:rsid\w*="[^"]*"|<w:rsid\b[^>]*/>', '', m.group(0))
            v['style ' + name] = ''.join(re.findall(r'<w:pPr>.*?</w:pPr>|<w:rPr>.*?</w:rPr>', body, re.S)) or '(none)'
    v.update(theme_fonts(read(z, 'word/theme/theme1.xml')))
    return which, v


def excel(path):
    z = zipfile.ZipFile(path)
    if 'Excel' not in app_name(z):
        return None
    sty, wb = read(z, 'xl/styles.xml'), read(z, 'xl/workbook.xml')
    font = (re.search(r'<fonts\b[^>]*>\s*<font>(.*?)</font>', sty, re.S) or [None, ''])[1]
    charset = attr(font, 'charset', 'val')
    sheet = attr(wb, 'sheet', 'name') or ''
    which = {'128': 'ja', '129': 'ko', '134': 'zh-CN', '136': 'zh-Hant'}.get(charset) or ('fr' if sheet.startswith('Feuil') else 'de' if sheet.startswith('Tabelle') else None)
    if not which:
        return None
    first = sorted(n for n in z.namelist() if n.startswith('xl/worksheets/sheet'))[:1]
    v = {'font': '%s %s charset=%s family=%s scheme=%s' % (attr(font, 'name', 'val'), attr(font, 'sz', 'val'), charset, attr(font, 'family', 'val'), attr(font, 'scheme', 'val')),
         'defaultRowHeight': attr(read(z, first[0]), 'sheetFormatPr', 'defaultRowHeight') if first else None,
         'first sheet': re.sub(r'\d+$', '', sheet)}
    v.update(theme_fonts(read(z, 'xl/theme/theme1.xml')))
    return which, v


def powerpoint(path):
    z = zipfile.ZipFile(path)
    if 'PowerPoint' not in app_name(z):
        return None
    pres = read(z, 'ppt/presentation.xml')
    lang = attr((re.search(r'<p:defaultTextStyle>.*?</p:defaultTextStyle>', pres, re.S) or [''])[0], 'a:defRPr', 'lang') or ''
    which = chinese(lang) or next((l for l in LANGS if lang.startswith(l)), None)
    if not which:
        return None
    v = {'slide size': '%s x %s %s' % (attr(pres, 'p:sldSz', 'cx'), attr(pres, 'p:sldSz', 'cy'), attr(pres, 'p:sldSz', 'type')), 'text lang': lang}
    v.update(theme_fonts(read(z, 'ppt/theme/theme1.xml')))
    return which, v


def main():
    for app, folder, ext, fn in (('Word', 'word', '.docx', word), ('Excel', 'excel', '.xlsx', excel), ('PowerPoint', 'powerpoint', '.pptx', powerpoint)):
        found = collections.defaultdict(list)
        d = os.path.join(CORPORA, folder)
        for f in sorted(os.listdir(d)) if os.path.isdir(d) else []:
            if not f.endswith(ext):
                continue
            try:
                r = fn(os.path.join(d, f))
            except Exception:   # a damaged file of the corpus
                continue
            if r:
                found[r[0]].append(r[1])
        for lang in LANGS:
            files = found[lang]
            print(f'\n== {app}, {lang}: {len(files)} files')
            for k in sorted({k for v in files for k in v}):
                common = collections.Counter(v.get(k) for v in files if v.get(k)).most_common(2)
                if common:
                    width = 600 if k.startswith('style ') else 110
                    print(f'  {k}: ' + '; '.join(f'{val!s:.{width}} ({n})' for val, n in common))


if __name__ == '__main__':
    main()
