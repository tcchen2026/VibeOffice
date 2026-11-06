#!/usr/bin/env python3
"""Writes public/oss-licenses.html: the licenses of the open source works this repository is made with, as a
browser's credits page (About ▸ Open source licenses in the Start Center and every app).

  python3 tools/oss-licenses.py            write the page
  python3 tools/oss-licenses.py --check    only check: the page lists the crates that ship, none needs copyleft

The site: VibeOffice itself (LICENSE), the fonts (public/fonts/LICENSES.txt), the proofing lists
(public/common/dict/LICENSES.txt), the ECMA-376 data in public/common/opc-order.js, and the code and data derived from
other projects (public/common/licenses/; each file says "Derived from <name> …"); their license files are read by the
page when an entry is opened, or kept in it. The desktop app: every crate compiled into the app and its installer
(the workspace's members and their normal dependencies, all the way down; not what only build scripts or procedural
macros use), on Windows (x86_64-pc-windows-gnu), Linux (x86_64-unknown-linux-gnu) or macOS (aarch64-apple-darwin; the
universal binary's x86_64 half has the same crates), with the license files from the
crate's package (cargo metadata; the sources must be downloaded, as after a build), and the DLLs that ship beside the
exe (desktop/licenses/; a crate whose package has no license file has the one from its repository in
desktop/licenses/crates/). A shipped crate whose licence needs copyleft stops it. Identical texts are kept once. Run it
again after changing a dependency. Standard library only."""
import hashlib, html, json, os, re, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'public', 'oss-licenses.html')
TARGETS = {'windows': 'x86_64-pc-windows-gnu', 'macos': 'aarch64-apple-darwin', 'linux': 'x86_64-unknown-linux-gnu'}
LICENSE_FILE = re.compile(r'^(licen[cs]e|copying|copyright|notice|unlicense)', re.I)

E = html.escape


def shipped_crates(target):
    """{package id: package} of every crate compiled into what ships for a target: the workspace's members (the app
    and its installer) and, along normal dependencies, what they link, with features resolved as the build resolves
    them (cargo tree; cargo metadata's graph also holds what only Tauri's build-time macros use). Our own path crates
    (desktop/patches/) are left out."""
    run = lambda *a: subprocess.run(['cargo', *a, '--locked'], cwd=os.path.join(ROOT, 'desktop'), capture_output=True, text=True, check=True).stdout
    tree = run('tree', '--target', target, '--workspace', '-e', 'normal,no-proc-macro', '--prefix', 'none', '--format', '{p}')
    shipped = {(m[1], m[2]) for m in re.finditer(r'^(\S+) v(\S+)(?! \((?:/|[A-Za-z]:\\))', tree, re.M)}
    meta = json.loads(run('metadata', '--format-version', '1', '--filter-platform', target))
    return {p['id']: p for p in meta['packages'] if p.get('source') and (p['name'], p['version']) in shipped}


COPYLEFT = re.compile(r'(MPL|GPL|EUPL|CDDL|EPL|OSL|CC-BY-SA)', re.I)


def without_copyleft(expr):
    """whether an SPDX licence expression can be met without a copyleft licence: one side of an OR is enough, every
    side of an AND is needed ("MIT OR Apache-2.0 OR LGPL-2.1-or-later" can; "MPL-2.0" can't)"""
    tokens = re.findall(r'\(|\)|[^\s()]+', expr.replace('/', ' OR '))
    pos = 0

    def alternatives():
        nonlocal pos
        ok = all_of()
        while pos < len(tokens) and tokens[pos].upper() == 'OR':
            pos += 1
            ok = all_of() or ok
        return ok

    def all_of():
        nonlocal pos
        ok = one()
        while pos < len(tokens) and tokens[pos].upper() == 'AND':
            pos += 1
            ok = one() and ok
        return ok

    def one():
        nonlocal pos
        tok = tokens[pos]
        pos += 1
        if tok == '(':
            ok = alternatives()
            pos += 1   # ')'
            return ok
        if pos < len(tokens) and tokens[pos].upper() == 'WITH':
            pos += 2   # an exception only widens the licence
        return not COPYLEFT.search(tok)

    return bool(tokens) and alternatives()


def license_texts(package):
    """the texts of the license files at the top of a crate's package (and in a LICENSES folder)"""
    folder = os.path.dirname(package['manifest_path'])
    names = []
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if os.path.isfile(path) and LICENSE_FILE.match(name) and not name.endswith('.rs'):
            names.append(path)
        elif os.path.isdir(path) and name.upper() == 'LICENSES':
            names += [os.path.join(path, n) for n in sorted(os.listdir(path)) if os.path.isfile(os.path.join(path, n))]
    if not names:   # none in the package: the one from its repository, kept in desktop/licenses/crates/
        kept = os.path.join(ROOT, 'desktop', 'licenses', 'crates', package['name'] + '.txt')
        if os.path.isfile(kept):
            names.append(kept)
    out = []
    for path in names:   # named as in the package, or as kept here
        with open(path, encoding='utf-8', errors='replace') as f:
            out.append((os.path.relpath(path, ROOT if path.startswith(ROOT + os.sep) else folder), f.read().strip()))
    return out


def main():
    texts = {}   # hash -> text, each kept once

    def keep(text):
        key = hashlib.sha1(text.encode()).hexdigest()[:12]
        texts[key] = text
        return key

    crates = {}   # (name, version) -> entry
    for platform, target in TARGETS.items():
        for package in shipped_crates(target).values():
            key = (package['name'], package['version'])
            if key not in crates:
                files = license_texts(package)
                crates[key] = {
                    'name': package['name'], 'version': package['version'], 'license': package.get('license') or package.get('license_file') or '',
                    'url': package.get('repository') or package.get('homepage') or f"https://crates.io/crates/{package['name']}",
                    'texts': [keep(f'{name}\n\n{text}') for name, text in files], 'platforms': [],
                }
            crates[key]['platforms'].append(platform)

    with open(os.path.join(ROOT, 'LICENSE'), encoding='utf-8') as f:
        apache = keep(f.read().strip())
    with open(os.path.join(ROOT, 'desktop', 'licenses', 'WebView2Loader.txt'), encoding='utf-8') as f:
        webview2 = keep(f.read().strip())
    with open(os.path.join(ROOT, 'desktop', 'licenses', 'appimage-runtime.txt'), encoding='utf-8') as f:
        appimage = keep(f.read().strip())

    def notice(name):
        with open(os.path.join(ROOT, 'public', 'common', 'licenses', name), encoding='utf-8') as f:
            return keep(f'{name}\n\n{f.read().strip()}')

    def work(name, lic, links, what, files='', texts_=(), src='', derived=False):
        head = f'<b>{E(name)}</b><span class="lic">{E(lic)}</span>' + ''.join(f'<a href="{E(u)}">{E(t)}</a>' for t, u in links)
        body = f'<p class="what">{what}</p>' if what else ''
        body += f'<p class="files">{E(files)}</p>' if files else ''
        if texts_ or src:
            attrs = f' data-texts="{" ".join(texts_)}"' if texts_ else f' data-src="{E(src)}"'
            body += f'<details{attrs}><summary>Show license</summary><pre></pre></details>'
        return f'<section class="work{" derived" if derived else ""}"><div class="head">{head}</div>{body}</section>\n'

    site = ''.join([
        work('VibeOffice.work', 'Apache License 2.0', [('github.com/tcchen2026/VibeOffice', 'https://github.com/tcchen2026/VibeOffice')],
             'Quire, Ledger, Lectern, the Start Center and everything they share, including the PDF writer, ZIP engine, metafile player, icons and clip art; and the desktop app around them. The parts that come from the projects below say so in their files.',
             texts_=[apache]),
        # code and data derived from other projects: each file says "Derived from <name> …" (tests/licenses.test.mjs)
        work('Apache OpenOffice', 'Apache License 2.0', [('openoffice.org', 'https://www.openoffice.org'), ('github.com/apache/openoffice', 'https://github.com/apache/openoffice')],
             'Rewritten in JavaScript from its spreadsheet code: the depreciation and interest functions DDB, IPMT, AMORLINC and AMORDEGRC, and the scroll bars\' range. Quire sizes each line by what it holds following its text layout (Writer\'s line portions). Its English AutoCorrect list is the base of the default list. Compound-file import follows its tolerant sector and directory reads; OfficeArt import follows its picture headers, native shape definitions, geometry formulas, anchors and fills. Binary Excel import follows its text runs, styles, worksheet layout, view records, formula tokens, names, workbook links, comments, hyperlinks, validation, conditional formats, filters, printing and protection. Binary PowerPoint import follows its document, master, notes and color-scheme inheritance records, text formatting, rulers, placeholders, bullets and fields, animation builds, timing behaviors and media links. Binary Word import follows its property operands, sections, headers, lists and document compatibility settings.',
             files='public/ledger/js/fn-fin.js, public/ledger/js/grid.js, public/quire/js/portions.js, public/common/spell.js, public/common/crypto.js, public/common/legacy/encryption-crypto.js, public/common/legacy/xls-crypto.js, public/common/legacy/doc-crypto.js, public/common/legacy/officeart.js, public/common/legacy/native-geometry.js, public/common/legacy/shape-path.js, public/common/legacy/shapes.js, public/common/legacy/xls-read.js, public/common/legacy/xls-style.js, public/common/legacy/xls-sheet.js, public/common/legacy/xls-write.js, public/common/legacy/xls-formula.js, public/common/legacy/xls-formula-book.js, public/common/legacy/xls-misc.js, public/common/legacy/xls-rules.js, public/common/legacy/xls-comments.js, public/common/legacy/xls-draw-read.js, public/common/legacy/xls-draw-write.js, public/common/legacy/xls-chart-read.js, public/common/legacy/xls-chart-style.js, public/common/legacy/xls-chart-write.js, public/common/legacy/ppt-structure.js, public/common/legacy/ppt-text-read.js, public/common/legacy/ppt-text.js, public/common/legacy/ppt-shapes.js, public/common/legacy/ppt-tables.js, public/common/legacy/ppt-links.js, public/common/legacy/ppt-features.js, public/common/legacy/ppt-timing.js, public/common/legacy/ppt-timing-behavior.js, public/common/legacy/ppt-builds.js, public/common/legacy/ppt-media.js, public/common/legacy/ppt-timing-write.js, public/common/legacy/doc-sprm.js, public/common/legacy/doc-structure.js, public/common/legacy/doc-lists.js, public/common/legacy/doc-markup.js, public/common/legacy/doc-fields.js, public/common/legacy/doc-table-props.js, public/common/legacy/doc-settings.js, public/common/legacy/doc-tables.js, public/common/legacy/doc-xml.js, public/common/legacy/doc-draw-read.js, public/common/legacy/doc-draw-write.js', texts_=[apache, notice('apache-openoffice-NOTICE.txt')], derived=True),
        work('Apache POI', 'Apache License 2.0', [('poi.apache.org', 'https://poi.apache.org'), ('github.com/apache/poi', 'https://github.com/apache/poi')],
             'Approximate spreadsheet lookup and database criteria; compound-file header, allocation and directory handling; binary Office import: BIFF cell records, shared strings, phonetic text, styles, worksheet views, formula tokens, function metadata, shared and array formulas, data tables, names and external caches, comments, hyperlinks, validation, differential and conditional formats, filters, printing and protection, Word FIB, text pieces, formatting pages, styles, fonts, lists, sections, headers and notes, PowerPoint edit and persist records, document sizes, slide and notes associations, fonts and color schemes, text runs, master styles, placeholders, rulers, bullets and text fields, OLE properties, OfficeArt pictures, properties, gradients and shape geometry, PowerPoint animation and media records. Rewritten in JavaScript from Apache POI at ae62bb51; Copyright The Apache Software Foundation.',
             files='public/ledger/js/fn-lookup.js, public/ledger/js/fn-eng.js, public/common/crypto.js, public/common/legacy/encryption.js, public/common/legacy/xls-crypto.js, public/common/legacy/doc-crypto.js, public/common/legacy/ppt-crypto.js, public/common/legacy/xls-read.js, public/common/legacy/xls-style.js, public/common/legacy/xls-sheet.js, public/common/legacy/xls-write.js, public/common/legacy/xls-formula.js, public/common/legacy/xls-formula-book.js, public/common/legacy/xls-functions.js, public/common/legacy/xls-misc.js, public/common/legacy/xls-dxf.js, public/common/legacy/xls-rules.js, public/common/legacy/xls-comments.js, public/common/legacy/xls-draw-read.js, public/common/legacy/xls-draw-write.js, public/common/legacy/xls-chart-read.js, public/common/legacy/xls-chart-style.js, public/common/legacy/xls-chart-write.js, public/common/legacy/doc-read.js, public/common/legacy/doc-sprm.js, public/common/legacy/doc-structure.js, public/common/legacy/doc-style.js, public/common/legacy/doc-lists.js, public/common/legacy/doc-markup.js, public/common/legacy/doc-table-props.js, public/common/legacy/doc-fields.js, public/common/legacy/doc-draw-read.js, public/common/legacy/ppt-read.js, public/common/legacy/ppt-structure.js, public/common/legacy/ppt-text-read.js, public/common/legacy/ppt-text.js, public/common/legacy/ppt-shapes.js, public/common/legacy/ppt-tables.js, public/common/legacy/ppt-links.js, public/common/legacy/ppt-features.js, public/common/legacy/ppt-builds.js, public/common/legacy/ppt-media.js, public/common/legacy/ppt-text-write.js, public/common/legacy/properties.js, public/common/legacy/officeart.js, public/common/legacy/shapes.js',
             src='common/legacy/POI-LICENSE.txt', texts_=[apache, notice('../legacy/POI-NOTICE.txt')], derived=True),
        work('Unicode', 'Unicode License v3', [('unicode.org', 'https://www.unicode.org'), ('ICU mapping data', 'https://github.com/unicode-org/icu/tree/release-77-1/icu4c/source/data/mappings')],
             'Single-byte OEM and Macintosh codepage tables used to read binary Office text and properties. Generated from Unicode\'s published Microsoft mapping data and ICU 77.1. Copyright 1991-2026 Unicode, Inc.',
             files='public/common/legacy/codepages.js', texts_=[notice('../legacy/UNICODE-LICENSE.txt')], derived=True),
        work('commonmark.js', 'BSD-2-Clause', [('github.com/commonmark/commonmark.js', 'https://github.com/commonmark/commonmark.js')],
             'Quire\'s Markdown reader is a condensed translation of its block and inline parsers. Copyright (c) 2014 John MacFarlane.',
             files='public/quire/js/markdown.js', texts_=[notice('commonmark.js-LICENSE.txt')], derived=True),
        work('puff', 'zlib License', [('github.com/madler/zlib', 'https://github.com/madler/zlib/tree/develop/contrib/puff')],
             'The ZIP engine\'s JavaScript inflater, used where the browser can\'t decompress, is an altered translation of puff.c. Copyright (C) 2002-2013 Mark Adler.',
             files='public/common/zip.js', texts_=[notice('puff-LICENSE.txt')], derived=True),
        work('SheetJS SSF', 'Apache License 2.0', [('github.com/SheetJS/ssf', 'https://github.com/SheetJS/ssf')],
             'Two helpers of the General number format, for trailing zeros and exponents. Copyright (C) 2013-present SheetJS LLC.',
             files='public/common/numfmt.js', texts_=[apache, notice('sheetjs-ssf-NOTICE.txt')], derived=True),
        work('Microsoft VBA documentation', 'CC BY 4.0', [('github.com/MicrosoftDocs/VBA-Docs', 'https://github.com/MicrosoftDocs/VBA-Docs')],
             'Ledger\'s function descriptions and argument help follow the Excel WorksheetFunction pages of this documentation, in places shortened or adapted to Ledger. Copyright (c) Microsoft Corporation.',
             files='public/ledger/js/fninfo.js', texts_=[notice('microsoft-vba-docs-LICENSE.txt')], derived=True),
        work('Open XML SDK', 'MIT License', [('github.com/dotnet/Open-XML-SDK', 'https://github.com/dotnet/Open-XML-SDK')],
             'From its schema data (commit 431ab05), the order of Word\'s run properties with their Office 2010 effects, and the limits of Word\'s revision and numbering IDs, in the generated element order. Version 3.3.0 schema data also informs property and compatibility element order in binary Word import and font child order in binary Excel import. Copyright (c) .NET Foundation and Contributors.',
             files='public/common/legacy/xls-style.js, public/common/opc-order.js, public/common/legacy/doc-xml.js, public/common/legacy/doc-tables.js, public/common/legacy/doc-settings.js, public/common/legacy/doc-fields.js, public/common/legacy/doc-draw-write.js', texts_=[notice('openxml-sdk-LICENSE.txt')], derived=True),
        work("Peter Norvig's spelling corrector", 'MIT License', [('norvig.com/spell-correct.html', 'https://norvig.com/spell-correct.html')],
             'The spelling checker\'s one-edit candidates (edits1), in JavaScript. Copyright (c) 2007-2016 Peter Norvig.',
             files='public/common/spell.js', texts_=[notice('norvig-spell-LICENSE.txt')], derived=True),
        work('Carlito, Caladea and Gelasio fonts', 'SIL Open Font License 1.1', [('github.com/google/fonts', 'https://github.com/google/fonts')],
             'Metric-compatible stand-ins for Calibri, Cambria and Georgia, for computers that do not have them. Copyright 2013 The Carlito Project Authors; 2012 The Caladea Project Authors; 2022 The Gelasio Project Authors.',
             files='public/fonts/*.woff2', src='fonts/LICENSES.txt'),
        work('English word lists and thesaurus', 'SCOWL license; WordNet 2.1 license',
             [('wordlist.aspell.net', 'http://wordlist.aspell.net'), ('LibreOffice dictionaries', 'https://github.com/LibreOffice/dictionaries')],
             'The spelling word lists (US and UK English) from SCOWL and the English Speller Database by Kevin Atkinson, and the thesaurus from LibreOffice\'s English dictionary package, derived from WordNet (Princeton University).',
             files='public/common/dict/en_US.words, en_GB.words, en.thes', src='common/dict/LICENSES.txt'),
        work('Office Open XML (ECMA-376)', 'Ecma International', [('ecma-international.org', 'https://ecma-international.org/publications-and-standards/standards/ecma-376/')],
             'The order of elements in each part of a file, and the shape presets\' adjustment values, generated from the schemas and preset definitions of ECMA-376 Office Open XML File Formats, fifth edition (Parts 1 and 4).',
             files='public/common/opc-order.js'),
    ])

    copyleft = sorted(f"{c['name']} {c['version']} ({c['license'] or 'no licence'})" for c in crates.values() if not without_copyleft(c['license']))
    if copyleft:
        sys.exit('Shipped crates whose licence needs copyleft (AGENTS.md, Code from other projects): ' + ', '.join(copyleft))
    if '--check' in sys.argv:
        with open(OUT, encoding='utf-8') as f:
            listed = set(re.findall(r'<b>([^<]+ \d[^<]*)</b>', f.read()))
        wanted = {f"{c['name']} {c['version']}" for c in crates.values()}
        missing, extra = sorted(wanted - listed), sorted(listed - wanted)
        if missing or extra:
            sys.exit(f'{os.path.relpath(OUT, ROOT)} is out of date (python3 tools/oss-licenses.py): missing {missing}, not shipped {extra}')
        print(f'{os.path.relpath(OUT, ROOT)}: lists the {len(wanted)} shipped crates; none needs a copyleft licence')
        return

    names = {'windows': 'Windows', 'macos': 'macOS', 'linux': 'Linux'}
    desktop = work('WebView2Loader.dll', 'BSD-3-Clause (Microsoft WebView2 SDK)', [('aka.ms/webview2', 'https://aka.ms/webview2')],
                   'Finds the system\'s WebView2 (Edge) runtime; ships beside vibeoffice.exe. Copyright Microsoft Corporation.', files='Windows', texts_=[webview2])
    desktop += work('AppImage runtime', 'MIT; with libfuse (LGPL-2.1), squashfuse, zstd, zlib and musl inside',
                    [('github.com/AppImage/type2-runtime', 'https://github.com/AppImage/type2-runtime/tree/20251108')],
                    'The first part of the Linux download (VibeOffice-Linux-x86_64.AppImage), unmodified: a program of the AppImage project (release 20251108) that mounts the rest and starts the app. Copyright probonopd and the AppImage contributors.',
                    files='Linux', texts_=[appimage])
    missing = []
    for c in sorted(crates.values(), key=lambda c: (c['name'].lower(), c['version'])):
        if not c['texts']:
            missing.append(c['name'])
        on = [names[p] for p in TARGETS if p in c['platforms']]
        where = ', '.join(on[:-1]) + ' and ' + on[-1] if len(on) > 1 else on[0]   # Windows, macOS and Linux
        desktop += work(f"{c['name']} {c['version']}", c['license'], [(re.sub(r'^https?://', '', c['url']).rstrip('/'), c['url'])], '',
                        files=where + ('' if c['texts'] else ' · no license file in the package: see its repository'), texts_=c['texts'])

    page = TEMPLATE.replace('{{SITE}}', site).replace('{{DESKTOP}}', desktop).replace('{{COUNT}}', str(len(crates)))
    page = page.replace('{{TEXTS}}', json.dumps(texts, ensure_ascii=False).replace('</', '<\\/'))
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(page)
    print(f'{os.path.relpath(OUT, ROOT)}: {len(crates)} crates, {len(texts)} distinct texts, {len(page) // 1024} KB'
          + (f'; no license file: {", ".join(missing)}' if missing else ''))


TEMPLATE = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self' data: blob: https://vibeoffice.work ipc: https://ipc.localhost; worker-src 'self'; manifest-src 'self'; frame-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Open source licenses - VibeOffice.work</title>
<link rel="icon" href="icons/icon.svg" type="image/svg+xml">
<script src="common/look.js"></script>
<link rel="stylesheet" href="common/looks.css">
<!-- Generated by tools/oss-licenses.py: edit that, not this page. -->
<style>
  /* the licenses of the works this repository is made with, as a browser's credits page: one entry per work,
     its license shown on request. Colours come from the look settings (common/looks.css). */
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--surface); color: var(--ink); font: 13px/1.5 var(--ui-font); }
  main { max-width: 860px; margin: 0 auto; padding: 24px 20px 40px; }
  h1 { display: flex; align-items: center; gap: 10px; margin: 0 0 6px; font-size: 20px; }
  h1 img { width: 32px; height: 32px; }
  h2 { margin: 26px 0 4px; font-size: 15px; }
  .intro { margin: 0 0 14px; color: var(--muted); }
  .work { padding: 9px 0; border-top: 1px solid var(--chrome-line); }
  .work:last-child { border-bottom: 1px solid var(--chrome-line); }
  .head { display: flex; flex-wrap: wrap; align-items: baseline; gap: 4px 12px; }
  .head b { font-size: 14px; }
  .head a { color: var(--pane-link); overflow-wrap: anywhere; }
  .lic, .files { color: var(--muted); }
  .what, .files { margin: 2px 0 0; }
  .files { font-size: 12px; }
  details { margin-top: 4px; }
  summary { cursor: pointer; color: var(--pane-link); width: max-content; }
  pre { margin: 6px 0 0; padding: 10px 12px; max-height: 420px; overflow: auto; white-space: pre-wrap; border: 1px solid var(--chrome-line);
    border-radius: 3px; background: var(--group-bg); font: 11px/1.45 ui-monospace, Consolas, monospace; }
</style>
</head>
<body>
<main>
  <h1><img src="icons/icon.svg" alt="">Open source licenses</h1>
  <p class="intro">VibeOffice.work is open source, and it is made with the open source works below.</p>

  <h2>The site</h2>
  <p class="intro">The files of the web pages: the apps, the Start Center and what they share.</p>
{{SITE}}
  <h2>The desktop app</h2>
  <p class="intro">VibeOffice for Windows and Linux, around the same pages: every Rust crate compiled into the app and its installer (Tauri and everything it brings, not the tools that only run while building) and the DLLs that ship with it.</p>
{{DESKTOP}}
</main>
<script type="application/json" id="texts">{{TEXTS}}</script>
<script src="oss-licenses.js"></script>
</body>
</html>
'''

if __name__ == '__main__':
    sys.exit(main())
