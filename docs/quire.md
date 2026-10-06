# Quire 2003 — Web Edition

A Word 2003–style word processor written in plain HTML and vanilla JavaScript (no build step, no
libraries). It reads and writes Office Open XML (`.docx`) files.

## Running it

The app is `public/quire/`; every path below is relative to it. Serve `public/` and open it:

    python3 tools/serve.py

then visit http://127.0.0.1:8760/quire/. Everything runs in the browser: there is no back end, no account
and no network service behind any feature. The proofing dictionaries are fetched at run time from
`../common/dict/` (shared with Ledger), which is why the site should be served rather than opened from
disk (a `file://` page cannot fetch them; editing still works, but spelling and the thesaurus stay off).
The metric-compatible fonts (see *Rendering notes*) currently come from Google Fonts when it can be reached.

## Proofing tools (offline)

| Feature | How it works |
| --- | --- |
| Spelling (as you type and Tools ▸ Spelling and Grammar) | Word lists for English (U.S.) and English (U.K.) in `common/dict/en_US.words` and `common/dict/en_GB.words`, expanded from the Hunspell/SCOWL dictionaries and ranked by word frequency so suggestions come out in a sensible order. Unknown words get red wavy underlines (CSS Custom Highlight API); right-click for suggestions, Ignore All and Add to Dictionary. |
| Grammar | A rule checker for the mistakes Word 2003 flagged most often: repeated words, a/an, capitalisation at sentence start, spacing around punctuation, commonly confused words, simple subject–verb agreement. Green wavy underlines. |
| Thesaurus and Research pane | `common/dict/en.thes`, built from the WordNet 3.1 thesaurus (MyThes format), with meanings grouped by part of speech, related words, antonyms and a Back history. "Search This Document" finds the word in the open document. |
| Custom dictionary | Tools ▸ Options ▸ Spelling & Grammar; words you add are stored in the browser. |
| Readability statistics | Tools ▸ Options ▸ Spelling & Grammar ▸ Readability Statistics (counts, Flesch Reading Ease, Flesch–Kincaid Grade Level, passive sentences). |

Licences for the word lists are in `public/common/dict/LICENSES.txt`.

## Rendering notes

- Layout follows Word's own rules where they differ from CSS: space before/after collapses to the
  larger of the two (unless the document asks for HTML-style spacing), space before is dropped at the
  top of a page reached by overflow, by "page break before" or (Word 2013+ documents) by a section
  break, a line with "multiple" line spacing only needs its single-spaced height to fit at the bottom
  of a page, and Word 2013+ justified text may squeeze spaces by up to 20% to fit a word on a line.
- More of Word's own habits: a right or centred tab stop placed beyond the right indent lets the page
  number of a table of contents run out to it; superscripts and footnote references do not push their
  line down; table rows are as tall as Word makes them even though the browser draws hairline borders a
  whole pixel wide; East Asian documents with a line grid snap their lines to it; a "next column"
  section break in a one-column layout starts a new page; "page break before" on the first paragraph
  of a table breaks before the table; Word 2013+ documents drop space before after a manual page break,
  and keep the paragraph mark of a page break that ends the document on the same page.
- Equations (Office Math, `m:oMath`) are rendered as MathML: fractions, scripts, radicals, delimiters,
  n-ary operators, accents, matrices and equation arrays. They are kept unchanged when saving.
- Content controls bound to custom XML or to the document properties show the bound values, as Word
  does when it opens a file. Linked text boxes continue their story from one box to the next.
- SmartArt is drawn from the picture Word stores with it; files saved without that picture (Word 2007
  and some generators) get a simple diagram built from the SmartArt data.
- Fonts are drawn with the fonts installed on your computer. When a document uses a font you don't
  have, Quire substitutes a metric-compatible one where one exists (Carlito for Calibri, Caladea for
  Cambria, Tinos/Arimo/Cousine for Times New Roman/Arial/Courier New, Gelasio for Georgia, TeX Gyre
  fonts for Palatino, Bookman, Century Schoolbook and Century Gothic), so line breaks and page counts
  stay close to Word's.

## File formats

| Format | Open | Save |
| --- | --- | --- |
| Word Document `.docx` / `.docm` | yes | yes |
| Word Template `.dotx` / `.dotm` | yes | yes |
| Web Page `.htm` / `.html` | yes | yes (single file, images inlined) |
| Rich Text Format `.rtf` | yes | yes |
| Plain Text `.txt` | yes | yes |
| PDF | — | yes (File ▸ Print, or Save As ▸ PDF) |
| Password-protected `.docx` | yes (Word 2007 "standard" and Word 2010+ "agile" encryption) | yes: Tools ▸ Options ▸ Security ▸ Password to open (AES-256, as Word 2013+) |

Encryption and decryption happen in the browser (WebCrypto for hashing, so the page must be served over
https or from localhost). A document opened with its password keeps that password when saved.

## Compatibility testing

Quire was checked against 2,808 `.docx` files that other people wrote, not files made for it:

| Source | Files |
| --- | --- |
| LibreOffice import/export regression documents | 1,703 |
| Open XML SDK test documents | 378 |
| docx-templates and docxtemplater fixtures | 279 |
| docx4j | 141 |
| Apache POI | 124 |
| pandoc | 87 |
| Apache Tika | 67 |
| Real documents from public repositories (theses, thesis templates, CVs, letters) | 30 |

Each file was opened in headless Chromium by the same code that ships here, laid out page by page,
saved back to `.docx`, and compared with three outside references:

- **Word itself.** When Word saves a file it records how many pages it had (`docProps/app.xml`).
  2,207 files carry that number.
- **LibreOffice 24.2**, converting the original file to PDF.
- **The round trip**: the file Quire saved, opened again by LibreOffice.

| Check | Result |
| --- | --- |
| Files opened | 2,799 of 2,808. The other nine are broken on purpose (fuzzing cases, truncated or damaged packages, two OpenDocument files renamed `.docx`, a 5,000-level nested table deeper than the browser's XML parser allows, and one encrypted test file whose password is not published) |
| Files saved without error | 2,799 of 2,799 |
| Page count equal to Word's | 2,059 of 2,207 (93.3%); within one page: 98.2%. LibreOffice on the same files: 92.9% and 97.6% |
| Page count equal to LibreOffice's | 2,593 of 2,765 (93.8%); within one page: 98.8% |
| Saved file opens in LibreOffice | Every file whose original LibreOffice can open (eight newer test files fail in LibreOffice 24.2 before and after) |
| Text kept through the round trip (word-level recall ≥ 98%) | 96.7% of comparable files; page count unchanged in 98.6% |
| Password-protected files | 9 of 10 test files open with their published passwords; files Quire encrypts open in LibreOffice with the right password and are refused with a wrong one |
| Opening time | median 25 ms, 90th percentile 73 ms, slowest 6.3 s (a 182-page table-heavy file) |

Scripted editing (typing, bold, a new paragraph, inserting a table, undo and redo, then save and reopen)
was also run on the real documents; every change survived the save and showed up when LibreOffice
opened the result.

Page counts are a blunt instrument, so the misses were also studied page by page against Word's own
page-break markers (`w:lastRenderedPageBreak`) and LibreOffice's PDF. The rules listed under
*Rendering notes* came out of that work. What is still known to differ: documents set in Palatino-like
fonts that are not installed (the substitute has slightly different widths), text wrapped tightly around
irregular picture outlines, vertical (East Asian) text direction, content-control bindings whose value is
rich text, and RC4-encrypted files from Office 97–2003.

## Source layout

| File | What it does |
| --- | --- |
| `index.html` | Window chrome, CSS (Office 2003 Luna Blue), script order |
| `js/core.js`, `js/zip.js` | Utilities, colour maths, storage, ZIP reader/writer, PDF writer, downloads |
| `js/crypto.js` | Password-protected documents: compound-file reader/writer, AES, ECMA-376 agile and standard encryption |
| `js/dmodel.js`, `js/ops.js` | Document model (paragraphs, runs, tables, sections, styles, numbering), undo history, editing operations |
| `js/render.js`, `js/layout.js` | Paragraph/run rendering, pagination, headers/footers, footnotes, columns, floats |
| `js/editor.js` | contentEditable bridge: selection mapping, typing, IME, clipboard, keyboard |
| `js/docx-read.js`, `js/docx-write.js`, `js/dml.js` | `.docx` import/export, DrawingML/VML shapes, charts, themes |
| `js/rtf.js`, `js/htmlio.js` | RTF and HTML import/export |
| `js/fields.js` | Field engine (PAGE, DATE, REF, SEQ, TOC, INDEX, IF, formulas, …), footnotes and endnotes |
| `js/review.js` | Track changes, comments, balloons, Reviewing pane, compare documents |
| `js/tables.js` | Tables and Borders toolbar, table styles, draw/erase, merge/split, sort, formulas |
| `js/drawing.js`, `js/drawtool.js`, `js/geometry.js`, `js/charts.js`, `js/metafile.js`, `js/clipart.js` | Pictures, AutoShapes, WordArt, diagrams, charts, WMF/EMF, clip art |
| `js/autocorrect.js`, `js/spell.js` | AutoCorrect, AutoFormat as you type, AutoText, Spelling and Grammar (offline) |
| `js/omml.js` | Office Math (OMML) to MathML for equations |
| `../common/dict/` | Spelling word lists, thesaurus, licences (shared with Ledger) |
| `js/outline.js`, `js/rulers.js`, `js/panes.js` | Outline view, rulers, task panes (Research/Thesaurus, Clip Art, Styles…), Document Map, thumbnails, Print Preview, Reading Layout |
| `js/mailmerge.js`, `js/templates.js` | Mail Merge wizard and toolbar, document templates |
| `js/ui.js`, `js/icons.js` | Menus, toolbars, dialogs, colour pickers, original 16×16 icon set |
| `js/commands.js`, `js/dialogs*.js`, `js/find.js`, `js/app.js` | Command registry, all dialogs, Find and Replace, application controller |

Quire is an independent program. Microsoft, Word and Office are trademarks of Microsoft Corporation;
Quire is not affiliated with Microsoft.
