# Quire 2003 — Web Edition

A Word 2003–style word processor written in plain HTML and vanilla JavaScript (no build step, no
libraries). It reads and writes Office Open XML (`.docx`, `.docm`, `.dotx`, `.dotm`) files.

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
| Spelling (as you type and Tools ▸ Spelling and Grammar) | Word lists for English (U.S.) and English (U.K.) in `common/dict/en_US.words` and `common/dict/en_GB.words`, expanded from the Hunspell/SCOWL dictionaries and ranked by word frequency so suggestions come out in a sensible order. Unknown words get red wavy underlines (CSS Custom Highlight API); right-click for suggestions, Ignore All and Add to Dictionary. Underlines are off when a document opens (and the dictionary is not loaded): click the spelling icon in the status bar to check that document as you type, click again to stop. |
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
| Markdown `.md` / `.markdown` (GitHub flavour) | yes | yes |
| Markdown with pictures `.zip` (a `.md` and its picture files) | yes | yes |
| PDF | — | yes (File ▸ Print, or Save As ▸ PDF) |
| Password-protected `.docx` | yes (Word 2007 "standard" and Word 2010+ "agile" encryption) | yes: Tools ▸ Options ▸ Security ▸ Password to open (AES-256, as Word 2013+) |

Encryption and decryption happen in the browser (WebCrypto for hashing, so the page must be served over
https or from localhost). A document opened with its password keeps that password when saved.

## Markdown

Quire reads and writes the Markdown GitHub uses: CommonMark plus tables, task lists, strikethrough,
extended autolinks, footnotes and alerts (`> [!NOTE]`).

| In the file | In Quire | Written back as |
| --- | --- | --- |
| `#` … `######` headings | Heading 1–6 | `#` headings (Title as `#`, Subtitle as `##`) |
| `**bold**`, `*italic*`, `~~strike~~`, `` `code` `` | Bold, italic, strikethrough, the HTML Code character style | The same; where a delimiter could not open or close, `<strong>`, `<em>`, `<del>` |
| Links, autolinks, `[text](#heading)` | Hyperlinks (titles as tips); `#heading` follows GitHub's heading anchors | The same |
| `-` / `1.` lists, nested, tight or loose | Bullets and numbering per list level (List Paragraph) | `-` and `1.`; tight unless the items have space between them |
| `- [ ]` / `- [x]` task lists | Check boxes that tick on click | `[ ]` / `[x]` |
| `>` quotes, nested; `> [!NOTE]` and the other four alerts | Quote style with a bar per level; alerts with their title and colour | `>` and `[!KIND]` |
| Fenced and indented code | HTML Preformatted paragraphs, the language kept | Fenced code with the language |
| Pipe tables with alignment | Tables (Markdown Table style: header row, banded rows) | Pipe tables; merged cells or several paragraphs in a cell as an HTML `<table>` |
| `---` | A paragraph with a bottom border | `---` |
| `[^label]` footnotes | Footnotes | `[^label]` and its definition |
| Pictures | Pictures from `data:` URLs, or from the `.zip`; others as a box with the alt text, the address kept | `![alt](path)`; a resized picture as `<img width>`; new pictures go into the `.zip` as `images/…` |
| `$…$` math, emoji `:shortcodes:`, Mermaid and other diagram code | Kept as written (math in the Markdown Source character style, diagrams as code) | Unchanged |
| Inline HTML GitHub allows (`<sub>`, `<sup>`, `<ins>`, `<kbd>`, `<br>`, `<img>`, `<a href>`, `<a name>`) | Subscript, superscript, underline, HTML Keyboard style, line break, picture, link, bookmark | The same tags |
| HTML blocks, YAML front matter, link reference definitions, comments | Blocks through the HTML reader; front matter, definitions and comments are kept but not shown | Unchanged unless edited |

Saving keeps the file as it was where it was not edited: each block of the file remembers its lines,
and a block whose content is unchanged is written back byte for byte (line endings, a byte-order mark,
`*` or `_`, setext headings, reference links and line wrapping included). Edited and new blocks are
written in one style. What Markdown cannot express (fonts, colours, alignment, page setup, headers and
footers) is not saved; a `.md` file cannot hold pictures, so saving one that has new pictures offers
the `.zip` form instead. A `.zip` is saved with every other file it held, unchanged. Unsaved versions
of a Markdown document (autosave, Document Recovery) are kept as Markdown.

Checks (tools in docs/testing.md):

| Check | Result |
| --- | --- |
| Reader against the CommonMark 0.31.2 spec examples (rendered as the reference implementation renders) | 652 of 652 |
| Reader against the GFM spec's extension examples (tables, task lists, strikethrough, autolinks, disallowed raw HTML) | 22 of 22 |
| Spec examples opened, every block rewritten, saved and read again, same HTML | 559 of 652; 507 of the 544 without raw HTML (the rest: nested emphasis of the same kind, `*a *b**`, which a document cannot hold; empty links; a quote that starts inside a list item's first line) |
| 24 real files (READMEs of React, Node.js, VS Code, Rust, Kubernetes, TensorFlow, Vue, Deno, Astro, Express, Go, TypeScript, cmark-gfm, markdown-it, awesome, public-apis, free-programming-books, the Airbnb style guide, GitHub's own formatting docs) saved unchanged | 24 of 24 byte for byte |
| The same files with one paragraph edited | Only the edited block's lines change in all 24 |
| The same files with every block rewritten, same HTML | 7 of 24; all 5 without raw HTML (the others turn logos, `<picture>` and centred `<p align>` blocks into formatting) |
| Opening time | 300 kB (public-apis): 80 ms; the spec examples twenty times over (300 kB): 70 ms to parse and render |

## What a save keeps, converts and drops

Quire regenerates document content from its model and carries original package dependencies alongside
it. Package preservation is implemented; content controls, opaque objects and unknown run/paragraph
properties are still being integrated. The Compatibility Checker reports recorded losses before a
user save. Its coverage is incomplete until those remaining object paths populate the loss ledger.

### Package preservation (2026-10-08)

Save, Save As and draft recovery support `.docx`, `.docm`, `.dotx` and `.dotm`, including encryption.
The main-part content type determines the imported variant. Macros remain in macro-enabled saves;
choosing a macro-free type reports their removal before download. Cancelling keeps the current file
identity and does not acknowledge the loss.

The pinned 2,902-file run saves and reopens **2,880**, with the same **22 failures** as the baseline.
Package/SDK comparison reports **2,792 OK and 110 failed**, up from 2,768 OK. Failed originals and
remaining content-writer diagnostics are counted, not accepted as passes. The independent package
audit checks original bytes, content types and relationship identities:

| Package feature | Preserved relationships |
|---|---:|
| Custom XML stores / their property parts | 965 / 965 each |
| Custom document properties, including empty bags and original types/pids | 353 / 353 |
| Attached templates | 102 / 102 |
| Referenced embedded fonts | 20 / 20 |
| VBA projects | 17 / 17 |
| Glossary | 193 / 195 |
| People / comment identity parts | 95 / 95; 15 / 15 |
| Sensitivity labels | 1 / 1 |

The two glossary exceptions have missing source dependencies. Orphaned parts with no incoming
relationship (including three font files) are not reattached; drops are recorded. Settings retain
unowned children, font-table entries retain embedding metadata, and existing comment paragraph IDs
and special footnote/endnote entries survive. General drawing/bookmark identity preservation remains
part of the next object stage.

Reports: `~/corpora/results/package-preservation-2026-10-08/`. Word samples are in
`~/Downloads/lossless-check/quire/package/`; Office acceptance is pending. The 24 Markdown fixtures
still save byte for byte unchanged, and each scripted edit stays within its edited block.

### Historical feature baseline

The following larger-corpus figures predate package preservation. `tools/quire/test/loss-audit.py`
compared 2,778 test documents with saved copies, feature by feature and word by word
(`python3 tools/quire/test/loss-audit.py originals/ saved/ out.json`, standard library only).

**Kept.** The text: 2,724 documents keep every word, and over the whole set 666,082 of 666,941 words
survive (the 859 missing sit in 54 documents, most of them deliberately malformed test files, comments
or footnotes that nothing in the text refers to, and text inside data-bound content controls). Styles,
lists, tables, sections and columns, headers and footers, footnotes and endnotes, comments with their
replies, tracked changes, fields, bookmarks, hyperlinks, pictures, text boxes, drawing shapes, equations,
charts, page borders, line numbers, legacy form fields, ruby text, custom document properties and
editing restrictions are all written back.

**Converted.** Embedded objects (Excel, Visio, PDF and other OLE packages, 81 documents) are saved as
their preview pictures, so they can no longer be opened for editing; SmartArt becomes grouped shapes
with its text (22); content controls are removed and their contents kept as ordinary text (296), which
also ends any binding to document data; smart tags and custom XML markup become plain text (48);
`HYPERLINK` fields become ordinary hyperlinks.

**Dropped.**

| What | Documents in the test set |
| --- | --- |
| Custom XML data parts (document-management metadata, content-control bindings) | 573 |
| Sensitivity labels in the newer `docMetadata/LabelInfo.xml` form (labels kept as `MSIP_Label` custom properties survive: 11 of 11) | 3 |
| Building blocks and AutoText stored in the document (glossary) | 129 |
| Link to the attached template | 77 |
| Page background with a picture or fill effect (coloured backgrounds are kept; 45 white ones are left out, which changes nothing) | 3 of 68 |
| Word 2010 text effects (glow, outline, shadow, ligatures) | 35 |
| Embedded fonts | 10 |
| Embedded HTML/RTF documents (`altChunk`) — their content is not shown, so it is lost on save | 4 |
| Permission ranges for restricted editing | 3 |
| VBA macros (a `.docm` is saved without its project) | — |

### Pinned lossless-save baseline (2026-10-07)

The reproducible three-source corpus in `tools/corpora.sh` was measured against application revision
`5abafdc`, before preservation changes. Of 2,902 inputs, 2,880 completed open → save → reopen,
22 failed and 0 were explicitly excluded. Completion is not a fidelity result. The independent
feature inventory compared 2,868 readable pairs; malformed/encrypted originals and missing
outputs remain in its accounting. This corpus differs from the earlier compatibility corpus above.

| Feature | Files containing it | Original inventory items | Saved inventory items | Text edit: kept / original |
|---|---:|---:|---:|---:|
| content controls | 379 | 2,696 | 0 | 0 / 2,594 |
| custom XML data parts | 646 | 966 | 0 | 0 / 814 |
| embedded OLE objects | 82 | 544 | 0 | 0 / 434 |
| VBA macros | 17 | 17 | 0 | 0 / 15 |

The text-edit inventory covers 2,236 readable saved pairs; files without editable text and failed
operations stay in the driver report, so its denominator differs from the unedited inventory.

Package checks plus the Office 2019 SDK comparison reported 2,768 attempts without new
automated diagnostics and 134 failures (including originals that could not be validated);
0 driver exclusions remain separate. These checks do not certify Office acceptance. Commands,
comparison policies and failure accounting are in [testing.md](testing.md#lossless-same-format-save-reproducible-baseline).

The text-edit/save/undo/redo baseline completed in 2,248 files, excluded 592 without a suitable
editable paragraph and failed in 62. Forty failures expose circular textbox-owner references in
JSON history snapshots; 22 are read failures.

The history prerequisite is now fixed: textbox owners are runtime-only back-references, and
original chart bytes live in the shared chart cache instead of JSON snapshots. All 40 previously
failing files complete text edit → save → undo → save → redo → save. Their undo packages match
the pre-edit baseline part for part, except generated core-property timestamps in one file.
The 40 edited outputs have no new package/SDK diagnostics compared with the pre-edit saves and
all open in LibreOffice. A chart regression also checks all four original chart parts byte for
byte across six states (open, edit, undo, redo, copy and snapshot restore). Samples are in
`~/Downloads/lossless-check/quire/chart-history/` and `textbox-history/`.

All 10 feature samples converted in LibreOffice with unchanged page counts. The full corpus
has not yet been rendered in LibreOffice. The baseline reports are retained outside git in
`~/corpora/results/lossless-baseline-2026-10-07/`; Office samples are in
`~/Downloads/lossless-check/baseline/`.

### Shared preservation core (2026-10-08)

The shared XML/OPC core now captures original AlternateContent before Quire normalises it. The writer
does not yet emit these records, so the feature-loss numbers above still apply. The same 2,902 inputs
again produced 2,880 successful open/save/reopen attempts and 22 failures. Compared with the earlier
saved outputs, 2,875 package/SDK comparisons completed with no new diagnostics; 5 earlier outputs
could not be SDK-validated and 22 had no saved output. Ten LibreOffice samples kept their page counts
relative to the earlier saves. Markdown remained byte-identical for all 24 unchanged files, with edits
confined to the edited block in all 24. Reports: `~/corpora/results/opc-core-2026-10-08/`.

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
| `../common/core.js`, `../common/zip.js` | Utilities, colour maths, storage, ZIP reader/writer, PDF writer, downloads |
| `../common/xml.js`, `../common/opc.js`, `../common/opc-order.js` | Shared OOXML package graph, identity reservations, AlternateContent capture and schema-order merges |
| `../common/sha.js`, `../common/crypto.js` | Password-protected documents: compound-file reader/writer, AES, ECMA-376 agile and standard encryption |
| `js/dmodel.js`, `js/ops.js` | Document model (paragraphs, runs, tables, sections, styles, numbering), undo history, editing operations |
| `js/render.js`, `js/layout.js` | Paragraph/run rendering, pagination, headers/footers, footnotes, columns, floats |
| `js/editor.js` | contentEditable bridge: selection mapping, typing, IME, clipboard, keyboard |
| `js/preserve.js` | Package ownership and settings, font-table and document-property merges |
| `js/docx-read.js`, `js/docx-write.js`, `../common/dml.js` | `.docx` import/export, DrawingML/VML shapes, charts, themes |
| `js/rtf.js`, `js/htmlio.js` | RTF and HTML import/export |
| `js/markdown.js`, `js/mdio.js` | Markdown reader (CommonMark + GitHub extensions, with source lines; also loads in Node) and Markdown/`.zip` import and export |
| `js/fields.js` | Field engine (PAGE, DATE, REF, SEQ, TOC, INDEX, IF, formulas, …), footnotes and endnotes |
| `js/review.js` | Track changes, comments, balloons, Reviewing pane, compare documents |
| `js/tables.js` | Tables and Borders toolbar, table styles, draw/erase, merge/split, sort, formulas |
| `js/drawing.js`, `js/drawtool.js`, `../common/geometry.js`, `../common/charts.js`, `../common/metafile.js`, `../common/clipart.js` | Pictures, AutoShapes, WordArt, diagrams, charts, WMF/EMF, clip art |
| `js/autocorrect.js`, `../common/spell.js`, `js/spell.js` | AutoCorrect, AutoFormat as you type, AutoText; the shared spelling engine and Quire's Spelling and Grammar (offline) |
| `js/omml.js` | Office Math (OMML) to MathML for equations |
| `../common/dict/` | Spelling word lists, thesaurus, licences (shared with Ledger) |
| `js/outline.js`, `js/rulers.js`, `js/panes.js` | Outline view, rulers, task panes (Research/Thesaurus, Clip Art, Styles…), Document Map, thumbnails, Print Preview, Reading Layout |
| `js/mailmerge.js`, `js/templates.js` | Mail Merge wizard and toolbar, document templates |
| `../common/ui.js`, `../common/icons.js`, `../common/luna.css` | Menus, toolbars, dialogs, colour pickers, original 16×16 icon set, the Luna look (shared) |
| `js/commands.js`, `js/dialogs*.js`, `js/find.js`, `js/app.js` | Command registry, all dialogs, Find and Replace, application controller |

Quire is an independent program. Microsoft, Word and Office are trademarks of Microsoft Corporation;
Quire is not affiliated with Microsoft.
