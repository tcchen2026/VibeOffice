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
| Strict Open XML `.docx` | yes | standard Transitional OOXML, with a Compatibility Checker notice |
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

Quire keeps package dependencies, control boundaries, text properties and opaque object XML beside
its editable model. Save and drafts retain DOCX/DOCM/DOTX/DOTM and encryption. The Compatibility
Checker reports conversions from the prepared output before a user download. Drafts retain notices
silently, and cancelling a conversion does not leave its notice on a later preserving save.

The final pinned **2,902-file** run uses writer `69be775`, resumed at `5247af8` after the byte-identical
XML patch performance correction. **2,880 save/reopen and draft-recovery attempts complete**, with
**22 failures** in each scenario. Text edit/save/undo/redo completes in **2,281**, with **599 exclusions
and 22 failures**. The failed inputs are password-protected, malformed or mislabeled. One timeout
from the initial slow writer has a passing targeted correction; its original failure remains recorded.

The full package/SDK comparison covers **19,028 emitted/attempted states**: **18,025 OK, 404 failed
and 599 excluded**. Save/reopen alone reports **2,823 OK and 79 failed**. Failed or unopenable originals
stay in those totals. No output adds package-consistency issues. The previous duplicate compatibility
attribute and drawing-ID defects are fixed; recovery and generated drawings now reserve their identities.
Six malformed-formatting inputs remain exceptions: `Tdf147485`, `tdf147485-forcepoint`, `tdf115212`,
`tdf91095`, `tdf57589_hashColor`, and the edited `tdf149198`. Reordering or splitting their already
invalid properties adds diagnostics; these are failed comparisons, not passes. Diagnostic matching
follows style IDs and unchanged property content when positions move, while retaining occurrence counts.
Newly generated drawings reserve both document drawing IDs and part-local shape IDs, so converting
one object cannot collide with the identity of an unrelated retained group.
An Office-found comment repair is also fixed: pre-release comment extensions retain their original
namespace and unchanged bytes, and a comment is written with one reference. Removing a duplicate
reference records a Compatibility Checker entry. The pinned scan finds one saved file with excess
anchors and 55 automatic extension-namespace promotions; the corrected point-comment save and a
current-format reply control both pass Word. The final full run includes these corrections.

Rewritten parts now hoist fragment namespaces and `mc:Ignorable` to the root. On the same 400-file
size sample, main XML shrinks from **22,183,474 to 4,770,570 bytes**, against **5,373,347 bytes** in
the originals. Independent expanded-XML comparisons preserve all content in **1,845 changed parts**;
there are no new package issues or inner-`mc:Ignorable` convention signatures. The compact Word
sample passes Word, SDK and LibreOffice. Opaque parts retain their bytes. This is an isolated
serialization measurement on existing saves, with fresh app saves for Office, not a new full corpus run.

The feature inventory compares **2,864 readable pairs**, with **38 failed attempts**. It retains
**776,207 of 776,869 counted words**; **2,811 documents** keep every counted word. Data-bound values,
malformed input, unreferenced stories and unsupported markup remain exceptions.

| Status | Feature | Current measured result / limit |
|---|---|---|
| Kept, with approved Strict conversions | Package metadata, custom XML and templates | 965/965 custom-XML graphs and property dependencies (960 exact, 5 namespace conversions); 353/353 typed custom-property bags, 102/102 template links (one namespace conversion), 17/17 VBA and 1/1 label relationships |
| Kept, with source exceptions | Glossary, embedded fonts and identity metadata | 193/195 glossary relationships (two missing original dependencies), 20/20 referenced fonts, 95/95 people and 15/15 comment-identity relationships; 3 orphan font parts are not reattached |
| Kept, with a remaining exception | Content controls | 2,675/2,676 wrappers; inline, block, row and cell controls retain properties and scalar bindings |
| Kept | OLE, SmartArt, charts and embedded documents | 516/516 OLE/object inventory nodes, 39/39 SmartArt data parts, 133/133 chart parts and 4/4 altChunk blocks; geometry edits retain payloads |
| Kept | Text effects, pictures in page backgrounds and drawing properties while untouched | 205/205 Word 2010 effect entries, 62/62 backgrounds, 1,054/1,054 text boxes and 1,803/1,803 DrawingML shape entries |
| Kept, with range/story exceptions | Comments, notes, bookmarks and permissions | 1,404/1,421 comments, 674/684 footnotes, 57/57 endnotes, 4,870/4,965 bookmarks and 50/54 permission ranges |
| Kept, with remaining losses | Revisions, fields, pictures and tables | 1,385/1,427 insertions/deletions, 1,085/1,103 formatting revisions, 2,590/2,711 fields, 1,239/1,242 pictures and 1,605/1,606 tables |
| Converted when edited | Unsupported contents inside an opaque frame; typed/complex controls | Edited or ungrouped converted content uses the model representation; incompatible typed-control edits unbind and report the conversion |
| Mostly converted, with a notice | Smart tags and inline custom-XML markup | 18/519 original wrappers remain; their text is represented as ordinary content |
| Kept; Office accepted | Per-property drawing edits, row/cell alternatives and watermarks | 62 focused checks and 48 save/edit/history/draft states pass; full corpus completed; 14 original batch-2 files pass Office and the comment repair has a confirmed correction |

Package counts verify bytes, types, rIds and targets; unreferenced parts have separate loss entries.
The prepared-save audit reports missing wrappers, fields, move/revision markup, controls, bookmarks,
permissions and notes/comments. Editing an imported chart records its formatting conversion.
UTF-16 XML parts use the shared decoder. The final full run includes these reporting changes.
The raw carry audit reports 2,795 exact passes and 107 findings: 38 failed driver/package reads,
67 approved Strict conversions and two incomplete glossary graphs. An independent namespace-aware
comparison confirms that the converted graphs retain their other content and targets. All 2,653
web-settings relationships survive (2,586 exact and 67 converted).
The edited inventory compares 2,270 readable pairs and retains 2,505/2,505 control wrappers,
809/809 text boxes and 1,046/1,046 DrawingML shape entries. Its 632 missing/unreadable copies include
the driver's 599 deliberately excluded edits; they are not counted as successful comparisons.
The user accepted the bound-text, formatted-effects, resized-SmartArt, pasted-OLE, macro and namespace
samples in Office batch 1. All 15 batch 2 files open in LibreOffice with unchanged counts. Strict
conversion passes **67 documents / 268 emitted states** against original package/SDK diagnostics and
keeps its notice through draft recovery; its Office review is included in batch 2.
Current reports and explicit exceptions: `~/corpora/results/final-suite-2026-10-09/quire/`. Earlier
feature checks remain in `~/corpora/results/quire-remainder-2026-10-08/`.
The Office checklist is `~/Downloads/lossless-check/office-batch-2/CHECKLIST.md`; file 12 records its
original repair and links the accepted corrected save. Both samples in
`~/Downloads/lossless-check/office-batch-2-comment-fix/` pass Word, package/SDK checks and LibreOffice,
with unchanged page counts.

### Content controls: edit behavior

| Edit | Save behavior |
|---|---|
| Text inside plain/rich controls | Keep the wrapper; update scalar custom-XML XPath targets, including attributes |
| Date/list value command or checkbox click | Keep typed properties and update the bound value |
| Ordinary typing inside a date/list/checkbox control | Convert to unbound text and report it |
| Formatting | Keep control properties; content locks prevent forbidden changes |
| Enter in a whole-paragraph control | Promote boundaries to a block control with the same identity |
| Enter in a partial-paragraph control | Split into unbound controls and report the conversion |
| Join across a multi-paragraph control boundary | Convert the affected wrapper to ordinary content and report it |
| Delete one boundary | Balance the remaining content and report conversion; deletion/content locks roll back the whole forbidden transaction |
| Copy/paste | Remap complete control/bookmark identities; carry XML stores and glossary dependencies; foreign copies receive separate store identities |
| Undo/redo and drafts | Restore boundaries, binding updates and pending conversions |

The context menu offers **Content Control Value…** for dates, lists and checkboxes. Complex XML/Flat
OPC bindings remain intact while untouched; editing them unbinds and reports the conversion. Copying
a built-in document-property binding into another document also unbinds it and reports it.

### Text properties, ranges and backgrounds: edit behavior

| Edit | Save behavior |
|---|---|
| Type or change an unrelated property | Keep the original property XML, compatibility alternatives and dependencies |
| Change font, colour, shadow, spacing or another modeled property | Replace that property; retain unrelated extensions and attributes in schema order |
| Clear all formatting | Remove direct formatting, including retained extensions |
| Copy a bookmark/permission range | Remap both endpoints; incomplete copies omit the markers |
| Delete one range endpoint | Omit the incomplete pair on save and report it |
| Set a page colour | Replace the original background; other edits keep its fill/picture |
| Undo/redo or drafts | Restore property fragments and range metadata |

### Opaque objects and compatibility alternatives: edit behavior

| Edit | Save behavior |
|---|---|
| Edit surrounding text | Keep the original DrawingML/VML, OLE, ActiveX, SmartArt or altChunk fragment and dependencies |
| Move, resize, rotate or flip one object | Update its anchor/transform in both compatibility branches |
| Resize a SmartArt preview group | Keep diagram data with scaled children or a changed outer extent |
| Edit or ungroup converted contents; replace a picture | Write the model representation and report conversion |
| Change picture crop, border, brightness/contrast, grayscale or opacity | Replace that property; retain recolouring, effects, blip extensions and sibling properties |
| Change shape fill, line, shadow or text-box layout | Merge the owning properties in schema order; a shadow edit keeps glow, reflection, soft edges and 3-D siblings |
| Change DrawingML wrapping or anchor reference mode | Merge anchor properties; changing inline/floating status retains the original graphic and identities |
| Change a complex VML fill, VML geometry/text layout, or a shadow in an effect graph | Convert the affected frame and report it |
| Duplicate or paste into another document | Carry preview and payload bytes; remap definitions and references together, including SmartArt's implicit drawing link |
| Delete all or part of a multi-member compatibility wrapper | Never resurrect deleted content; convert surviving members and report it when the wrapper is incomplete |
| Undo and drafts | Restore frame/dependency identities |

Embedded documents and objects without a usable preview display a placeholder. Run, inline, block,
row and cell alternatives are captured before normalization and emitted once. Empty table choices
remain attached to adjacent members, retaining their fallback without adding a visible row or cell.
Moving one of several drawings owned by a single wrapper currently converts that wrapper.

Original text/picture watermarks stay in their own header/footer locations with their XML and
relationships. Editing the watermark replaces those originals; removing it omits them. Undo restores
the originals, and drafts retain the current state. Eight focused cases cover drawing properties,
row/cell alternatives (including empty choices) and watermarks: **62 checks and 48 emitted states**
pass, with no added package/SDK diagnostics. The two watermark inputs retain all six original VML
watermark elements byte for byte before editing. Office review of these changes is still pending.

### Package and file identity: edit behavior

| Edit | Save behavior |
|---|---|
| Change settings, custom properties or font use | Merge owned properties; retain unrelated settings, property types/pids and referenced embedding metadata |
| Save as a macro-free variant | Report removed VBA before download; cancellation keeps the original filename/type and acknowledgements |
| Recover an encrypted or macro-enabled draft | Keep its main-part variant, password and pending loss entries |
| Save an invalid original dependency or orphaned part | Report unavailable/unreferenced content instead of silently claiming it was retained |

## Compatibility testing

The reproducible OOXML run uses the pinned LibreOffice, Apache POI and Open XML SDK inputs from
`tools/corpora.sh`. Quire runs in Chromium through CDP; independent package checks, SDK diagnostics
and LibreOffice compare the saved file with its original. Page counts, Word's saved page-break markers
and reference PDFs help diagnose layout. Missing fonts, tight wrapping around irregular pictures,
vertical text and complex data bindings remain known differences. RC4-encrypted legacy binary files
are outside the supported OOXML encryption path.

Markdown checks remain separate: **24/24 unchanged documents save byte for byte**, edits stay inside
the edited block in **24/24**, and CommonMark/GFM specifications pass **652/652 and 22/22** cases.
Commands and comparison policies are in [testing.md](testing.md); earlier broad-corpus layout and
compatibility results are archived with the run history rather than mixed into current save numbers.

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
| `js/preserve.js` | Package ownership/merges; controls and XML bindings; property/range preservation; opaque frames, compatibility alternatives and clipboard dependencies |
| `js/preserve-drawing.js` | Property-level DrawingML/VML merges, including compatibility choices, picture effects and anchor edits |
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
