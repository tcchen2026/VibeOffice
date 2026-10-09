# Ledger 2003 — Web Edition

A spreadsheet in the image of Excel 2003, written in plain HTML and vanilla JavaScript, with
no libraries, no frameworks, no build step at run time. It opens and saves Office Open XML workbooks
(`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, password-protected ones included), CSV and text files, and XML Spreadsheet
2003, and it exports web pages and PDF.

The look is Office 2003 on Windows XP (Luna Blue): menus with their access keys, the Standard and
Formatting toolbars, the Name Box and formula bar, sheet tabs with the four scroll arrows, the task
pane, the status bar with AutoCalculate, and the dialogs where you expect them (Format Cells, Page
Setup, Sort, Chart Wizard, Insert Function, Text Import Wizard and some forty more).

## Using it

Open https://vibeoffice.work/ledger/. The app is `public/ledger/`; paths below are relative to it unless
they start with `tools/`. To run it locally, see [testing.md](testing.md). Everything runs in the browser; there is no back end and no network
service behind any feature. Two things are fetched at run time: the spelling word lists in
`../common/dict/`, shared with Quire (which is why the site should be served rather than opened from disk — from a `file://` page
everything works except Tools ▸ Spelling), and, from Google Fonts when it can be reached, the
metric-compatible faces Carlito, Caladea, Arimo, Tinos and Cousine, which stand in for Calibri, Cambria,
Arial, Times New Roman and Courier New so that text fits cells the way it does in Excel. Offline, the
browser's own fonts are used.

`index.html` is generated: edit `tools/ledger/build/template.html`, `tools/ledger/build/*.css` or the
script list in `tools/ledger/build/make.py`, then run `python3 tools/ledger/build/make.py`.

## What is in it

**File.** New (blank, or from eight built-in templates: Loan Amortization, Expense Statement, Sales
Invoice, Timecard, Personal Monthly Budget, Task Tracker, Shopping List, Weekly Schedule), Open, Close, Save, Save As with every format listed under
*File formats*, Save as Web Page, Page Setup, Print Area, Print Preview, Print (to PDF), Send To,
Properties. Several workbooks can be open at once (Window menu).

**Edit.** Multi-level Undo and Redo, including structural changes (inserting or deleting rows,
sorting, moving sheets). Cut, Copy, Paste, Paste Special (values, formulas, formats, comments,
validation, column widths, operations, skip blanks, transpose, paste link), the Office Clipboard (24
items), Fill (Down, Right, Up, Left, Across Worksheets, Series with linear, growth, date and AutoFill
types, Justify), Clear, Delete, Delete Sheet, Move or Copy Sheet, Find and Replace (Find All list,
within sheet or workbook, by rows or columns, formulas, values or comments, match case or entire cell),
Go To and Go To Special, Edit Links. Copying to and from other programs carries tab-separated text and
HTML tables.

**Entering data.** Excel's own input rules: numbers, percentages, currency, fractions (`1 1/2`),
dates and times in the usual short forms, a leading apostrophe for text, formulas with point mode
(click or use the arrow keys to insert references, F4 to cycle `$`), colour-coded reference highlights,
AutoComplete from the column, AutoCorrect, AutoFill by dragging the fill handle (numbers, dates,
weekdays, months, quarters, custom lists), Ctrl+Enter and Alt+Enter.

**View.** Normal and Page Break Preview (drag the blue page breaks), the task pane, the eight
toolbars (Standard, Formatting, Borders, Chart, Drawing, Formula Auditing, Reviewing, List), formula
bar, status bar, Header and Footer, Comments, Full Screen and Zoom (10–400 %, Fit Selection).

**Insert.** Cells, Rows, Columns, Worksheet, Chart (the four-step Chart Wizard), Symbol, Page Break,
Function (Insert Function and Function Arguments with live results and help for every function),
Name (Define, Paste, Create, Apply), Comment, Picture from file, AutoShapes, Text Box, Hyperlink.

**Format.** Format Cells with all six tabs (Number with the full category list and custom formats,
Alignment with indent, rotation, wrap, shrink and merge, Font, Border, Patterns, Protection), Row and
Column height, width, AutoFit, Hide and Unhide, Sheet rename, hide, background and tab colour,
AutoFormat (the sixteen Excel 2003 table formats), Conditional Formatting (three conditions, cell value
or formula), Style (Normal, Comma, Currency, Percent and your own).

**Tools.** Spelling (offline, English U.S. and U.K.), Error Checking, Protection (sheet and workbook,
with passwords), Goal Seek, Scenarios (with summary report), Formula Auditing (trace precedents,
dependents and errors with arrows, Evaluate Formula step by step, Watch Window, Show Formulas),
Calculate Now (F9), AutoCorrect Options, Options (View, Calculation with manual mode and iteration,
Edit, General, Spelling…).

**Data.** Sort (three keys, by rows or columns, case-sensitive, custom orders), AutoFilter with the
Excel 2003 drop-down (values, Top 10, Custom AutoFilter), Advanced Filter (criteria ranges, copy to,
unique records), Form, Subtotals, Validation (lists with in-cell drop-down, whole number, decimal,
date, time, text length, custom formula; input and error messages; Circle Invalid Data), Table (one-
and two-variable data tables), Text to Columns, Consolidate, Group and Outline (automatic outline,
show and hide detail), PivotTable report, Import External Data ▸ Import Text File (the three-step Text
Import Wizard), List (Create List, Total Row, Convert to Range).

**Window and Help.** Split, Freeze Panes, switching between open workbooks; a Help task pane with
searchable topics, the keyboard-shortcut reference and a function reference for all 519 functions.

### Formulas and calculation

- 519 worksheet functions: every function in Excel 2003 including the Analysis ToolPak, and the later
  ones that turn up in today's workbooks — IFS, SWITCH, MAXIFS, TEXTJOIN, CONCAT, XLOOKUP, XMATCH,
  FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, LET, LAMBDA, MAP, REDUCE, BYROW, TEXTSPLIT, TEXTBEFORE,
  VSTACK, HSTACK, TAKE, DROP and more.
- Array formulas (Ctrl+Shift+Enter), dynamic arrays that spill, `A1#` spill references and the `@`
  operator; structured references to lists (`Table1[Amount]`, `[@Price]`, `[#Totals]`); defined names
  at workbook and sheet level, including relative names; 3-D references across sheets.
- A dependency graph recalculates only what changed; volatile functions update on open and on F9;
  circular references are reported, or iterated when iteration is turned on; the 1900 and 1904 date
  systems; A1 or R1C1 reference style.
- Number formats follow Excel to the letter: sections and conditions, colours, `[h]:mm` elapsed time,
  fractions, scientific and engineering notation, `_` and `*` padding, locale and currency tags, and
  the General format that shows as many digits as the column has room for (otherwise `####`).

### Charts, pictures and drawing

- The Chart Wizard makes Column, Bar, Line, Pie, XY (Scatter), Area and Doughnut charts in the Excel
  2003 palette, embedded or on a chart sheet. Charts follow their data as you type.
- Charts in opened files are drawn from their own definitions (3-D charts flat; radar and stock charts
  as lines; bubble charts as scatter) and are saved back unchanged unless you edit them.
- Pictures (PNG, JPEG, GIF, BMP, SVG and Windows metafiles — EMF and WMF are drawn by Ledger's own
  metafile player), AutoShapes and text boxes: move, resize, copy, order, Format Object.
- Conditional formatting from newer files is shown too: data bars, colour scales and icon sets
  (including the Excel 2010 sets and custom icon combinations), and sparklines.

### Printing

Page Setup (orientation, scaling or fit to pages, paper size, margins, centring, headers and footers
with the `&P`, `&N`, `&D`, `&T`, `&F`, `&A` codes and fonts, print titles, print area, gridlines,
headings, page order), automatic and manual page breaks (drag them in Page Break Preview), Print
Preview with zoom and margin guides, and Print, which produces a PDF of the chosen pages because a web page cannot drive the
printer itself.

## File formats

| Format | Open | Save |
| --- | --- | --- |
| Excel Workbook `.xlsx`, Macro-Enabled `.xlsm`, Templates `.xltx` / `.xltm` | yes | yes (VBA projects are kept, not run) |
| Password-protected `.xlsx` / `.xlsm` | yes: Office 2007 "standard" and Office 2010+ "agile" encryption (AES with SHA-1/256/384/512) | yes: Tools ▸ Options ▸ Security, or Save As ▸ General Options (AES-256 with SHA-512, as Excel 2013+) |
| Strict Open XML `.xlsx` | yes | standard Transitional OOXML, with a Compatibility Checker notice |
| CSV `.csv`, text `.txt` / `.tsv` / `.prn` | yes: UTF-8, UTF-16 and Windows-1252 detected; the Text Import Wizard for delimited and fixed-width text | yes: CSV (UTF-8) and tab-delimited text, what each cell shows, as Excel does |
| XML Spreadsheet 2003 `.xml` | yes | yes |
| Web Page `.htm` / `.html` | yes: each table on the page becomes a sheet | yes: one self-contained page per workbook |
| PDF | — | yes |
| Excel 97–2003 `.xls` | no — Ledger explains how to convert it | no |

Passwords are handled in the browser (WebCrypto plus Ledger's own AES and SHA code), so the page must
be served over https or from localhost. A workbook opened with its password keeps it when saved.

## What a save keeps, converts and drops

The full LibreOffice pass on the versioned `984eaf3` artifacts renders **930/935 unedited saves**;
**three** readable pairs change page count, down from 37 after the page-setup and layout corrections,
with no new differences. Across **6,589** attempted states, **6,326** pass the paired render/count
check, **246** fail and **17** are excluded. Failures include
unrenderable originals, missing outputs, conversion errors/timeouts and count differences. This does
not establish visual fidelity for every workbook. All five saved-render failures also fail or time
out in their originals. Reports: `~/corpora/results/lossless-final-984eaf3/ledger/`. The layout
correction restores the original counts for `tdf66668` (two pages) and `pivottable_tabular_mode`
(four). Signed drawing offsets are retained, and empty cells retain their original style identity
even when it maps to the model's default style. These cells follow row shifts and undo; clearing
formats or all content removes the retained style. All 74 Ledger tests, 19 command package/SDK
comparisons and six browser draft/recovery comparisons pass. The `984eaf3` full refresh confirms both
count corrections. Office batch 7 file 07 passed. File 08 and its untouched Excel source both offer
repair; removing the source recovery flag does not resolve it. An Office-authored replacement with
the same blank-cell style behavior (`tableStyle.xlsx`, B7/C7) passes Excel, package/SDK and LibreOffice.
The source failure remains a documented exception. Focused evidence is in
`~/corpora/results/ledger-layout-final-2026-10-09/`.
The isolation and accepted replacement are recorded in `~/corpora/results/office-batch-7-repair-2026-10-09/`.
The remaining three count differences are `tdf118668` (invalid source sheet visibility),
`Photo Formats - CGM-O12-XL-Pictures` (unresolved reflow despite retained picture formatting), and
`clusterfuzz-testcase-minimized-POIXSSFFuzzer-5937385319563264` (a damaged source ZIP member).

Worksheet-format preservation retains authored automatic row heights and unread attributes;
changing a default width or height replaces only that property. Cell edits, sheet copies, history
and drafts keep the retained settings. Literal `true`/`false` view flags retain reading direction.
Pre-release worksheet-size markup converts to current settings with a notice. The 2005/8 vocabulary
uses twip row heights and zero-based columns with widths in 1/256-character units; the reader now
converts these and keeps the original automatic default height. The 2006/2 vocabulary retains its
current-format units. The conversion also replaces obsolete workbook/style/theme namespaces,
translates shared-string entries and moves the legacy colour palette into the theme. It restores ten
previously empty text cells in the 2006 preview sample. A 2005 theme is rebuilt from readable colours
and fonts, with a notice for unsupported details. The corpus scan identifies two affected inputs;
23 unreadable archives remain recorded separately. All 78 Ledger tests and six save/edit/draft
attempts pass. Their 14 saved states pass the SDK independently with zero errors and add no package
issues; comparisons against the unreadable originals remain failed. Both saves open in LibreOffice:
the readable original keeps its one page, while the unreadable original has no comparable count.
The pre-release conversion passed Excel in Office batch 7 files 03–04. Evidence:
`~/corpora/results/ledger-preview-schema-2026-10-09/`. The automatic-height sample passed Excel in
Office batch 5.

Ordinary pictures now retain image effects, recolouring, geometry, fills, borders, styles and
extensions. Cell edits leave these properties intact; moving/resizing changes geometry, crop/border
and alternative-text edits replace only their property, and copied pictures remap identities.
Deleting a picture removes its retained properties. Replacing its image removes image-specific
extensions/effects with a notice. Two real-file regressions and 14 package/SDK states pass through
edits, copies, undo/redo and recovery. The picture-format workbook opens in LibreOffice but retains
its separate 19→17-page layout difference. Invalid source sheet visibility, which accounts for another
remaining count difference, now has an explicit conversion notice. These corrections are included in
the `984eaf3` full refresh; Office batch 7 file 06 passed Excel. Focused evidence is in
`~/corpora/results/lossless-evening-review-2026-10-09/`.

Page setup now retains its original absence, per-property values and printer-setting relationships.
Cell edits no longer introduce a portrait orientation into a sheet that omitted page setup. Seven
public save/edit/draft inputs produce 49 passing package/SDK states, with all seven unedited render
pairs retaining their original page counts. Sixteen distinct command states cover paper/quality
edits, chart sheets, history and copies.
The `984eaf3` refresh completes all 2,853 driver attempts and 6,589 package/SDK states with
unchanged statuses and zero added diagnostics; independent cell comparisons are also unchanged.
Its full LibreOffice pass resolves 34 earlier count differences. Office batch 6 files 01–04 cover absence,
paper/quality edits, copied printer relationships and scoped namespace directives; all four pass
focused package/SDK and LibreOffice checks and are accepted in Excel. Evidence:
`~/corpora/results/ledger-page-setup-2026-10-09/`, `ledger-final-3cb0cd2/` and
`office-final-corrections-2026-10-09/`.

Ledger regenerates worksheets and keeps original package dependencies, workbook settings, style
indices and shared strings. Save and drafts retain all four OOXML variants and encryption. The
Compatibility Checker reports conversions from the prepared output before a user download,
including edited imported charts and opaque drawings. Drafts retain notices silently.

The pinned 951-file corpus saves and reopens **935 workbooks**, with **11 failures and 5 exclusions**.
Cell edit/save/undo/redo completes in **933**, with **11 failures and 7 exclusions**; draft recovery
completes in **935**, with **11 failures and 5 exclusions**. The current full driver, validation and
cell run uses writer `984eaf3`. The two-input `7f85450` preview-schema correction has its separately
hashed 14-state validation/render evidence above; unchanged corpus inputs are not rerun.
Package/SDK comparison reports **911 OK, 35 failed and
5 excluded** for unedited saves. Across **6,589** save/edit/history/draft states, **6,371 pass,
201 fail and 17 are excluded**; none adds diagnostics. Malformed or unvalidatable originals
remain failures in the reports. All validation statuses match the preceding run. The shared review
suite and audit regression pass; all 78 Ledger unit tests pass.

The feature inventory compares **922 readable pairs**. The independent openpyxl cell comparison
covers **889 workbooks / 1,649,147 cells**: **883 workbooks** keep every compared value/formula;
6 differ, and 2 differ in formatting. Exceptions include adversarial shared strings, out-of-range
cells, a template and dynamic-array formulas. The other **62 attempts** remain reported failures,
including six saved copies openpyxl cannot read. The edited comparison excludes only the scripted
A1 target: **879 of 887 workbooks** match every other value/formula; six have existing differences
and two recalculate dependent array-result cells. No blanket lossless claim is made for these inputs.
The large workbook that exceeded the ordinary 40-second cell-comparison limit now completes both
scenarios with a 150-second limit. All cell coverage above is from current output; the initial
timeouts remain in the report. The initial run without openpyxl is recorded as a harness failure;
the corrected audit uses the installed environment. The feature audit rejects an invalid UTF-8
part in `tdf76115` explicitly.

| Status | Feature | Current measured result / limit |
|---|---|---|
| Kept | Custom XML, typed custom properties, external links | 97/97 custom-XML relationships and property dependencies, 114/114 custom-property relationships, 49/49 external links (one converted from Strict); 97/99 XML parts in the inventory because unreferenced parts are not reattached |
| Kept | VBA, people, labels and web extensions | 26/26 VBA, 4/4 persons, 2/2 label and 3/3 web-extension relationships |
| Kept, with source exception | Connections | 35/36 relationships, including one Strict conversion; one original dependency is missing. All 58/58 actual connection/query parts remain; ZIP directory markers are excluded from counts |
| Kept; Excel samples accepted | Pivot tables and caches | 156/156 tables and 131/131 caches. Twelve focused cases / 34 saved states cover sources, shifts, refresh, copy and deletion; dependent slicers follow deleted sources and shared pivots |
| Kept; Excel samples accepted | Threaded comments | 6/6 corpus records across four workbooks; five thread parts and four person parts remain byte-identical. Focused edits retain replies, mentions, person IDs and shifted cell references |
| Kept | Notes, tables, merges, validation and names | 373/373 notes, 230/230 tables, 56,263/56,263 merges, 259/259 validations, 2,106/2,106 names |
| Kept; Excel samples accepted | Page setup, absent defaults and printer settings | Seven public fixtures / 49 save/edit/history/draft states add no package/SDK diagnostics; all seven unedited render pairs retain their page counts. Paper and print-quality edits retain sibling options and binary printer settings |
| Kept, with exceptions | Charts, pictures and shapes | 155/157 charts, 164/165 pictures and 492/493 shapes; unedited raw shapes retain nested compatibility alternatives and relationships |
| Kept; Excel samples accepted | Worksheet extensions and newer conditional formatting | The 65-workbook extension sweep retains 100/100 extension entries with unchanged expanded XML: 23/23 extended bars, 53/53 extended rules, 20/20 extended validations, 9/9 sparkline groups and 63/63 unknown entries. Package checks add no issues; SDK comparisons pass in 64, with one malformed original unvalidatable |
| Kept; Excel samples accepted | Controls, OLE, non-comment VML and sheet custom properties | All 40 relevant workbooks save without new package/SDK errors: 344/344 control references, 370/370 control/ActiveX parts, 26/26 OLE entries, 367/367 non-note VML shapes and 31/31 sheet properties. All 43 VML parts retain their original bytes |
| Kept; Excel samples accepted | Slicers and timelines | Three pinned supplemental workbooks retain 7/7 views and 12/12 definition/cache parts byte for byte; edits cover shared pivots, table copies, source deletion and drafts. OLAP cases remain unmeasured |
| Kept; Excel samples accepted | Query tables and column metadata | Ten public workbooks retain 23/23 query parts and 10/10 connections parts byte for byte, with 20/20 query tables and 68/68 column identities. Eight focused cases cover column edits, copies, deletion and history |
| Converted, with a notice | Edited imported charts, opaque drawing contents and macro/dialog sheets | Regenerate the edited model content; report its conversion before download |
| Dropped, with a notice | Missing or unreferenced package dependencies; VBA in macro-free Save As | One source connection dependency is missing; two unreferenced custom-XML parts are not reattached. Macro removal requires the selected format and successful save |

Package figures compare bytes, content types, original rIds and intended targets. The user accepted
both Excel files in Office batch 1 (XLSM and custom XML/external links); the namespace sample also
passed, but tested repeated declarations only, not inner `mc:Ignorable`. Batch 2's namespace follow-up
confirms both the old `calcPr mc:Ignorable` form and the compact save in Excel. Rewritten parts now
hoist fragment declarations to their root; opaque dependencies remain byte-identical.
All 15 suite batch files open in LibreOffice with unchanged counts. Strict conversion passes
all **23 Strict workbooks / 92 emitted states** against original package/SDK diagnostics and keeps
its notice through draft recovery; the Strict Excel sample passed Office batch 2.
Reports, failed attempts and historical measurements: `~/corpora/results/next-plan-head-2026-10-08/`
and `~/corpora/results/strict-transitional-2026-10-08/`. Current full results:
`~/corpora/results/lossless-final-984eaf3/ledger/`; feature edit checks are in
`~/corpora/results/ledger-group-2026-10-09/` and `~/corpora/results/review-fixes-2026-10-09/`.
The convention scan has 192 signatures: three additions to the preceding writer retain a Strict
picture's round-rectangle adjustment and duotone colour/order in Transitional vocabulary. The
subtrees otherwise match the source and passed Excel in batch 7. Scoped
`mc:PreserveAttributes`/`mc:Ignorable` on worksheet-format and page-setup properties retain their
original processing scope; the page-setup form passed Excel in batch 6. The rebuilt 15-file **Office batch 3**
uses `56de065` and passes package/SDK and LibreOffice with expected counts, with explained changes
against the previous handoff; **all 15 samples passed Excel**. File 04 also checks the root-comment
edit correction measured in `~/corpora/results/thread-root-edit-2026-10-09/`; its updated sample keeps
the reply and mention after editing, shifting and copying the thread. A fresh LibreOffice check
of the exact accepted batch-3 and batch-5 files also passes with expected page counts; hashes and
results are in `~/corpora/results/office-accepted-2026-10-09/`.

### Package, styles and strings: edit behavior

| Edit | Save behavior |
|---|---|
| Type or format an ordinary cell | Keep unrelated opaque parts; append changed styles/strings without rebinding old indices |
| Change document properties or workbook settings | Replace owned properties; retain other values and custom-property types/pids |
| Change default row height or column width | Replace only the changed worksheet-format attributes; keep authored automatic heights, unread attributes and alternatives confined to that property |
| Change paper size, orientation or print quality | Replace only the owning page-setup attributes; keep other settings and printer relationships. Untouched absent page setup stays absent; sheet copies and history retain it |
| Copy a sheet, edit cells or undo dimensions | Retain the worksheet-format fragment with the sheet; drafts recover its pending settings. A mixed alternative spanning other sheet properties converts with a notice to avoid restoring deleted content |
| Save as a macro-free variant | Report removed VBA before download; cancellation keeps the original filename/type and acknowledgements |
| Edit an imported macro/dialog sheet | Convert that sheet to a worksheet and report it |
| Undo/redo or recover a draft | Restore model properties, dependency identities and pending loss entries |

### Drawings: edit behavior

| Edit | Save behavior |
|---|---|
| Edit cells around an unedited raw shape | Keep its original XML, compatibility alternatives and relationships |
| Move a drawing | Update its worksheet anchor |
| Edit chart contents or a converted shape | Regenerate the modeled content with a Compatibility Checker notice; general per-property preservation remains pending |
| Copy/duplicate or delete | Controls and OLE follow the ownership rules below; other opaque-frame ownership remains part of the remaining Ledger work |

### Pivots, threads, controls and extensions: current edit limits

| Edit | Current behavior / remaining work |
|---|---|
| No edit to pivot/thread content | Package parts are carried; there is no pivot or threaded-comment editor |
| Edit pivot source cells, insert/delete rows or columns, rename/delete a sheet | Pivot sources, output locations and shared caches follow the rules below; slicer/timeline connections are pruned when their source disappears |
| Move or edit a threaded comment | References follow the note; the comment editor updates the root text and retains replies, people and thread identities |
| Edit conditional formatting | Keep extended rule identities and unchanged properties; update base/extended bar links together, and retain unknown worksheet extensions by URI |
| Save a sheet containing controls, OLE or non-comment VML | Keep their worksheet entries, previews, VML and dependencies together; they remain uneditable. Each sheet reserves its own `o:idmap` blocks, extending only when fresh identities need space. Notes, controls and nested group members use those blocks; unchanged VML retains its bytes |
| Edit a query table | Retain external field identities; added columns are unbound and never reuse a column ID from the file, deleted fields stay excluded from refresh, and copied tables get independent query parts |

### Pivots: edit behavior

Pivot definitions and cached records remain original package data. The model holds source and
output locations, cache identities and pending changes; Ledger does not calculate a pivot or execute
an external connection. Excel performs any requested refresh when it opens the saved file.

| Source or edit | Required save behavior |
|---|---|
| Several pivots share a cache | Keep one cache definition, records part and workbook `cacheId`; every surviving pivot refers to that identity. A source change marks the shared cache once. Deleting one output retains the cache for the others and updates shared slicer connections. |
| Worksheet range source | Resolve the source sheet to its model identity and keep the original range until an edit affects it. Sheet renames update the source name without changing the cache identity. |
| Table source | Retain `worksheetSource@name` and the table identity; follow the table's current range after row/column edits. Do not replace the table name with a fixed range. |
| Defined-name source | Retain the name and its scope, and use the current name formula when it resolves to a range. Keep formula-based names intact; if a source cannot be bounded safely, request refresh after workbook value or structure changes instead of inventing a range. |
| External source | Keep the connection, external source metadata and cached records unchanged. Local edits do not execute the connection or rewrite its source. |
| Consolidation source | Keep every source area and its page-field metadata. Shift or rename each local area by its owning sheet; retain external areas unchanged. |
| Change a value or formula in a local source | Keep the current cache records and set `refreshOnLoad="1"` on the cache definition. Formatting-only edits do not request refresh. The Compatibility Checker explains that Excel must refresh the cached result. |
| Insert/delete source rows or columns | Shift each affected area and request refresh; a whole-column (`A:D`) or whole-row source stays whole along that axis. Keep field identities and cached records. Deleting a used field removes its affected pivot definition with a notice, retaining the remaining result cells. |
| Insert/delete rows or columns around the output | Shift the pivot `location` with its displayed cells. Lines inserted or deleted inside the output shrink or grow `location`, request refresh so Excel rebuilds the layout, and are reported. If the whole output is deleted, remove that pivot and report it. |
| Edit a value or formula inside the output | Allow the cell edit and retain the pivot. Report that Excel can overwrite that edit on refresh; never silently discard the edit while saving. |
| Delete a pivot sheet, source sheet, whole source range, source table or source name | Remove the affected pivots and report the loss; retain shared caches for surviving pivots. Deleting one local consolidation area invalidates its dependent pivot. Slicers/timelines remove obsolete connections and disappear if their source is gone. |
| Copy the output sheet | Give the copied pivot a fresh part, name and extension identity; share its source cache. Undo restores the original membership. |
| Undo/redo a cell or structural edit | Restore source/output locations, refresh flags, cache membership and loss entries together with the cells. `snapSheet`/`restoreSheet` must cover the preserved state, including changes to caches shared with another sheet. |
| Save, reopen or recover a draft | Write the current pivot state and original identities; drafts retain the same refresh requests and pending notices without showing or acknowledging the checker. |

The pinned inventory contains 105 pivot workbooks, including 14 with shared caches, but no
defined-name or consolidation source. `tools/ledger/test/pivot-fixtures.py` supplies those two
structural fixtures from a pinned range-source workbook. Both have zero package/SDK diagnostics;
an independent reader confirms their source kinds, and the consolidation's ten input values match
its cache and output. Twelve focused pivot cases pass, covering all five source kinds, shared caches,
format-only edits, whole-row/column sources, shifts, copy/deletion and undo/redo. All **34 emitted states** introduce no
package/SDK diagnostics; unchanged definitions and records retain their bytes. The real browser
draft/recovery hooks keep a source-edit refresh request and its notice. Two samples open in
LibreOffice with the same two-page counts. These are focused measurements, separate from the frozen
full-corpus figures above; The batch-3 Excel samples are accepted. Reports:
`~/corpora/results/review-fixes-2026-10-09/current-pivots/`.

### Threaded comments: edit behavior

| Edit | Save behavior |
|---|---|
| Keep the comment text | Carry every reply, mention, person record and unknown child; threads without a legacy note are displayed too |
| Insert/delete rows or columns, sort or move cells | Every reply follows its owning note's cell; legacy note references and thread links stay consistent |
| Delete a note or its row | Remove the whole thread; undo restores it |
| Copy a note or sheet | Give copied comments, parent links, mentions and legacy note identities fresh IDs together through the shared copy helper; keep shared persons, remapping conflicting person IDs on paste |
| Paste into another workbook | Include referenced authors and mentioned persons with their original metadata |
| Edit the displayed comment text | Edit the root text, retaining replies, original thread/person IDs and unknown metadata. Untouched mentions stay attached and shift with their text; changed or ambiguous mentions are removed with a notice. Undo restores the original thread and mention positions |
| Save or recover a draft | Keep current references, replies and mentions; draft recovery retains the pending notices |

Six focused cases cover the four pinned thread workbooks and an authored mention/no-legacy-note
fixture. All **22 emitted states** pass package/SDK comparison. Actual browser root editing, Paste Special,
undo/redo and draft recovery keep two threads, four comments and their mention/parent links without
console errors. Two LibreOffice samples retain their one-page counts. The batch-3 Excel samples are accepted. The authored fixture
and failed pre-fix SDK attempts are retained with the reports; mention coverage is not claimed from
the public corpus. Two current browser draft/recovery states also pass package/SDK comparison;
the root-edit Office sample opens in LibreOffice with the same one-page count.
Reports: `~/corpora/results/review-fixes-2026-10-09/guid-threads/` and
`~/corpora/results/thread-root-edit-2026-10-09/`.

### Controls, embedded objects and VML: edit behavior

| Edit | Save behavior |
|---|---|
| Leave the object untouched | Keep its worksheet XML, both compatibility branches, preview and dependent parts; VML-only objects remain retained without an editing UI |
| Move/resize or insert/delete rows and columns | Update linked worksheet, drawing and VML anchors; fixed-size objects move without shrinking |
| Copy an object or sheet | Remap worksheet/drawing/VML identities together, including nested controls and sheets containing earlier copies; copy mutable payload parts independently. Copied ActiveX names are valid VBA identifiers and match their VML shape IDs; form-control captions retain their spacing |
| Paste into another workbook | Transfer preview and payload dependencies; IDs remain valid in the destination |
| Delete the displayed object | Remove all its representations and unreachable dependencies; undo restores them |
| Edit converted drawing content | Convert only that object to its displayed drawing, with a compatibility notice |
| Add, edit or remove a note | Merge note VML while keeping non-comment shapes and their dependencies |
| Undo/redo or recover a draft | Retain object ownership, payloads, current anchors and pending notices |

Sixteen focused tests produce **63 passing package/SDK states**, including ActiveX copies with and
without a numeric suffix, copies of copies, copied sheets and recovery. Both selected copied-control
saves open in LibreOffice with unchanged counts; Office batch 7 file 05 passes Excel.
The earlier sweep also has **two passing browser draft states**
after cross-workbook paste and undo/redo. Two edited samples open in LibreOffice with unchanged
one-page counts. Core checks pass and all four pages load without console errors (existing
font requests remain). The 40-workbook feature sweep keeps **396 dependent parts byte-identical**;
13 ActiveX XML parts in a Strict input undergo the approved Transitional conversion. One unrelated
unsupported drawing in `stress020.xlsx` remains reported. The batch-3 Excel samples are accepted.
Current focused reports: `~/corpora/results/review-fixes-2026-10-09/current-objects/`;
the broader sweep and browser reports remain in `~/corpora/results/ledger-preservation-2026-10-08/`.

### Slicers and timelines: edit behavior

| Edit | Save behavior |
|---|---|
| Leave a view untouched | Keep its drawing frame, compatibility fallback, view definition, source cache and sheet/workbook extension links |
| Move/resize or insert/delete rows and columns | Update its anchor while retaining the interactive filter payload; the app displays the original fallback |
| Copy a view in the workbook | Give the drawing and definition matching fresh names/identities; retain its source cache |
| Copy a table sheet | Create an independent table and slicer cache, named as Excel names copies (`Slicer_Region2`, view "Region 1") with its own workbook name; deleting the original sheet keeps the copy |
| Copy or delete one of several connected pivots | Update the cache's pivot connections; remaining views and pivots stay connected |
| Delete the source or last view | Remove the affected view, cache, its workbook name and extension references, with a loss entry |
| Edit the fallback drawing | Convert only that view to the displayed shape, with a notice; undo restores it |
| Paste a view into a different workbook | Keep the fallback and report conversion because its source table/pivot was not copied |
| Undo/redo or recover a draft | Restore definitions, source connections, copied identities and pending notices |

Eight focused cases produce **21 passing package/SDK states**, with **two more passing browser draft
states** after real copy/paste and undo/redo. The table-sheet copy and moved timeline open in
LibreOffice (4 and 3 pages respectively); timeline undo/redo also opens. Copies of copied views keep
independent GUIDs, stable across repeated saves. The three supplemental inputs are pinned by hash and
commit through `tools/corpora.sh DIR ledger-features`, separate from the 951-file main corpus.
Current focused reports are in `~/corpora/results/review-fixes-2026-10-09/guid-slicers/` and
`guid-browser/`. The batch-3 Excel samples are accepted.

### Worksheet extensions: edit behavior

| Edit | Save behavior |
|---|---|
| Edit unrelated cells | Keep extended bars, custom icons, validation metadata, sparkline properties and unknown extension entries |
| Change a bar property | Replace that property in the extended rule and its base representation; keep unchanged negative-border colors, automatic thresholds and other details |
| Insert/delete rows or columns | Shift the rule ranges and formula thresholds; external-sheet references participate in undo |
| Copy a sheet or split a rule range | Remap duplicated extended rule IDs and validation/sparkline UIDs; base/extended bar links remain paired |
| Delete a rule or sparkline group | Remove its owned XML without reintroducing it from the retained extension container |
| Edit a validation or ordinary conditional rule | Retain its extension metadata through the existing dialog; replace the edited properties |
| Undo/redo or recover a draft | Restore the rules, extension fragments and matching identities |

Seven focused cases produce **23 passing package/SDK states**, plus **two browser draft/recovery
states** after editing, copying and undo/redo. Edited bars and moved sparklines open in LibreOffice
with their original 3/5 page counts; history states also open. Copied rule GUIDs are assigned at copy
time and stay stable across repeated saves. The extension inventory records all 951 attempts, including 25 ZIP/XML failures.
`empty_ext_defined_name.xlsx` remains an SDK failure because its original external-link content
type is invalid; it is not counted as a passing comparison. Current focused reports are in
`~/corpora/results/review-fixes-2026-10-09/guid-extensions/`; the broader inventory and browser reports
remain in `~/corpora/results/ledger-preservation-2026-10-08/extension-*`. The batch-3 Excel samples are accepted.

Namespace declarations still move to the part root. A local `mc:Ignorable` list remains where
`PreserveAttributes`, `PreserveElements` or `ProcessContent` requires it on the same element; the
14 affected Mac-authored workbooks in this sweep now pass SDK comparison. These scoped lists do
not reintroduce the repeated namespace declarations removed by the namespace follow-up.

### Query tables: edit behavior

| Edit | Save behavior |
|---|---|
| Edit unrelated cells or save without edits | Keep query and connection parts byte for byte, including refresh settings; retain table/column attributes and original field IDs |
| Edit a result cell or column heading | Keep the external field identity; write the local value/heading without changing the external query or requesting refresh |
| Insert/delete columns through a table | Keep retained column IDs, mark new columns as unbound, and record deleted bound fields so a later refresh does not restore them |
| Insert/delete rows or columns | Shift table, filter and sort ranges together |
| Copy a table's worksheet | Give the table independent query parts and fresh UIDs; use numeric query-name suffixes. Retain field links and duplicate sheet-scoped destination names for range queries |
| Delete a table or its worksheet | Remove its query parts with a Compatibility Checker entry; keep workbook-owned connections |
| Undo/redo or recover a draft | Restore table membership, columns, scoped names and query changes; saving does not mutate model headings |

Eight focused cases produce **23 unique passing package/SDK states**, plus **two actual browser
draft/recovery states**. Edited and copied samples and three history states open in LibreOffice
with unchanged page counts. The ten-file query sweep has no new package issues; nine SDK
comparisons pass and one original (`queryTableExport.xlsx`) has a missing dependency and cannot
be validated. Ledger does not execute a query or enable refresh. Current focused reports:
`~/corpora/results/review-fixes-2026-10-09/guid-queries/`; sweep and browser reports:
`~/corpora/results/ledger-preservation-2026-10-08/query-*`.
The full Ledger group measurement is reported above; the batch-3 Excel samples are accepted.

## Compatibility testing

`tools/corpora.sh` pins the LibreOffice, Apache POI and Open XML SDK inputs used above. The Node and
browser drivers exercise the shipping readers/writers, real undo operations and suite draft hooks.
Independent ZIP/XML checks, Open XML SDK diagnostics and openpyxl comparisons are separate from the
app's own reopen check. Commands and comparison policies are in [testing.md](testing.md).

Formula recalculation, CSV, XML Spreadsheet 2003 and UI harnesses remain in `tools/ledger/test/`.
Their earlier broad-corpus measurements are archived with the run history, rather than mixed into
the current OOXML preservation figures. Current preservation does not imply exact print layout or
support for every formula or external data source.

## Known limitations

- Macros are kept in `.xlsm` files but never run; there is no VBA editor.
- PivotTables: Data ▸ PivotTable builds a static summary report. Pivot tables in opened files keep
  their cached cells and original pivot/cache parts. Source/output edits follow the rules above;
  slicer dependents are retained and the batch-3 Excel samples are accepted.
- Power Pivot models and Power Query connections are carried as package dependencies; Ledger does not
  execute them or maintain their query-specific edit semantics.
- Chart editing covers the Chart Wizard types; other chart types from files are shown and kept but
  can only be replaced, not edited in place. Surface charts are not drawn.
- One window per workbook (Window ▸ New Window is not available); the binary `.xls` format is not read.
- Spelling is English only (U.S. and U.K.).
- Fonts are drawn with what the browser has. Calibri, Cambria, Arial, Times New Roman and Courier New
  are drawn with their metric-compatible stand-ins (installed, or loaded from Google Fonts), so column
  fits and `####` decisions match Excel's; with other substitutes a number that just fits in Excel may
  show as `####`, and a few fonts have no stand-in at all.

## Source layout

| File | What it does |
| --- | --- |
| `index.html`, `tools/ledger/build/` | Window chrome, Ledger's CSS after the shared `../common/luna.css`, script order (generated by the build script) |
| `js/pivots.js` | Preserved pivot/cache identities, source and output edits, shared-cache history and refresh notices |
| `js/threads.js` | Complete threads beside notes, root-text edits, mentions/persons and copied identities |
| `js/objects.js` | Controls/OLE ownership, linked DrawingML/VML identities, payload copies and note/VML merging |
| `js/slicers.js` | Slicer/timeline frames, source-cache ownership, sheet/workbook extension links and copied view identities |
| `js/extensions.js` | Extended rule/validation/sparkline ownership, per-property merges, copied identities and unknown worksheet extension entries |
| `js/tables.js` | Table/column fragments and identities, query fields and copies, deleted-field retention and structural edits |
| `../common/core.js`, `../common/zip.js` | Utilities, file saving, PDF writer, ZIP reader/writer |
| `../common/xml.js`, `../common/opc.js`, `../common/opc-order.js`, `js/xml.js` | Shared XML parser and OOXML preservation core; `js/xml.js` aliases the shared tree as `L.xml`. Original package ownership, settings merges and style/string identities are implemented. |
| `../common/sha.js`, `../common/crypto.js` | SHA-1/256/384/512 and AES, Office document encryption (standard and agile) |
| `js/model.js`, `js/ops.js` | Workbook model (sparse rows, shared styles, names, tables), undo history, every editing operation |
| `js/formula.js`, `js/calc.js` | Formula language (A1, R1C1, structured references, `@`, spills), calculation engine and dependency graph |
| `js/fn-core.js`, `js/fn-lookup.js`, `js/fn-stat.js`, `js/fn-fin.js`, `js/fn-eng.js`, `js/fninfo.js` | The 519 worksheet functions and their catalogue for Insert Function |
| `../common/numfmt.js`, `js/layout.js`, `js/render.js`, `js/cf.js`, `js/styles.js` | Number formats, text measurement, the canvas renderer, conditional formats, list styles |
| `js/grid.js`, `js/editor.js`, `js/clipboard.js` | Selection, scrolling, keyboard and mouse; in-cell editing and point mode; clipboard |
| `js/preserve.js` | Package ownership, workbook/property merges and retained style indices |
| `js/xlsx-read.js`, `js/xlsx-write.js`, `js/csv.js`, `js/xmlss.js` | File formats |
| `../common/dml.js`, `../common/charts.js`, `js/xchart.js`, `js/drawing.js`, `../common/geometry.js`, `../common/metafile.js` | DrawingML, chart model and SVG drawing, Chart Wizard, pictures and AutoShapes, WMF/EMF |
| `js/filter.js`, `js/print.js`, `../common/spell.js`, `js/spell.js`, `js/panes.js` | AutoFilter, printing and Print Preview, the shared spelling engine with Ledger's Spelling dialog and AutoCorrect, task panes and templates |
| `../common/ui.js`, `../common/icons.js`, `../common/luna.css`, `js/commands.js`, `js/dialogs*.js`, `js/app.js` | Menus, toolbars and dialogs; original 16×16 icons; the command registry; the application controller |
| `../common/dict/` | Spelling word lists (English U.S. and U.K.) and their licences (shared with Quire) |
| `tools/ledger/test/` | The test harnesses described above (Node.js and Playwright) |

Ledger is an independent program. Microsoft, Excel and Office are trademarks of Microsoft Corporation;
Ledger is not affiliated with Microsoft.
