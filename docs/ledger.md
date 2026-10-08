# Ledger 2003 — Web Edition

A spreadsheet in the image of Excel 2003, written in plain HTML and vanilla JavaScript, with
no libraries, no frameworks, no build step at run time. It opens and saves Office Open XML workbooks
(`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, password-protected ones included), CSV and text files, and XML Spreadsheet
2003, and it exports web pages and PDF.

The look is Office 2003 on Windows XP (Luna Blue): menus with their access keys, the Standard and
Formatting toolbars, the Name Box and formula bar, sheet tabs with the four scroll arrows, the task
pane, the status bar with AutoCalculate, and the dialogs where you expect them (Format Cells, Page
Setup, Sort, Chart Wizard, Insert Function, Text Import Wizard and some forty more).

## Running it

The app is `public/ledger/`; paths below are relative to it unless they start with `tools/`. Serve
`public/` and open it:

    python3 tools/serve.py

then visit http://127.0.0.1:8760/ledger/. Everything runs in the browser; there is no back end and no network
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

Ledger regenerates worksheets and keeps original package dependencies, workbook settings, style
indices and shared strings. Save and drafts retain all four OOXML variants and encryption. The
Compatibility Checker reports recorded conversions before a user download; reporting is incomplete
for the remaining object features.

The pinned 951-file corpus saves and reopens **935 workbooks**, with **11 failures and 5 exclusions**.
Cell edit/save/undo/redo completes in **933**, with **11 failures and 7 exclusions**; draft recovery
completes in **935**, with **11 failures and 5 exclusions**. The full run uses writer `cf1b313`.
Package/SDK comparison reports **910 OK, 36 failed and 5 excluded** for unedited saves. One equation
wrapper diagnostic found in that run is fixed in `bc3a04e`; its seven save/edit/history/draft states
have no new diagnostics. Malformed or unvalidatable originals remain failures in the reports.

The feature inventory compares **923 readable pairs**. The independent openpyxl cell comparison
completes in **887 workbooks / 991,581 cells**: **881 workbooks** keep every compared value/formula;
6 differ, and 2 differ in formatting. Exceptions include adversarial shared strings, out-of-range
cells, a template and dynamic-array formulas. The other **64 attempts** remain reported failures,
including 8 saved copies openpyxl cannot read. No blanket lossless claim is made for these inputs.

| Status | Feature | Current measured result / limit |
|---|---|---|
| Kept | Custom XML, typed custom properties, external links | 97/97 custom-XML relationships and property dependencies, 114/114 custom-property relationships, 49/49 external links; 97/99 XML parts in the inventory because unreferenced parts are not reattached |
| Kept | VBA, people, labels and web extensions | 26/26 VBA, 4/4 persons, 2/2 label and 3/3 web-extension relationships |
| Kept, with source exception | Connections | 35/36 relationships; one original dependency is missing; connections/query-part inventory 58/59 |
| Kept on an unedited save; edit rules pending | Pivot tables and caches | 156/156 table parts and 131/131 cache definitions; source shifts, refresh flags and deletion rules remain unfinished |
| Kept on an unedited save; edit rules pending | Threaded comments | 6/6 thread records and their persons; thread/mention identity and reference updates still need integration with edits |
| Kept | Notes, tables, merges, validation and names | 373/373 notes, 230/230 tables, 56,263/56,263 merges, 258/259 validations, 2,106/2,106 names |
| Kept, with exceptions | Charts, pictures and shapes | 155/157 charts, 164/165 pictures and 492/493 shapes; unedited raw shapes retain nested compatibility alternatives and relationships |
| Converted | Newer conditional formatting | 733/752 base conditional-format entries and 7/30 data-bar/icon-set extensions; some newer properties fall back to the base representation |
| Dropped from worksheet content | Controls and embedded OLE objects | 0/344 control references and 0/26 OLE objects; carrying 370/371 control-related package parts does not make the controls usable |
| Pending / unmeasured | Slicers, timelines, query-table edit rules and unknown worksheet extensions | The pinned inventory has no slicer/timeline samples; query-table parts alone do not establish preservation of table fields or refresh behavior |

Package figures compare bytes, content types, original rIds and intended targets. The user accepted
both Excel files in Office batch 1 (XLSM and custom XML/external links); the namespace sample also
passed, but tested repeated declarations only, not inner `mc:Ignorable`. Batch 2's namespace follow-up
confirms both the old `calcPr mc:Ignorable` form and the compact save in Excel. Rewritten parts now
hoist fragment declarations to their root; opaque dependencies remain byte-identical.
All 15 suite batch files open in LibreOffice with unchanged counts. Strict conversion passes
all **23 Strict workbooks / 92 emitted states** against original package/SDK diagnostics and keeps
its notice through draft recovery; the Strict Excel sample passed Office batch 2.
Reports, failed attempts and historical measurements: `~/corpora/results/next-plan-head-2026-10-08/`
and `~/corpora/results/strict-transitional-2026-10-08/`.

### Package, styles and strings: edit behavior

| Edit | Save behavior |
|---|---|
| Type or format an ordinary cell | Keep unrelated opaque parts; append changed styles/strings without rebinding old indices |
| Change document properties or workbook settings | Replace owned properties; retain other values and custom-property types/pids |
| Save as a macro-free variant | Report removed VBA before download; cancellation keeps the original filename/type and acknowledgements |
| Edit an imported macro/dialog sheet | Convert that sheet to a worksheet and report it |
| Undo/redo or recover a draft | Restore model properties, dependency identities and pending loss entries |

### Drawings: edit behavior

| Edit | Save behavior |
|---|---|
| Edit cells around an unedited raw shape | Keep its original XML, compatibility alternatives and relationships |
| Move a drawing | Update its worksheet anchor |
| Edit chart contents or a converted shape | Regenerate the modeled content; general per-property preservation remains pending |
| Copy/duplicate or delete | Existing drawing operations apply; complete opaque-object ownership and reference rules are part of the remaining Ledger work |

### Pivots, threads, controls and extensions: current edit limits

| Edit | Current behavior / remaining work |
|---|---|
| No edit to pivot/thread content | Package parts are carried; there is no pivot or threaded-comment editor |
| Edit source cells, insert/delete rows or columns, rename/delete a sheet | Pivot source/location/cache and threaded-comment reference rules are not yet implemented |
| Edit conditional formatting | The model writes its base form; extended bars and unknown entries need property-level merging |
| Save a sheet containing controls, OLE or non-comment VML | Their worksheet/VML content still needs integration; retained package parts alone are insufficient |
| Edit a query table | Modeled table fields are written; query-specific field identity and refresh rules remain pending |

### Pivot preservation rules for the next implementation

These are the agreed behavior rules for the pending pivot work, not additional supported behavior.
Pivot definitions and cached records remain original package data. The model will hold source and
output locations, cache identities and pending changes; Ledger will not calculate a pivot or execute
an external connection. Excel performs any requested refresh when it opens the saved file.

| Source or edit | Required save behavior |
|---|---|
| Several pivots share a cache | Keep one cache definition, records part and workbook `cacheId`; every surviving pivot refers to that identity. A source change marks the shared cache once. Remove it only when no pivot or other retained dependent uses it. |
| Worksheet range source | Resolve the source sheet to its model identity and keep the original range until an edit affects it. Sheet renames update the source name without changing the cache identity. |
| Table source | Retain `worksheetSource@name` and the table identity; follow the table's current range after row/column edits. Do not replace the table name with a fixed range. |
| Defined-name source | Retain the name and its scope, and use the current name formula when it resolves to a range. Keep formula-based names intact; if a source cannot be bounded safely, request refresh after workbook value or structure changes instead of inventing a range. |
| External source | Keep the connection, external source metadata and cached records unchanged. Local edits do not execute the connection or rewrite its source. |
| Consolidation source | Keep every source area and its page-field metadata. Shift or rename each local area by its owning sheet; retain external areas unchanged. |
| Change a value or formula in a local source | Keep the current cache records and set `refreshOnLoad="1"` on the cache definition. Formatting-only edits do not request refresh. The Compatibility Checker explains that Excel must refresh the cached result. |
| Insert/delete source rows or columns | Apply `O.shiftRange` to each affected area and request refresh. Preserve field identities and cached records; if a deleted source column is used by a pivot field, report that field's loss rather than silently remapping it to another column. |
| Insert/delete rows or columns around the output | Shift or shrink the pivot `location` with its displayed cells. Keep offsets and field layout consistent with the surviving range. If the whole output is deleted, remove that pivot and report it. |
| Edit a value or formula inside the output | Allow the cell edit and retain the pivot. Report that Excel can overwrite that edit on refresh; never silently discard the edit while saving. |
| Delete a pivot sheet, source sheet, whole source range, source table or source name | Remove the affected pivots and their slicers, and report the loss. Keep caches still used elsewhere; remove dependencies that no surviving object references. Deleting one local consolidation area invalidates its dependent pivot rather than changing the meaning of the consolidation. |
| Undo/redo a cell or structural edit | Restore source/output locations, refresh flags, cache membership and loss entries together with the cells. `snapSheet`/`restoreSheet` must cover the preserved state, including changes to caches shared with another sheet. |
| Save, reopen or recover a draft | Write the current pivot state and original identities; drafts retain the same refresh requests and pending notices without showing or acknowledging the checker. |

The pinned inventory contains 105 pivot workbooks, including 14 with shared caches, but no
defined-name or consolidation source. `tools/ledger/test/pivot-fixtures.py` supplies those two
structural fixtures from a pinned range-source workbook. Both have zero package/SDK diagnostics;
an independent reader confirms their source kinds, and the consolidation's ten input values match
its cache and output. Office acceptance and Ledger mutation checks remain pending.

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
  their cached cells and original pivot/cache parts. Range changes, source edits and deletion behavior
  are not yet safe; see the edit limits above.
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
