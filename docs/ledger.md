# Ledger 2003 — Web Edition

A spreadsheet in the image of Excel 2003, written in plain HTML and vanilla JavaScript: 44 script files,
no libraries, no frameworks, no build step at run time. It opens and saves Office Open XML workbooks
(`.xlsx`, `.xlsm`, `.xltx`, password-protected ones included), CSV and text files, and XML Spreadsheet
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
| Excel Workbook `.xlsx`, Macro-Enabled `.xlsm`, Template `.xltx` | yes | yes (VBA projects are kept, not run) |
| Password-protected `.xlsx` / `.xlsm` | yes: Office 2007 "standard" and Office 2010+ "agile" encryption (AES with SHA-1/256/384/512) | yes: Tools ▸ Options ▸ Security, or Save As ▸ General Options (AES-256 with SHA-512, as Excel 2013+) |
| Strict Open XML `.xlsx` | yes | as transitional `.xlsx` |
| CSV `.csv`, text `.txt` / `.tsv` / `.prn` | yes: UTF-8, UTF-16 and Windows-1252 detected; the Text Import Wizard for delimited and fixed-width text | yes: CSV (UTF-8) and tab-delimited text, what each cell shows, as Excel does |
| XML Spreadsheet 2003 `.xml` | yes | yes |
| Web Page `.htm` / `.html` | yes: each table on the page becomes a sheet | yes: one self-contained page per workbook |
| PDF | — | yes |
| Excel 97–2003 `.xls` | no — Ledger explains how to convert it | no |

Passwords are handled in the browser (WebCrypto plus Ledger's own AES and SHA code), so the page must
be served over https or from localhost. A workbook opened with its password keeps it when saved.

## What a save keeps, converts and drops

Ledger reads a workbook into its own model and writes a new file from it. Anything the model has no
place for is left out of the saved file, and today there is no warning when that happens.
`tools/ledger/test/loss-audit.py` measures it: for each of 2,941 test workbooks it counts every feature in the
original and in Ledger's saved copy, and reads every cell of both with openpyxl, an independent reader,
comparing value, formula, number format, font, fill, borders and alignment
(`python3 tools/ledger/test/loss-audit.py originals/ saved/ out.json`).

**Kept.** Of 3,092,390 cells in 2,875 workbooks, every value and formula is unchanged in 2,861
workbooks. The other 14 are malformed test files (strings of a million characters, cells beyond column
XFD), dates written as text that Ledger stores as real dates, and spilled dynamic-array cells that
openpyxl reports differently in the two files. Number formats, fonts, fills, borders and alignment agree
except where a file leaves the font unspecified. Merged cells, column widths and row heights, hidden and
grouped rows and columns, frozen panes, defined names, tables, filters, data validation, conditional
formatting, notes, hyperlinks, sparklines, charts and chart sheets, pictures (linked ones keep their
link), shapes and text boxes, page setup, print headers and footers, sheet and workbook protection,
external links, custom document properties, scenarios and the VBA project of an `.xlsm` are all written
back.

**Converted.**

- Pivot tables (118 workbooks) keep their last values as ordinary cells; the pivot table itself — layout,
  fields and cache — is not saved, so Excel sees a plain range that cannot be refreshed.
- Threaded comments (6) survive as the classic notes Excel stores alongside them; the thread structure
  and @mentions do not.
- Excel 2010 data bars and icon sets (12 of 15 workbooks) fall back to their Excel 2007 form: negative-bar
  colours, solid fills, bar borders and the newer icon sets are lost.

**Dropped.**

| What | Workbooks in the test set |
| --- | --- |
| Form controls and ActiveX controls (buttons, check boxes, drop-downs) | 25 |
| Data connections and Power Query queries (the data stays, refreshing it does not) | 31 |
| Custom XML data parts (document-management metadata) | 41 |
| Sensitivity labels in the newer `docMetadata/LabelInfo.xml` form (labels kept as `MSIP_Label` custom properties survive: 35 of 35) | 27 |
| Embedded OLE objects | 5 |
| Slicers and timelines | 1 |

## Compatibility testing

Ledger was tested against workbooks other people wrote — the test suites of sixteen spreadsheet
libraries and applications, not files made for Ledger:

| Source | `.xlsx` / `.xlsm` files |
| --- | --- |
| XlsxWriter (Python) | 979 |
| ClosedXML (.NET) | 382 |
| Apache POI (Java) | 350 |
| PhpSpreadsheet | 287 |
| LibreOffice import/export regression files | 263 |
| rust_xlsxwriter | 228 |
| ExcelDataReader (.NET) | 183 |
| calamine (Rust) | 64 |
| pandas | 60 |
| readxl (R) | 43 |
| ExcelJS | 36 |
| tealeg/xlsx (Go) | 32 |
| xlsx-populate | 28 |
| EPPlus | 25 |
| Microsoft Power BI sample workbooks | 10 |
| excelize (Go) | 9 |
| **Total** | **2,979** |

plus 584 CSV files: 502 data sets published by FiveThirtyEight, the CSV test files of PhpSpreadsheet,
pandas, LibreOffice, ExcelDataReader and ExcelJS, the csv-spectrum edge cases and Our World in Data's
CO₂ table (14 MB).

Every file was opened by the same code that ships here (in Node.js for the bulk runs, in headless
Chromium for the interface tests), saved again, and compared with outside references:

| Check | Result |
| --- | --- |
| Workbooks opened | 2,964 of 2,979. Of the 34 password-protected test files, the 20 whose password is published open with it; 8 use ciphers Excel never writes (DES, 3DES, RC2 or MD5 hashing) and are reported as unsupported; the passwords of 6 are not published. The last file is a deliberately corrupt ZIP |
| Encrypted workbooks Ledger writes | open in LibreOffice with the right password and are refused with a wrong one |
| Saved without error | 2,964 of 2,964 |
| Round trip (open → save → open, then compare every cell's value, formula and style, row heights, column widths, merges, names, validation, conditional formats, comments, hyperlinks, tables, drawings, print settings) | 2,919 of the 2,944 unencrypted workbooks identical, 13,708,945 cells compared; the other 25 differ by design — see below |
| Saved files read by openpyxl | 2,931 of 2,939 (all files under 3 MB); the 8 failures are files openpyxl cannot read in their original form either |
| Saved files opened by LibreOffice 24.2 and compared cell by cell | a sample of 240 workbooks: all open; of 75,277 cells, 214 differ and none because of the file — 199 sit in one workbook using the `TRIMRANGE` functions LibreOffice lacks, 11 depend on `RANDBETWEEN`, 2 use `_xlfn.SINGLE`, 1 is an external link LibreOffice cannot follow, and 1 is a value the source file never calculated |
| Recalculation: every formula recalculated by Ledger and compared with the value Excel stored in the file | 117,649 of 120,944 (97.3 %) across the 446 workbooks that carry stored results. Two files account for 78 % of the differences, and in both the stored values do not follow from the cells (one refers to cells that are empty in the file; the other was not last saved by Excel). Without them, 99.3 %. Most of the rest are files whose stored values were written by the library that made them rather than by Excel |
| CSV | 584 of 584 files parsed exactly as Python's `csv` module parses them (29.4 million cells); writing each back and reading it again changes no General-format value. Values shown in a number format are written as shown, as Excel does — `3.73929E-09` typed into a cell is saved as `3.74E-09` |
| XML Spreadsheet 2003 | 297 sampled workbooks written and read back: 287 identical; the rest differ by design (the format has no chart sheets, external references are written by file name, dates keep milliseconds as Excel's do). 50 of them opened in LibreOffice: 14,325 of 14,335 values agree once LibreOffice's habit of dropping the time from date-times is allowed for |
| Interface | all 213 menu and toolbar commands run without a script error; 44 scripted end-to-end checks pass in Chromium (typing and entry rules, formulas, formats, AutoSum, AutoFill, insert row with undo, sort, AutoFilter, Chart Wizard, save and reopen with the chart, nine real workbooks, CSV import, print preview, validation circles, dragging a page break, opening a web page, and opening, refusing and re-saving a password-protected workbook) |
| Speed (Chromium) | a 15 MB workbook with 2.1 million cells opens in 4.7 s and saves in 5.6 s; 3 million cells open in 9.7 s; a 40 MB CSV with 5.2 million cells imports in 12 s; scrolling stays under 20 ms a screen |

The 25 round-trip differences are deliberate: invalid page-setup values in the source are normalised
(paper size 0 becomes Letter, scale 0 becomes 100 %) in 20 files; text longer than Excel's
32,767-character cell limit is cut to the limit in three (two of them "XML bomb" tests, whose entities
Ledger refuses to expand); cells beyond column XFD are dropped, as Excel drops them; and a frozen pane
at A1, which freezes nothing, is not written.

## Known limitations

- Macros are kept in `.xlsm` files but never run; there is no VBA editor.
- PivotTables: Data ▸ PivotTable builds a static summary report. Pivot tables in opened files keep
  their last values, but only as ordinary cells: the pivot table itself (layout, fields, cache) is not
  saved, so after a save Excel sees a plain range that cannot be refreshed.
- Power Pivot data models and Power Query connections are not loaded and are not saved back.
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
| `../common/core.js`, `js/xml.js`, `../common/zip.js` | Utilities, file saving, PDF writer, a small fast XML parser, ZIP reader/writer |
| `../common/sha.js`, `../common/crypto.js` | SHA-1/256/384/512 and AES, Office document encryption (standard and agile) |
| `js/model.js`, `js/ops.js` | Workbook model (sparse rows, shared styles, names, tables), undo history, every editing operation |
| `js/formula.js`, `js/calc.js` | Formula language (A1, R1C1, structured references, `@`, spills), calculation engine and dependency graph |
| `js/fn-core.js`, `js/fn-lookup.js`, `js/fn-stat.js`, `js/fn-fin.js`, `js/fn-eng.js`, `js/fninfo.js` | The 519 worksheet functions and their catalogue for Insert Function |
| `../common/numfmt.js`, `js/layout.js`, `js/render.js`, `js/cf.js`, `js/styles.js` | Number formats, text measurement, the canvas renderer, conditional formats, list styles |
| `js/grid.js`, `js/editor.js`, `js/clipboard.js` | Selection, scrolling, keyboard and mouse; in-cell editing and point mode; clipboard |
| `js/xlsx-read.js`, `js/xlsx-write.js`, `js/csv.js`, `js/xmlss.js` | File formats |
| `../common/dml.js`, `../common/charts.js`, `js/xchart.js`, `js/drawing.js`, `../common/geometry.js`, `../common/metafile.js` | DrawingML, chart model and SVG drawing, Chart Wizard, pictures and AutoShapes, WMF/EMF |
| `js/filter.js`, `js/print.js`, `../common/spell.js`, `js/spell.js`, `js/panes.js` | AutoFilter, printing and Print Preview, the shared spelling engine with Ledger's Spelling dialog and AutoCorrect, task panes and templates |
| `../common/ui.js`, `../common/icons.js`, `../common/luna.css`, `js/commands.js`, `js/dialogs*.js`, `js/app.js` | Menus, toolbars and dialogs; original 16×16 icons; the command registry; the application controller |
| `../common/dict/` | Spelling word lists (English U.S. and U.K.) and their licences (shared with Quire) |
| `tools/ledger/test/` | The test harnesses described above (Node.js and Playwright) |

Ledger is an independent program. Microsoft, Excel and Office are trademarks of Microsoft Corporation;
Ledger is not affiliated with Microsoft.
