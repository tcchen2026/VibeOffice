# Lectern 2003 — Web Edition

A presentation editor modeled on PowerPoint 2003 (Luna Blue), written in plain HTML and vanilla JavaScript with no frameworks and no build step. It opens and saves real `.pptx` files.

The app is `public/lectern/` (`python3 tools/serve.py`, then http://127.0.0.1:8760/lectern/); every path below is relative to it, except `tools/`. It runs in any modern browser (Chrome, Edge, Firefox, Safari). Everything runs locally; nothing is uploaded, including the passwords of protected presentations.

## What's in the box

| File | Responsibility |
|---|---|
| `index.html` | Window chrome, Luna CSS, slide rendering styles, script order |
| `../common/core.js` | Utilities, colors, fonts, file I/O, media store, PDF writer (shared, see docs/suite.md) |
| `../common/sha.js`, `../common/crypto.js` | Password-protected presentations: ECMA-376 Agile and Standard encryption (AES, SHA-1/SHA-2), read and write |
| `../common/zip.js` | ZIP reader/writer (CompressionStream with a pure-JS inflate fallback) |
| `../common/xml.js`, `../common/opc.js`, `../common/opc-order.js` | Shared OOXML preservation graph, identities, namespace capture and schema order |
| `../common/geometry.js` | 92 AutoShape presets, adjust handles, text rectangles, connectors |
| `js/model.js` | Presentation model, 14 design templates, 21 slide layouts, undo history |
| `js/render.js` | DOM/SVG renderer for slides, fills, patterns, gradients, WordArt |
| `../common/metafile.js` | WMF/EMF player, so old clip art, logos and pasted Excel/Visio graphics display |
| `../common/numfmt.js` | Excel number formats (dates, currency, percentages, sections) for chart axes and data labels |
| `../common/charts.js` | Chart model, the PowerPoint 2003 look for charts made here, datasheet and chart options dialog (17 chart types with `js/chart-draw.js`) |
| `js/chart-draw.js` | Drawing of charts that come from PowerPoint 2007 and later (see *Charts* below) |
| `../common/ui.js`, `../common/icons.js`, `../common/clipart.js`, `../common/luna.css` | Menus, command bars, combo boxes, dialogs, color pickers, tooltips |
| `js/editor.js` | Slide editing surface: select, move, resize, rotate, crop, draw, grid |
| `js/textedit.js` | Rich text editing, bullets, levels, AutoCorrect, tables |
| `js/dialogs.js` | Font, Bullets, Format AutoShape, Fill Effects, Header & Footer, Options (with Security), etc. |
| `js/panes.js` | Slides/Outline tab, notes, sorter, notes page, print preview, task panes |
| `js/slideshow.js` | Full-screen show, 58 transitions, 50+ entrance/emphasis/exit/motion-path effects, pen, rehearse timings |
| `js/preserve.js` | Package ownership and settings merges; model-attached media, playback, property invalidation and clipboard dependencies |
| `js/pptx-read.js` | PresentationML import (themes, masters, layouts, placeholders, table styles, charts, SmartArt, custom geometry, animations; repairs damaged packages) |
| `js/pptx-write.js` | PresentationML export (validated against the OOXML schema) |
| `js/commands.js` | Every menu/toolbar command |
| `js/app.js` | Application wiring, keyboard shortcuts, clipboard, sample deck |
| `js/templates.js` | Content templates (Project Status Report, Team Meeting, Lesson, Event Announcement) for the Start Center and File ▸ New |
| `tools/lectern/test/` | The real-file test harness described below |

## Opening

File ▸ Open, drag a file onto the window, or *Open a presentation* in the Getting Started pane. `.pptx`, `.pptm`, `.ppsx`, `.potx` and their macro variants open; Strict Open XML files open too. Damaged packages are repaired where possible, and legacy binary `.ppt` files get a clear message.

**Password-protected presentations** ask for the password, as PowerPoint does, and say so when it is wrong. Decryption happens in the browser (Agile encryption from Office 2010 on, Standard encryption from Office 2007). A presentation opened with a password is saved with the same password; Tools ▸ Options ▸ Security sets, changes or removes it (AES-256 with SHA-512, the scheme PowerPoint 2013 and later use).

## Charts

Charts read from a file are drawn the way PowerPoint draws them:

* column, bar (clustered, stacked, 100 %), line (with or without markers, smoothed), area, pie, exploded pie, doughnut, XY scatter, bubble, radar, stock (high-low lines) and combinations of them, with secondary axes;
* 3-D columns, bars and pies are drawn in an oblique view with walls, rows in depth and a pie rim;
* Office's automatic axis scaling (zero kept in view unless the data stay away from it, 5 % head-room, at most ten major units), fixed minimum/maximum/major unit, logarithmic and reversed axes, display units, number formats (dates included), axis titles, crossing points;
* gap width and overlap, series and point colours (including colours that come from the chart style or a theme override), marker symbols and sizes, line widths and dashes, data labels (value, percentage, category, series name, custom text with fields, values from cells), legends in every position or placed by hand, manual plot-area layout, multi-level category labels, wrapped or turned category labels.

A chart that is not edited in Lectern is saved exactly as it came: its XML, the embedded Excel workbook, the chart style and colour parts and any theme override travel with it, so PowerPoint can still open the data in Excel. Once the datasheet or chart options are changed (double-click the chart), the chart is written from Lectern's model, keeping colours, labels, axis settings and gap width. New charts keep the PowerPoint 2003 look.

## Slide size

File ▸ Page Setup. Choose *Widescreen (16:9)* (13.33 × 7.5 in, today's PowerPoint default) or *On-screen Show (16:9)* (10 × 5.63 in). Pictures, shapes, charts and tables keep their proportions; placeholders reflow to the new frame. Tick *Use for new presentations* to make it the default.

## Saving

* **.pptx** — standard PresentationML. Exports pass the OOXML schema validator and open in LibreOffice and python-pptx; they are meant for PowerPoint 2007 and later, Keynote and Google Slides. With a password to open, the file is encrypted.
* **.ppsx / .potx** — show and template variants (standalone page only).
* **PDF** — one slide per page, or exactly what Print Preview shows (handouts, notes pages, outline).
* **Single-file web page**, **PNG** of the current slide, **plain-text outline**.

Pictures survive a round trip in every form they come in: SVG pictures keep the vector original next to a PNG copy, linked pictures keep their link, pictures whose image is missing keep their frame, and a transparent colour (Picture toolbar ▸ Set Transparent Color, or `a:clrChange` in a file) is stored as a setting, not baked into the image.

Imported video and audio retain their original media bytes, links, pictures and playback timing on
save. Lectern displays their poster pictures; playback in PowerPoint still needs Office acceptance.
Moving, resizing, cropping, duplicating, copying between decks and draft recovery preserve the media.
Replacing the media picture converts it to an ordinary picture and records that conversion. The
Compatibility Checker dialog is the next integration stage; these records are not yet shown on Save.

## Testing against real presentations

The importer was tested on **1,672 distinct presentations (253 MB)** collected from the public test suites and sample folders of 20 open-source projects: LibreOffice (577), Aspose.Slides examples for Java and .NET (346), the Open XML SDK (211), Apache POI (109), pandoc (103), ShapeCrawler (98), python-pptx (66), pptx-automizer (43), Open-Xml-PowerTools (38), Apache Tika (29) and ten smaller ones. Many are deliberately hard: fuzzed and truncated packages, Strict Open XML, password-protected files, decks from PowerPoint 2007 betas, and generator output from a dozen libraries.

For every file, `tools/lectern/test/corpus.js` opens the deck in headless Chromium, loads it into the editor, renders every slide, rasterises the first six, saves it, opens the saved copy again and compares slide by slide (shape count, shape types, text). The saved files are then checked by external tools: python-pptx opens them, the OOXML schema validator checks them, and LibreOffice converts them. `test/viscompare.py` compares Lectern's slide images with LibreOffice's renderings of the same 670 slides (400 decks), and `test/triptych.py` builds before/after/LibreOffice montages of the slides that differ most, which is how most of the fixes below were found.

| | Before this round | Now |
|---|---|---|
| Decks that open | 1,662 | 1,665 (3 more: password-protected, with the passwords their test suites document) |
| Errors loading, rendering, saving or reopening | 0 | 0 |
| Decks whose shapes, text or pictures change on save (round trip; other content: see *What a save keeps* below) | 14 (SVG, linked and broken pictures dropped) | 0 |
| Slide counts different from python-pptx (1,505 decks compared) | 0 | 0 |
| Charts saved with their workbook and styles | 0 of 441 (rewritten from a summary) | all unedited charts |
| Saved files that python-pptx opens | | 1,665 of 1,665 (3 after decryption) |
| Saved files that pass the OOXML schema validator (random 240) | | 240 of 240 |
| Edited charts written from the model that pass the validator | | 430 charts in 257 decks, all |
| Saved files LibreOffice converts (random 30) | | 28 of 30, same page counts; the other two hold only a hidden slide, which LibreOffice will not export from the original file either |
| Difference from LibreOffice's rendering, median / 90th percentile (670 slides) | 1.21 / 15.2 | 1.09 / 8.8 |
| Slides that differ strongly from LibreOffice (score > 12 / > 20) | 89 / 47 | 42 / 19 |
| Median time to open and render a deck | 43 ms | 43 ms |

The seven files that still do not open are four that are not presentations at all (a zero-byte file, two 22-byte stubs and an exploit sample from LibreOffice's CVE tests), two encrypted decks whose passwords are not published, and one fuzzed package with no presentation part.

Fixes found this way, besides charts and passwords: shapes filled with the slide background (`useBgFill`), theme background textures and their duotone recolouring, pictures in theme fill styles, the focus point of radial gradients, stretched fills that reach past the shape (`fillRect`), percentages written as `"90%"` (Strict Open XML), pictures with a transparent colour, empty placeholders left out of thumbnails and the slide show, custom date fields that pin literal text (think-cell), ActiveX controls shown by their preview pictures, Wingdings 2 and 3 bullets, and better layouts for SmartArt saved without its drawing (block lists, horizontal and grouped lists, matrices, process, chevrons, cycles, org charts with connectors) in the diagram's own colour style.

To run the harness: serve Lectern (`python3 tools/serve.py`) and the folder above the decks (`python3 -m http.server 8766`; set `HOST` and `CORPUS_URL` to use other addresses), then `node tools/lectern/test/corpus.js <folder-or-list> results.jsonl --png png/ --save saved/` and `node tools/lectern/test/summarize.js results.jsonl`. `tools/lectern/test/ui-features.js` drives the password dialog, Tools ▸ Options ▸ Security and the chart dialog in the real UI; `tools/lectern/test/chartedit.js` writes every chart from the model for validation; `tools/lectern/test/decrypt.py` checks encrypted saves with an independent implementation.

## What a save keeps, converts and drops

Lectern regenerates slide content from its model and carries original media properties alongside it.
The media and package stages below measure the current implementation. General object preservation
is still in progress; the older baseline tables remain relevant to those unfinished paths.
`tools/lectern/test/loss-audit.py` compares original and saved packages feature by feature and word by
word (`python3 tools/lectern/test/loss-audit.py corpus/ saved/ out.json`).

### Package preservation (2026-10-08)

Save, Save As and draft recovery keep all six variants: `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`,
`.potm`. Password-protected drafts remain encrypted. VBA is preserved in macro-enabled formats;
macro-free Save As reports its removal in the Compatibility Checker before download. A cancelled
save neither changes the file identity nor acknowledges the warning.

Custom XML, tags, custom-property types/pids, labels, embedded fonts, comment authors, handout and
notes masters are carried through their original relationships. Unedited notes pages keep their
header/footer text; duplicating a slide copies its notes page and redirects the backlink. Unedited
designs used by slides retain their original master, all its layouts and theme. Carrying masters that
no slide uses, design-edit invalidation, section/custom-show edits and general opaque frames remain
work for the next Lectern stage.

The pinned run remains **840 save/reopen, 3 failed, 1 excluded**. Package/SDK comparison remains
**822 OK, 21 failed, 1 excluded**, with no additional diagnostics in that measured run. Independent
checks retain **111 custom XML, 70 custom-property, 33 tag, 29 comment-author, 12 font, 7 VBA and 1
label relationships**, plus **97 handout-master and 932 notes-master references**, with their bytes,
content types, rIds and targets intact. The repeated media check passes **110 of 110** operations.

The checker reports entries recorded by the implemented preservation paths. Complete reporting of
content conversions depends on the remaining object work. Reports are retained in
`~/corpora/results/package-preservation-2026-10-08/`; samples are in
`~/Downloads/lossless-check/lectern/package/`. PowerPoint acceptance and the slide-6 repair diagnosis
remain pending. Full-corpus LibreOffice rendering has not been completed.

The following older baseline used 1,652 decks, before preservation was added. Its media counts have
not been remeasured on that larger corpus; the pinned three-source corpus below is reproducible.

**Kept.** The words on the slides (all but 24 of 111,243, in two decks: a text box holding an equation, saved as the picture PowerPoint stores for older readers, and text typed into connector-shaped AutoShapes) and in the speaker notes (7,274 of 7,282); pictures, shapes, tables, group structure; animations (2,873 of 2,876 effects) and transitions (143 of 144); links to web pages, e-mail addresses and other slides; unedited charts byte for byte with their workbooks.

**Converted** — still visible, no longer editable the original way: SmartArt becomes grouped shapes (129 decks); embedded Excel, Word or Visio objects become their preview pictures (46 decks); Office 2016 charts (waterfall, sunburst, funnel…) become the picture stored with them (9 decks); text with equations becomes a picture.

**Dropped.**

| What | Decks in the test set |
|---|---|
| Video and audio, and the animations that play them | 43 (29) |
| Comments | 21 |
| Sections, custom shows | 29, 6 |
| Handout master; slide layouts and masters that no slide uses | 154; 8,668 layouts in 1,316 decks |
| Header, footer and date text on notes pages (often a copyright line) | 214 decks, 752 words |
| 3-D bevel and extrusion, glow, reflection, soft edges, inner shadow, text effects, picture recolouring | 50, 27, 11, 16, 8 |
| Action settings that run a program or macro; links to files or other presentations | 39; 70 of 536 links |
| Ink, 3-D models, embedded fonts | 5, 2, 15 |
| VBA macros (a `.pptm` is saved as `.pptx`) | 11 |
| Custom XML parts and tags, custom document properties (document IDs, add-in data) | 317, 126 |
| Sensitivity labels, in both forms (`MSIP_Label` custom properties and `docMetadata/LabelInfo.xml`) | 10 and 16 |
| Auto-advance set to 0 seconds | 14 of 47 timings |

### Pinned lossless-save baseline (2026-10-07)

The reproducible three-source corpus in `tools/corpora.sh` was measured against application revision
`5abafdc`, before preservation changes. Of 844 inputs, 840 completed open → save → reopen,
3 failed and 1 were explicitly excluded. Completion is not a fidelity result. The independent
feature inventory compared 830 readable pairs; malformed/encrypted originals and missing
outputs remain in its accounting. This corpus differs from the earlier compatibility corpus above.

| Feature | Files containing it | Original inventory items | Saved inventory items | Text edit: kept / original |
|---|---:|---:|---:|---:|
| video / audio on a slide | 17 | 29 | 0 | 0 / 5 |
| SmartArt (editable diagram data) | 106 | 184 | 0 | 0 / 180 |
| comments | 6 | 20 | 0 | 0 / 7 |
| VBA macros | 7 | 7 | 0 | 0 / 7 |

The text-edit inventory covers 545 readable saved pairs; files without editable text and failed
operations stay in the driver report, so its denominator differs from the unedited inventory.

Package checks plus the Office 2019 SDK comparison reported 822 attempts without new
automated diagnostics and 21 failures (including originals that could not be validated);
1 driver exclusions remain separate. These checks do not certify Office acceptance. Commands,
comparison policies and failure accounting are in [testing.md](testing.md#lossless-same-format-save-reproducible-baseline).

The seven-slide sample associated with the slide-6 repair report passes the package checker
and SDK validation and converts to seven pages in LibreOffice before and after saving. The
PowerPoint result for the six diagnostic copies remains **pending**; these automated results do
not establish a repair fix.

`tools/lectern/test/timing-repair.mjs` now regenerates that diagnostic set and can use the exact
reported sample as its source. The original and four isolated slide-6 variants pass both package
checks and SDK validation. LibreOffice's resaved control passes the SDK but introduces a duplicate
shape ID in `ppt/slideLayouts/slideLayout1.xml`, which the package checker reports. The latest copies
and validation manifest are in `~/Downloads/lossless-check/lectern/slide6-diagnostics/`.
All six copies convert to seven pages in LibreOffice.
The slide-6 fix remains pending the PowerPoint results; no speculative timing change has been made
to the writer. Independent work on the shared preservation core proceeds while acceptance is pending.
The text-edit/save/undo/redo scenario completed in 551 files, excluded 290 without editable
slide text, and failed in 3.

Of 12 feature samples, 11 converted in LibreOffice with unchanged slide counts. LibreOffice
failed to open the original `tdf173266.pptx`; that pair remains a failed attempt. The full corpus
has not yet been rendered in LibreOffice. The baseline reports are retained outside git in
`~/corpora/results/lossless-baseline-2026-10-07/`; Office samples are in
`~/Downloads/lossless-check/baseline/`.

### Shared preservation core (2026-10-08)

The reader captures original AlternateContent before selecting a branch and normalising percentages.
At this stage the writer did not emit those records, so the feature-loss numbers above still applied. The same
844 inputs again produced 840 successful open/save/reopen attempts, 3 failures and 1 exclusion.
All 840 saved outputs completed package/SDK comparisons against the earlier saves with no new
diagnostics. Twelve LibreOffice samples kept their slide counts relative to the earlier saves;
the original `tdf173266.pptx` remains outside that comparison because LibreOffice cannot open it.
Reports: `~/corpora/results/opc-core-2026-10-08/`. PowerPoint acceptance remains pending.

### Media preservation (2026-10-08)

The pinned corpus now keeps **29 of 29 audio/video/media relationships in all 17 relevant decks**,
up from 0 of 29. Original picture frames include AlternateContent Choice and Fallback, p14 media,
playback and interactive sequences. All 16 media play/pause behaviors also survive, including
CD-track audio with no media relationship. Transition sounds, click sounds and external sound URLs are
carried too. Slide, shape and timing identities stay connected; converted SmartArt children receive
their own slide IDs instead of reusing the diagram's ID space.

| Edit to a media object | Preserved data |
|---|---|
| Text elsewhere on the slide | Keep the frame, effects, media and playback |
| Move, resize, rotate, or group coordinates | Update the outer transform; keep the media and inner properties |
| Crop, line, geometry or picture adjustment | Replace that property only |
| Shadow | Replace `outerShdw`; keep glow, reflection, bevel and 3-D properties |
| Duplicate or copy into another deck | Import dependencies and remap copied shape/timing identities together |
| Delete or replace the media picture | Remove its playback; replacement records a conversion |
| Undo/redo or draft recovery | Restore the preserved properties and dependency references |

Independent ZIP/XML checks pass all **120 real-file operations** (12 media-picture decks × 10
scenarios), plus **20 operations** on Transitional and Strict AlternateContent/effects fixtures derived
from an authored deck. They compare media bytes, intended relationship targets, unedited properties, playback and exact
undo/redo package restoration apart from core timestamps. Keyboard copy transports dependencies into
a separate browser tab. Another 117 saved states across all 18 decks keep sound properties and bytes
through repeat-save, text edits and draft recovery.

All 844 corpus attempts retain the baseline result: **840 save/reopen, 3 failed, 1 password exclusion**.
Package/SDK comparison against originals reports **822 OK, 21 failed, 1 excluded**, with no new
diagnostics. The 21 include malformed packages and originals the SDK cannot open; they are not counted
as passes. Across the media operations, 374 of 408 saved states compare successfully; 34 have the same
unvalidatable original. All 68 AlternateContent fixture states compare successfully. The saved states
of two malformed media originals (40 artifacts) also pass standalone SDK validation after repairs to
missing media content types and escaped Unicode part URIs.

LibreOffice opens all **18 of 18** final media samples with unchanged slide counts. Full-corpus
LibreOffice rendering and PowerPoint acceptance remain pending, including the separate slide-6
repair report. Reports and saved files are retained in
`~/corpora/results/lectern-media-2026-10-08/`; Office samples are in
`~/Downloads/lossless-check/lectern/media/`. The audit compares 827 readable pairs; three bad-CRC
originals previously counted as readable are now explicit failures. The text-loss inventory is
unchanged. Transitions improve from 111/112 to 112/112; other unsupported actions and general design,
object, comment and package preservation follow
this stage.

## Known limits

* Legacy binary `.ppt` files must be re-saved as `.pptx` first. Macro-enabled OOXML variants keep their VBA; see *What a save keeps, converts and drops* for the remaining preservation gaps.
* SmartArt imports as grouped shapes. When a file carries only the diagram data (PowerPoint 2007 saved no drawing), Lectern lays it out itself with one of its built-in layout families, so unusual layouts look simpler than in PowerPoint.
* 3-D charts are drawn in a fixed oblique view (no perspective, no lighting); surface charts, trendlines and error bars are not drawn (they are kept in unedited charts).
* Embedded OLE objects and ActiveX controls import as their preview pictures.
* Fonts that aren't installed fall back to metric-compatible web fonts (Arimo, Tinos, Cousine, Carlito, Caladea).
* Spelling uses the browser's spell checker.
