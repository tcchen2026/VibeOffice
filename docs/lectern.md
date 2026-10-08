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

Strict inputs save as standard Transitional OOXML, with a Compatibility Checker notice. Same-format
saves retain the macro, template or slideshow variant and any opening password.

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
save. Lectern displays their poster pictures; the video sample passed Office batch 1.
Moving, resizing, cropping, duplicating, copying between decks and draft recovery preserve the media.
Replacing the media picture converts it to an ordinary picture and records that conversion. The
Compatibility Checker shows recorded conversions before the user downloads a saved file.

## Testing against real presentations

`tools/corpora.sh` pins the LibreOffice, Apache POI and Open XML SDK files used in the current figures
below. `tools/lectern/test/corpus.js --lossless` exercises save/reopen, actual history operations and
suite draft recovery. The rendering mode also rasterizes slides for comparison with LibreOffice;
`viscompare.py` and `triptych.py` inspect layout differences. Open XML SDK and independent ZIP/XML
checks compare each output with its original. Earlier broad-corpus results are archived with the
run history, rather than presented as measurements of the current writer.

The UI/password/chart harnesses remain in `tools/lectern/test/`. Commands, failure accounting and
Office review batches are documented in [testing.md](testing.md).

## What a save keeps, converts and drops

Lectern regenerates slides and keeps package dependencies and media properties beside the model.
Save and drafts retain all six OOXML variants and encryption. The Compatibility Checker reports
recorded conversions before a user download; complete coverage depends on the remaining object work.

The pinned 844-file corpus completes **840 save/reopen attempts**, with **3 failures and 1 password
exclusion**. Text edit/save/undo/redo completes in **551**, excludes **290** without editable text and
fails in **3**. Draft recovery completes in **840**, with **3 failures and 1 exclusion**. The full run
uses writer `cf1b313`; package/SDK comparison reports **822 OK, 21 failed and 1 excluded** for saves,
and no new diagnostics in comparable outputs. Unvalidatable originals are failures, not passes.

The independent feature inventory compares **827 readable pairs**; **17 attempts fail**. Its text
inventory retains **57,931 of 58,131 words**: **812 of 827 decks** keep every counted word. Part and
relationship counts do not establish that an unsupported frame remains attached to its slide.

| Status | Feature | Current measured result / limit |
|---|---|---|
| Kept | Video/audio and playback | 29/29 media relationships in 17 decks; 16/16 media play/pause behaviors; original frames, links and media bytes survive |
| Kept | Package metadata and fonts | 111/111 custom-XML, 67/67 XML-property, 33/33 tag, 70/70 custom-property, 12/12 font, 7/7 VBA and 1/1 label relationships |
| Kept | Notes and handout dependencies | 932/932 notes-master references, 97/97 handout masters and 29/29 comment-author relationships; unedited notes keep header/footer text |
| Kept as records; edit rules pending | Comments, sections and custom shows | 20/20 comment inventory items, 32/32 sections and 8/8 custom shows; membership and slide deletion/insertion rules remain unfinished |
| Kept, with remaining losses | Animation and transitions | 639/642 general effects, 112/112 transitions and 29/43 auto-advance timings |
| Kept for used designs; unused designs incomplete | Masters and layouts | 864/896 masters and 7,209/7,249 named layouts; all layouts of retained used masters survive |
| Converted to displayed groups/pictures | SmartArt, OLE and newer charts | 0/184 SmartArt frames, 0/71 OLE frames and 0/4 chartEx frames; their 184 diagram relationships, 51 OLE relationships and 4 chartEx parts remain in the package, which alone does not preserve editability |
| Dropped from slide content | Ink and 3-D models | 0/44 ink/content-part entries and 0/2 3-D models |
| Dropped on general non-media shapes | Extended visual effects | 0/64 glow/reflection/soft-edge, 0/109 3-D effect, 0/36 inner-shadow, 0/4 recolour/artistic and 0/24 run-effect inventory items |
| Partly kept | Links, program/macro actions and shape tags | 369/422 hyperlinks and 14/29 action settings; shape-level tag ownership and unsupported action types remain unfinished |

The custom-XML/tag part inventory is **234/244** even though all selected package relationships pass:
unreferenced parts and modeled content have separate ownership rules. The Compatibility Checker
reports losses known to the implemented paths; the remaining conversions are still being covered.

The user accepted the PPTM video, VBA and authors/handout samples in Office batch 1. The slide-6
repair is fixed and confirmed in PowerPoint: presets write complete adjustment lists and masters
share an identity space with layouts. All 15 suite batch files open in LibreOffice with unchanged
counts. Strict conversion passes **85 presentations / 340 emitted states** against original
package/SDK diagnostics; the Strict presentation sample passed Office batch 2. Rewritten parts now
hoist fragment namespace declarations and `mc:Ignorable` to the root. The fresh namespace sample
passes PowerPoint, package/SDK and LibreOffice with unchanged slide count. Opaque parts retain their
bytes. Full-corpus
LibreOffice rendering has not been completed. Reports and history:
`~/corpora/results/next-plan-head-2026-10-08/` and
`~/corpora/results/strict-transitional-2026-10-08/`.

### Media: edit behavior

| Edit | Save behavior |
|---|---|
| Text elsewhere on the slide | Keep the frame, effects, media and playback |
| Move, resize, rotate or change group coordinates | Update the outer transform; keep media and inner properties |
| Crop, line, geometry or picture adjustment | Replace that property only |
| Shadow | Replace `outerShdw`; keep glow, reflection, bevel and 3-D properties |
| Duplicate or copy into another deck | Import dependencies and remap copied shape/timing identities together |
| Delete or replace the media picture | Remove its playback; replacement records a conversion |
| Undo/redo or draft recovery | Restore properties, dependencies and pending losses |

### Package, notes and designs: edit behavior

| Edit | Save behavior |
|---|---|
| Edit slide text | Keep unrelated package parts, custom-property types/pids and unedited notes |
| Duplicate a slide | Copy its notes page and redirect its backlink |
| Leave a used design unchanged | Retain its original master, every layout and theme |
| Change a design, or leave a master unused | Complete invalidation/retention rules are pending |
| Save as a macro-free variant | Report VBA removal; cancellation leaves file identity and acknowledgements unchanged |

### Sections, comments and opaque frames: current edit limits

| Edit | Current behavior / remaining work |
|---|---|
| Insert, delete or reorder slides in sections/custom shows | Original lists remain; membership and stable-identity updates need integration |
| Change or delete a commented slide | Comment parts are carried; complete legacy/modern comment ownership rules remain pending |
| Move, resize, edit or ungroup converted SmartArt/OLE/new-chart content | The displayed conversion is saved; original-frame ownership and geometry merging remain pending |
| Change a non-media shape's shadow, fill or picture treatment | Regenerate modeled properties; unknown sibling effects still need per-property preservation |
| Copy shape tags or unsupported file/program actions | Dependency and reference remapping still needs integration |

## Known limits

* Legacy binary `.ppt` files must be re-saved as `.pptx` first. Macro-enabled OOXML variants keep their VBA; see *What a save keeps, converts and drops* for the remaining preservation gaps.
* SmartArt currently saves as grouped shapes; its carried diagram parts do not restore the original frame. When a file carries only the diagram data (PowerPoint 2007 saved no drawing), Lectern lays it out itself with one of its built-in layout families, so unusual layouts look simpler than in PowerPoint.
* 3-D charts are drawn in a fixed oblique view (no perspective, no lighting); surface charts, trendlines and error bars are not drawn (they are kept in unedited charts).
* Embedded OLE objects and ActiveX controls import as their preview pictures.
* Fonts that aren't installed fall back to metric-compatible web fonts (Arimo, Tinos, Cousine, Carlito, Caladea).
* Spelling uses the browser's spell checker.
