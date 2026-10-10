# Lectern 2003 — Web Edition

A presentation editor modeled on PowerPoint 2003, written in plain HTML and vanilla JavaScript with no frameworks and no build step. It opens and saves real `.pptx` files. In the suite's Paper 2016 look (Themes on the Start Center, or Tools ▸ Options ▸ General ▸ Theme) a ribbon with PowerPoint 2016's tabs (Home, Insert, Design, Transitions, Animations, Slide Show, Review, View and the tool tabs) takes the place of the menus and toolbars.

Open https://vibeoffice.work/lectern/ (to run it locally, see [testing.md](testing.md)). The app is `public/lectern/`; every path below is relative to it, except `tools/`. It runs in any modern browser (Chrome, Edge, Firefox, Safari). Everything runs locally; nothing is uploaded, including the passwords of protected presentations.

## What's in the box

| File | Responsibility |
|---|---|
| `index.html` | Window chrome CSS, slide rendering styles, script order |
| `../common/core.js` | Utilities, colors, fonts, file I/O, media store, PDF writer (shared, see docs/suite.md) |
| `../common/sha.js`, `../common/crypto.js` | Password-protected presentations: ECMA-376 Agile and Standard encryption (AES, SHA-1/SHA-2), read and write |
| `../common/zip.js` | ZIP reader/writer (CompressionStream with a pure-JS inflate fallback) |
| `../common/xml.js`, `../common/opc.js`, `../common/opc-order.js` | Shared OOXML preservation graph, identities, namespace capture and schema order |
| `../common/geometry.js` | 92 AutoShape presets, adjust handles, text rectangles, connectors |
| `js/designs.js` | Imported master/layout library, selective design changes and independent design copies |
| `js/properties.js` | Shape/text properties, tags, actions and transitions retained until their owning property changes |
| `js/model.js` | Presentation model, 14 design templates, 21 slide layouts, undo history |
| `js/render.js` | DOM/SVG renderer for slides, fills, patterns, gradients, WordArt |
| `../common/metafile.js` | WMF/EMF player, so old clip art, logos and pasted Excel/Visio graphics display |
| `../common/numfmt.js` | Excel number formats (dates, currency, percentages, sections) for chart axes and data labels |
| `../common/charts.js` | Chart model, the PowerPoint 2003 look for charts made here, datasheet and chart options dialog (17 chart types with `js/chart-draw.js`) |
| `js/chart-draw.js` | Drawing of charts that come from PowerPoint 2007 and later (see *Charts* below) |
| `../common/ui.js`, `../common/search.js`, `../common/icons.js`, `../common/clipart.js`, `../common/looks.css`, `../common/ui.css` | Menus, command bars, combo boxes, dialogs, color pickers, tooltips, Search commands |
| `js/ribbon.js`, `../common/ribbon.js` | The Paper 2016 look's ribbon: PowerPoint's tabs and groups over Lectern's commands ([suite.md](suite.md#the-ribbon)) |
| `js/editor.js` | Slide editing surface: select, move, resize, rotate, crop, draw, grid |
| `js/textedit.js` | Rich text editing, bullets, levels, AutoCorrect, tables |
| `js/dialogs.js` | Font, Bullets, Format AutoShape, Fill Effects, Header & Footer, Options (with Security), etc. |
| `js/panes.js` | Slides/Outline tab, notes, sorter, notes page, print preview, task panes |
| `js/slideshow.js` | Full-screen show, 58 transitions, 50+ entrance/emphasis/exit/motion-path effects, pen, rehearse timings |
| `js/preserve.js` | Package ownership and settings merges; model-attached media, playback, property invalidation and clipboard dependencies |
| `js/comments.js` | Legacy/modern comment ownership, author dependencies, copied identities and anchors |
| `js/frames.js` | Opaque frame membership, geometry, diagram/VML dependencies and preview clipboard transport |
| `js/pptx-read.js` | PresentationML import (themes, masters, layouts, placeholders, table styles, charts, SmartArt, custom geometry, animations; repairs damaged packages) |
| `js/pptx-write.js` | PresentationML export (validated against the OOXML schema) |
| `js/commands.js` | Every menu/toolbar command |
| `js/app.js` | Application wiring, keyboard shortcuts, clipboard, sample deck |
| `js/templates.js` | Content templates (Project Status Report, Team Meeting, Lesson, Event Announcement) for the Start Center and File ▸ New |
| `tools/lectern/test/` | The real-file test harness described below |

## Opening

File ▸ Open, drag a file onto the window, or *Open a presentation* in the Getting Started pane. `.pptx`, `.pptm`, `.ppsx`, `.potx` and their macro variants open; Strict Open XML files open too. Damaged packages are repaired where possible, and legacy binary `.ppt` files get a clear message.

Strict inputs save as standard Transitional OOXML (recorded, not shown). Same-format
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
Replacing the media picture converts it to an ordinary picture; the Compatibility Checker tells the
user before the download that it will no longer play.

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

The full LibreOffice pass on the versioned `984eaf3` artifacts renders **836/837 unedited saves**,
with **zero** page-count differences among pairs that render. Across **5,302** attempted states,
**4,939** pass the paired render/count check, **71** fail and **292** are excluded. Failures include
missing outputs, unrenderable originals and LibreOffice crashes; they remain recorded separately
from package/SDK results. All 15 batch-4 samples and the batch-7 corrections passed PowerPoint,
including the corrected fixture described below. Reports: `~/corpora/results/lossless-final-984eaf3/lectern/`.

Lectern regenerates slides and keeps package dependencies and unsupported properties beside the model.
Save and drafts retain all six OOXML variants and encryption. The Compatibility Checker shows, before
a user download, only the losses that cost the user content or a working feature (see docs/suite.md, What the user is told);
drafts retain those notices silently.

The pinned **844-file** run completes **837 saves and 837 draft recoveries**, with **6 failures and
1 password exclusion** in each scenario. Text edit/save/undo/redo completes in **548**, excludes
**290** without editable text and fails in **6**. Across **5,302 emitted or attempted states**, package/SDK
comparison has **4,914 OK, 96 failed and 292 excluded**, with **no new diagnostics**. Failed originals
and incomplete recovered decks remain failures. The current full writer is `984eaf3`.
Two initial text/draft timeouts pass a focused retry with a 240-second deadline; the original failed
attempts remain recorded. SDK evidence is reused for 4,608 states only after the original hash and
every saved ZIP member match the previous validated output; package checks are fresh. Four PDF
count comparisons are corrected for authored hidden slides, with unchanged slide identities,
visibility and PDF hashes; those original failed comparisons also remain in the primary report.

The independent inventory compares **827 readable pairs** and records **17 failed comparisons**.
All **58,022 counted words** and every inventoried feature remain after an unchanged save, including
**896 masters, 7,265 named layouts, 642 effects, 232 transition branches and 184 SmartArt references**.
Counts alone do not establish rendering fidelity; the edit checks below verify ownership and dependencies.

| Status | Feature | Current measured result / limit |
|---|---|---|
| Kept | Video/audio and playback | 29/29 media relationships in 17 decks; 16/16 media play/pause behaviors; original frames, links and media bytes survive |
| Kept | Package metadata and fonts | 111/111 custom-XML, 67/67 XML-property, 33/33 tag, 70/70 custom-property, 12/12 font, 7/7 VBA and 1/1 label relationships |
| Kept | Notes and handout dependencies | 932/932 notes-master references, 97/97 handout masters and 29/29 comment-author relationships; unedited notes keep header/footer text |
| Kept; PowerPoint samples accepted | Sections and custom shows | All 18 relevant decks retain 32/32 sections and 8/8 custom shows with unchanged expanded XML. Eight focused edit cases / 26 saved states pass package/SDK comparison |
| Kept; PowerPoint samples accepted | Comments | The 30 relevant decks retain 39/39 comment/author parts: 37 byte-identical, two Strict parts converted to Transitional. Copy/delete/paste ownership and modern anchors are implemented |
| Kept; PowerPoint samples accepted | Animation and transitions | Original timing and all 482 transition elements in the 214-deck property sweep remain; editing a modeled animation retains unrelated interactive sequences |
| Kept; PowerPoint samples accepted | Masters and layouts | All 54 masters and 276 layouts in the 13 unused-master decks remain; 52 masters and every layout retain exact bytes. Two damaged source masters require conversion |
| Kept; PowerPoint samples accepted | SmartArt, OLE and newer charts | 184/184 SmartArt, 71/71 OLE entries and 4/4 chartEx frames remain attached to their slides across all 135 relevant inputs; three missing drawing-cache references are removed with a notice |
| Kept; PowerPoint samples accepted | Ink and 3-D models | 44/44 ink/content-part entries and 2/2 3-D models; unavailable previews have editable placeholder frames |
| Kept; PowerPoint samples accepted | Extended visual effects | The property sweep retains all 18 glow, 30 reflection, 16 soft-edge, 36 inner-shadow, 136 scene3d, 123 sp3d and 6 picture-treatment entries; unknown run/paragraph details remain with their properties |
| Kept; PowerPoint samples accepted | Program/macro actions and shape tags | All 29 action settings and 13 slide/shape tag lists remain in the property sweep; copied tag parts and shape identities are independent |
| Converted, with a notice | Edited opaque contents, imported charts and notes formatting | The text script converts 67 selected SmartArt previews and removes two effects targeting one converted object's parts; untouched inventoried features remain |
| Dropped, with a notice | Missing source dependencies, invalidated signatures and VBA in macro-free Save As | Damaged source parts cannot be reconstructed; explicit deletion and format changes remove their owned parts |

The custom-XML/tag inventory retains **234/234 parts**, excluding ZIP directory markers. Package
carry checks pass byte comparison in **818 files**; **9 Strict inputs** deliberately convert retained
XML and relationship vocabulary. Sixteen malformed/failed inputs and one password exclusion remain
separate. The edit inventory covers **545 readable pairs**: **67 SmartArt objects whose preview text
the script edits** are converted, with **two part-animation effects** on one of those objects removed
with their own notice. Other inventoried features remain. Run-boundary edits and those converted
previews account for the word-count differences. Specific frame notices now explain their opaque
dependencies too; generic unreferenced-part notices fall from 375 to 58 in the edited run, with
unrelated losses still reported.

Documented exceptions are damaged or missing source dependencies, signatures invalidated by a save,
Strict-to-Transitional conversion, edited chart/opaque-frame content, and edited notes-page formatting.
Form-control or field markup that requires conversion is also reported. The full-corpus LibreOffice
results are above. The **15 PowerPoint batch-4 samples rebuilt from `56de065`** pass package/SDK
and LibreOffice with expected counts. Their previous-writer differences are reviewed, and the user
accepted all 15 in PowerPoint. Fresh LibreOffice checks of those exact handed-off files also pass;
hashes and results are in `~/corpora/results/office-accepted-2026-10-09/`. The accepted checklist is
`~/Downloads/lossless-check/office-batch-4/CHECKLIST.md`.

The user accepted the PPTM video, VBA and authors/handout samples in Office batch 1. The slide-6
repair is fixed and confirmed in PowerPoint: presets write complete adjustment lists and masters
share an identity space with layouts. All 15 suite batch files open in LibreOffice with unchanged
counts. Strict conversion passes **85 presentations / 340 emitted states** against original
package/SDK diagnostics; the Strict presentation sample passed Office batch 2. Rewritten parts now
hoist fragment namespace declarations and `mc:Ignorable` to the root. The fresh namespace sample
passes PowerPoint, package/SDK and LibreOffice with unchanged slide count. Opaque parts retain their
bytes. Current full reports and exceptions are in `~/corpora/results/lossless-final-984eaf3/lectern/`.
The convention scan has 163 signatures, with no additions; removing rotation/flip attributes from
graphic-frame transforms removes three signatures from the preceding scan. Focused checks remain in
`~/corpora/results/lectern-group-2026-10-09/` and `~/corpora/results/review-fixes-2026-10-09/`.

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
| Edit notes text | Regenerate the notes page and report its formatting conversion |
| Leave a design unchanged, unused, or delete its last slide | Retain its original master, every layout and theme |
| Insert a slide using a missing layout | Generate that layout under the selected master; reserve its structural shape ID before assigning placeholders |
| Change colors or fonts | Update only changed color slots or Latin font entries; keep theme effects, other scripts and layouts |
| Edit a master/layout background, decoration or placeholder frame | Update the owning property; retain unread shapes, unrelated siblings, stacking positions and original shape identities |
| Delete, insert or reorder design decorations | Change only the modeled decorations; unread siblings stay in their source slots |
| Change master text styles, footer font or title alignment | Replace edited attributes/properties; keep unknown siblings and other levels |
| Apply colors to only some slides | Create an independent master/layout/theme family for those slides |
| Undo/redo or recover a draft | Restore the design library and its preserved properties |
| Save as a macro-free variant | Report VBA removal; cancellation leaves file identity and acknowledgements unchanged |

The 13-deck design sweep, with four targeted corrections, has 11 passing package/SDK comparisons;
two originals with missing dependencies remain failed comparisons. Sixteen focused cases / 50 saved
states pass package/SDK checks and cover geometry, theme changes, styles, independent copies,
last-slide deletion, new layouts, interleaved unread shapes, stacking order, undo/redo and actual drafts.
The corrected design-style case is included. Three earlier edited samples and the unread-decoration
regression open in LibreOffice with expected slide counts.
Office batch 7 file 13 and its synthetic source both offered repair. Removing the invented graphic
type passed PowerPoint with both decoration edits intact. The fixture now guards that test object
inside an unsupported compatibility choice with an empty fallback. Its actual writer output passes
PowerPoint, package/SDK and LibreOffice; two movement scenarios retain save/undo/redo behavior.
The checker rejects the unguarded test type, and its regression fails on the previous checker.
No production writer changed for this fixture defect. Evidence: `~/corpora/results/office-batch-7-repair-2026-10-09/`.
All 15 batch-4 samples passed PowerPoint. Reports:
`~/corpora/results/lectern-preservation-2026-10-09/designs-*` and
`~/corpora/results/review-fixes-2026-10-09/designs*`.

### Sections and custom shows: edit behavior

| Edit | Save behavior |
|---|---|
| Save unchanged slides | Keep original slide IDs, section names/IDs, empty sections and custom-show order, including repeated show entries |
| Insert or copy slides | Use fresh IDs above the source maximum; join the preceding slide's section (the following section when inserting first). Copies are not automatically added to custom shows |
| Move slides | Join the destination section and retain original slide IDs; custom shows keep their own playback order |
| Delete slides | Remove their section and custom-show references; retain empty section definitions |
| Delete every slide in a custom show | Remove that show with a recorded entry; a selected removed show switches to all slides |
| Undo/redo or recover a draft | Restore slide membership and original identities; draft recovery retains pending notices |

The 18-deck sweep and all 26 edit/history/draft states introduce no package/SDK diagnostics.
Large slide IDs and empty sections are covered by a public one-slide deck with ten sections.
Two edited samples and four history states open in LibreOffice with the expected inserted/deleted
slide counts. The browser edit driver and app load report no console errors.
The full group measurement is above; all 15 batch-4 samples passed PowerPoint. Reports:
`~/corpora/results/lectern-preservation-2026-10-09/membership-*`.

### Comments: edit behavior

| Edit | Save behavior |
|---|---|
| Text elsewhere, slide move or shape geometry | Keep comments, replies, authors and task metadata; retain unchanged comment parts byte for byte |
| Copy a slide or paste a slide from another document | Independent comment parts and identities; remap authors, per-author indices, reply links, slide/shape anchors and creation identities together |
| Duplicate only a shape | Existing comments remain on the original shape |
| Delete a commented shape or edit its anchored text | Keep the thread attached to the slide with its slide ID and creation ID; report the anchor conversion |
| Delete a slide | Remove its comment parts; retain parts owned by other slides and copied comments |
| Undo/redo or recover a draft | Restore ownership and dependencies; retain pending notices |

All 30 public comment/author inputs save and reopen; package/SDK comparison passes for 29, with one
original ZIP CRC failure explicitly retained as a failed comparison. Thirteen focused cases / 41
saved states pass package/SDK comparison and cover
legacy and Strict copies, modern replies/tasks, shape and text anchors, detached clipboard transfer,
undo/redo and actual drafts. The modern fixture is schema-authored because the pinned corpus has
modern authors but no modern comment threads; it passes the SDK's Microsoft365 target. Copies carry
only their own thread relationship, and re-anchored comments must match a declared slide creation ID.
Two earlier edited samples, four history states and the corrected text-anchor sample open in
LibreOffice with expected slide counts.
All 15 batch-4 samples passed PowerPoint. Reports:
`~/corpora/results/lectern-preservation-2026-10-09/comments-*` and
`~/corpora/results/review-fixes-2026-10-09/*comments*`.

### Opaque frames: edit behavior

| Edit | Save behavior |
|---|---|
| Text outside an opaque object | Keep its original frame and dependent parts |
| Move or resize its preview | Update the outer frame; retain each member's orientation, internal 3-D coordinates, data and effects |
| Rotate or flip a preserved preview or chart | Rotation and flip controls are disabled, including for groups containing these objects; the existing orientation remains |
| Ungroup a rotated group containing a graphic frame | Reposition the frame without adding rotation or flips; retain its original payload |
| Copy or paste the object | Independent dependent parts and remapped identities; import preview renditions, diagram cache links and OLE/VML preview identities |
| Change converted content or ungroup SmartArt | Save the edited group/picture and record the conversion; animations of its parts are removed with their own notice |
| Delete the object | Remove its original frame; do not restore it from package parts |
| Delete or copy only some members of a compatibility fallback | Save those remaining/copied members as converted content; emit the original wrapper only when all members remain intact, ordered and contiguous |
| Undo/redo or recover a draft | Restore original membership, dependencies and pending notices |

The 135 relevant decks all save and reopen. Package/SDK comparison passes for 133; two originals
have missing parts and remain failed comparisons. Twenty focused cases / 62 saved states cover
geometry, copying, clipboard previews, content edits, ungrouping, deletion and actual draft recovery,
including a schema-authored fallback with two displayed members. The two OLE identity corrections
and the missing-cache correction were rerun separately; the original measurements remain recorded.
Three orientation regressions add nine passing save/undo/redo package/SDK comparisons: disabled
preview rotation/flip and moved wrappers whose members have different rotations and flips.
The stronger single-preview fixture checks different Choice/Fallback orientations and fails on the
old frame code. Together with chart/OLE ungrouping, it adds 14 passing package/SDK comparisons.
Single-frame placeholders now retain both flips; `p:xfrm` never receives rotation/flip attributes.
Modern comment extensions retain their original sibling slots, including when a slide needs a new
creation ID; both added comment cases pass package/SDK checks. All five selected PowerPoint saves
open in LibreOffice with unchanged slide counts. A scan finds seven illegal frame transforms in
five preceding full-run saves and none in readable originals; the writer correction prevents them.
These changes are measured in the `984eaf3` refresh. Office batch 7 files 09–12 and 14 passed;
file 13's corrected master/layout fixture also passed, as described above. Evidence:
`~/corpora/results/lossless-evening-review-2026-10-09/`.
Of 959 diagram, embedding, ink and model dependency parts, 928 are byte-identical, ten are Strict
conversions, three remove missing optional cache references, and 18 are unreferenced orphan parts
in one damaged source. Five edited samples open in LibreOffice with expected slide counts.
All four pages load without console errors.
All 15 batch-4 samples passed PowerPoint. Reports:
`~/corpora/results/lectern-preservation-2026-10-09/frames-*`.

An object conversion/removal notice also accounts for its unreachable opaque dependency parts;
these no longer add generic part-loss notices beside the specific explanation. The focused SmartArt,
animated SmartArt and OLE deletion cases pass all nine saved/history package/SDK comparisons.
Unrelated orphan parts and distinct animation conversions remain reported.

### Shape properties, tags, actions and timing: edit behavior

| Edit | Save behavior |
|---|---|
| Change shadow, fill, line, crop or picture treatment | Replace that property only; retain sibling effects, 3-D and extensions |
| Change text or bold/size/alignment | Keep unknown run/paragraph/body properties, including empty formatted runs and line text bodies |
| Move or resize a group | Keep its effects and update the group/child coordinate transforms together |
| Copy shapes, slides or paste into another deck | Copy tags independently; remap actions, shape identities and dependencies; include preview bytes |
| Edit an action or delete its slide target | Replace the action or remove the dangling jump with a notice |
| Change transition speed/advance time | Update those attributes in every retained compatibility branch; preserve its original effect |
| Replace a transition | Use the chosen transition and report the conversion |
| Edit a modeled animation | Replace the modeled main sequence; retain unrelated interactive/media sequences and consistent build IDs |
| Delete an animated shape | Remove its targets and build entries without regenerating other effects |
| Undo/redo or recover a draft | Restore property fragments, tags, timing and pending notices |

Save also reports imported-chart regeneration, converted form controls/text fields and incomplete
source-part recovery. These notices are recomputed from the prepared file. Cancelling a conversion
does not leave its notice on a later preserving save.

The 214-deck sweep plus targeted corrections saves/reopens every input; 212 package/SDK comparisons
pass and two damaged originals remain failed comparisons. Eighteen focused cases / 56 emitted
states pass, including text edits, geometry, copies, deletion, detached clipboard and actual drafts.
Four edited samples open in LibreOffice with unchanged counts. UTF-16 package XML is decoded
correctly, retaining a six-slide template previously read as empty; its save/edit/history states
also pass SDK and LibreOffice. The corpus driver now compares original, opened and saved slide
counts. Four pages load without console errors. All 15 batch-4 samples passed PowerPoint. Reports: `~/corpora/results/lectern-preservation-2026-10-09/properties-*`.

## Known limits

* Legacy binary `.ppt` files must be re-saved as `.pptx` first. Macro-enabled OOXML variants keep their VBA; see *What a save keeps, converts and drops* for the remaining preservation gaps.
* SmartArt is displayed as grouped shapes and keeps its original frame until the group's content is edited or ungrouped. When a file carries only diagram data, Lectern draws it with a built-in layout family, so unusual layouts can look simpler in the editor.
* 3-D charts are drawn in a fixed oblique view (no perspective, no lighting); surface charts, trendlines and error bars are not drawn (they are kept in unedited charts).
* Embedded OLE objects and ActiveX controls import as their preview pictures.
* Fonts that aren't installed fall back to metric-compatible ones: Carlito, Caladea and Gelasio (for Calibri, Cambria and Georgia) are served with the site, and Save as Web Page embeds the ones its slides name; Arial, Times New Roman and Courier New use the computer's own copies or the Liberation/Arimo/Tinos/Cousine fonts of the same widths where installed.
* Spelling uses the browser's spell checker.
