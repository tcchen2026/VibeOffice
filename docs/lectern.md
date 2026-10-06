# Lectern 2003 — Web Edition

A presentation editor modeled on PowerPoint 2003 (Luna Blue), written in plain HTML and vanilla JavaScript with no frameworks and no build step. It opens and saves real `.pptx` files.

The app is `public/lectern/` (`python3 tools/serve.py`, then http://127.0.0.1:8760/lectern/); every path below is relative to it. It runs in any modern browser (Chrome, Edge, Firefox, Safari). Everything runs locally; nothing is uploaded.

## What's in the box

| File | Responsibility |
|---|---|
| `index.html` | Window chrome, Luna CSS, slide rendering styles, script order |
| `js/core.js` | Utilities, colors, fonts, file I/O, media store, PDF writer |
| `js/zip.js` | ZIP reader/writer (CompressionStream with a pure-JS inflate fallback) |
| `js/geometry.js` | 92 AutoShape presets, adjust handles, text rectangles, connectors |
| `js/model.js` | Presentation model, 14 design templates, 21 slide layouts, undo history |
| `js/render.js` | DOM/SVG renderer for slides, fills, patterns, gradients, WordArt |
| `js/metafile.js` | WMF/EMF player, so old clip art, logos and pasted Excel/Visio graphics display |
| `js/charts.js` | 12 chart types (column, bar, line, area, pie, doughnut, scatter) with a datasheet |
| `js/ui.js` | Menus, command bars, combo boxes, dialogs, color pickers, tooltips |
| `js/editor.js` | Slide editing surface: select, move, resize, rotate, crop, draw, grid |
| `js/textedit.js` | Rich text editing, bullets, levels, AutoCorrect, tables |
| `js/dialogs.js` | Font, Bullets, Format AutoShape, Fill Effects, Header & Footer, etc. |
| `js/panes.js` | Slides/Outline tab, notes, sorter, notes page, print preview, task panes |
| `js/slideshow.js` | Full-screen show, 58 transitions, 50+ entrance/emphasis/exit/motion-path effects, pen, rehearse timings |
| `js/pptx-read.js` | PresentationML import (themes, masters, layouts, placeholders, table styles, charts, custom geometry, animations; repairs damaged packages) |
| `js/pptx-write.js` | PresentationML export (validated against the OOXML schema) |
| `js/commands.js` | Every menu/toolbar command |
| `js/app.js` | Application wiring, keyboard shortcuts, clipboard, sample deck |

## Opening

File ▸ Open, drag a file onto the window, or *Open a presentation* in the Getting Started pane. Tested against the 95 decks in Apache POI's real-world test corpus: 94 open (the remaining one has no presentation part at all), slide counts match python-pptx, and every re-saved copy passes the OOXML schema validator. Damaged packages are repaired where possible; password-protected and binary .ppt files get a clear message.

## Slide size

File ▸ Page Setup. Choose *Widescreen (16:9)* (13.33 × 7.5 in, today's PowerPoint default) or *On-screen Show (16:9)* (10 × 5.63 in). Pictures, shapes, charts and tables keep their proportions; placeholders reflow to the new frame. Tick *Use for new presentations* to make it the default.

## Saving

* **.pptx** — standard PresentationML. Exports pass the OOXML schema validator and open in LibreOffice and python-pptx; they are meant for PowerPoint 2007 and later, Keynote and Google Slides.
* **.ppsx / .potx** — show and template variants (standalone page only).
* **PDF** — one slide per page, or exactly what Print Preview shows (handouts, notes pages, outline).
* **Single-file web page**, **PNG** of the current slide, **plain-text outline**.

## Known limits

* Legacy binary `.ppt` files must be re-saved as `.pptx` first.
* SmartArt imports as grouped shapes; embedded OLE objects import as their preview picture.
* Fonts that aren't installed fall back to metric-compatible web fonts (Arimo, Tinos, Cousine, Carlito, Caladea).
* Spelling uses the browser's spell checker.
