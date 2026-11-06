# Better PDF export: a text layer in Quire, Ledger and Lectern PDFs

Status: Phase 0 ready for Lead, 2026-11-07; Phases 1–5 pending.
Planned 2026-10-15; design revised that day: pages stay pictures, with an invisible text layer on
top; no font files read or embedded. Read first: this file, `public/common/core.js` (`L.buildPDF`, `L.embedFonts`),
Quire `public/quire/js/app.js` (`A.rasterizePage`, `A.buildPDF`), Lectern `public/lectern/js/app.js` (`A.rasterize`,
`A.buildPDF`), Ledger `public/ledger/js/print.js` (`P.renderPage`, `P.exportPDF`). Work the phases in order and mark
each task when it is done and checked.

## Where we are (2026-10-15)

All three apps make the same kind of PDF: each page is drawn into a canvas, saved as a JPEG (quality 0.92), and
`L.buildPDF` (core.js:226) writes one full-page image per page. There's no text in the PDF.

| App | How a page becomes a picture | Resolution |
|---|---|---|
| Quire | `A.rasterizePage` (app.js:887): the page's DOM cloned into an SVG `foreignObject`, drawn into a canvas | 2.2 × CSS px, about 211 dpi |
| Lectern | `A.rasterize` (app.js:493): the same `foreignObject` approach, for slides and print-preview pages | 2.5 px per pt, 180 dpi |
| Ledger | `P.renderPage` (print.js:193): cells drawn on a canvas with `fillText` etc.; charts and shapes are SVGs drawn into it as images | 2 px per pt, 144 dpi |

What this costs the user:
- Text can't be found, selected or copied.
- Screen readers get nothing.
- There are no links and no bookmarks.

**Defect, measured 2026-10-15:** pictures made through an SVG image can't use the fonts we ship in `public/fonts/`.
In Chromium 150, text set in Carlito inside the SVG produced the same pixels as a font that doesn't exist. An
installed font (DejaVu Sans) did render. So on a computer without Calibri, a Calibri document is laid out on screen
in Carlito but drawn in the PDF in a fallback font, and its lines can wrap differently there. This affects:
- Quire's and Lectern's pages;
- the text in Ledger's charts and shapes;
- Quire's Recent Files thumbnails (app.js:1366), which use the same rasterizer.

## Goal

Every PDF from the three apps:
- looks as it does today, with the right fonts and at 300 dpi, a common printer's resolution, in all three apps
  (Phase 0);
- has every word on the page in an invisible text layer, so it can be found, selected and copied (with spaces between
  words and line breaks between lines, in reading order), and screen readers get the text;
- has links that click, and bookmarks;
- has no print dialog and nothing from outside the machine;
- produces the same bytes for the same pictures and text.

## Design (decided)

- **We write the PDF ourselves**, never through the browser's own print (Decisions).
- **A page is today's picture, unchanged, plus an invisible text layer.**
  - The picture still holds everything visible, text included. Whatever the browser can draw comes out as it does
    now: every script, every font, equations, LettersArt, vertical text.
  - Over it, each word is written with text render mode 3 (invisible) at its place. A viewer finds and selects it
    there but draws nothing, so nothing on the page can be doubled or lost.
- **No font files.** The layer's text uses one font we generate ourselves in `pdf.js`: a TrueType font of a few hundred
  bytes, written from the OpenType and PDF specifications, with no glyph shapes.
  - Every character shows the same empty glyph with a fixed advance (half an em).
  - The PDF font is a Type0 font with `Identity-H`. Each distinct character in the document gets its own CID, and a
    `CIDToGIDMap` stream sends every CID to the empty glyph.
  - A ToUnicode CMap maps each CID to its character, as a full code point, so emoji and rare CJK characters copy
    correctly.
  - The font's Ascent is 800 and its Descent −200. With the font size set to the word's box height and the baseline
    at 80% of the box, a viewer's selection highlight covers exactly the word's box.
- **Word placement.** Each word is one text object at the word's box: a `Tm` at its left edge and baseline, the font
  size equal to the box height, and horizontal scaling (`Tz`) so its characters fill exactly the box's width.
  - Between words on a line, a space character is written to fill the gap, so copied text has spaces in every
    viewer.
  - Lines come one after another in reading order.
  - Rotated text (Ledger's rotated cells, Lectern's rotated boxes) takes the rotation in `Tm` where the app knows it.
    Elsewhere, the word's axis-aligned box is used: it's still found and copied, only its highlight is approximate.
- **Positions come from what the apps already draw:**
  - Quire and Lectern: the live page's DOM ranges.
  - Ledger: the canvas renderer's own `fillText` calls.
- **What apps hand to the writer.** `L.buildPDF(pages, info)` keeps its name; pages and info gain optional fields:
  ```js
  pages: [{ jpeg, iw, ih, pw, ph,                            // as today
            lines: [[{ text, x, y, w, h, angle? }, ...], ...],  // words per line, in reading order; page points, top-left origin
            links: [{ x, y, w, h, url } | { x, y, w, h, page, top }] }]
  info:  { title, author, lang, outline: [{ title, page, top, level }] }
  ```
- **Shared code** lives in `public/common/` (names provisional):
  - `pdf.js`: the writer, with the generated font. `L.buildPDF` moves there from core.js, or core.js calls it.
  - `pdftext.js`: collects words from a DOM page, and records `fillText` from a canvas.
  - Both load on the first export (a script tag added then), not at app start. The service worker's site zip
    includes them by itself.

### Which text goes into the layer

Only text a reader can see on the page:
- Skip formatting marks (`.format-marks`), section marks, page-break marks, crop marks, balloons, and anything with
  no rectangle, `visibility: hidden` or `display: none`.
- Text clipped away by an ancestor (`overflow` not visible: a cell with an exact row height, an overflowing text
  box): only the words inside the ancestor's box. Otherwise a search would find text no one can see.
- Text drawn by CSS rather than a text node (tab leaders from `::before`, list labels if they're pseudo-elements)
  isn't text a reader would copy; leave it out.
- SVG text (charts, LettersArt, IntelliArt) and MathML go in where `Range.getClientRects` gives boxes for them.
  Check this in Phase 2; if it doesn't, they stay picture-only for now.

## Tasks

Phase 0: the page pictures (the right fonts, 300 dpi)
- [x] Add `L.embedFonts` rules (core.js:512; Lectern's HTML export already uses it, lectern app.js:556) to the CSS
      that Quire's `A.rasterizePage` and Lectern's `A.rasterize` put inside the SVG. Build it once per export from the
      fonts the pages name, not once per page.
- [x] Do the same for the chart and shape SVGs Ledger draws into its print canvas (print.js:244, print.js:313), as a
      `<style>` inside each SVG.
- [x] Check: Carlito text inside a picture differs from the fallback font (the probe from 2026-10-15: the same text
      drawn through an SVG image in 'Carlito' and in a missing font must give different pixels). Export a Calibri
      document on a machine without Calibri and compare the PDF page with the screen.
- [x] Time a long document (181 pages) before and after; the base64 fonts are repeated in every page's SVG.
- [x] **300 dpi in all three apps**, a common printer's resolution. Text stays a picture, so its sharpness is the
      picture's; today it's 211 dpi in Quire, 180 in Lectern and 144 in Ledger.
  1. One constant where `L.buildPDF` lives (`PDF_DPI = 300`). Each app derives its scale from it:
     - Quire: `A.rasterizePage(p.el, 300 / 96)`, instead of 2.2;
     - Lectern: `300 / 72` pixels per point for slides and preview pages, instead of 2.5;
     - Ledger: `renderAny(…, 300 / 72)` in `P.exportPDF`, instead of 2.
  2. **A cap for very large pages:** lower the resolution for a page whose picture would pass 40 megapixels (a Letter
     page at 300 dpi is 8.4; A3 is 17; A0 would be 139). Chromium limits a canvas's area, and each picture is held
     in memory four bytes per pixel while it's encoded.
  3. Recent Files thumbnails keep their own small scale.
  4. **Check:** export time and file size before and after on the 181-page document, five decks and five workbooks.
     JPEG size grows with the pixels: about ×2 for Quire, ×2.8 for Lectern and ×4.3 for Ledger. If that's too much,
     lower the JPEG quality (0.92 now) for text-only pages before lowering the resolution. Compare a printed or
     zoomed page with today's.

### Phase 0 report — 2026-11-07

10 regression checks fail on 5422906 and pass now. All three apps export at 300 dpi;
actual A0 exports stay under 40 MP; Quire's Recent thumbnail remains 400 × 518.
The Calibri control (no Calibri installed) matches the screen's Carlito text pixels
and line breaks; the old picture adds a line. Font CSS is prepared once per PDF;
Ledger releases the SVG images afterward. JPEG quality stays 0.92; API/load order unchanged.

| Controls | Pages | Seconds before → after | MB before → after |
|---|---:|---:|---:|
| Long Word control (formerly 181 pages) | 180 | 10.57 → 21.69 | 83.77 → 134.45 |
| Five corpus decks | 15 | 0.95 → 2.87 | 1.62 → 3.13 |
| Five corpus workbooks | 16 | 0.57 → 1.01 | 0.78 → 2.34 |

12 of 13 originals open in Office 2016 without repair and in LibreOffice; the extra
Divino_Revelado has a corrupt ZIP directory and both refuse it. 11 of the 12 accepted
inputs match the suite's page count; bug65649 remains Quire 180 / Word 177 / LO 178.

Limit: 180 repeated Calibri pages with all four faces take 112.67 s (was 6.90 s),
59.09 MB (was 39.62 MB); 1.06 MB of font CSS is parsed/rasterized per image.
Limit: the known Word pagination gap and Lectern rotated-table clipping remain.

Final `tools/quick.mjs` summaries, with a private English gather and catalogs blocked:

```text
✖ Node tests: 190 passed, 1 failed (414 s)
✔ tests/ui/smoke.test.mjs: 4 passed (18 s)
✔ tests/ui/lectern.test.mjs: 39 passed (127 s)
✖ tests/ui/ledger.test.mjs: 38 passed, 1 failed (124 s)
✔ tests/ui/quire.test.mjs: 57 passed (206 s)
2 of 5 failed.
```

The two failures reproduce before these changes: translated locale labels are
unavailable under the no-catalog-access rule; the Ledger toolbar test clicks a
hidden Drawing toolbar. They are recorded, not counted as passes.

[Detailed report](../tmp/PDF-0-REPORT.md), [Lead findings](../tmp/PDF-LEAD-FINDINGS.md);
evidence in `tmp/pdf/artifacts` on the share. No commits or edits to Lead's files.

**PDF.0 READY FOR LEAD**

Phase 1: the writer (`public/common/pdf.js`)
- [ ] **Objects.**
  - Catalog, a pages tree, pages, content streams, and Info (Title, Author, Producer `L.APP`, CreationDate).
  - `/ID` from a SHA-256 of the content (`common/sha.js`), so the same input gives the same bytes; the date is
    fixable for tests.
  - Header `%PDF-1.7`.
  - Streams compressed with `CompressionStream('deflate')` (zlib, as FlateDecode wants), left uncompressed where
    there's none. This makes `L.buildPDF` async: Quire and Lectern already return it from async functions, and
    Ledger's `P.exportPDF` (print.js:341) must `await` it.
- [ ] **Page content.** The JPEG as today (DCTDecode, bytes copied as they are), then the text layer:
      `BT 3 Tr /F1 … Tf` and per word `Tz … Tm <cids> Tj`, then `ET`.
- [ ] **The generated font.**
  - The TrueType bytes: head, hhea, maxp, hmtx, loca, glyf (glyph 0 and one empty glyph), cmap, name, OS/2 and post.
  - The Type0 font, the CIDFontType2 with `DW 500`, the `CIDToGIDMap` stream, and the FontDescriptor (Ascent 800,
    Descent −200, flags non-symbolic).
  - The ToUnicode CMap, written as ranges where CIDs and code points run together.
- [ ] **Links and bookmarks** (data from the apps in Phase 5):
  - `/Link` annotations with a `/URI` action, or a destination `[page /XYZ left top null]`;
  - an `/Outlines` tree nested by level;
  - `/Lang` in the catalog;
  - `/ViewerPreferences << /DisplayDocTitle true >>` when there's a title.
- [ ] **Tests** (`tests/pdf.test.mjs`, Node):
  - Two pages with a JPEG each and lines of words, including Japanese, an emoji, right-to-left text and a rotated
    word.
  - pdftotext gives the words with single spaces and the lines in order.
  - pdffonts lists one embedded font with ToUnicode.
  - Every xref offset points at its `obj`, and the same input gives the same bytes.
  - pdftoppm draws the page with the layer pixel for pixel like the page without it: the layer is invisible.
  - Then once by hand in Chromium's viewer: select a line and copy it, and Find locates a word on page 2 with its
    highlight on the word.

Phase 2: Quire
- [ ] **Collector** (`pdftext.js`, `collectDOM(pageEl, scale)`).
  - Walk the page's text nodes in document order, which is Quire's reading order: body, then columns, text boxes,
    footnotes and headers and footers as the page holds them. Apply the skip rules above.
  - Find the line breaks inside each node with `Range.getClientRects` and the binary search `LY.lines` uses
    (layout.js:105).
  - Split each line piece into words at spaces, and take one rectangle per word. CJK has no spaces, so one box per
    line piece, with its characters spread evenly over it.
  - Group the words into lines by their vertical position.
  - Positions are relative to the page element, divided by `LY.zoom`, then turned into points.
- [ ] **Output.** `A.buildPDF` collects each page's lines before rasterizing it and passes `lines` with the JPEG.
      The picture is unchanged.
- [ ] **Checks.** The sample document plus about ten corpus files of different kinds: tables, columns, headers and
      footers, footnotes, lists, East Asian on a grid, right-to-left, equations, LettersArt, rotated text boxes, and
      a table cell with clipped text.
  - pdftotext word-diffed against the document's text.
  - Rendered pages identical to today's export.
  - Export time on the 181-page document: the collector should add little to the rasterizing.
  - One Chromium test (`tests/quire/`): export the sample; its text extracts with spaces and line breaks.
- [ ] **For the user to check:** Acrobat Reader and macOS Preview, a handful of files, with a checklist: Find finds
      words; selecting a paragraph highlights its lines; copying gives the text with spaces; nothing visible changed.

Phase 3: Lectern
- [ ] **Collector on slides.** `A.slideCanvas` renders a slide element that isn't in the document, so it has no
      rectangles. Attach it off-screen (fixed position, off the viewport) while collecting, then detach it.
- [ ] **Transforms.** Autofit and thumbnail scaling are already in the rectangles. For rotated boxes, pass the
      rotation (`angle`, from the `DOMMatrix` of the ancestors' computed `transform`) with the box's words.
- [ ] **Print-preview pages** (`.pv-page`: handouts, notes pages, outline) are already in the document: collect them
      as Quire does.
- [ ] **Checks:** five corpus decks, as in Phase 2, plus slide size and page count unchanged.

Phase 4: Ledger
- [ ] **A recording 2D context** (`pdftext.js`, `recordCanvas(ctx)`): a proxy that forwards every call to the real
      canvas, `fillText` included, so the picture is unchanged. For each `fillText` it records:
  - the text, `ctx.font` size, `textAlign` and `textBaseline`;
  - `getTransform()` (rotated cells, render.js:731, give the `angle`);
  - the current clip: rectangles from `rect()` plus `clip()` through the `save`/`restore` stack (cell clipping,
    render.js:649 and 518). Only the part inside the clip goes into the layer;
  - `maxWidth`, which squeezes text to fit.

  The width comes from `measureText`. Use it in `P.renderPage` (print.js:193), headers and footers included.
- [ ] **Reading order:** row by row, cells left to right, as the renderer draws them; check that it does.
- [ ] **Checks:** a workbook with numbers, wrapped text, overflowing text, clipped text, rotated text, shrink-to-fit,
      and headers and footers. pdftotext gives the cell values; Find in a viewer highlights the right cell; clipped
      text is found only as far as it shows.

Phase 5: links, bookmarks, language
- [ ] **Quire:**
  - hyperlinks become link areas from their rectangles;
  - internal links and cross-references go to their target's page and position;
  - bookmarks come from headings (outline level, page and position from `LY.pages` and the fragment's rectangle),
    nested by level;
  - `lang` comes from the document's language.
- [ ] **Lectern:** links on text and shapes (URL, or a jump to another slide); bookmarks from slide titles.
- [ ] **Ledger:** cell hyperlinks; bookmarks from sheet names.
- [ ] **Checks:** links click through in Chromium's viewer and in Firefox, and the bookmarks show in their side panes.

When it's done
- [ ] Docs:
  - docs/quire.md, docs/ledger.md and docs/lectern.md: the PDF row (pages as pictures, with text you can find, select
    and copy, links and bookmarks);
  - docs/suite.md: the shared `pdf.js` and `pdftext.js`;
  - AGENTS.md: the list of `public/common/` libraries;
  - docs/testing.md: the PDF checks.
- [ ] Each app loaded and exported once after the change in common code (AGENTS.md: a common change is checked in
      all three apps), and `node tools/quick.mjs`.


## Not in this plan

- **The browser's own print, in any form:** `window.print()`, print stylesheets, CDP `Page.printToPDF`, the desktop
  web views' print-to-PDF. Not for export and not for checks. Today Lectern's print preview has a Print button that
  calls `window.print()` (panes.js:418), and Quire and Ledger have print stylesheets; this plan neither uses nor
  fixes them.
- **Visible real text with embedded fonts** (decided against, 2026-10-15). It would make text sharp at any zoom and
  files smaller, and needs:
  - reading font files (installed ones only through Local Font Access in Chrome and Edge; our WOFF2 files decoded);
  - subsetting, and CFF fonts;
  - licence bits;
  - a per-character font order, since PDF has no fallback list.

  Reopen only if sharpness or file size becomes a complaint. Picture improvements that need the text out of the
  picture (no image for blank pages, cropping, lossless pictures) would come with it.
- Quire on a canvas, and our own layout engine ("Project Re:Layout", 2026-10-15): not planned; Quire stays on the
  DOM. Reopen only if most of the corpus files whose page count differs from Word's turn out to differ because of
  line breaking CSS can't express.
- **Later, if wanted:** vector borders, shading, shapes and charts; tagged PDF; PDF/A; the source `.docx` embedded on
  request.
- **Belongs to the layout plan, not this one** (todos/QUIRE_LAYOUT_ALIGNMENT.md, not added there yet):
  - Quire kerns all text (`.pc { font-kerning: normal }`, quire index.html:94) and only turns kerning off for
    `w:kern` 0 (render.js:175).
  - Word kerns only from `w:kern`'s size and uses no ligatures by default.
  - In Carlito, ordinary prose is 0.56% narrower with Chromium's kerning and ligatures than with plain advance widths:
    about 2.6pt on a 6.5-inch line.
  - Worth one page-count run over the corpus with kerning and ligatures off.

## Decisions (the user, 2026-10-15)

- Write the PDF ourselves; Chromium's print-to-PDF has behaviour we don't control (headers and footers, among
  others).
- Don't use the browser's own print at all, for the export or to check it.
- One exporter for all three apps.
- Quire probably won't move to a canvas: do what's needed now, on the DOM as it is.
- No fonts: pages stay pictures, with an invisible text layer on top.
- Pictures at a common printer's resolution, 300 dpi, in all three apps.
