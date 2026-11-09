# Better PDF export: a text layer in Quire, Ledger and Lectern PDFs

Status: Phases 0–4 committed (705019a, ca1d7cd, 1ed6acb, 31a03dd, eb03301); Phase 5 complete and reviewed by Lead, including the final URI filter.
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
            lines: [[{ text, x, y, w, h, angle?, visual? }, ...], ...],  // words in reading order; page points, top-left origin
            links: [{ x, y, w, h, url } | { x, y, w, h, page, top }] }]
  info:  { title, author, lang, outline: [{ title, page, top, level }] }
  ```
  `visual` optionally supplies the DOM's resolved visual character order for a mixed bidi word;
  `text` retains its logical text. Callers without DOM geometry keep the writer's existing fallback.
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
- [x] **Objects.**
  - Catalog, a pages tree, pages, content streams, and Info (Title, Author, Producer `L.APP`, CreationDate).
  - `/ID` from a SHA-256 of the content (`common/sha.js`), so the same input gives the same bytes; the date is
    fixable for tests.
  - Header `%PDF-1.7`.
  - Streams compressed with `CompressionStream('deflate')` (zlib, as FlateDecode wants), left uncompressed where
    there's none. This makes `L.buildPDF` async: Quire and Lectern already return it from async functions, and
    Ledger's `P.exportPDF` (print.js:341) must `await` it.
- [x] **Page content.** The JPEG as today (DCTDecode, bytes copied as they are), then the text layer:
      `BT 3 Tr /F1 … Tf` and per word `Tz … Tm <cids> Tj`, then `ET`.
- [x] **The generated font.**
  - The TrueType bytes: head, hhea, maxp, hmtx, loca, glyf (glyph 0 and one empty glyph), cmap, name, OS/2 and post.
  - The Type0 font, the CIDFontType2 with `DW 500`, the `CIDToGIDMap` stream, and the FontDescriptor (Ascent 800,
    Descent −200, flags non-symbolic).
  - The ToUnicode CMap, written as ranges where CIDs and code points run together.
- [x] **Links and bookmarks** (data from the apps in Phase 5):
  - `/Link` annotations with a `/URI` action, or a destination `[page /XYZ left top null]`;
  - an `/Outlines` tree nested by level;
  - `/Lang` in the catalog;
  - `/ViewerPreferences << /DisplayDocTitle true >>` when there's a title.
- [x] **Tests** (`tests/pdf.test.mjs`, Node):
  - Two pages with a JPEG each and lines of words, including Japanese, an emoji, right-to-left text and a rotated
    word.
  - pdftotext gives the words with single spaces and the lines in order.
  - pdffonts lists one embedded font with ToUnicode.
  - Every xref offset points at its `obj`, and the same input gives the same bytes.
  - pdftoppm draws the page with the layer pixel for pixel like the page without it: the layer is invisible.
  - Then once by hand in Chromium's viewer: select a line and copy it, and Find locates a word on page 2 with its
    highlight on the word.

### Phase 1 report — 2026-11-07

26 PDF checks pass (12 writer, 14 raster). 211 / 211 corpus page JPEGs remain
byte-identical. 10 writer and 3 batching/encoding assertions fail before the change;
the SVG-reference isolation test also fails without its guard.

| Control | Pages | Before PDF.0 s | PDF.0 s¹ | PDF.1 s | MB PDF.0 → PDF.1 |
|---|---:|---:|---:|---:|---:|
| Long Word control | 180 | 10.57 | 18.45 | 15.01 | 134.45 → 134.45 |
| Five decks | 15 | 0.95 | 2.97 | 2.08 | 3.15 → 3.15 |
| Five workbooks | 16 | 0.57 | 0.98 | 0.95 | 2.34 → 2.34 |
| Calibri, four faces | 180 | 6.90 | 117.82 | 8.83 | 59.09 → 59.09 |

¹ Remeasured with the committed export path and current rendering dependencies.

1.28 × the original Calibri time, below the 3 × target. Peak renderer RSS falls
from 1,071 MB to 553 MB; 180 different pages take 13.97 s / 845 MB. Four-page SVG
batches reach 33.66 MP; actual A0 exports in all apps stay below 40 MP. Immediate
JPEG encoding avoids Chromium's idle callback delay; quality remains 0.92.

The asynchronous writer loads on first export; Ledger awaits it. Its generated
empty glyph retraces a line for measurable bounds: a zero-contour glyph lost
single-character and vertical words in PDFium. PDFium and Poppler confirm zero
ink even in visible mode. Chromium copies the line, retains Unicode/rotation and
highlights the page-2 search word; LibreOffice opens the two-page control. Existing
Office/LO references and the two Lead-owned layout findings are retained.

Review correction: the scanner now treats pdf.js as a writer: PDF/font syntax is
excluded and explicit L.t errors remain. The new regression fails before and
passes after. The requested real English gather adds only the exporter-load
message, plus refreshed locations; all 8 i18n checks pass against that catalog.

Final `tools/quick.mjs`, after the requested English gather and scanner correction:

```text
✖ Node tests: 219 passed, 1 failed (29 s)
✔ tests/ui/smoke.test.mjs: 4 passed (18 s)
✔ tests/ui/lectern.test.mjs: 39 passed (125 s)
✔ tests/ui/ledger.test.mjs: 39 passed (130 s)
✔ tests/ui/quire.test.mjs: 57 passed (227 s)
1 of 5 failed.
```

Limit: the existing locale assertion requires blocked catalog data; it is not counted as a pass.
Limit: this headless viewer uses native text-fragment search; it has no keyboard Find bar.
Limit: app text collectors remain for Phases 2–4; current app exports still contain pictures only.

[Detailed report](../tmp/PDF-1-REPORT.md), [findings](../tmp/PDF-LEAD-FINDINGS.md).
Evidence: `tmp/pdf/artifacts/phase1`; final logs `tmp/pdf/phase1-correction-*.log`.
No commits or translation edits; the English catalog was gathered as requested.
All final code/test/catalog hashes stayed unchanged during the gate; all 11 original
tabs remain and server 8769 is stopped. Phase files are frozen until committed.

**PDF.1 READY FOR LEAD**

Phase 2: Quire
- [x] **Collector** (`pdftext.js`, `collectDOM(pageEl, scale)`).
  - Walk the page's text nodes in document order, which is Quire's reading order: body, then columns, text boxes,
    footnotes and headers and footers as the page holds them. Apply the skip rules above.
  - Find the line breaks inside each node with `Range.getClientRects` and the binary search `LY.lines` uses
    (layout.js:105).
  - Split each line piece into words at spaces, and take one rectangle per word. CJK has no spaces, so one box per
    line piece, with its characters spread evenly over it.
  - Group the words into lines by their vertical position.
  - Positions are relative to the page element, divided by `LY.zoom`, then turned into points.
- [x] **Output.** `A.buildPDF` collects each page's lines before rasterizing it and passes `lines` with the JPEG.
      The picture is unchanged.
- [x] **Checks.** The sample document plus about ten corpus files of different kinds: tables, columns, headers and
      footers, footnotes, lists, East Asian on a grid, right-to-left, equations, LettersArt, rotated text boxes, and
      a table cell with clipped text.
  - pdftotext word-diffed against the document's text.
  - Rendered pages identical to today's export.
  - Export time on the 181-page document: the collector should add little to the rasterizing.
  - One Chromium test (`tests/quire/`): export the sample; its text extracts with spaces and line breaks.
- [x] **For the user to check:** Acrobat Reader and macOS Preview, a handful of files, with a checklist: Find finds
      words; selecting a paragraph highlights its lines; copying gives the text with spaces; nothing visible changed.
      [Six PDF pairs and checklist prepared](../tmp/pdf/artifacts/phase2/reader-review.zip); these two viewers remain unchecked.

### Phase 2 report — 2026-11-07

194 / 194 page JPEGs and rendered pages are unchanged across 14 controls (12 corpus,
sample, clipped cell). All 9 new Quire tests fail on ca1d7cd; 35 PDF checks pass now.
The final focused gate, including licences and i18n after English gather:

```text
ℹ tests 46
ℹ pass 46
ℹ fail 0
```

| Controls | Pages | Seconds before → after | Collector seconds | MB before → after |
|---|---:|---:|---:|---:|
| Long Word control | 180 | 16.22 → 17.14 | 0.72 | 134.45 → 135.40 |
| Other 13 | 14 | 1.72 → 2.03 | 0.03 | 2.89 → 2.93 |

14 original controls open in Office 2016 without repair and in LibreOffice;
Office input hashes match, including the reused long reference. Word differences
against source XML and native PDFs are retained. Chromium copies a selected line,
finds the page-2 phrase and copies DrawingML warped text once. All three apps load
and export without errors or external requests.

`L.pdfModule(true)` lazily loads the DOM collector. Quire supplies lines before
rasterization; optional `word.visual` retains browser-resolved mixed bidi order.
No app script-order change or new UI strings; English gather completed.

Limit: older VML LettersArt has no unsliced text body; renderer finding left for Lead.
Limit: equations copy linearly; rotated/warped highlights approximate the picture.
Limit: Poppler joins one narrow space in the long control; PDFium preserves it.
Limit: Acrobat/Preview checks remain for the user; the review packet is prepared.

Final `node tools/quick.mjs`:

```text
✖ Node tests: 202 passed, 1 failed (26 s)
✔ tests/ui/smoke.test.mjs: 4 passed (18 s)
✔ tests/ui/quire.test.mjs: 57 passed (207 s)
1 of 3 failed.
```

The existing locale-label assertion needs blocked translation catalogs; it is not
counted as a pass. Own code/test/doc hashes stayed unchanged through the gate.
All 11 original tabs remain; server 8769 is stopped. No commits or Lead-file edits.

[Detailed report](../tmp/PDF-2-REPORT.md), [Lead finding](../tmp/PDF-LEAD-FINDINGS.md).
Phase files are frozen until committed.

**PDF.2 READY FOR LEAD**

Phase 3: Lectern
- [x] **Collector on slides.** `A.buildPDF` renders a slide element that isn't in the document, so it has no
      rectangles. Attach it off-screen (fixed position, off the viewport) while collecting, then detach it.
- [x] **Transforms.** Autofit and thumbnail scaling are already in the rectangles. For rotated boxes, pass the
      rotation (`angle`, from the `DOMMatrix` of the ancestors' computed `transform`) with the box's words.
- [x] **Print-preview pages** (`.pv-page`: handouts, notes pages, outline) are already in the document: collect them
      as Quire does.
- [x] **Checks:** five corpus decks, as in Phase 2, plus slide size and page count unchanged.

### Phase 3 report — 2026-11-07

19 / 19 page JPEGs and rendered pages are unchanged: eight corpus decks,
six preview exports and the rotation fixture. Page counts and sizes match.
264 / 264 collected corpus tokens extract through Poppler and PDFium.
All six new Lectern tests fail with 1ed6acb's export modules. Final focused gate,
including licences and i18n after a private English gather:

```text
ℹ tests 52
ℹ pass 52
ℹ fail 0
```

| Controls | Pages | Seconds before → after | Collector seconds | MB before → after |
|---|---:|---:|---:|---:|
| Eight corpus decks | 9 | 1.542 → 1.519 | 0.029 | 1.419 → 1.439 |
| Handouts, notes, outline | 8 | 1.008 → 1.055 | 0.019 | 1.074 → 1.091 |

8 / 8 inputs open in Office 2016 without repair and in LibreOffice, including
Mac producers and a template (submitted as a plain PPTX copy to the Office API).
Input hashes match. Chromium finds and mouse-copies the 45-degree phrase with
its space, and copies LettersArt once. All three apps load/export without errors.

`collectDOM` gains an opt-in `{ transforms: true }` third argument for Lectern;
Quire keeps its existing path. No script-order change, new UI text or catalog edits.

Limit: Poppler raw extraction splits diagonal letters; PDFium keeps the phrase.
Limit: warped text, SVG-only rotation, skew/reflection and 3D CSS have approximate bounds.
Limit: Acrobat/Preview remain unchecked; [six PDF pairs and checklist are ready](../tmp/pdf/artifacts/phase3/reader-review.zip).

Final `node tools/quick.mjs`:

```text
✖ Node tests: 198 passed, 1 failed (25 s)
✔ tests/ui/smoke.test.mjs: 4 passed (18 s)
✔ tests/ui/lectern.test.mjs: 39 passed (127 s)
✔ tests/ui/quire.test.mjs: 57 passed (210 s)
1 of 4 failed.
```

The sole failure is the existing locale-label assertion requiring blocked
translation catalogs; it is not counted as a pass. Own hashes stayed unchanged
through the gate. All 11 original tabs remain; server 8769 is stopped. No commits.
[Report](../tmp/PDF-3-REPORT.md), [renderer findings](../tmp/PDF-LEAD-FINDINGS.md).
Phase files are frozen until committed.

**PDF.3 READY FOR LEAD**

Phase 4: Ledger
- [x] **A recording 2D context** (`pdftext.js`, `recordCanvas(ctx)`): a proxy that forwards every call to the real
      canvas, `fillText` included, so the picture is unchanged. For each `fillText` it records:
  - the text, `ctx.font` size, `textAlign` and `textBaseline`;
  - `getTransform()` (rotated cells, render.js:731, give the `angle`);
  - the current clip: rectangles from `rect()` plus `clip()` through the `save`/`restore` stack (cell clipping,
    render.js:649 and 518). Only the part inside the clip goes into the layer;
  - `maxWidth`, which squeezes text to fit.

  The width comes from `measureText`. Use it in `P.renderPage` (print.js:193), headers and footers included.
- [x] **Reading order:** row by row, cells left to right, as the renderer draws them; check that it does.
- [x] **Checks:** a workbook with numbers, wrapped text, overflowing text, clipped text, rotated text, shrink-to-fit,
      and headers and footers. pdftotext gives the cell values; Find in a viewer highlights the right cell; clipped
      text is found only as far as it shows.

### Phase 4 report — 2026-11-07

36 / 36 page JPEGs and rendered pages are unchanged: 19 corpus pages, five feature
and print-title pages, twelve large-control pages. Sizes/counts match the previous
export. Poppler extracts 1,469 / 1,469 collected corpus tokens. All eight new tests
fail with 31a03dd's export modules. Final PDF/licence/i18n gate after private gather:

```text
ℹ tests 60
ℹ pass 60
ℹ fail 0
```

| Control | Pages | Seconds before → after | MB before → after |
|---|---:|---:|---:|
| Seven public workbooks | 19 | 1.198 → 1.269 | 4.065 → 4.092 |
| 34.3 MB workbook, first 12 pages | 12 | 1.363 → 1.564 | 17.974 → 18.024 |

951 MB → 742 MB peak renderer RSS on the large control. Temporary canvases are
released after copying; recorded measurements are consumed per page. The recorder
follows actual fillText geometry, clipping, squeeze and rotation through child
canvases; repeated titles copy by row. Printed SVG labels are collected too.

8 / 8 controls open in Excel 2016 without repair and in LibreOffice. Both Office
rounds include Mac, missing-metadata and old-template controls; input hashes match
(the template uses a recorded plain-XLSX copy). Chromium finds the page-4 cell and
mouse/clipboard copies it and the diagonal phrase. All three apps load/export
without errors or external requests. No script-order change or new UI strings.

Final `node tools/quick.mjs`:

```text
✔ Node tests: 183 passed (14 s)
✔ tests/ui/smoke.test.mjs: 4 passed (18 s)
✔ tests/ui/ledger.test.mjs: 39 passed (132 s)
✔ tests/ui/quire.test.mjs: 57 passed (212 s)

All 4 passed.
```

Limit: existing overlapping chart labels join four spaces in PDFium; Poppler keeps all tokens.
Limit: PDFium/Chromium mishandle an emoji beside Arabic; the isolated case copies correctly in Poppler.
Limit: SVG subword clips use approximate advances; curved canvas clips are not measured.
Limit: Acrobat/Preview remain unchecked; [six PDF pairs and checklist are ready](../tmp/pdf/artifacts/phase4/reader-review.zip).

Own hashes stayed unchanged through the gate; all 11 original tabs remain and
server 8769 is stopped. No commits, catalog edits or Quire edits. [Report](../tmp/PDF-4-REPORT.md),
[Lead findings](../tmp/PDF-LEAD-FINDINGS.md). Phase files are frozen until committed.

**PDF.4 READY FOR LEAD**

Phase 5: links, bookmarks, language
- [x] **Quire:**
  - hyperlinks become link areas from their rectangles;
  - internal links and cross-references go to their target's page and position;
  - bookmarks come from headings (outline level, page and position from `LY.pages` and the fragment's rectangle),
    nested by level;
  - `lang` comes from the document's language.
- [x] **Lectern:** links on text and shapes (URL, or a jump to another slide); bookmarks from slide titles.
- [x] **Ledger:** cell hyperlinks; bookmarks from sheet names.
- [x] **Checks:** links click through in Chromium's viewer and in Firefox, and the bookmarks show in their side panes.

### Phase 5 implementation notes — planned 2026-11-07, authorized 2026-11-09

The release freeze was lifted and I25 was cancelled before any edits. Phase 5's
implementation is confined to shared PDF code and the existing export hooks;
Lead's desktop, Quire layout/drawing and UI work remains untouched. The following
plan guided implementation; the final report below records the measured results.

1. **Reuse the writer.** `common/pdf.js` already accepts page `links` and info
   `outline`/`lang`; `tests/pdf.test.mjs` checks URI escaping, destinations and nested
   outlines. The missing work is app metadata and geometry. Keep `L.buildPDF` and
   the word collectors' existing return values. Shared DOM link geometry belongs
   in `pdftext.js`, using the same visibility, clipping and transform rules as text;
   each app resolves its own targets. Measure while export DOM is attached.
2. **Map the actual exported pages first.** Destinations use zero-based output
   indexes, not source page/slide numbers. Build maps after page-range and hidden-item
   filtering, including Lectern handout slots and Ledger page order. Do not invent a
   destination for an omitted or unresolved target. Keep links out of clipped-away
   content and keep wrapped link rectangles separate instead of covering their gaps.
3. **Quire.** Text and image links already expose `data-link`; internal targets
   start with `#`. `L.fields.bookmarks()` supplies paragraph/offset positions;
   `LY.fragFor`, `LY.posToDom` and `LY.pageOf` locate them without editor navigation.
   Use `D.headingLevel` and `D.plainText` for outlines, deduplicating paragraphs split
   across pages and excluding non-exported fragments. Resolve a bookmark inside a
   split paragraph against its offset, not the paragraph's first page. Document
   language is available as `doc.props.language` and `doc.defaults.rPr.lang`; validate
   the document tag rather than using the UI locale. No layout/render/model edits
   are proposed.
4. **Lectern.** Text links have JSON `data-link`, but shape links are only attached
   in show mode; export uses thumbnail mode. Resolve model shape links from
   `data-sid` in the export hook, including grouped shapes, without changing the
   renderer. Model targets are `{url}`, `{slide}` or `{action}`. Map explicit slide
   IDs and deterministic first/last/next/previous jumps; last-viewed/end-show actions
   need a documented static-PDF policy rather than an arbitrary page. Titles come
   from `M.slideTitle`; `pres.lang` is available. Preview thumbnails retain
   `data-slide`; outline preview currently emits plain `.pv-text` with no slide ID,
   so associate its title rows with the filtered slide order, not unique title text.
5. **Ledger.** `sh.links` holds `{ref, target, location, display, tooltip}`. Compute
   link rectangles from the same `LY.geo`, print scaling, margins, centring and four
   repeated-title/body blocks used by `P.renderPage`; intersect each with its printed
   block. Cover merged cells, hidden rows/columns and repeated title copies. Resolve
   internal sheet/cell/named-range references through the formula/reference helpers,
   without calling `A.goToRef` (which changes selection and active sheet). Add one
   sheet bookmark at its first exported page, including chart sheets. A `HYPERLINK`
   formula currently returns its display value without populating `sh.links`; probe
   that separately before claiming formula-backed links. No calculation-engine edits
   are proposed. The workbook reader does not expose a document language; do not
   substitute the UI locale.
6. **Regression and reference cases.** Add failing-before app checks for wrapped
   and clipped links, Unicode/percent-escaped URLs, nested headings, duplicate titles,
   split-paragraph bookmarks, partial exports, omitted targets, grouped/rotated slide
   links, handouts/notes/outline, merged cells and repeated print titles. Check relative
   file links and non-static slide actions against Office/LibreOffice before settling
   their export behavior; never resolve a document-relative URL against the app's web
   origin. Candidate corpus controls: Word `Hyperlink - internal.docx`,
   `hyperlink-in table.docx`, `footer-contain-hyperlink.docx`; PowerPoint
   `hyperlinktest.pptx`, `hyperlinkOnImage.pptx`; Excel `ExcelWithHyperlinks.xlsx`,
   `hyperlink_formula.xlsx`. Select exact files under `/work/corpora`, plus the usual
   two or three independent metadata/producer controls per Office round.
7. **Gate after implementation is authorized.** Inspect PDF annotations/outlines
   with an independent reader; click local-fixture links and outline entries in native
   Chromium and Firefox (both Firefox binaries are installed). Use 9223/8769 and an
   isolated Firefox profile; preserve existing tabs and make no external requests.
   Verify JPEG/pixel identity, text extraction, selected-page destination indexes,
   document language, deterministic output and long-document/large-workbook cost.
   Run the full focused PDF/licence/i18n set after private English gather, all-three-app
   export checks and `tools/quick.mjs`; paste summaries. Finish the remaining docs and
   common-library list, report limits, then mark **PDF.5 READY FOR LEAD**. No Phase 5
   completion boxes were checked by the planning pass.

### Phase 5 report — 2026-11-09

357 / 357 page JPEGs, rendered pages and sizes match the preceding exporter;
25 / 25 PDF pairs extract identical text. The controls cover 13 public originals,
seven single-change derivatives, three navigation fixtures and the long/large
timing controls. All 11 new tests fail with 8f93f0f's frozen export modules.
Final PDF/licence/i18n gate after private English gather:

```text
ℹ tests 71
ℹ pass 71
ℹ fail 0
```

| Control | Pages | Seconds before → after |
|---|---:|---:|
| 20 corpus/derivative controls | 155 | 10.066 → 10.268 |
| Long Word control | 180 | 11.547 → 11.265 |
| 34.3 MB workbook, first 12 pages | 12 | 0.984 → 1.236 |

30 / 30 Office jobs open without repair and produce PDFs, with hashes and
unchanged inputs checked; LibreOffice opens the same controls. Each Office round
includes unusual metadata/producer controls. Native Chromium and Firefox follow
links and sidebar bookmarks in all three apps; Quire's headings nest correctly.
Relative URIs stay literal. Page-range exports omit unavailable destinations.
Three fixed-date fixture exports are byte deterministic. All three app exports
and screenshots have no errors or external requests.

Final `node tools/quick.mjs`:

```text
✖ Node tests: 247 passed, 2 failed (145 s)
✔ tests/ui/smoke.test.mjs: 4 passed (16 s)
✔ tests/ui/lectern.test.mjs: 39 passed (116 s)
✔ tests/ui/ledger.test.mjs: 39 passed (117 s)
✔ tests/ui/quire.test.mjs: 57 passed (192 s)

1 of 5 failed.
```

The two failures are the desktop test's hard-coded port 8760 (Intern uses 8769)
and the locale-label assertion requiring blocked translation catalogs. Neither
is counted as passing. Two raster/preview test fixtures now supply the required
shape list and await the queued repaint; 20 corrected old checks pass on the
baseline, and all 11 new checks still fail there.

Limit: nested HYPERLINK formulas are not exported as links.
Limit: rotated shape links use bounding rectangles.
Limit: absent document language is omitted; Ledger has no document-language field.
Limit: the formula workbook keeps its prior pagination difference (122 vs Excel 116/LO 79).
Limit: Acrobat Reader/macOS Preview remain unchecked.

No writer API/script-order change, new UI strings, commits, catalog access or
Lead-file edits. I25 had no edits to revert; its todo is intact. Own hashes stayed
unchanged through the gate; all 11 original tabs remain, and 8769/9224 are stopped.
[Report](../tmp/PDF-5-REPORT.md), [findings](../tmp/PDF-LEAD-FINDINGS.md).
Phase files are frozen pending review.

**PDF.5 READY FOR LEAD**

### Phase 5 review correction — 2026-11-09

3 / 3 hostile-link regressions fail before the correction and pass after it;
the complete focused PDF/licence/i18n gate now passes 74 / 74. The shared writer
drops `javascript:`, `vbscript:` and `data:` links, including mixed case and ASCII
controls. HTTP, HTTPS, mail, FTP, file and relative links retain their spelling;
the stricter HTML rule is unchanged. In the affected public Word control, six
script annotations are omitted (27 → 21 links); all other destinations, 18
bookmarks, language, eight page pictures and extracted text are unchanged.

All three fresh app exports/screenshots pass without errors or external requests.
Final `node tools/quick.mjs` on the correction:

```text
✖ Node tests: 209 passed, 1 failed (18 s)
✔ tests/ui/smoke.test.mjs: 4 passed (17 s)
✔ tests/ui/lectern.test.mjs: 39 passed (117 s)
✔ tests/ui/ledger.test.mjs: 39 passed (120 s)
✔ tests/ui/quire.test.mjs: 57 passed (236 s)

1 of 5 failed.
```

The sole failure needs the prohibited translation catalogs; Lead confirmed it is
an environment constraint. Lead authorized the Phase 5 commit after this fix.
Only Phase 5 files and the two PDF library rows in `docs/suite.md` are included;
Lead's desktop and other documentation hunks remain untouched.

When it's done
- [x] Docs:
  - docs/quire.md, docs/ledger.md and docs/lectern.md: the PDF row (pages as pictures, with text you can find, select
    and copy, links and bookmarks);
  - docs/suite.md: the shared `pdf.js` and `pdftext.js`;
  - AGENTS.md: the list of `public/common/` libraries;
  - docs/testing.md: the PDF checks.
- [x] Each app loaded and exported once after the change in common code (AGENTS.md: a common change is checked in
      all three apps), and `node tools/quick.mjs` (the two environment-related failures are recorded above).


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
