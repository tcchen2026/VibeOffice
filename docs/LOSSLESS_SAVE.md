# Lossless same-format save

**Goal.** Open a .docx/.xlsx/.pptx (or a macro, template or slideshow variant), edit it and save it in
the same format: what the user did not touch comes back as it was, and what cannot come back is
reported before the save. Conversions to other formats (PDF, CSV, Markdown, .ppt) are out of scope.

**Status:** all planned preservation features are in, and Office batches 1–7 are all accepted. Known
gaps, including a few files that still lose content without a warning, are under To do.
Known exceptions are documented in each app doc's "What a save keeps, converts and drops". Work from
here starts from a specific failing input, not another broad run.

## Design

The contracts are in [suite.md, Preservation](suite.md#preservation); the per-feature edit rules are in each app doc.

- **Baseline and identities.** The opened package stays as an immutable baseline (`doc.pkg`, outside
  undo snapshots). Writers reuse original part names, relationship IDs and object IDs for content that
  came from the file. Duplicates get fresh IDs, and their references are remapped together. New IDs
  and names follow what Office writes.
- **Part classes** (`common/opc.js`):
  - **opaque**: original bytes;
  - **merged**: settings-like parts, original DOM with only the model's properties replaced in schema
    order;
  - **regenerated**: content parts rebuilt from the model, plus the preserved fragments attached to it.
- **AlternateContent.** It is captured from the raw XML before normalisation and written once, at the
  first surviving object. Namespaces and `mc:Ignorable` are hoisted to the part root.
- **Edits.** Kept data is stored per property. Edits replace only what they change. Moves and resizes
  update anchors without converting.
- **Undo, copy and drafts** carry fragments and identity reservations. Drafts keep the document's
  variant and encryption.
- **Losses.** Every conversion and drop is recorded in `doc.losses`. The Compatibility Checker shows
  only real, important losses (`notify: true`), in plain words with a place the user recognises; the
  rest is recorded for tests (suite.md, What the user is told).
- **Office conventions.** A `tools/ooxml/package.py` rule exists for each defect Office found.
  `tools/ooxml/conventions.py` flags markup our saves add that Office-authored files never contain.

## What was done

- **All three apps:**
  - package carry with ownership classes;
  - all 14 save variants;
  - encrypted drafts and recovery;
  - Strict input saved as Transitional;
  - the Compatibility Checker;
  - signatures reported, because saving invalidates them;
  - VBA kept in macro-enabled formats.
- **Quire:**
  - content controls and bindings;
  - text, paragraph and section properties;
  - opaque objects (OLE, SmartArt, controls, altChunk);
  - drawing-property merges;
  - row and cell alternatives;
  - watermarks;
  - comments.
- **Ledger:**
  - pivot tables (source and output edits, refresh, shared caches);
  - threaded comments with mentions;
  - controls, OLE objects and VML;
  - slicers and timelines;
  - extensions;
  - query tables;
  - custom XML.
- **Lectern:**
  - media;
  - sections and custom shows;
  - comments;
  - SmartArt, OLE, ink, 3-D and chartEx frames;
  - unused masters and design-edit invalidation;
  - tags, actions and effects.
- **Final measurement (`984eaf3`):**
  - **Corpus:** the pinned corpus of 4,697 files; 14,091 driver attempts; 30,919 package/SDK states,
    with 29,341 OK, 670 failed (failures relative to the original count as failures) and 908 excluded.
  - **LibreOffice:** renders 2,869/2,880 Word, 930/935 Excel and 836/837 PowerPoint unedited saves.
    Page counts differ for 12 Word and 3 Excel files.
  - **Text:** Word keeps 776,207/776,869 words. PowerPoint keeps all 58,022 words.
  - **Cells:** Excel matches exactly in 883/889 saves.
  - **Evidence:** `~/corpora/results/lossless-final-984eaf3/`.

## Working rules

These are in addition to AGENTS.md.

- **Office batches:** at most 15 files in `tmp/lossless/<batch>/` (the repository's ignored `tmp/`), each batch with a
  `CHECKLIST.md`. A repair prompt blocks the step that produced the file.
- **Office 2016 automation:** `tools/win_test_server/` on `win10` (`C:\vo`) provides HTTP normal-open/repair
  results and native DOCX/XLSX/PPTX PDFs; setup and endpoints are in [Testing](testing.md#office-2016-vm-http-openrepair-and-pdf-checks).
  Compare original and saved files. An inconclusive result is not a pass; a clean open/PDF does not replace
  the checklist's content and editability checks.
- **Docs:** they describe the current state. Run details go in `~/corpora/results/<run>/README.md` and
  in commit bodies.
- **Commit bodies:** what changed for the user, the measured numbers (and that the test fails on the
  previous commit), and what is pending.
- **Plan status:** update this file at milestones only.

## To do

1. **Silent losses** (2026-11-01; the alarm plus a count of drawing objects per sheet, story and slide in every
   original and its save). Fixed: SmartArt, Office add-ins and other objects Ledger can't show are kept (Ledger
   kept frames); part names compare case-insensitively (17 Word template thumbnails); a relationship to a part the
   original lacks no longer drops its owner (a glossary, nine charts; Office's `Target="NULL"` written back, nine
   Excel templates' themes); fields whose code holds nested fields or pictures; bookmarks and permission ranges
   between table rows or cells; a picture's PowerPoint-for-Mac PDF version; Quire's built-in styles written beside
   a file's own styles of the same name (Word in other languages: `berschrift1` is heading 1), 565 of 2,768 saves,
   where LibreOffice gave the file's paragraphs the second one's look; a PowerPoint 97-2003 organization chart
   (`com:legacyDrawing`) that became an "Embedded object" rectangle; Lectern's ActiveX controls, saved as dead
   pictures (84 in 14 decks), now kept unedited, moved or resized; what the user deletes is recorded as
   `deleted:`, so `unreferenced:` stays exact in edit runs too (Lectern 44 cases, Ledger 13, Quire 9 → 0); Quire's
   tracked moves, which became a deletion and an insertion (`w:moveFrom`/`w:moveTo` and their named ranges; 16 of 17
   files now keep every move, the 17th has no ranges; moved-away text stays `w:t`, which Word requires), a simple field's lock (a locked PRINTDATE showed today's date),
   link ScreenTips with quotes (cut at the first `\"`) and a link's target frame (`w:tgtFrame`); a Word file
   that names no default formatting (119 corpus files) is shown and saved with Word's, Calibri 11 pt, 8 pt after and
   1.08 lines (Word 2016), where Times New Roman 10 pt was written; the styles part, which Quire rebuilt on every save
   (found on the Fluent Ribbon walkthrough): the latent styles (2,088 of 2,132 Word-authored files), each style's
   `w:unhideWhenUsed` (2,083; without it Heading 2–6 stay hidden in Word's gallery where used), `w:locked`, `w:rsid` and
   a table style's row properties are kept, the styles stay in the file's order, and a built-in style the file doesn't
   define is added only once something uses it or it is changed, with its latent settings, as Word writes it (about 25
   unused ones were added to every file). Over 2,310 Word files, the unedited saves' `styles.xml` went from 2,523,618
   lines differing from the original to 41,908 (1,299 within 3 lines), with no new package findings or losses; Word
   2016 also opens `fdo76101` again (4,091 styles: the 47 added built-ins passed its limit of 4,092, rule
   `too-many-styles`), and `fdo78904` gets its table styles' repeated header rows back; the ids of paragraphs, runs and
   table rows (`w:rsid*` in 2,029 of those files; `w14:paraId`/`w14:textId`, 572), which Quire dropped and merged runs
   across: Word 2016 places text per rsid run, so the save's text moved by fractions of a pixel (the walkthrough: 18 of
   19 pages, 18 → 3 after the fix; removing only the rsids from the original gave the same shift), and Compare, Combine
   and comments use them; a copied paragraph doesn't repeat its paragraph id. Text is written with `xml:space="preserve"`
   only where it needs it, as Word does: on every `w:t`, it moved Word's text a fraction too (7 of the walkthrough's
   pages); what rows and cells have that Quire doesn't model (`w:cnfStyle`, `w:wBefore`/`w:wAfter`, a row's
   `w:tblCellSpacing`, cell revisions and property changes), a cell's width as the file gives it (a percentage, auto:
   written in twips before; kept while the table's grid is unchanged), a row's exceptions to the table's properties
   (`w:tblPrEx`, 1,146 in 39 files: Word moved a row's text 54 pt without them) and the columns' tracked change
   (`w:tblGridChange`), the last two found by Intern's part-by-part measurement (I19); a run's complex-script size
   (`w:szCs`), which was written equal to its size where the file had none, changing the size of Arabic, Hebrew and
   other complex-script text (27,025 times in 683 files; 1,715 left, on runs the file gave no properties). Intern's review
   of these (I20) found five edit and copy cases, fixed with them: a changed style's base style, a new style's own
   settings, Insert Rows copying cell history, a copy before its source taking its ids, a row exception's revision id. The
   comment extensions Quire adds (`commentsExtended.xml`) use the draft w15 namespace in a file from Word 2013's
   preview that binds it (64 Word files, 3 with comments and no extensions of their own): Word 2016 refused the three
   saves (`Comment012`, `Comment035`, `Comment050`), which now open, rule `w15-namespaces`. A comment's reference run
   keeps its properties and ids (all 1,340 in the 110 Word files with comments, from 846 ids and 990 properties) and its
   place after the other comments' marks (1,058, from 634; Word puts a few after a picture or a proofing mark, and
   those still follow their range's end): Word 2016 placed the text after them a fraction of a point differently
   (`Comment035`, `Comment050`, now pixel-identical in Word 2016; 7 of 8 files of the batch are); a field code's runs,
   their properties, ids and split, while the code is unchanged (2,748 of the 3,476 instruction runs in 364 Word files
   as the original, from 1,048; most of the rest in one file and in HYPERLINK fields, which become hyperlinks): they
   were one run with the field's first run's properties, and Word formats the result from the code (`wDateValueFormat`'s
   page numbers printed at 12 pt, every page) and places text by their ids (`form_footnotes`, found by Intern, I21); a
   comment read without a date gets none (57 in the corpus got the save's time). Unedited
   saves losing a part the original used: Ledger 22 files → 1, Quire 31 → 2, Lectern 4 → 1, all damaged originals. The alarm now reports
   `unreferenced:` only for parts the original used (`orphan:` for the rest), so that count is direct.
   Still to do:
   - **Lectern ActiveX copies:** a copied control (the shape or its slide) is saved as its picture, with a notice;
     keeping it needs new ids across the slide, its VML drawing and its ActiveX parts, as Ledger does.
   - **Quire field codes** Word wouldn't write that way (one run, across paragraphs, around a bookmark) and
     **direction wrappers** (`w:dir`, `w:bdo`) are converted, now with a notice.
   - `tdf93284.docx`: the original gives two relationships the same id (a picture and the font table); the save
     keeps the picture's reference pointing at the font table.
   - Quire: text pasted as HTML, RTF or Markdown into a file from Word in another language keeps the built-in
     styles' ids (Heading 1, Table Grid, Quote); it is saved with the file's own style of that name, but Quire shows
     the built-in's look until the file is reopened.
   - Quire text Word places differently, from Intern's bisections (I21, `tmp/losses/i21/REPORT.md`): adjacent runs
     with the same formatting and ids are merged (`D.normalize`; `tdf81345`, `Strict01`); a field's code runs go
     outside the hyperlink that holds the field (`docx-write.js` hyperlink groups leave field markers out; `Strict01`'s
     table of contents, and `stress004`'s: there, keeping the code's own runs crosses some limit in Word that moves
     Persian text on 12 later pages by a fraction, which any four fields written as one run undo); text with an
     interior double space gets `xml:space="preserve"` (`form_footnotes`).
   - Quire revision history it doesn't model, dropped silently on an unedited save: a HYPERLINK field whose address
     was changed with Track Changes (`tdf124472`: the old address and the instruction's revisions go); a revision on
     a comment's reference mark (`tdf105485`, `Comment024`); a bound content control's cached text with revisions,
     replaced by one deletion of the bound value (`SdtContent`'s header). Keeping them needs the reference mark and
     the field instruction to carry revisions in the model, and Accept/Reject to handle them.
   - Quire, a damaged original that LibreOffice recovers and the save doesn't: footnotes whose ids don't match their
     references (`tdf153255`: LibreOffice shows six, the save three). Word 2016 repairs the original and opens the save
     cleanly, so it waits. (Two final section definitions, `tdf108849`: fixed, the last one wins, as in Word 2016.)
2. **Lectern checker numbers:** run the Lectern corpus driver, which needs Playwright, to measure how
   often unedited saves show the checker. Ledger shows it for 5 of 888 files, and the Quire sample
   for 3 of 106.
3. **Office conventions** (low risk; fix when touching that code):
   - A hyperlink containing `w:del`/`w:ins`: checked in Word 2016 on 2026-11-03 (`tdf115212`, `tdf150443`): the saves
     open without repair and their PDFs match the originals' pixel for pixel.
4. **Related, separate plan:** SmartArt editing in Lectern. It replaces the `frames.js`
   "edited preview → plain group" path and is waiting on the PowerPoint spike results.
