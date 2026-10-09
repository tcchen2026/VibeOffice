# Lossless same-format save

**Goal.** Open a .docx/.xlsx/.pptx (or a macro, template or slideshow variant), edit it and save it in
the same format: what the user did not touch comes back as it was, and what cannot come back is
reported before the save. Conversions to other formats (PDF, CSV, Markdown, .ppt) are out of scope.

**Status: done.** All planned preservation features are in, and Office batches 1–7 are all accepted.
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

- **Office batches:** at most 15 files in `~/Downloads/lossless-check/<batch>/`, each batch with a
  `CHECKLIST.md`. A repair prompt blocks the step that produced the file.
- **Docs:** they describe the current state. Run details go in `~/corpora/results/<run>/README.md` and
  in commit bodies.
- **Commit bodies:** what changed for the user, the measured numbers (and that the test fails on the
  previous commit), and what is pending.
- **Plan status:** update this file at milestones only.

## To do

1. **Silent real losses** that the unreferenced-part alarm found on unedited saves:
   - **Ledger SmartArt:** Ledger drops SmartArt on a sheet entirely
     (`libreoffice__12275492e8fc__tdf83671_SmartArt_import.xlsx`,
     `libreoffice__6dda90d75ceb__tdf151818_SmartartThemeFontColor.xlsx`).
   - **Other Ledger drawings:** 16 more Ledger files lose a sheet drawing. Separate the empty ones from
     the real ones (`poi__3c51c202a1fc__npe.xlsx` loses a drawing and its picture).
   - **Lectern:** `libreoffice__9d97505735d4__import-characters.pptx` loses `ppt/media/image1.png`,
     and `poi__793baf3ae217__2411-Performance_Up.pptx` loses `ppt/media/image3.pdf`.
   - **Required with each fix:** keep the object (or give it its own notice), and add a regression
     test. In corpus runs, an `unreferenced:` entry for a part the original used is a defect.
2. **Lectern checker numbers:** run the Lectern corpus driver, which needs Playwright, to measure how
   often unedited saves show the checker. Ledger shows it for 5 of 888 files, and the Quire sample
   for 3 of 106.
3. **Office conventions** (low risk; fix when touching that code):
   - `p:penClr` uses `a:srgbClr` where Office uses `a:prstClr`.
   - Quire WordArt writes `a:normAutofit`.
   - A hyperlink containing `w:del`/`w:ins` has never been in an Office batch.
4. **Related, separate plan:** SmartArt editing in Lectern. It replaces the `frames.js`
   "edited preview → plain group" path and is waiting on the PowerPoint spike results.
