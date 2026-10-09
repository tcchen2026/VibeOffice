# Next plan: lossless same-format save

The approved plan stands. This file changes the order
and the habits, based on the progress review of commits b06d938..ffc6814.

## Status

**The planned preservation features and review fixes are implemented. The final Excel page-setup
correction, remaining render exceptions and Office acceptance still need closeout.**

Full runs are complete for Word (`7c25bd7`), Excel (`3716a3b`) and PowerPoint (`56de065`), with
**14,091 driver attempts / 30,919 package/SDK states** and no added validator diagnostics.
Word's later one-input break recovery and Excel's two pre-release-format conversions have separate
focused evidence; the versioned full artifacts remain unchanged. Five Excel pagination probes have
identified another real preservation gap: absent page setup becomes portrait on save. Its correction
passes 71 unit tests, 49 focused corpus states and seven render pairs; the final Excel refresh follows.

- Word and package-preservation Office batches 1–2 are accepted, including corrected comments and
  all namespace follow-ups. The slide-6 PowerPoint repair is fixed and accepted (`60c26b4`).
- [Office batch 3: Excel, 15 files](~/Downloads/lossless-check/office-batch-3/CHECKLIST.md): pending.
- [Office batch 4: PowerPoint, 15 files](~/Downloads/lossless-check/office-batch-4/CHECKLIST.md): pending.
- [Office batch 5: Word/Excel, six files](~/Downloads/lossless-check/office-batch-5/CHECKLIST.md): pending.

The handed-off files remain unchanged. Failed/excluded inputs and the documented content/layout
exceptions remain explicit. **The plan is not signed off**, and automated checks do not establish Office acceptance.

## Changed habits
1. **Office is the gate.**
   - Before starting a step that adds a new kind of preserved XML, the patterns it builds on must
     have passed an Office batch.
   - A batch has at most 15 files, in `~/Downloads/lossless-check/office-batch-N/`, with a
     `CHECKLIST.md`: file, feature, what to look for, and a result column (✓ / repair prompt /
     looks wrong).
   - A repair prompt blocks the step that produced the file.
2. **Docs describe the current state.**
   - Each `docs/<app>.md` gets one "What a save keeps, converts and drops" section: one table of
     kept / converted / dropped, with the current numbers, plus one edit-behaviour table per feature.
   - No dated increment sections and no superseded tables. History and run details go in
     `~/corpora/results/<run>/README.md` and in commit bodies.
3. **Commit bodies.** Every commit body says three things:
   - what changed for the user;
   - the measured numbers, from a full run or focused;
   - what is still pending.
4. **AGENTS.md stays short.**
   - Move the six preservation rule paragraphs to a "Preservation" section in docs/suite.md, or a
     new docs/preservation.md that holds the opc.js contracts and the per-app edit matrices.
   - Keep one line in AGENTS.md pointing there.
5. **Full runs before features.** After each app's group of steps, run the full corpus, not only
   focused fixtures. Never let the documented corpus numbers predate the code by more than one step.

## Order
1. **Office batch 1** (now, about 15 files):
   - Quire: one content control with a binding edit, one Word 2010 text effect after a bold edit,
     one SmartArt that was resized, one OLE object that was pasted, one `.docm`;
   - Lectern: one `.pptm` with video, one deck with comment authors and a handout master;
   - Ledger: one `.xlsm`, one workbook with custom XML and external links.
   Hand it to the user, then continue with steps 2–3 while waiting.
2. **Full reruns on HEAD:**
   - the Word, Excel and PowerPoint corpora: save/reopen, package/SDK, loss audit, text edit with
     undo/redo, draft recovery;
   - `mdroundtrip.mjs` and `mdspec.js`;
   - LibreOffice on the batch-1 files.
   Update the numbers.
3. **Docs cleanup** as in habit 2, and the AGENTS.md move from habit 4.
4. **Fix what batch 1 finds.**
   - Any repair prompt is fixed, and its pattern re-checked across the corpus, before step 5.
5. **Quire remainder:**
   - merge drawing-property edits instead of converting the whole frame;
   - alternatives around table rows and cells;
   - the watermark path.
   Then a full rerun, then **Office batch 2** (Quire).
6. **Ledger**, in order of harm:
   - (a) Pivot tables. Write the rules table in docs first: shared caches, each source kind (range,
     table, name, external, consolidation), source row/column changes, edits inside the output, and
     sheet or source deletion. Then the code, with undo through `snapSheet`.
   - (b) Threaded comments: all threads, mentions, person ids.
   - (c) Controls, embedded objects and non-comment VML shapes.
   - (d) Slicers and timelines.
   - (e) Data-bar extensions and unknown `extLst` entries.
   - (f) Query tables.
   Then a full rerun, then **Office batch 3** (Excel).
7. **Lectern:**
   - (a) sections and custom shows;
   - (b) comments;
   - (c) SmartArt, OLE, ink, 3-D and new-chart frames;
   - (d) unused masters and design-edit invalidation;
   - (e) tags, actions, and effects on non-media shapes.
   Then a full rerun, then **Office batch 4** (PowerPoint).
8. **Finish:**
   - every remaining conversion path records its loss in the Compatibility Checker;
   - final full runs for all three apps;
   - the docs give the final numbers and the documented exceptions.

## Done means (per step)
- Unit tests pass (`node --test tools/ooxml/*.test.mjs tools/quire/test/*.test.js
  tools/ledger/test/*.test.js`).
- All four pages load with no console errors.
- No new validator diagnostics against the originals.
- Unchanged opaque parts are byte-identical.
- Edit, undo/redo and draft scenarios pass.
- LibreOffice opens the samples.
- The docs and the commit body carry the numbers.
- Office acceptance is either checked or listed as pending in the next batch, never assumed.

## Budget
About 16–19 steps, roughly 20–30 agent hours, plus the user's Office rounds. If a step runs past
twice its estimate, stop and report rather than widen the scope.

## Office-convention bugs (found 2026-10-08, do before step 5)
Found with `python3 tools/ooxml/conventions.py ORIGINALS SAVED`, which lists markup our saves add that
Office-authored corpus files never contain. Rerun it on every full corpus run; each confirmed finding
becomes a `package.py` rule with a test, per AGENTS.md "Office acceptance".

1. **WordArt fill outside `w14:textFill`** (Quire, 21 files).
   - Fixed in `27e9f73`: `docx-write.js` wraps the fill.
   - The package checker now rejects fills outside `textFill`/`textOutline`, with a regression test.
2. **Inserted-then-deleted text nested `<w:del><w:ins>`** (Quire, 9 files).
   - Fixed in `27e9f73`: `wrapRunGroup` writes `<w:ins><w:del>`, as Word does.
   - The package checker rejects reversed nesting, with a regression test and an authored batch-1 save/reopen sample.
3. **Strict files written with Transitional vocabulary** (67 Word files; also some Excel and PowerPoint).
   - The cause: `K.strictXML` swaps only namespace URIs, so `w:ind w:right`, tab `left`, `0`/`1`
     on/off values and the like are left in Strict files.
   - Implemented:
     - save a file opened as Strict in Transitional;
     - convert its preserved fragments Strict→Transitional (namespaces; DrawingML `N%` → thousandths;
       Transitional also accepts Strict's `start`/`end` and `true`/`false`);
     - record a "saved in the standard format" entry in the Compatibility Checker.
   - Includes conditional table flags, tab measures and relative drawing percentages. The 175 Strict
     inputs save/reopen; 700 saved/draft states introduce no package/SDK diagnostics. The SDK does open
     these Strict originals; malformed/unopenable originals elsewhere remain failures, not passes.
     **User approved Transitional output with a notice on 2026-10-08.** The conversion samples
     passed Office batch 2.
4. **Fragments repeat root declarations on inner elements.**
   - Each kept fragment re-declares the namespaces and `mc:Ignorable`: Excel `calcPr`, `phoneticPr`
     and other sheet children in 364 of 935 saves; Word footnote separators in 650 of 2,880.
   - This is allowed by the markup-compatibility rules, but Office never writes it.
   - Proposed fix: when emitting a fragment into a part, drop declarations identical to the
     destination root's; add the missing ones, and the union of `mc:Ignorable`, to the root instead.
   - Batch 1 accepted repeated declarations and inner `mc:Ignorable` in Word. Its Excel sample
     repeated declarations only; the Excel `calcPr` form was not checked in that batch.
   - Implemented in shared `opc.js`: rewritten parts hoist namespaces and `mc:Ignorable` to the
     root, retaining real local prefix rebindings. Opaque parts retain their original bytes.
     On the same 400-document subset, main XML falls from 22,183,474 to 4,770,570 bytes, compared
     with 5,373,347 original bytes. All 1,845 changed XML parts preserve expanded XML content;
     package checks add no issues and conventions report no inner `mc:Ignorable` signatures.
     All four batch-2 follow-up samples pass Office: the original Excel form and compact Excel,
     Word and PowerPoint saves.
5. **Low risk; make them match Office when touching that code:**
   - Lectern `defaultTextStyle` writes `a:ea`/`a:cs typeface="+mn-lt"`; Office writes `+mn-ea`/`+mn-cs`.
   - `p:penClr` uses `a:srgbClr` where Office uses `a:prstClr`.
   - Quire WordArt writes `a:normAutofit` (no Office example to compare).
   - Hyperlinks containing `w:del`/`w:ins` (no Office example either way): include one in an Office
     batch.

## Current execution

- Versioned suite evidence: **14,091 driver attempts / 30,919 package/SDK states**. Validation has
  **29,341 OK, 670 failed and 908 excluded**, with no added package issues or SDK diagnostics.
  Unvalidatable originals and missing outputs remain failures. Scope and hashes are recorded in
  each run's README and reports.
- Quire (`quire-final-7c25bd7/`): save/draft each **2,880 OK / 22 failed**; text **2,281 OK /
  22 failed / 599 excluded**. Its **19,028** validation states have **18,056 OK / 373 failed /
  599 excluded**. This full refresh includes sections, picture bullets and list metadata. Counts
  now retain **3,404/3,405 sections, 175/175 multi-column sections, 9,202/9,202 list definitions,
  81/81 picture-bullet definitions, 106/106 picture-bullet references, 8,923/8,923 list identities,
  8,910/8,910 template entries, 232/232 legacy settings and 103/103 cleanup entries**. Independent
  word counts remain **776,207/776,869**; stale bound values, unread stories and unsupported or
  malformed content remain explicit exceptions. Driver statuses match the preceding full run.
- Quire's malformed-break correction (`e7387c7`, `quire-break-recovery-2026-10-09/`) retains seven
  misplaced breaks and restores **four pages instead of one**. Seven command states and seven
  corpus states add no diagnostics; all seven corpus states render as four pages. The initial
  scope scan's 33 archive/XML failures remain recorded separately. The correction supplies seven
  separately hashed render rows; the frozen full report remains unchanged.
- Ledger (`ledger-final-3716a3b/`): save/draft each **935 OK / 11 failed / five excluded**; text
  **933 OK / 11 failed / seven excluded**. Its **6,589** validation states have **6,371 OK /
  201 failed / 17 excluded**. All independent cell coverage is fresh: **889 saves / 1,649,147 cells**,
  with **883** exact value/formula matches; **887 edits / 1,648,615 cells**, with **879** exact outside
  the scripted target. The initial missing-openpyxl attempt and large-file timeouts remain failures
  in their raw reports; the installed environment and a single-input extended-timeout comparison
  provide corrected evidence. Existing value/formula and formatting exceptions remain documented.
- Ledger's pre-release worksheet-format correction (`3a93371`, `ledger-legacy-format-2026-10-09/`)
  covers two inputs / six worksheet properties. All six browser save/edit/draft attempts pass;
  14 SDK comparisons remain failed because their originals are unreadable. Both saves render;
  the readable original retains one page and the unreadable original has no comparable count.
  The correction reports conversion; all 68 Ledger tests passed at that step.
- Ledger page setup now keeps absence, individual attributes and printer dependencies. The seven
  public fixtures produce **21 successful driver attempts / 49 passing package/SDK states**; all
  seven unedited render pairs retain page counts, fixing five diagnosed differences. Thirteen
  command states cover paper/quality edits, history and sheet copies, with three additional chart-sheet
  states passing separately. The **71-test Ledger suite and new chart-sheet regression** pass.
  Evidence: `ledger-page-setup-2026-10-09/`. The final full Excel refresh and Office check remain.
- Lectern (`review-final-56de065/lectern/`): save/draft each **837 OK / six failed / one excluded**;
  text **548 OK / six failed / 290 excluded**. Its **5,302** validation states have **4,914 OK /
  96 failed / 292 excluded**. All **827** independently readable unedited pairs retain every
  inventoried feature and **58,022** words. Scripted SmartArt-preview conversions and the two
  part-animation effects they invalidate have explicit notices. Parallel SmartArt work is untouched.
- Current render evidence, including the seven-row Word correction: **30,919 states**, **29,027 OK /
  984 failed / 908 excluded**. Unedited saved renders complete for **Quire 2,869/2,880, Ledger
  930/935 and Lectern 836/837**. Every unsuccessful unedited saved render also fails or times out
  in its original. Readable unedited page-count differences are **12 Word / 37 Excel / zero
  PowerPoint** before the final Excel page-setup refresh. Missing outputs, original failures and
  count mismatches are separate; these are not lossless-save percentages. The renderer reuses only
  successful, hash-verified PDFs, and the initial dead-worker harness failure remains archived.
- Convention scans: **Word 193** (three additions retaining scoped MC rules and Strict list alignment,
  two removed defaults); **Excel 187** after the pre-release correction (two added scoped MC rules);
  **PowerPoint 166**, with no additions. Each signature has a recorded review, not an automatic pass.
- Office batches 3–4 have 15 files each and pass automated package/SDK/LibreOffice checks. Their
  previous-writer differences are reviewed; batch-3 file 04 includes the root-comment edit fix.
  Batch 5 has six passing automated samples for Word sections/numbering and Excel row heights.
  Exact files, hashes and superseded preparation cases are recorded in `office-followup-2026-10-09/`.
  Actual Office acceptance of all three batches remains pending.
- All four pages load without console errors; known font requests remain. Markdown evidence remains
  **652/652 CommonMark, 22/22 GFM and 24/24 exact saves with confined edits**. Shared preservation,
  identity and Compatibility Checker checks retain their prior passing evidence; unchanged suites
  are not repeated solely to generate another total.

Next: freeze and measure the Excel page-setup correction, finish the remaining render-exception
review and current docs, then obtain Office acceptance before signoff. Original failures, hashes,
correction overlays and earlier measurements remain with the versioned report directories.

## Review of 6d24c43..275406e (2026-10-08)
Followed well: bounded Office batches with previous-writer comparison, checker rules for the WordArt
and nesting fixes, Strict→Transitional (verified: no Strict namespaces left except a custom-XML
schema reference, which is data), consolidated docs, commit bodies with numbers, paragraph-order fix
(the frozen run's 143 files with `w:pPr` after a run in comments/headers/footnotes: 25 of 25 correct
at HEAD). Corrections:

1. **Excel inner `mc:Ignorable` is not Office-checked.**
   - "Office batch 1 accepted both forms" holds for Word: the batch files carry `mc:Ignorable` on
     `w:rPr`, `w:pPr` and `w:footnote`.
   - It does not hold for Excel: `12-sheet-namespaces.xlsx` repeats only `xmlns` declarations.
   - `calcPr mc:Ignorable="…"` is still written at HEAD, in 364 of 935 saved workbooks.
   - Put one such workbook in Office batch 2, or remove the pattern (item 2).
2. **Hoist fragment declarations; the trade-off is wrong.**
   - Every retained `rPr`/`pPr` now re-declares about 30 namespaces plus `mc:Ignorable`: 2,009 and
     1,998 of 2,873 saves.
   - On 400 documents, `document.xml` grows from 5.3 MB (originals) to 21.9 MB (saves); zipped
     11.1 → 9.9 MB, compared with 8.7 MB before property preservation. Word parses 4× the XML.
   - Fix in `common/opc.js` emit:
     - drop a fragment declaration identical to one in scope at its destination;
     - declare missing prefixes, and the union of `mc:Ignorable`, on the part root;
     - keep a local declaration only for a real prefix conflict.
   - Gate: byte size back near the originals, `conventions.py` shows no inner `mc:Ignorable`, no new
     package/SDK diagnostics, and one Word and one Excel sample in Office batch 2.
3. **(Resolved: batch 2 was finished and accepted.) The Office gate is being skipped.** The working tree has uncommitted Ledger pivot code
   (`pivots.js`, changes to `xlsx-read/write`, `ops.js`, `preserve.js`) while Office batch 2 is
   pending and this plan says no Ledger feature code before it. Finish batch 2 (with items 1–2) first,
   or record why the gate was changed.
4. **Run `tools/ooxml/conventions.py` on every full run** and list its new signatures in the commit
   body; the property-preservation step would have shown item 2.
5. **(Resolved.) Minor:** `public/index.html` still says "placeholder address" above the Feedback link, which now
   has the real address.

## Review of 3b92fe9..9f18001 (2026-10-09)
The approach is right; the weak spot is identity allocation (local max+1 or `_Copy<hex>` instead of
what Office writes and what the whole file uses). Five items are fixed in `1a33733`, with tests and
the package rules `duplicate-vml-id`, `shared-vml-preview`, `duplicate-vml-block` and
`animation-part-target`. They cover copied OLE previews, SmartArt part animations, reused query
column IDs, note VML blocks, and slicer copy names and workbook names.

**Rebuilt:** batches 3 and 4 now use `56de065`, with every previous-writer difference reviewed.
The earlier handoffs are archived; the current copies await Excel and PowerPoint acceptance.

Also fixed afterwards:
- namespace processing cost: a fast path in `K.hoistNamespaces` with identical output on 11,595
  corpus parts. On a 30 MB sheet, namespace processing alone went from 2,956 ms to 3 ms; total save
  time was not measured.
- whole-column and whole-row pivot sources: they keep their extent, and field bookkeeping still runs;
- pivot outputs reshaped by inserted or deleted lines: refresh plus a notice;
- sheets needing more than one VML id block: every needed block is reserved, and `duplicate-vml-id`
  is workbook-wide for spreadsheets.

Review items, now implemented:
1. **Lectern rotation (fixed).** Rotation and flip commands, handles and the size dialog disable
   unsupported transforms on retained frames and groups containing them. Moving/resizing a
   multi-member wrapper keeps each original rotation and flip. Three focused scenarios and their
   nine saved/undo/redo states pass package/SDK comparison; the edit matrix describes the limitation.
2. **Lectern designs (fixed).** Retained master/layout shapes are tied to their displayed models.
   Edits replace only those shapes, retaining unread siblings and their stacking positions. Sixteen
   scenarios / 50 saved states pass package/SDK comparison, including interleaved unread frames,
   deletion, explicit reorder, addition, style edits, undo/redo and drafts.
3. **Lectern comments (fixed).** Re-anchored comments use the slide's creation ID; copies remap
   definitions and anchors together. Replaced extensions do not first carry their old comment
   relationship. Thirteen scenarios / 41 saved states pass package/SDK comparisons, including the
   explicit creation-ID check and deletion of an anchored shape on a copied slide.
4. **Ledger copies (fixed).** Controls/OLE, notes and nested VML group members allocate linked
   identities from the destination sheet's blocks. Originals reserve all their blocks; copies extend
   only when needed, and a reopened 512-note sheet keeps its bytes. Copied query/control names use
   numeric suffixes; range-query names retain their sheet scope. Fifteen object cases / 38 saved
   states and eight query cases / 23 saved states pass package/SDK comparison.
5. **Hygiene.**
   - (Fixed.) Tables, extensions, slicers and threads use shared GUID copying/remapping in `K.duplicate`.
     All 141 unit tests pass; 83 focused saved states plus two actual browser draft states add no
     package/SDK diagnostics. Copy-of-copy identities remain independent and repeat saves keep them.
   - (Fixed.) Pivots and threads use the writer's sheet-to-part map.
   - (Fixed.) Pivot source-kind tests generate missing fixtures from the pinned corpus, with no dated path.
   - (Fixed.) docs/ledger.md uses the current per-feature test counts.
   - (Fixed.) `namespace-size.mjs` counts deliberate scoped `Ignorable` separately. The three MC
     directive cases pass without treating their necessary local declaration as a failure.
   - The former `designs.mjs` `style` failure is fixed: paragraph preservation now writes an edited
     default run property while retaining its unknown siblings.
   - (Fixed.) An explicit frame conversion/deletion notice accounts for its opaque dependency tree;
     unrelated losses remain visible. Three real-file cases / nine saved states pass package/SDK
     comparison without the redundant notices. The change is in the loss audit, not SmartArt conversion.


Leave SmartArt to the parallel SmartArt work: it replaces the "edited preview → plain group" path
in `frames.js` and the diagram parts' read/write. Don't change those.
