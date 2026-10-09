# Next plan: lossless same-format save

The approved plan stands. This file changes the order
and the habits, based on the progress review of commits b06d938..ffc6814.

## Status

**The planned preservation features, review corrections and final measurements are complete.**
Office batches 1–6 passed. Batch 7 has twelve accepted original handoffs and two accepted replacements;
no Office check remains pending. Documented content and layout exceptions remain.

The pinned corpus has **4,697 files** (2,902 Word, 951 Excel, 844 PowerPoint). The final refresh has
**14,091 primary driver attempts plus four timeout retries**, and **30,919 unique saved/attempted
package/SDK states**: **29,341 OK, 670 failed and 908 excluded**, with no added diagnostics.
The full writer is `984eaf3`; `7f85450`'s two-input legacy Excel correction is separately measured
in six attempts / 14 states. Original-relative failures remain failures even when a corrected save
passes standalone validation.

LibreOffice renders **2,869/2,880 Word, 930/935 Excel and 836/837 PowerPoint** unedited saves.
Readable-pair page-count differences are **12 Word / three Excel / zero PowerPoint**. Across all
unique states the paired render/count result is **29,265 OK, 746 failed and 908 excluded**.
These figures do not establish lossless saving or exact visual fidelity for every input.

Batch-7 isolation showed that the untouched Excel source also repairs, and that PowerPoint accepts
the decoration edits when the invalid synthetic objects are removed. A real Excel blank-style sample
and a guarded PowerPoint fixture both passed the final [two-file Office check](~/Downloads/lossless-check/office-batch-7-followup/CHECKLIST.md),
package/SDK and LibreOffice. The original failing files remain recorded; no production writer change
was needed for these source defects.

Final evidence: `~/corpora/results/lossless-final-984eaf3/README.md`,
`consolidated/summary.json` and `RENDER-EXCEPTIONS.md`. The app docs describe the kept/converted/dropped
content and edit rules. Full runs were completed once for the feature group; the fixture correction
used focused regressions and an Office recheck.

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
3. **Commit bodies.** A few lines, saying three things:
   - what changed for the user;
   - the measured numbers, from a full run or focused, and that the new test fails on the previous
     commit;
   - what is still pending.
4. **AGENTS.md stays short.**
   - Move the six preservation rule paragraphs to a "Preservation" section in docs/suite.md, or a
     new docs/preservation.md that holds the opc.js contracts and the per-app edit matrices.
   - Keep one line in AGENTS.md pointing there.
5. **Full runs before features.** After each app's group of steps, run the full corpus, not only
   focused fixtures. Never let the documented corpus numbers predate the code by more than one step.
   Run it once per feature group and per Office batch, not after every fix; a fix needs its focused
   regression test.

Added after the 2026-10-09 reviews (same quality in less time):

6. **A fix's test must fail on the previous commit.** Run it against the old code before committing.
   Two tests passed on the old code (the frame-orientation and whole-column cases).
7. **New identities and names follow Office.** Before allocating any new ID or name (list `nsid`,
   VML ids and blocks, slicer/cache/query names, control names, creation IDs), look up how
   Office-authored corpus files do it and copy that. Each review found this mistake again.
8. **Notices say exactly what happened.** "Converted" only when the value was really converted;
   otherwise name what was dropped.
9. **The plan's status is updated at milestones.** That means a feature group done, or an Office
   batch prepared or answered, not after every commit. No commits whose only change is plan totals.
10. **Other agents' work in progress.** Don't edit or commit files that hold another session's
    uncommitted changes; ask first, or work around them.

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

The implementation and acceptance work above is complete. Final scope, failed attempts, retries,
source hashes and corrected expectations are recorded in `~/corpora/results/lossless-final-984eaf3/`.
Raw reports remain unchanged; `consolidated/` resolves unique cases with explicit provenance.

| App | Save / draft successes | Text-edit successes | Package/SDK OK / failed / excluded |
|---|---:|---:|---:|
| Quire | 2,880 / 2,880 | 2,281 | 18,056 / 373 / 599 |
| Ledger | 935 / 935 | 933 | 6,371 / 201 / 17 |
| Lectern | 837 / 837 | 548 | 4,914 / 96 / 292 |

- No-edit inventories match the preceding full measurements. Word retains **776,207/776,869**
  counted words; PowerPoint retains all **58,022** words and every inventoried feature in **827**
  readable pairs. Known malformed-input, binding and unsupported-content exceptions remain explicit.
- Excel independently compares **889 saves / 1,649,147 cells**, with **883** exact value/formula
  matches; **887 edits / 1,648,615 cells** have **879** exact matches outside the scripted target.
  The two legacy namespace inputs have separately measured conversion and saved-only validation.
- Convention scans: **193 Word**, **192 Excel**, **163 PowerPoint** signatures. Word adds none;
  Excel's three additions preserve Strict-source picture properties; PowerPoint removes three
  illegal frame-transform signatures. The delta reviews are retained with the reports.
- All Office handoffs are resolved, including the batch-7 replacements. Their eight focused
  package/SDK comparisons and two LibreOffice render/count comparisons pass. The new checker test
  fails on `7f85450`; the added real Excel regression fails on the pre-fix writer `9352691`.
- All four pages retain their passing load checks. Markdown remains **652/652 CommonMark,
  22/22 GFM and 24/24 exact saves with confined edits**. Unchanged suites were not rerun to produce
  another total. Other sessions' work remains outside these commits.

The remaining fidelity exceptions are documented in the app docs and final report; none are
represented as a blanket lossless-save guarantee. Further fidelity work is outside this completed
preservation step, and should begin with a specific failing input rather than another broad rerun.

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
The earlier handoffs are archived; the current copies have passed Excel and PowerPoint.

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

## Review of 1a33733..b8a0af5 (2026-10-09, evening)
The work is careful and the plan is honest. Batches 3–4 were rebuilt after the fixes, and their
differences are explained. Fix these before sign-off, each with a test that fails on the current
code and a `package.py` rule where Office would repair the file:

**Implementation status:** all ten items are complete, including the full refresh and Office
acceptance described above. The focused review's 181 passing tests and 158 package/SDK comparisons
remain in `~/corpora/results/lossless-evening-review-2026-10-09/`; its ten unreadable-original
comparisons remain failures. The subsequent legacy Excel correction independently validates all
14 saved states. The two batch-7 source/fixture failures and accepted replacements are recorded in
`~/corpora/results/office-batch-7-repair-2026-10-09/`. Performance evidence and convention-rule scans
remain with those versioned reports; no unchanged full run is repeated for this closure.

1. **Word: copied list definitions reuse `w:nsid`** (new in `7c25bd7`).
   - Cause: `D.addNum` (`dmodel.js:274`) duplicates a kept `abstractNum`, but `nsid` is not a
     copied identity in `opc.js` `identity()`. Both definitions are written with the same nsid.
   - Reached by Customize then Restart/New list, the Bullets toggle (`A.lastList`) and
     Compare/Insert File (`mergeNumbering`).
   - Fix: give copies a fresh nsid. Add an Office sample with two definitions that came from one
     source.
2. **PowerPoint: rotation still reaches graphic frames.**
   - `M.canRotate` allows charts, and the chart writer puts `rot` and flips on `p:xfrm`
     (`pptx-write.js:89`, `357`).
   - Ungrouping a rotated group that contains a preserved frame adds `g.rot` with no `canRotate`
     guard (`app.js:994`).
   - Disable rotation for charts too (PowerPoint has none for them). Keep frames unrotated when
     ungrouping, or convert them with a notice. Correct the docs/lectern.md row.
3. **Excel: the pre-release (2005/8 schema) conversion is partial** (`3a93371`). Reproduce with
   `openxml-sdk__9f7806671fc8__ProjectStatusReport_TP10094814.xltx`.
   - Row `ht` is in twips but is written as points (375 → 20× too tall; 657 is over 409.5).
   - `defaultRowHeight="300"` becomes 12.75.
   - Zero-based `<col min="0" defaultWidth=…>` columns are dropped.
   - The notice says the sizes "were converted".
   - Convert rows and columns too, or report exactly what is dropped. Add an Office sample, or
     list it as pending.
4. **Shared pre-scan parses ordinary text** (`opc.js:924`). `\b(?:ins|del|moveFrom|moveTo)\b`
   matches the words in sheets, slides and sharedStrings (Spanish or Italian text), so they are
   fully parsed on every save for nothing. Match only `w:`-prefixed tags, or Word parts only.
   Recheck the large-sheet benchmark.

Smaller:
5. **Lectern orientation test.** It passes on the code before `0bcaec5`: `L.unionBounds` has no
   `rot`, so the multi-member path never wrote orientation. Add a fixture with one model member over
   an AlternateContent Choice/Fallback whose shapes differ in rot and flips; that case was the real
   bug.
6. **Lectern flips.** Moving a single-shape preserved frame resets an existing flip to 0:
   `F.placeholder` (`frames.js:33`) reads `rot` but not `flipH`/`flipV`.
7. **Lectern extension order.** Slide `p:extLst` is reordered on every save with modern comments.
   `properties.js:240` deletes, then appends, `commentRel`/`creationId`; blank them in place
   instead. Also test the path that creates a new `p14:creationId` (the fixture always has one).
8. **Ledger ActiveX copies.** Copied ActiveX controls get numeric VML `v:shape id` values and names
   like `cmdOK 1`. Office uses the control name as the VML id, and a space is not a valid VBA
   identifier. Follow Office (`objects.js:105`).
9. **Quire index with columns.** Inserting an index with columns moves the section's page break: the
   new section is `continuous` and the original start type moves after the index
   (`fields.js:656`).
10. **Plan hygiene.**
    - List the pre-release Excel conversion and the master/layout decoration edits as Office-pending.
      Neither is in a batch.
    - Update the Excel figures under "Current execution" (`ledger-final-3716a3b`/187 vs
      `3cb0cd2`/189 in docs/ledger.md).
