# KuTools Audit — mistakes found, improvements applied

Three parallel audits of `kutools-source/` (cleanup loops, footnote conversion,
robustness). Rule: never port a KuTools behavior without asking whether it is
correct first. Findings below; "Applied" = changed in this repo.

## Bugs with fixes applied

| # | KuTools mistake | Location | Applied |
|---|---|---|---|
| 1 | `^p`/`^l` replace and empty-para delete run inside tables, corrupting cell structure (no `wdWithInTable` guard in space/LF handlers) | `Connect.cs` space + LF/PAR handlers | `removeEmptyParagraphs` and `removeHorizontalLines` now skip in-table paragraphs via `parentTableOrNullObject`; `paragraphMarksToLineBreaks` refuses table selections |
| 2 | `while(true)` probe+ReplaceAll loops with no iteration cap; probe/replace scope disagreement → premature exit or potential infinite loop on self-matching patterns | `Connect.cs:4907/5085` (`_0084_0006`) | Our transforms are single-pass by construction (regex/OOXML, no probe loops); no loop without a bound exists in `src/` |
| 3 | Collapsed caret (`wdSelectionIP`) escalates to whole-document wipe | all cleanup handlers | `getSelectionOrThrow` rejects empty selections; selection-mode ops never fall through to document scope |
| 4 | `Endnote→Text` synthesizes markers from loop index (`Roman(i)`), covering 2 of ~9 number styles; `>3999` throws and aborts | `Connect.cs:3451`, `--\--.cs:4067` | Deferred to Phases 25–27: read the actual marker at convert time; cover all number styles; per-note `errors[]` (see below) |
| 5 | `.Text`-only transfer drops formatting/hyperlinks/fields/images | `Connect.cs:3418/3452` | Deferred: `insertFootnote` takes text only (API limit) — documented; block text captured verbatim, never re-typed char-by-char |
| 6 | `wdToggle` (9999998) assumes superscript-ON; fragile state-dependent flip | `Connect.cs:3468` | `setSuperscript/setSubscript` set both flags explicitly; toggles read-then-invert, never toggle constants |
| 7 | 188 catches, ~3 surface errors; half-converted docs with no message | `Connect.cs` throughout | `safeWordRun` returns `{error}` on every module fn; panels surface via `setError`; counts returned (`deleted`, `converted/failed`) |
| 8 | No undo grouping (N undo steps per click); selection clobbered by `TypeText`/`GoTo` | throughout; `Connect.cs:464-487` | One `Word.run` + batched `sync()` per command; range-based inserts, no `select()`/cursor moves anywhere in `src/` |
| 9 | Dead WdColor branch (signed `>=` vs `0xFF000000` makes Yellow→Black guard unreachable) | `--\--.cs:3852` | Not ported; font color passes through as hex/`"auto"` |
| 10 | `AdjustParagraphEnd` heuristic near-dead for CJK (`lastLen <= lines`), width-blind char counts, whole-para direct-format side effects | `Connect.cs:458-550` | Not ported — no auto font-shrink feature; widow/orphan control only |

## Risks noted, no action (out of scope or API-limited)

- Sticky `Find` state / missing `ClearFormatting`+`Wrap` in selection helpers — we use no `Find.Execute` loops.
- `SelectEmptyParagraphs` table-stuck popup and `+len(StyleName)` skip — no such feature planned.
- `Footnote→Endnote` drops numbering style/separators; conversion without `UndoRecord` — informs converter verification step (re-scan + count check after convert).
- `FormFont` preview-vs-paint mismatch; `FormDocFormatConverter` dead format case, non-bitmask attribute check — UI-phase cautionary tales (single render path, no dead options).
- `CanRunAction` silent veto — panels must explain disabled states (Phases 16/24).
- Whole-document ops must confirm with match counts in UI (Phases 16/24), never silent-escalate.

## Converter requirements hardened (for Phases 25–27)

1. Re-scan at convert time (no stored ranges); descending `paragraphIndex` order.
2. Per-note try/catch → `errors[]`; one bad note never aborts the run.
3. Two-phase with verification: insert descending → delete descending → re-scan, count-check, abort-with-errors on mismatch (no duplicate window as in `Connect.cs:3417/3423`).
4. Explicit `superscript:false/subscript:false` on inserted markers; save/restore selection.
