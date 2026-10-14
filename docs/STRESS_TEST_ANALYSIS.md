# Stress-Test Analysis — 15-page before/after docs (2026-10-14)

Method: OOXML diff of `docs/Word_Toolkit_15_Page_Stress_Test*.docx`
(274 → 212 paras, 20714 → 19721 chars). User applied features manually;
attribution below is evidence-graded: **proven** (mechanism visible in XML),
**likely** (consistent, one explanation), **unknown** (untestable headlessly).

## Proven harmful (fixed this round)

1. **Footnote duplicates + body-as-footnotes.** Eight footnotes share one body
   text ("1. INTRODUCTION…", ids 1,2,10–15); pairs duplicated (3/4, 5/6, 7/9);
   body parentheticals (1)–(4) converted; source body paras deleted.
   Causes: all matches defaulted accepted (Accept-All converted mediums),
   no one-to-one enforcement (8 markers × 1 block), no block hygiene
   ("1. INTRODUCTION" heading matched the decimal block pattern).
   Fixes: high-only auto-accept, both-directions one-to-one (extras unchecked),
   block hygiene (headings, native lists, table cells excluded + chain breaks).
2. **Case transforms wiped direct bold** (39 paras, bold-only drops, italic kept).
   Cause: `Range.insertText(Replace)` regenerates runs. Fix: `rewriteTextPreservingRuns`
   OOXML surgery (positional for case, deletion-exact for trim/collapse),
   applied to transformCase/trim/collapse/leading-tabs + repair prefix edits.

## Proven working

- `removeEmptyParagraphs` 23 → 16 (protected/table/structural paras correctly kept).
- `tableToText` (2 tables → 0, cell text preserved).
- `Remove All Indents` (manual-numbered paras lost `w:ind`).
- Footnote mechanics (insert → verify → delete) functioned; genuine note untouched.

## No observable effect (untested vs silently failed — unknown)

Collapse spaces (8 triple-runs remain), page-break removal (14 remain),
hyperlinks/bookmarks/pictures untouched, ordinals not superscripted, numbering
dup+gap still present, line spacing `w:line="552"` unchanged. If any of these
WAS clicked, report it — the panel now toasts counts/errors per action, so a
rerun distinguishes "not clicked" from failure. Line-spacing-single is first
suspect for a real failure (value mapping unverified in Word).

## Benign / user-caused

- `endnotes.xml` skeleton (Word side effect of footnote insertion; no content).
- Title "15-Page" → "5-Page" (manual edit).
- ~62 removed paras only partly attributable (blocks, tables, empties, edits).

## Retest protocol

On a fresh copy: scan (expect mediums unchecked, citation cap, hygiene skips) →
check 1–2 rows → convert → verify counts → rerun converts 0. Case-convert a
bold/italic-mixed selection → emphasis must survive.
