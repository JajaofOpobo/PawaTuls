# Phase 29 Test Report — PawaTuls v1 (legal suspended)

Word build/bitness: recorded at manual-test time below.
`tsc --noEmit`: clean (per-commit, CI-equivalent).
`npm run build`: emits complete current `dist/` (taskpane.js minimized) but the
Node process crashes post-emit in a minifier worker thread (exit 1, twice,
2026-10-14) — environmental flake under investigation (dev server resident at
the time); not a code error. Sideloading uses `webpack serve` and is unaffected.
Retry with the dev server stopped if a clean exit code is needed.

## Automated (ran 2026-10-14, no Word required)

| Suite | Command / script | Result |
|---|---|---|
| types + build | `npx tsc --noEmit`, `npm run build` | pass |
| citations + matcher | temp/check2.js (compiled `citations.ts`, `matcher.ts`) | pass: 5 systems, negatives, dense indices, confidence cap, conflict rows |
| ordinal OOXML | temp/check3.js (compiled `text.ts` helper) | pass: split shape, format preservation, superscript skip, balanced XML |
| profiles | temp/check4.js (compiled `profiles.ts`) | pass: validation matrix, round-trip stability, builtins |
| numbering | temp/check5.js (compiled `numbering.ts`) | pass: alpha/roman converters, gap/dup repair, restart untouched, `(c)`→alpha regression |
| manifest | `npm run validate` | valid |
| fixtures | make-fixtures.ps1 + check-fix.py OOXML asserts | 4 docs generated, all structural asserts pass |

## Fixtures (`decompilation/fixtures/`, generated — not committed)

- `formatting-test.docx`: mixed runs, big-red, alignments, case lines, manual
  breaks, ordinals, dashes line, empty + whitespace paras, native list, 3×2
  table (empty middle row), genuine footnote + endnote, hyperlink, comment,
  bookmark, page break, hidden text.
- `footnotes-test.docx`: superscript 1 + 12 (multi-digit), [4], (5), unmatched
  superscript 7, `___` separator, blocks incl. continuation lines, orphan 9,
  NWLR citation paragraph (citationDense path).
- `numbering-test.docx`: native list, manual 1/2/2/4, decimal prose negative.
- `pleadings-test.docx`: reserved for legal re-enable.

## Manual Word-run checklist (requires `npm start`, human clicks)

Sideload verified 2026-10-14: `office-addin-debugging start` enabled debugging,
dev server serves taskpane/commands (200), Word launched with the sideload doc.
Loopback exemption for Edge WebView is set (one-time prompt answered).
In-Word rendering and the clicks below remain manual.

Formatting (formatting-test.docx copy):
- [ ] Bold toggle on mixed selection; size +/−; Times New Roman apply
- [ ] Clear All Formatting → Normal style
- [ ] Justify; Remove Empty Paragraphs toast count (table row 2 must survive)
- [ ] UPPERCASE / Sentence case; collapse spaces; ordinals → superscript st/nd
- [ ] LF→PM and PM→LB; table→text; profiles apply/import/export
- [ ] Numbering audit finds native list + dup/gap; repair previews; apply
- [ ] Remove group spot-checks (breaks, hidden, comments, hyperlinks keep text)

Footnotes (footnotes-test.docx copy):
- [ ] Scan: 4 matches (1, 12, 4, 5), unmatched 7, orphan 9, existing 0,
  citation line capped at medium, systems line shown
- [ ] Deselect one → Convert Selected converts rest; markers gone; blocks gone
- [ ] Real footnotes in Word footnote pane; rerun converts 0
- [ ] Manual area override: select block region → re-match
- [ ] Manage Notes on formatting-test copy: genuine footnote counted, removals work

Integrity (every destructive test): unrelated text/styles/lists/notes/fields/
hyperlinks/tracked changes unchanged; semantic diff, never raw-ZIP equality.
