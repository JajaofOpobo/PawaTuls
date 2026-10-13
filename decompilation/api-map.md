# KuTools COM Interop → Office.js API Map

Derived from `kutools-source/` (ilspycmd 9.0, `Kutools/Kutools.dll` v7.7.102.0) + raw
metadata survey. External calls are SmartAssembly House-of-Cards proxied
(`global::_XXXX._007E_YYYY`), so COM members are proven via proxy-delegate shapes in
`--.cs`/`-.cs` + `MemberRef` table + call-site context. See `FINDINGS.md` for code refs.

| # | Operation | KuTools COM call (VSTO) | Office.js equivalent |
|---|---|---|---|
| 1 | Bold set | `Selection.Font.Bold = 1` (int; `Connect.cs:11667`) | `range.font.bold = true` |
| 2 | Italic set | `Font.Italic` int setter | `range.font.italic = true` |
| 3 | Underline | `Font.Underline = WdUnderline.wdUnderlineNone` to clear (`Connect.cs:703`) | `range.font.underline = Word.UnderlineType.*` |
| 4 | Strikethrough | `Font.StrikeThrough` int setter | `range.font.strikeThrough` |
| 5 | Superscript | `Font.Superscript` (bool; `= wdToggle` once at `Connect.cs:3468`) | `font.superscript = true` + `font.subscript = false` |
| 6 | Subscript | `Font.Subscript` int setter | `font.subscript = true` + `font.superscript = false` |
| 7 | AllCaps / Hidden | `Font.AllCaps`, `Font.Hidden` (`9999999` = inherit on ListLevel) | `font.allCaps`; `font.hidden` |
| 8 | Font color | `Font.Color = WdColor` (only bullet `wdColorGray40` found) | `font.color = "#RRGGBB"` / `"auto"` |
| 9 | Font name / size | `Font.Name`, `Font.Size` (float) | `font.name`, `font.size` |
| 10 | Character spacing | `Font.Spacing` float (tighten −0.8/−0.5/−0.2) | `font.spacing` |
| 11 | Symbol insert | `Selection.InsertSymbol(CharNumber, Font, Unicode, Bias)` | No direct equivalent — `insertText()` the unicode char |
| 12 | Alignment | `Paragraph.ParagraphFormat.Alignment = WdParagraphAlignment` | `paragraph.alignment = Word.Alignment.*` |
| 13 | Remove indents | `FirstLineIndent/LeftIndent/RightIndent = 0f` (+`CharacterUnit*` twins) | `= 0` on each indent property |
| 14 | Keep with next | `ParagraphFormat.KeepWithNext` | `paragraph.keepWithNext` |
| 15 | Find / Replace | `Find.Text`, `Replacement.Text`, `Find.Execute(..., wdReplaceAll)`, `Wrap`, `MatchWildcards` | `body.search()` + per-hit `insertText(replace)`; no replace-all primitive |
| 16 | Reset find state | `Find.ClearFormatting`, `Replacement.ClearFormatting` | No equivalent needed — searches are fresh objects |
| 17 | Selection clear | `Selection.ClearFormatting` (MemberRef exists, site unproven) | No equivalent — reset properties individually |
| 18 | Clear char/para all | **Absent** — KuTools resets 14 `ListLevel.Font` props individually | Individual resets (plan approach confirmed) |
| 19 | Remove empty paras | `Find "^p^p" → "^p"`, `wdReplaceAll` | No direct equivalent (empty paras unsearchable) — reverse-iterate `body.paragraphs` and `delete()` |
| 20 | Remove empty lines | `Find "^l^l" → "^l"`, `wdReplaceAll` | Search `"\v"`, reverse-iterate and delete/merge |
| 21 | LF ⇄ PAR | `Find "^l" ⇄ "^p"` swap replace | `"\v"` → `"\n"` via `insertText`; reverse for selection |
| 22 | Collapse spaces | Looped `Find.Execute` probes + `wdReplaceAll` until clean | `text.replace(/ {2,}/g, " ")` + write back |
| 23 | Case conversion | **Absent** — no `ChangeCase`/`ToUpper`/title-case anywhere | JS string transforms (plan approach; no KuTools precedent) |
| 24 | Mixed selection | `9999999` (`wdUndefined`) **written** = inherit; never read-checked; no toggle logic | `font.bold === null` → set `true`, else invert |
| 25 | Toggle constant | `9999998` (`wdToggle`) used once (superscript off) | No equivalent — read + invert |
| 26 | Footnote → endnote | `Footnotes.Item(i)` copy fwd, `Endnotes.Add(range)`, `Footnote.Delete` bwd | No equivalent needed (not a PawaTuls feature) |
| 27 | Endnote → text | `Endnotes` fwd transcribe, `Endnote.Reference.Delete` bwd | No equivalent needed |
| 28 | Insert footnote | **Absent** — no `Footnotes.Add` proxy exists | `markerRange.insertFootnote(text)` (novel to PawaTuls) |
| 29 | Scan inline markers | **Absent** — no superscript/bracket digit scan | `body.search()` wildcards (novel to PawaTuls) |
| 30 | Scope selection-vs-doc | `Selection.Type == wdSelectionNormal ? selection : whole document` + confirm prompt | `getSelectionOrThrow` + explicit `applyTo` params |
| 31 | Batching | `Application.ScreenUpdating = false` wrap + per-item `try/catch` | Batch work, single `await context.sync()`; `safeWordRun` catches |
| 32 | Remove para borders | `Borders(wdBorderBottom).LineStyle = wdLineStyleNone` | `paragraph.getBorder(Word.BorderLocation.bottom).type = "none"` |
| 33 | Table convert/diagonal | `ConvertToText`, `ColorFormat.RGB`, `Sort`, `Tables.Item.Delete` (bwd) | Out of PawaTuls scope except convert module patterns |

Notes:
- WdColor normalization helper (`--\--.cs:3852`) is border/shape color code, not font color:
  `Automatic → Black`, mask `& 0xFFFFFF`, yellow-on-white guard. Not ported.
- No `SpaceBefore/SpaceAfter/LineSpacing/LineSpacingRule/WidowControl/PageBreakBefore`
  members exist in the DLL — KuTools never touches paragraph spacing; our spacing APIs
  come from Office.js semantics alone.
- `HighlightColor`/`SmallCaps` on `Font`: absent (only `Find.HitHighlight` search-UI).
