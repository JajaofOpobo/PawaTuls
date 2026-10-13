# Integration Notes — KuTools findings → Office.js implementation

Source: `FINDINGS.md` + `api-map.md` + `grep-*.txt` (ilspycmd 9.0 on `Kutools.dll` v7.7.102.0).
Read this before implementing Phases 10, 11, 13, 15. Caveat: SmartAssembly
House-of-Cards proxies external calls and encrypts Find-pattern strings, so exact
Find literals and some member identities are inferred from handler shape + supertips,
not proven. Anything marked UNCERTAIN below must not be treated as ground truth.

## 1. Mixed selection handling

KuTools does **not** read-check `Font.Bold == 9999999` and has **no bold/italic/underline
toggle logic at all** in this build — it writes values directly (`Bold = 1`,
`Underline = None`). The only `9999999` writes set `ListLevel.Font` properties to
"inherit/clear" when building bullet galleries (`Connect.cs:692-712`); the only
`9999998` (`wdToggle`) use turns superscript off after typing a footnote marker
(`Connect.cs:3468`). Office.js equivalent (plan stands, no KuTools change):
a mixed property reads back `null` (NOT `9999999` — that is COM-only).
`toggleBold` and friends: load `font.bold`; if `null`, set `true`; else invert.
Never use a toggle constant — Office.js has none.

## 2. Iteration order for deletions

Confirmed: **every** destructive KuTools loop iterates **backward**
(`for (i = Count; i >= 1; i--)` — footnotes `Connect.cs:3420`, endnotes `:3462`,
fields `:3500/:3507`, tables `:4668`, shapes `:6079`, comments `:5703`, hyperlinks
`:993+`). Forward copy first where content must survive (footnote→endnote `:3409`
copies forward, then deletes backward). The plan's reverse-order rule
(`items.slice().reverse()`) matches exactly — keep it for `removeEmptyParagraphs`,
block deletion in the footnote converter (descending `paragraphIndex`), and
`removeHorizontalLines`. Divergence noted: KuTools removes empty paragraphs via
`Find "^p^p → ^p" (wdReplaceAll)`, NOT paragraph iteration — but Office.js cannot
search for empty paragraphs, so reverse iteration remains the correct Office.js
workaround (documented in `api-map.md` #19).

## 3. Clear formatting implementation

KuTools does **not** call `ClearCharacterAllFormatting` / `ClearParagraphAllFormatting`
— both are absent from binary and source. It resets properties individually
(14 `ListLevel.Font` props + Underline/Color/Size/Name in the bullet builder) and uses
`Find.ClearFormatting` / `Selection.ClearFormatting` only to reset *search* state.
The plan's approach (individual resets in `clearCharacterFormatting`,
`removeAllIndents` + `removeAllSpacing` + single spacing, `applyStyle("Normal")` for
"all") matches KuTools. No change.

## 4. Case conversion implementation

KuTools has **no case-conversion feature at all** — no `ChangeCase`, no `ToUpper`/
`ToLower`/title-case usage, no ribbon button. There is no KuTools pattern to follow.
The plan's approach (JavaScript string transforms + `insertText(replace)`) stands on
its own and is the only viable Office.js path (Office.js exposes no `Range.Case`).

## 5. Findings that change the planned implementation

1. **Footnote detection/conversion is novel — no KuTools precedent.** KuTools only
   operates on existing `Footnotes/Endnotes` collections (→endnote, →text, remove);
   it never scans body text for inline markers, never calls `Footnotes.Add`, never
   touches separators. Phases 25–27 design is unconstrained by decompilation.
2. **Empty-paragraph removal differs, plan still correct for Office.js** (see §2).
   Also handle whitespace-only paragraphs under `collapseMultipleSpaces`/trim rather
   than the empty-para remover — KuTools treats them as separate ops (`btnRemoveSpaces`).
3. **Indent removal sets BOTH point and CharacterUnit properties** (6 setters for 3 axes
   — covers CJK `UseCharacterUnit` documents). Office.js `Word.ParagraphFormat` has no
   CharacterUnit properties, so plain point sets are the complete Office.js equivalent.
   No action, but do not go looking for a CharacterUnit API.
4. **KuTools never touches paragraph spacing** (`SpaceBefore/After`, `LineSpacing*`
   members absent). Our spacing APIs have no KuTools edge cases; follow Office.js docs.
5. **Scope pattern to adopt:** KuTools branches every destructive op on
   selection-vs-document with a confirm prompt. Our `applyTo: "selection" | "document"`
   params (pleadings) and `getSelectionOrThrow` cover this; keep the confirm/prompt
   behavior in UI panels, not modules.
6. **Batching pattern:** KuTools wraps bulk ops in `ScreenUpdating=false` + per-item
   `try/catch{}`. Office.js equivalent: do all loads/mutations, then a single
   `await context.sync()`; per-item failures in converters accumulate into the
   `errors[]` result instead of throwing.
7. **`WdColor` helper is a red herring** (`--\--.cs:3852` feeds shape-border RGB, not
   font color; likely-unsigned comparison artifact — do not replicate). Font color
   passes through as hex; `"auto"` maps to automatic.
