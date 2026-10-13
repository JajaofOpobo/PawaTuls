# Decompilation Findings — KuTools for Word 8.2.119.0

Source binary: `../wordkutools_v8.2.exe` (Kutools for Word Setup, 17.7 MB).
Installed to: `C:\Program Files (x86)\ExtendOffice\WordKutools\`
Reference document: `kutools-test.docx` (created by `make-testdoc.ps1` — Heading 1,
3 body paragraphs, 1 mixed bold/italic paragraph, 1 two-by-three table).

## Phase 1: Behavioral inspection status — PARTIAL (automation infeasible)

- KuTools 8.2 installed silently (`/VERYSILENT`), registered as `WordKutools.Connect`
  + `WordKutoolsCheck.Connect` (HKCU + HKLM) and connects in headless Word (`Connect=True`).
- `probe-kutools.ps1` proves KuTools exposes **no automation API** (`COMAddIn.Object` is null)
  and creates **no CommandBar controls** (Ribbon-XML UI only). Its buttons cannot be
  driven headlessly — the 20+ before/after operation observations in the plan require
  interactive Word UI clicks and remain a **manual step** (open `kutools-test.docx`,
  copy per test, unzip and diff `word/document.xml` before/after each op).
- Decision: the decompiled COM-interop calls (Phases 2–3 below) are the ground truth
  for the Office.js port; manual XML observations can be backfilled for validation only.

## Installation layout (largest DLLs)

| Path | Size | Role |
|---|---|---|
| `Kutools/Kutools.dll` | 1973 KB | Main managed assembly (.NET Fx 2.0, `v2.0.50727`), decompiled in Phase 2 |
| `Kutools4/Kutools.dll` | 2056 KB | .NET 4.x build variant of the same assembly |
| `Kutools/Microsoft.Office.Interop.Word.dll` | 762 KB | Word COM interop assembly KuTools codes against |
| `Kutools/Office.dll`, `Utils.dll` | 406/69 KB | Shared Office + utility assemblies |
| `x64/WordKutools.dll`, `x86/WordKutools.dll` | 6037/4023 KB | Native COM shims (regsvr32), not decompilable with ILSpy |
| `x64/JMV.vmp.dll` (+ `JMA.dll`) | 5340 KB | VMProtect licensing VM — out of scope |

## Config-file intelligence (`Kutools/Kutools.dll.config`)

User settings leak feature defaults worth cross-checking against decompiled code:

- `ShortParaMaxChars = 25` — "short paragraph" threshold used by merge/split paragraph tools.
- `SingObjParaWordObj = True`, `HeadParaEmpPara = True` — paragraph-selection options.
- `ConvertTableToTextSeperator = "-"` — table-to-text default separator.
- Caption defaults: table/figure horizontal = 1, vertical = 1, equation vertical = 2;
  empty prefix/suffix defaults (user fills per document).
- `HighlightSameWords = True`, AutoCorrect/MathAutoCorrect defaults True.
- `SnapshotPrompt = True` — KuTools prompts before snapshot/undo-style operations.

## Phase 2: Decompilation record

- `ilspycmd 9.0.0.7889` (pinned; unversioned latest fails with a bad-package error) on
  `Kutools/Kutools.dll` → `decompilation/kutools-source/` (**68 `.cs` files**, gitignored).
- Obfuscation: **SmartAssembly, heavy but navigable.** UI/form classes readable
  (`FormFont`, `KutoolsAddIn`, `Connect` dispatch, full ribbon XML in `--.--.resx`
  with button ids/supertips). Core logic in `--.cs`/`-.cs`/`--\--.cs` with
  House-of-Cards proxied external calls (`global::_XXXX._007E_YYYY`) + encrypted
  Find-pattern strings (`_0098(N)`). External API *types* stay readable via proxy
  delegate signatures + raw `MemberRef` table. `Kutools4/` variant not decompiled
  (same feature set, newer framework; string survey only). Native shims + VMProtect
  licensing DLLs out of scope.
- Phase 3a greps on source: `decompilation/grep-{font,paragraph,style,find,footnotes,delete,case}.txt`
  (`delete`/`case` empty — findings in themselves, headers explain).

## Phase 3: analysis summary (see `api-map.md` for the full table)

- **Font:** direct writes only, no toggles, no mixed read-checks (`Bold = 1`,
  `Underline = None`, `Superscript = wdToggle` once). `9999999` = inherit-write on
  `ListLevel.Font`, never a comparison. No `Selection.Font.Color/Highlight/SmallCaps`
  sets. `Clear*AllFormatting` absent → individual resets.
- **Paragraph:** indent removal = `0f` on First/Left/Right (+CharacterUnit twins);
  no spacing/line-spacing members exist. Empty paras/lines removed via
  `Find "^p^p→^p"` / `"^l^l→^l"` (`wdReplaceAll`), never paragraph iteration.
  All destructive loops backward (`Count..1`). LF⇄PAR = swapped `^l`/`^p` replace.
  Collapse-spaces = looped Find probes until clean. `AdjustParagraphEnd` shrinks
  last-line font (−0.8/−0.5/−0.2pt by script) instead of reflowing.
- **Case conversion: feature absent.** No `ChangeCase`/casing helpers/buttons.
- **Footnotes:** only Footnote→Endnote (copy fwd + `Endnotes.Add` + delete bwd),
  Endnote→Text (transcribe fwd + `Reference.Delete` bwd), remove buttons
  (impl. unrecoverable — likely Find `^f`/`^e`), hyperlink-strip in note stories.
  No marker scanning, no `Footnotes.Add`, no separator handling. Full ribbon/button
  inventory in the footnote sub-report scope.
- **Edge cases documented:** mixed-write vs mixed-read (§1 of integration notes),
  reverse iteration universal, per-item `try/catch` + `ScreenUpdating=false` batching,
  selection-vs-document scope branches with confirm prompts, `WdColor` helper is
  shape-border code (not font), yellow-guard artifact.

## Interim static analysis (pre-ilspycmd, byte-level, provisional)

Reflection inventory failed: `GetTypes()` throws `ReflectionTypeLoadException` with
empty `LoaderExceptions` even for `Utils.dll` — the .NET 5.1 reflection-only path
cannot enumerate these assemblies (consistent with SmartAssembly metadata tampering;
the literal string `SmartAssembly` is present in `Kutools.dll`). ILSpy reads raw
metadata and usually tolerates this — pending SDK install.

Namespace/form inventory (readable — **not** name-obfuscated): `Kutools.Connect`,
`Kutools.Dialogs`, `FormFont`, `FormDocFormatConverter`, `FormConvertTableToText`,
`FormBatchInsertCaptions`, `FormToCaption`, `FormSelectPages/SelectOptions/SelectLastParaLines`,
`FormDisplay`, `FormHyperlink(Manager)`, `FormDictionary`, `FormAutoCorrect`,
panes `UCAutoTextPane/UCBookmark/UCCaptionPane`, `FormExportImages/ExportToImages/ExportTableToImages`,
`FormDocumentSplit/Merge`, `FormDiagonalHeader`, `FormInsertPictures`, `FormWebPage`.

Word-interop member references found (`decompilation/grep-*.txt`):
- Font: Bold/Italic/Underline/StrikeThrough/Superscript/Subscript/AllCaps/HighlightColor, `WdUnderline`
- Paragraph: ParagraphFormat/Alignment/First/Left/RightIndent/KeepWithNext, `WdParagraphAlignment`
- Styles: `get_Style`, `WdStyleType`
- Find: Replacement/HitHighlight/`WdFindWrap`/FindText
- Footnotes: `Footnote`, `Footnotes` referenced — but **`InsertFootnote` absent**
- Range model: Range/Paragraphs/Words/Characters/Select/Collapse/MoveEnd/`wdStory`/Content,
  TypeText/TypeParagraph/InsertAfter, Sort/ConvertToText/Tables/Cell/PageSetup/Borders
- **`wdUndefined` + literal `9999999` present** → KuTools checks mixed-selection state
- `ClearCharacterAllFormatting` / `ClearParagraphAllFormatting` **absent** → likely resets
  properties individually (matches plan hypothesis; confirm post-decompile)
- `WdCharacterCase`/ChangeCase **absent** → likely string processing in code, not Word's
  Case property (confirm post-decompile)

Caveat: SmartAssembly encrypts string *literals*, so late-bound call names could hide;
metadata member names above are not encryptable and their presence/absence is meaningful.
