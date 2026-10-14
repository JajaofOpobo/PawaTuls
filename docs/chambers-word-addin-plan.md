# PawaTuls — AGENT IMPLEMENTATION PLAN

**Version:** 2.0  
**Target platform:** Windows, Microsoft Word for Microsoft 365 (desktop)  
**Minimum Word API:** WordApi 1.5  
**Runtime:** Office.js, React 18, TypeScript 5  

---

## HOW TO USE THIS PLAN

1. Open OpenCode: `opencode`
2. At the start of every session, paste the **MASTER CONTEXT BLOCK** (Section 0) as your first message.
3. Then paste the **Resume Prompt** for the current phase.
4. Consult **Agents.md** to decide which model to load before each session.
5. After every phase, manually run `npm start` and test in Word before marking the phase complete.
6. Tick every item in the **Completion Test** before advancing.
7. Phases are sequential. Do not skip.
8. Appendices D–H (merged from the v1 plan) are normative: new Phases 31–34 build
   new features, and Appendix H amends several unbuilt phases — the agent running
   an amended phase must read Appendix H first.

---

## SECTION 0: MASTER CONTEXT BLOCK

> Copy everything between the fences and paste it as your **first message** at the start of every OpenCode session.

---

```
You are helping build a Microsoft Word task pane add-in called "PawaTuls".

Stack: Office.js (WordApi 1.5 minimum), React 18, TypeScript 5, Fluent UI v9.
Platform: Windows, Word for Microsoft 365 desktop only.
Project root: chambers-word-toolkit/

The add-in has three feature areas:
  1. Full Formatting Engine — font, paragraph, styles, text transforms, conversions,
     clear formatting.
  2. Nigerian Legal Document Workflows — pleadings numbering, affidavit formatting,
     exhibit marking, reliefs formatting, cause headings, legal citations, schedules.
  3. Footnote Detection — scans documents for footnotes that were pasted inline from
     PDFs or conversion tools, matches them by number, and converts them to proper
     Word footnotes using Range.insertFootnote() (WordApi 1.5).

Early phases of this project also involved decompiling KuTools for Word (a VSTO
add-in) to understand its implementation patterns. The findings from that decompilation
are documented in decompilation/FINDINGS.md. Consult that file whenever you are
implementing a feature that KuTools also provides.

Coding rules that apply to every file you write:
- Wrap all Word object model calls inside Word.run(async (context) => { ... })
- Always call await context.sync() after loading properties and after mutations
- Always handle the empty-selection case with a clear user-facing error message
- Export all functions individually — no default exports from module files
- Use types from src/shared/types.ts; do not redefine them elsewhere
- Use helpers from src/shared/wordUtils.ts; do not duplicate them

The plan document is the single source of truth for what to build and how.
Read the current phase in full before writing any code.
```

---

---

## SECTION 1: TECH STACK

| Decision | Choice | Rationale |
|---|---|---|
| Add-in type | Office.js Task Pane | Windows-only target; covers full required API surface |
| Language | TypeScript 5 | Type safety; agents produce more reliable TS than plain JS |
| UI framework | React 18 | Component reuse; Footnote Detection UI requires controlled state |
| Component library | Fluent UI v9 | Matches Word's native visual design |
| Build tool | Webpack (Yeoman default) | Standard for Office.js; well-documented |
| Manifest format | XML v1.1 | Wider tooling support than Unified Manifest; Yeoman default |
| Package manager | npm | Avoids version conflicts with Office Addin CLI |
| Node version | 18 LTS | Minimum required by current Yeoman Office generator |
| State management | React useState / useReducer | No external state library needed at this scale |

---

## SECTION 2: REPOSITORY STRUCTURE

```
chambers-word-toolkit/
├── manifest.xml
├── package.json
├── tsconfig.json
├── webpack.config.js
├── .eslintrc.json
├── decompilation/                ← Phases 1–3 outputs
│   ├── FINDINGS.md
│   ├── kutools-source/           ← ilspycmd decompile output
│   └── api-map.md
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── src/
    ├── commands/
    │   ├── commands.html
    │   └── commands.ts
    ├── shared/
    │   ├── types.ts              ← Phase 7
    │   └── wordUtils.ts          ← Phase 8
    ├── formatting/
    │   ├── font.ts               ← Phase 10
    │   ├── paragraph.ts          ← Phase 11
    │   ├── styles.ts             ← Phase 12
    │   ├── text.ts               ← Phase 13
    │   ├── convert.ts            ← Phase 14
    │   └── clear.ts              ← Phase 15
    ├── legal/
    │   ├── pleadings.ts          ← Phase 17
    │   ├── affidavit.ts          ← Phase 18
    │   ├── exhibit.ts            ← Phase 19
    │   ├── reliefs.ts            ← Phase 20
    │   ├── causeHeading.ts       ← Phase 21
    │   ├── citations.ts          ← Phase 22
    │   └── schedule.ts           ← Phase 23
    ├── footnotes/
    │   ├── scanner.ts            ← Phase 25
    │   ├── matcher.ts            ← Phase 26
    │   └── converter.ts          ← Phase 27
    └── taskpane/
        ├── taskpane.html
        ├── index.tsx
        ├── taskpane.css
        └── components/
            ├── App.tsx               ← Phase 9 (modify)
            ├── FormattingPanel.tsx   ← Phase 16
            ├── LegalPanel.tsx        ← Phase 24
            ├── FootnotePanel.tsx     ← Phase 28
            └── shared/
                └── ErrorBanner.tsx   ← Phase 9
```

---

---

## PHASE 0: PREREQUISITES

**Prerequisites:** None  
**Resume Prompt:** `Verify all prerequisites for the chambers-word-toolkit project. Follow Phase 0 of the plan exactly.`

### Instructions

Verify each item. Do not proceed until all pass.

1. Confirm Node.js 18 LTS:
   ```
   node --version
   ```
   Must return `v18.x.x`. If not: https://nodejs.org/en/download (LTS tab).

2. Confirm npm:
   ```
   npm --version
   ```
   Must return `9.x.x` or higher.

3. Install Yeoman and Office Add-in generator:
   ```
   npm install -g yo generator-office
   ```

4. Confirm Word (Microsoft 365) is installed. Open Word manually to verify it loads.

5. Confirm VS Code is installed: https://code.visualstudio.com

6. Confirm OpenCode is installed:
   ```
   opencode --version
   ```
   If not: `npm install -g opencode-ai`

7. Create the project root folder:
   ```
   mkdir chambers-word-toolkit
   cd chambers-word-toolkit
   mkdir decompilation
   ```

### Completion Test

- [ ] `node --version` returns 18.x.x
- [ ] `npm --version` returns 9.x or higher
- [ ] `yo --version` returns any version
- [ ] `opencode --version` returns any version
- [ ] Word opens on this machine
- [ ] `decompilation/` folder exists

---

## PHASE 1: KUTOOLS BEHAVIORAL INSPECTION

**Prerequisites:** Phase 0; KuTools for Word installed (trial is sufficient)  
**Resume Prompt:** `Document the behavioral output of KuTools for Word formatting features. Follow Phase 1 of the plan exactly. You will be reading and analysing documents, not writing code.`

### Purpose

Before decompiling any binary, observe KuTools in action. The goal is to record exactly what KuTools does to a Word document for each formatting operation — what property values it sets, what the document XML looks like before and after. This gives you ground-truth reference data that the decompiled code must explain.

### Instructions

#### 1a. Install KuTools for Word

Download from https://www.extendoffice.com/download/kutools-for-word.html  if not already provided.
Install with default settings. A 30-day trial requires no payment.

#### 1b. Create a test document

Create a new Word document named `kutools-test.docx`. Add:
- Three paragraphs of normal body text
- One heading paragraph (Heading 1)
- A short table (2 rows, 3 columns)
- A paragraph with mixed formatting (some words bold, some italic)

Save it. This is your reference document. Do not modify it during testing — make a copy for each test.

#### 1c. Inspect before/after document XML

Word `.docx` files are ZIP archives. After each KuTools operation:
1. Save the modified document as a new `.docx` file
2. Rename to `.zip`
3. Extract it
4. Open `word/document.xml` in VS Code
5. Compare with the pre-operation XML

The relevant XML elements to note:
- Font properties: `<w:rPr>` (run properties) — contains `<w:b/>`, `<w:i/>`, `<w:color>`, `<w:sz>`, `<w:vertAlign>`, `<w:spacing>`, `<w:highlight>`
- Paragraph properties: `<w:pPr>` — contains `<w:jc>` (alignment), `<w:ind>` (indent), `<w:spacing>`, `<w:pStyle>`
- Style definition: `word/styles.xml`

#### 1d. Test and document these operations

For each operation below, record: (a) the OOXML change, (b) any unexpected behaviour.

**Font group:**
- Bold toggle on a mixed selection (some bold, some not)
- Set font color to red (`#FF0000`)
- Set character spacing to 2pt expanded
- Superscript and then subscript on the same text
- Apply highlight (yellow)
- All caps vs. small caps

**Paragraph group:**
- Apply "Justify" to a paragraph
- Set line spacing to "Exactly 14pt"
- Set first-line indent to 1.27cm (0.5 inch)
- Set hanging indent to 1.27cm
- Space before 12pt, space after 6pt
- Remove all indents
- Remove all spacing

**Styles group:**
- Create a new style from a selection
- Apply a named style to a paragraph

**Remove/clear group:**
- Clear character formatting only
- Clear all formatting (font + paragraph)
- Remove empty paragraphs

**Text group:**
- UPPERCASE conversion
- Title Case conversion
- Collapse multiple spaces

#### 1e. Record findings

Create `decompilation/FINDINGS.md` and document each operation in this format:

```markdown
## Operation: [Name]

**Observed behavior:** [What visually happened]

**XML before (relevant excerpt):**
[paste the OOXML snippet]

**XML after (relevant excerpt):**
[paste the OOXML snippet]

**Notes:** [Anything unexpected]
```

### Completion Test

- [ ] `decompilation/FINDINGS.md` exists with entries for all 20+ operations listed above
- [ ] At least one XML before/after pair documented per operation
- [ ] Unexpected behaviors noted

---

## PHASE 2: KUTOOLS DECOMPILATION

**Prerequisites:** Phase 1 complete  
**Resume Prompt:** `Decompile the KuTools for Word assembly using ilspycmd. Follow Phase 2 of the plan exactly.`

### Legal note

Decompiling a commercial application for the purpose of understanding its API usage for interoperability with your own independently written software is permitted under Nigerian copyright law (Copyright Act 2022, Section 43 — acts permitted in relation to computer programs) and under the EU Software Directive principle that has influenced many jurisdictions. You are not copying KuTools code. You are studying it to understand which Word API calls to make in your own independent implementation.

### Instructions

#### 2a. Locate the KuTools for Word installation

KuTools for Word installs as a VSTO add-in. The typical installation path is:
```
C:\Program Files (x86)\Kutools for Word\
```
or:
```
C:\Users\[username]\AppData\Local\Kutools for Word\
```

Find the main DLL. It will be named something like `KutoolsForWord.dll`, `KWforWord.dll`, or `ExtendOffice.Word.dll`. Use:
```
dir "C:\Program Files (x86)\Kutools for Word\" /s /b *.dll
```

There may be multiple DLLs. Identify the largest one — that is typically the main assembly. List them all and note their sizes.

#### 2b. Install ilspycmd

ilspycmd is the command-line version of the ILSpy .NET decompiler. It is free and open source.

```
dotnet tool install -g ilspycmd
```

If .NET SDK is not installed: https://dotnet.microsoft.com/en-us/download/dotnet/8.0

Verify installation:
```
ilspycmd --version
```

#### 2c. Decompile to C# source

Create the output directory:
```
mkdir -p decompilation/kutools-source
```

Decompile the main DLL:
```
ilspycmd -p -o decompilation/kutools-source "C:\Program Files (x86)\Kutools for Word\KutoolsForWord.dll"
```

The `-p` flag outputs to individual `.cs` files organised by namespace. The `-o` flag sets the output directory.

If the DLL name differs from above, substitute the correct filename found in step 2a.

If decompilation reports an error about dependencies, add the DLL's containing folder as a reference path:
```
ilspycmd -p -o decompilation/kutools-source --referencepath "C:\Program Files (x86)\Kutools for Word\" "C:\Program Files (x86)\Kutools for Word\KutoolsForWord.dll"
```

#### 2d. Assess obfuscation level

Open VS Code on the decompiled output:
```
code decompilation/kutools-source
```

Check whether class and method names are readable (e.g., `FormatFont`, `ApplyParagraphSpacing`) or obfuscated (e.g., `a`, `b`, `ᵃ`).

**If readable:** proceed directly to Phase 3.

**If obfuscated:** the code structure is scrambled but the external API calls (Word COM interop calls) remain readable because they reference external assemblies that cannot be renamed. You can still identify the Word API calls even when the surrounding logic is obfuscated. Search for:
```
grep -r "Selection.Font" decompilation/kutools-source/
grep -r "ParagraphFormat" decompilation/kutools-source/
grep -r "Footnotes" decompilation/kutools-source/
```

#### 2e. Also decompile any secondary DLLs

Repeat step 2c for any other DLLs found in step 2a that are larger than 50KB. These may contain helper classes or feature-specific modules.

### Completion Test

- [ ] `decompilation/kutools-source/` contains `.cs` files
- [ ] At least one file is readable (either proper class names or legible Word API calls)
- [ ] Obfuscation level assessed and noted in `FINDINGS.md`

---

## PHASE 3: DECOMPILATION ANALYSIS

**Prerequisites:** Phase 2 complete  
**Resume Prompt:** `Analyse the decompiled KuTools for Word C# source in decompilation/kutools-source/. Find the classes and methods that handle the formatting operations listed in Phase 3 of the plan. Document your findings in decompilation/FINDINGS.md and decompilation/api-map.md.`

### Instructions

The agent reads the decompiled `.cs` files and locates the code responsible for each formatting operation. The Word COM interop API uses `Microsoft.Office.Interop.Word` — search for its types.

#### 3a. Find the formatting classes

Search for the following strings in the decompiled source. These are Word COM interop property and method names that KuTools must use:

```
grep -rn "\.Font\." decompilation/kutools-source/ > decompilation/grep-font.txt
grep -rn "ParagraphFormat\." decompilation/kutools-source/ > decompilation/grep-paragraph.txt
grep -rn "\.Style\b" decompilation/kutools-source/ > decompilation/grep-style.txt
grep -rn "\.Find\." decompilation/kutools-source/ > decompilation/grep-find.txt
grep -rn "Footnotes\." decompilation/kutools-source/ > decompilation/grep-footnotes.txt
grep -rn "\.Delete()" decompilation/kutools-source/ > decompilation/grep-delete.txt
grep -rn "ChangeCase" decompilation/kutools-source/ > decompilation/grep-case.txt
```

Open each grep output file and trace back to the class and method that contains each call.

#### 3b. Document the COM-to-Office.js mapping

Create `decompilation/api-map.md` with a table mapping every COM interop call found to its Office.js equivalent:

```markdown
# KuTools COM Interop → Office.js API Map

| Operation | KuTools COM call (VSTO) | Office.js equivalent |
|---|---|---|
| Bold | `selection.Font.Bold = true` | `range.font.bold = true` |
| Font color | `selection.Font.Color = WdColor.wdColorRed` | `range.font.color = "#FF0000"` |
| Alignment | `para.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify` | `para.alignment = Word.Alignment.justified` |
| ... | ... | ... |
```

Fill this table for every operation found. If a COM call has no direct Office.js equivalent, note `"No direct equivalent — workaround needed"` and describe the workaround.

#### 3c. Specific patterns to look for

Beyond simple property assignments, look for:

**How KuTools handles mixed selections** (e.g., a selection where some text is bold and some is not):
- COM interop returns `Word.WdConstants.wdUndefined` (= 9999999) for mixed properties
- Note how KuTools checks for this value and decides whether to set to true or false

**How KuTools iterates paragraphs:**
- Does it iterate forward or backward when deleting?
- Does it use `Selection.Paragraphs` or `Document.Paragraphs` or a range?

**How KuTools implements case conversion:**
- Does it use Word's built-in `Selection.Range.Case = WdCharacterCase.wdUpperCase` or does it process the string in C# and then reinsert it?

**How KuTools removes empty paragraphs:**
- Does it use Find and Replace with a regex?
- Does it iterate paragraphs and delete?

**How KuTools handles the "Remove all formatting" operation:**
- Does it set each property individually?
- Does it use `Selection.ClearCharacterAllFormatting()` and `Selection.ClearParagraphAllFormatting()`?
- Note: these two COM methods have direct equivalents in Office.js.

#### 3d. Update FINDINGS.md

For each behavioral finding from Phase 1, add a new section linking it to the code found:

```markdown
## Operation: Bold Toggle

**Observed behavior:** [from Phase 1]

**XML change:** [from Phase 1]

**Code found (decompiled):**
[paste the relevant C# method or key lines]

**COM calls used:**
- selection.Font.Bold

**Office.js equivalent:**
- range.font.bold

**Edge case handling found:**
[e.g., "Checks for wdUndefined (9999999) and sets to true if mixed"]
```

### Completion Test

- [ ] `decompilation/api-map.md` exists with at least 15 operation rows filled in
- [ ] `decompilation/FINDINGS.md` updated with code references for each behavioral observation
- [ ] Edge cases documented (mixed selection handling, iteration order, etc.)
- [ ] At least one finding changes or refines the planned Office.js implementation (document what changed)

---

## PHASE 4: PROJECT SCAFFOLD

**Prerequisites:** Phase 0 complete (decompilation phases can run in parallel)  
**Resume Prompt:** `Scaffold the chambers-word-toolkit Office.js Word add-in project using Yeoman. Follow Phase 4 of the plan exactly.`

### Instructions

1. Navigate to the project root (the `chambers-word-toolkit/` folder created in Phase 0):
   ```
   cd chambers-word-toolkit
   ```

2. Run the Yeoman generator:
   ```
   yo office
   ```

3. Answer the prompts:
   - **Project type:** `Office Add-in Task Pane project using React framework`
   - **Script type:** `TypeScript`
   - **Add-in name:** `PawaTuls`
   - **Office application:** `Word`

   Yeoman generates files into the current directory.

4. Install Fluent UI v9:
   ```
   npm install @fluentui/react-components
   ```

5. Create the module directories:
   ```
   mkdir -p src/shared src/formatting src/legal src/footnotes src/taskpane/components/shared
   ```

6. Verify the dev loop works:
   ```
   npm start
   ```
   Word should open with the default "Hello World" task pane. Close Word and stop the server (`Ctrl+C`) after confirming.

### Completion Test

- [ ] `npm start` opens Word with a working task pane
- [ ] `npx tsc --noEmit` reports no errors on the scaffolded code
- [ ] All `src/` subdirectories exist

---

## PHASE 5: MANIFEST CONFIGURATION

**Prerequisites:** Phase 4 complete  
**Resume Prompt:** `Edit manifest.xml in the chambers-word-toolkit project to configure the add-in metadata and set the minimum Word API version to 1.5. Follow Phase 5 of the plan exactly.`

### Instructions

Open `manifest.xml` in the project root and make these changes:

1. Update display name and description:
   ```xml
   <DisplayName DefaultValue="PawaTuls" />
   <Description DefaultValue="Formatting, Nigerian legal workflows, and footnote detection for Word." />
   ```

2. Set the minimum API requirement to WordApi 1.5:
   ```xml
   <Requirements>
     <Sets DefaultMinVersion="1.5">
       <Set Name="WordApi" MinVersion="1.5" />
     </Sets>
   </Requirements>
   ```

3. Confirm the `<Hosts>` block contains `<Host Name="Document" />`.

4. Save and run `npm start` to confirm Word still loads the add-in.

### Completion Test

- [ ] `manifest.xml` contains `<Set Name="WordApi" MinVersion="1.5" />`
- [ ] `npm start` opens Word with a working task pane

---

## PHASE 6: DECOMPILATION INTEGRATION REVIEW

**Prerequisites:** Phases 3 and 4 complete  
**Resume Prompt:** `Read decompilation/api-map.md and decompilation/FINDINGS.md. Identify any edge cases or implementation patterns found in the KuTools decompilation that must be incorporated into the Office.js modules. Produce a short integration notes file at decompilation/integration-notes.md. Follow Phase 6 of the plan exactly.`

### Instructions

This phase bridges the decompilation work and the build work. The agent reads the decompilation outputs and produces a concise reference that all future build sessions will consult.

Create `decompilation/integration-notes.md` with the following sections:

**Section 1: Mixed selection handling**  
How KuTools detects and responds to a mixed selection (e.g., some text is bold, some is not). State the Office.js equivalent approach. A mixed property in Office.js returns `null` (not `9999999` as in COM) — confirm this is noted.

**Section 2: Iteration order for deletions**  
Whether KuTools iterates forward or backward when deleting paragraphs. Confirm the plan's approach (reverse order) matches or note if KuTools uses a different approach.

**Section 3: Clear formatting implementation**  
Whether KuTools uses individual property resets or a single clear-all method. Note whether `Selection.ClearCharacterAllFormatting()` and `Selection.ClearParagraphAllFormatting()` map directly to an Office.js equivalent (they do not have exact equivalents — the plan uses individual resets).

**Section 4: Case conversion implementation**  
Whether KuTools uses Word's built-in `Case` property or processes the string in code. Note the chosen approach for the Office.js implementation.

**Section 5: Anything found in the decompiled code that changes the planned implementation**  
If the decompilation revealed an approach or edge case not anticipated in the plan, document it here with the implication for the build.

### Completion Test

- [ ] `decompilation/integration-notes.md` exists with all five sections filled in
- [ ] Any changes to planned implementations noted clearly

---

## PHASE 7: SHARED TYPES

**Prerequisites:** Phase 4 complete  
**Resume Prompt:** `Create src/shared/types.ts in the chambers-word-toolkit project. This file defines all TypeScript interfaces and types used across all three feature areas. Follow Phase 7 of the plan exactly. Create only this file.`

### Instructions

Create `src/shared/types.ts`. These types are the contract that every other module imports. Do not modify this file from other phases — if a new type is needed, add it here and update this file.

```typescript
// src/shared/types.ts
// Central type definitions for PawaTuls.
// All three feature areas import from here. Do not define types elsewhere.

// ─── FORMATTING ENGINE ────────────────────────────────────────────────────────

export type UnderlineStyle =
  | "none" | "single" | "double" | "dotted"
  | "dashed" | "wave" | "thick" | "thickDotted" | "thickDashed";

export type LineSpacingRule =
  | "single" | "onePtFive" | "double"
  | "atLeast" | "exactly" | "multiple";

export type ClearScope = "character" | "paragraph" | "all";

export type CaseTransform =
  | "upper" | "lower" | "title" | "sentence" | "toggle";

export interface FontOptions {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  strikeThrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  color?: string;           // hex "#RRGGBB" or "auto"
  highlightColor?: string;  // Word English color name or "none" to clear
  allCaps?: boolean;
  smallCaps?: boolean;
  spacing?: number;         // character spacing in points
  hidden?: boolean;
}

export interface ParagraphOptions {
  alignment?: "left" | "centered" | "right" | "justified";
  firstLineIndent?: number;     // points
  leftIndent?: number;          // points
  rightIndent?: number;         // points
  lineSpacing?: number;         // points (used with atLeast, exactly, multiple)
  lineSpacingRule?: LineSpacingRule;
  spaceBefore?: number;         // points
  spaceAfter?: number;          // points
  keepWithNext?: boolean;
  keepLinesTogether?: boolean;
  pageBreakBefore?: boolean;
}

export interface StyleCreationOptions {
  name: string;
  styleType: "paragraph" | "character";
  basedOn?: string;
  font?: FontOptions;
  paragraph?: ParagraphOptions;
}

export interface BorderOptions {
  color?: string;   // hex
  width?: number;   // points
  type?: "single" | "double" | "dotted" | "dashed" | "none";
}

// ─── LEGAL WORKFLOWS ─────────────────────────────────────────────────────────

export type NumberingScheme =
  | "decimal-decimal-decimal"    // 1. / 1.1 / 1.1.1
  | "decimal-alpha-roman";       // 1. / (a) / (i)

export type LawReporter =
  | "NWLR" | "LPELR" | "SCNJ" | "FWLR"
  | "NLR" | "ANLR" | "NSCC" | "UILR" | "JELR";

export type CourtName =
  | "Supreme Court of Nigeria"
  | "Court of Appeal"
  | "Federal High Court"
  | "High Court of the Federal Capital Territory"
  | "High Court of Justice"
  | "National Industrial Court";

export interface NigerianCaseCitation {
  year: string;
  volume?: string;
  reporter: LawReporter;
  part?: string;
  page: string;
}

export interface CauseHeadingData {
  court: CourtName | string;
  division?: string;
  suitNumber: string;
  claimant: string;
  defendant: string;
  actionType?: string;
}

export interface StatutoryCitation {
  section: string;
  actName: string;
  year: string;
}

// ─── FOOTNOTE DETECTION ──────────────────────────────────────────────────────

export type MarkerPattern = "superscript" | "bracketed" | "parenthetical";

export interface FootnoteMarker {
  number: number;
  pattern: MarkerPattern;
  paragraphIndex: number;
  contextBefore: string;   // up to 80 chars before the marker
  contextAfter: string;    // up to 80 chars after the marker
}

export interface FootnoteBlock {
  number: number;
  text: string;
  paragraphIndex: number;
}

export type MatchConfidence = "high" | "medium" | "low";

export interface FootnoteMatch {
  id: string;
  marker: FootnoteMarker;
  block: FootnoteBlock;
  confidence: MatchConfidence;
  accepted: boolean;
}

export interface ScanResult {
  matches: FootnoteMatch[];
  unmatchedMarkers: FootnoteMarker[];
  orphanedBlocks: FootnoteBlock[];
  separatorLineFound: boolean;
  totalMarkersFound: number;
  totalBlocksFound: number;
  scanDurationMs: number;
}

export interface ConversionResult {
  converted: number;
  failed: number;
  errors: string[];
}
```

### Completion Test

- [ ] `src/shared/types.ts` exists
- [ ] `npx tsc --noEmit` returns zero errors

---

## PHASE 8: WORD API UTILITIES

**Prerequisites:** Phase 7 complete  
**Resume Prompt:** `Create src/shared/wordUtils.ts in the chambers-word-toolkit project. This file contains shared helpers for all Office.js Word API calls. Follow Phase 8 of the plan exactly. Create only this file.`

### Instructions

Create `src/shared/wordUtils.ts`. Add this file header:

```typescript
/**
 * wordUtils.ts
 * Shared helpers for Office.js Word API operations.
 * All Word.run wrappers and unit converters live here.
 * Import these in every formatting, legal, and footnote module.
 */
```

**Implement these exported functions:**

**`getSelectionOrThrow(context: Word.RequestContext): Promise<Word.Range>`**  
Calls `context.document.getSelection()`, loads `"text"`, syncs. If `selection.text.trim().length === 0`, throws `new Error("No text selected. Select text before applying this operation.")`. Returns the range.

**`getSelectionParagraphsOrThrow(context: Word.RequestContext): Promise<Word.ParagraphCollection>`**  
Calls `context.document.getSelection().paragraphs`, loads `"text"`, syncs. Throws if `paragraphs.items.length === 0`.

**`safeWordRun<T>(fn: (context: Word.RequestContext) => Promise<T>): Promise<{ data?: T; error?: string }>`**  
Wraps `Word.run(fn)` in try/catch. On success returns `{ data }`. On error extracts `err.message` and returns `{ error: message }`.

**`getDocumentBodyOrThrow(context: Word.RequestContext): Promise<Word.Body>`**  
Returns `context.document.body` after loading `"text"`.

**`ptToCm(pt: number): number`**  
Returns `pt / 28.3465`.

**`cmToPt(cm: number): number`**  
Returns `cm * 28.3465`.

**`inchToPt(inch: number): number`**  
Returns `inch * 72`.

### Completion Test

- [ ] `src/shared/wordUtils.ts` exists
- [ ] `npx tsc --noEmit` passes with zero errors
- [ ] All six functions are exported

---

## PHASE 9: TASK PANE SHELL

**Prerequisites:** Phases 4, 7, 8 complete  
**Resume Prompt:** `Replace the default App.tsx with a three-tab navigation shell and create ErrorBanner.tsx. Follow Phase 9 of the plan exactly.`

### Instructions

#### 9a. Create ErrorBanner

Create `src/taskpane/components/shared/ErrorBanner.tsx`:
- Props: `{ message: string | null; onDismiss: () => void }`
- Renders nothing if `message` is null
- Uses Fluent UI v9 `MessageBar` with `intent="error"`
- Dismiss button calls `onDismiss`

#### 9b. Replace App.tsx

Replace the Yeoman-generated `App.tsx`:
- Wrap everything in Fluent UI `<FluentProvider theme={webLightTheme}>`
- Three tabs using Fluent UI `TabList` and `Tab`: **Format**, **Legal**, **Footnotes**
- State: `const [activeTab, setActiveTab] = useState<"format" | "legal" | "footnotes">("format")`
- State: `const [error, setError] = useState<string | null>(null)`
- Render `<ErrorBanner message={error} onDismiss={() => setError(null)} />` at the top
- Render the active panel below the tabs, passing `setError` as a prop
- Active panel mapping:
  - `"format"` → `<FormattingPanel setError={setError} />`
  - `"legal"` → `<LegalPanel setError={setError} />`
  - `"footnotes"` → `<FootnotePanel setError={setError} />`

#### 9c. Create placeholder panels

Create three stub files so `App.tsx` imports resolve:

`src/taskpane/components/FormattingPanel.tsx`:
```typescript
import React from "react";
export function FormattingPanel({ setError }: { setError: (e: string | null) => void }) {
  return <div>Formatting Panel — Phase 16</div>;
}
```

`src/taskpane/components/LegalPanel.tsx`:
```typescript
import React from "react";
export function LegalPanel({ setError }: { setError: (e: string | null) => void }) {
  return <div>Legal Panel — Phase 24</div>;
}
```

`src/taskpane/components/FootnotePanel.tsx`:
```typescript
import React from "react";
export function FootnotePanel({ setError }: { setError: (e: string | null) => void }) {
  return <div>Footnote Panel — Phase 28</div>;
}
```

### Completion Test

- [ ] `npm start` opens Word showing a task pane with three tabs
- [ ] Clicking each tab switches content
- [ ] `npx tsc --noEmit` passes with zero errors

---

## PHASE 10: FORMATTING — FONT MODULE

**Prerequisites:** Phases 6, 7, 8 complete  
**Resume Prompt:** `Create src/formatting/font.ts. Before writing code, read decompilation/integration-notes.md for any findings that affect font implementation. Follow Phase 10 of the plan exactly.`

### Instructions

Create `src/formatting/font.ts`. Consult `decompilation/integration-notes.md` Section 1 (mixed selection handling) before implementing toggle functions.

**Implement these exported functions:**

**`applyFont(options: FontOptions): Promise<{ error?: string }>`**  
Gets selection, applies every non-undefined property in `FontOptions` to `selection.font`.  
- Map `UnderlineStyle` values to `Word.UnderlineType.*` enum values
- `color` passes as hex string directly: `range.font.color = options.color`
- `highlightColor` uses Word's English color names (`"yellow"`, `"turquoise"`, `"pink"`, `"violet"`, `"darkBlue"`, `"darkRed"`, `"darkYellow"`, `"darkGreen"`, `"teal"`, `"darkCyan"`, `"black"`, `"blue"`, `"darkMagenta"`, `"darkGray"`, `"lightGray"`, `"white"`, `"none"`)

**`increaseFontSize(by: number = 1): Promise<{ error?: string }>`**  
Loads `selection.font.size`. Adds `by`. If `size` is null (mixed selection), use 12 as base. Clamps minimum to 1.

**`decreaseFontSize(by: number = 1): Promise<{ error?: string }>`**  
Same but subtracts. Clamps minimum to 1.

**`clearCharacterFormatting(): Promise<{ error?: string }>`**  
Resets: bold false, italic false, underline none, strikeThrough false, superscript false, subscript false, allCaps false, smallCaps false, color "auto", highlightColor "none", spacing 0, hidden false.

**`toggleBold(): Promise<{ error?: string }>`**  
Loads `selection.font.bold`. If `null` (mixed), set to `true`. Otherwise invert.

**`toggleItalic(): Promise<{ error?: string }>`**  
Same pattern as `toggleBold` for `font.italic`.

**`toggleStrikethrough(): Promise<{ error?: string }>`**  
Same pattern for `font.strikeThrough`.

**`setSuperscript(): Promise<{ error?: string }>`**  
Sets `font.superscript = true` and `font.subscript = false` in the same sync.

**`setSubscript(): Promise<{ error?: string }>`**  
Sets `font.subscript = true` and `font.superscript = false` in the same sync.

### Completion Test

- [ ] `src/formatting/font.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 9 functions exported

---

## PHASE 11: FORMATTING — PARAGRAPH MODULE

**Prerequisites:** Phases 6, 7, 8 complete  
**Resume Prompt:** `Create src/formatting/paragraph.ts. Before writing code, read decompilation/integration-notes.md for any findings that affect paragraph implementation — particularly iteration order and empty paragraph removal. Follow Phase 11 of the plan exactly.`

### Instructions

Create `src/formatting/paragraph.ts`.

**`applyParagraphFormat(options: ParagraphOptions): Promise<{ error?: string }>`**  
Gets selection paragraphs. Applies every non-undefined property in `ParagraphOptions` to each paragraph item.  
- Map alignment strings to `Word.Alignment.*`
- Map `lineSpacingRule` strings to `Word.LineSpacingRule.*`
- `firstLineIndent`, `leftIndent`, `rightIndent` are in points — pass directly

**`setLineSpacing(rule: LineSpacingRule, value?: number): Promise<{ error?: string }>`**  
Sets `lineSpacingRule` on each selected paragraph. For `"exactly"` and `"atLeast"`, `value` (points) is required. For `"multiple"`, `value` is the multiplier (e.g. 1.5). For `"single"`, `"onePtFive"`, `"double"`, `value` is ignored.

**`removeAllSpacing(): Promise<{ error?: string }>`**  
Sets `spaceBefore = 0` and `spaceAfter = 0` on all selected paragraphs.

**`removeAllIndents(): Promise<{ error?: string }>`**  
Sets `firstLineIndent = 0`, `leftIndent = 0`, `rightIndent = 0` on all selected paragraphs.

**`setFirstLineIndent(cm: number): Promise<{ error?: string }>`**  
Converts cm to points and sets `paragraph.firstLineIndent`.

**`setHangingIndent(cm: number): Promise<{ error?: string }>`**  
Sets `firstLineIndent = -cmToPt(cm)` and `leftIndent = cmToPt(cm)`.

**`removeEmptyParagraphs(): Promise<{ data?: number; error?: string }>`**  
Loads all body paragraphs with `"text"`. Iterates in reverse order using `items.slice().reverse()`. Deletes paragraphs where `paragraph.text.trim() === ""`. Returns count deleted.

**`setParagraphBorder(location: "top" | "bottom" | "left" | "right" | "all", options: BorderOptions): Promise<{ error?: string }>`**  
For each selected paragraph, calls `paragraph.getBorder(Word.BorderLocation.*)` for the specified location. Sets `border.color`, `border.width`, `border.type`. For `"all"`, applies to top, bottom, left, right.

### Completion Test

- [ ] `src/formatting/paragraph.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 8 functions exported

---

## PHASE 12: FORMATTING — STYLES MODULE

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/formatting/styles.ts. Follow Phase 12 of the plan exactly.`

### Instructions

Create `src/formatting/styles.ts`.

**`applyStyle(styleName: string): Promise<{ error?: string }>`**  
Gets selected paragraphs. Sets `paragraph.style = styleName` on each. If the style does not exist, Word throws — let `safeWordRun` catch it.

**`getAvailableStyles(): Promise<{ data?: string[]; error?: string }>`**  
Calls `context.document.getStyles()`, loads `"nameLocal,type"`. Filters to `Word.StyleType.paragraph`. Returns array of `nameLocal` strings sorted alphabetically. The return must use `return { data: styleNames }` inside `safeWordRun` — not set state.

**`createStyleFromSelection(options: StyleCreationOptions): Promise<{ error?: string }>`**  
Calls `context.document.addStyle(options.name, styleType)`. Applies `options.font` properties to `newStyle.font` and `options.paragraph` properties to `newStyle.paragraphFormat`. If `options.basedOn` is provided, sets `newStyle.baseStyle = options.basedOn`.

**`getSelectionStyle(): Promise<{ data?: string; error?: string }>`**  
Loads the first paragraph of the selection's `style` property. Returns the style name string.

### Completion Test

- [ ] `src/formatting/styles.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 4 functions exported

---

## PHASE 13: FORMATTING — TEXT TRANSFORMS

**Prerequisites:** Phases 6, 7, 8 complete  
**Resume Prompt:** `Create src/formatting/text.ts. Before writing the case conversion function, check decompilation/integration-notes.md Section 4 for whether KuTools uses Word's built-in Case property or string processing. Follow Phase 13 of the plan exactly.`

### Instructions

Create `src/formatting/text.ts`.

**`transformCase(transform: CaseTransform): Promise<{ error?: string }>`**  
Loads `selection.text`. Applies the JavaScript transformation:
- `"upper"`: `text.toUpperCase()`
- `"lower"`: `text.toLowerCase()`
- `"title"`: `text.replace(/\b\w/g, c => c.toUpperCase()).replace(/\B\w/g, c => c.toLowerCase())`
- `"sentence"`: lowercase all, then uppercase after `. `, `? `, `! `, and the first character
- `"toggle"`: for each character, uppercase → lowercase, lowercase → uppercase, other → unchanged

Uses `selection.insertText(result, Word.InsertLocation.replace)` to write back.

**`trimSelectionSpaces(): Promise<{ error?: string }>`**  
Loads `selection.text`. Trims with `text.trim()`. Replaces selection.

**`collapseMultipleSpaces(): Promise<{ error?: string }>`**  
Loads `selection.text`. Applies `text.replace(/ {2,}/g, " ")`. Replaces selection.

**`trimParagraphLeadingTabs(): Promise<{ error?: string }>`**  
Gets selected paragraphs. For each paragraph, strips leading tabs with `text.replace(/^\t+/, "")`. Uses `paragraph.insertText(stripped, Word.InsertLocation.replace)`.

### Completion Test

- [ ] `src/formatting/text.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 4 functions exported

---

## PHASE 14: FORMATTING — CONVERSION TOOLS

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/formatting/convert.ts. Follow Phase 14 of the plan exactly.`

### Instructions

Create `src/formatting/convert.ts`.

**Agent context:** In Office.js, a manual line break (Shift+Enter) is detected by searching for `"\v"` (vertical tab) in body text. Inserting `"\n"` into a range creates a new paragraph mark. Inserting `"\v"` creates a manual line break.

**`lineBreaksToParagraphMarks(): Promise<{ data?: number; error?: string }>`**  
Searches body for `"\v"`. For each match, calls `match.insertText("\n", Word.InsertLocation.replace)`. Returns count.

**`paragraphMarksToLineBreaks(): Promise<{ data?: number; error?: string }>`**  
Applies to current selection only — not the whole document.  
Gets selected paragraphs. For all but the last one, inserts `"\v"` at the end of each paragraph's range and deletes the paragraph mark by merging the paragraph with the next. Returns count.

**`fieldCodesToText(): Promise<{ data?: number; error?: string }>`**  
Loads `context.document.body.fields` with `"type,result"`. For each field, inserts `field.result.text` at the field's location, then calls `field.delete()`. Returns count.

**`listToText(): Promise<{ error?: string }>`**  
Gets selected paragraphs. For paragraphs with list formatting, calls `paragraph.detachFromList()` (WordApi 1.3). Sets `leftIndent = 0` and `firstLineIndent = 0`.

### Completion Test

- [ ] `src/formatting/convert.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 4 functions exported

---

## PHASE 15: FORMATTING — CLEAR FORMATTING

**Prerequisites:** Phases 6, 10, 11, 12 complete  
**Resume Prompt:** `Create src/formatting/clear.ts. Before writing, read decompilation/integration-notes.md Section 3 for how KuTools implements clear formatting. Follow Phase 15 of the plan exactly.`

### Instructions

Create `src/formatting/clear.ts`. This module calls functions from other formatting modules — import them.

**`clearFormatting(scope: ClearScope): Promise<{ error?: string }>`**  
- `"character"`: calls `clearCharacterFormatting()` from `font.ts`
- `"paragraph"`: calls `removeAllIndents()` and `removeAllSpacing()` from `paragraph.ts`, then sets line spacing to single via `setLineSpacing("single")`
- `"all"`: calls both character and paragraph clearing, then calls `applyStyle("Normal")` from `styles.ts`

**`removeHorizontalLines(): Promise<{ data?: number; error?: string }>`**  
Loads all body paragraphs with `"text"`. Deletes paragraphs whose text matches `/^[-_=]{3,}$/.test(para.text.trim())`. Also checks for paragraph bottom borders (via `getBorder(Word.BorderLocation.bottom)`) and deletes those paragraphs. Iterates in reverse. Returns count.

**`removeExtraDocumentSpacing(): Promise<{ data?: number; error?: string }>`**  
Loads all body paragraphs. For each where `spaceBefore > 12` or `spaceAfter > 12`, sets both to 6. Returns count adjusted.

### Completion Test

- [ ] `src/formatting/clear.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 3 functions exported

---

## PHASE 16: FORMATTING PANEL UI

**Prerequisites:** Phases 9, 10, 11, 12, 13, 14, 15 complete  
**Resume Prompt:** `Replace the FormattingPanel.tsx stub with the full formatting panel component. Follow Phase 16 of the plan exactly.`

### Instructions

Replace `src/taskpane/components/FormattingPanel.tsx`. Props: `{ setError: (e: string | null) => void }`.

**Structure:** Three collapsible `Accordion` sections using Fluent UI v9.

**Section 1 — Font**  
Font family `Dropdown` (options: Times New Roman, Arial, Calibri, Garamond, Book Antiqua, Cambria). Font size `SpinButton` (min 6, max 72). Bold / Italic / Underline / Strikethrough as `ToggleButton` row. Superscript / Subscript as `Button` row. Text color `Input type="color"`. Highlight color `Dropdown` (Word named colors). Character spacing `SpinButton` (min -20, max 200, step 0.5 points). Clear character formatting `Button`.

**Section 2 — Paragraph**  
Alignment: four `ToggleButton` icons, one active at a time. Line spacing `Dropdown` (Single, 1.5 Lines, Double, At Least, Exactly, Multiple). Line spacing value `SpinButton` — visible only when "At Least", "Exactly", or "Multiple" is selected. Space before / Space after `SpinButton` in points. First-line indent / Hanging indent `SpinButton` in cm. Remove All Spacing `Button`. Remove All Indents `Button`. Remove Empty Paragraphs `Button` — shows count in a Fluent UI `Toast` after run.

**Section 3 — Styles and Tools**  
Apply style `Dropdown` — populated by calling `getAvailableStyles()` when the section opens. Case conversion: five `Button` items (UPPERCASE, lowercase, Title Case, Sentence case, tOGGLE). Collapse spaces `Button`. Line Breaks → Paragraph Marks `Button`. Paragraph Marks → Line Breaks `Button`. Clear All Formatting `Button`.

**Error handling:** Each `onClick` calls the relevant module function and on error calls `setError(message)`.

**Rule:** No Office.js calls inside this component. Every button calls a pre-built function from a module.

### Completion Test

- [ ] Formatting tab renders all three sections
- [ ] Sections expand and collapse
- [ ] Bold with selected text works (test manually in Word)
- [ ] Error banner appears when a formatting button is clicked with no selection
- [ ] `npx tsc --noEmit` passes

---

## PHASE 17: LEGAL — PLEADINGS FORMATTER

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/legal/pleadings.ts. Follow Phase 17 of the plan exactly.`

### Instructions

Create `src/legal/pleadings.ts`.

**Context:** Nigerian court pleadings use plain numbered paragraph text, not Word's List feature, because List formatting behaves unpredictably across different machines and court filing systems. The formatter prefixes paragraph text directly.

**`applyPleadingsNumbering(scheme: NumberingScheme, applyTo: "selection" | "document"): Promise<{ error?: string }>`**  
Gets paragraphs from selection or full body. Detects indentation level from `paragraph.leftIndent`:
- Level 1: less than 18 points
- Level 2: 18–36 points
- Level 3: more than 36 points

Tracks counters: `const counters = [0, 0, 0]`. On entering a new level, resets all lower levels to 0. Increments the current level counter. Prepends the prefix to each paragraph's text.

Prefix formats:
- `decimal-decimal-decimal`: Level 1 → `1.`, Level 2 → `1.1`, Level 3 → `1.1.1`
- `decimal-alpha-roman`: Level 1 → `1.`, Level 2 → `(a)`, Level 3 → `(i)`

Alpha sequence: a, b, c ... z, aa, ab ... Roman: i, ii, iii, iv, v, vi, vii, viii, ix, x, xi, xii, xiii, xiv, xv ...

Uses `paragraph.insertText(prefix + " " + existingText, Word.InsertLocation.replace)`.

**`removePleadingsNumbering(applyTo: "selection" | "document"): Promise<{ error?: string }>`**  
For each paragraph, strips the leading number prefix using this regex:
```
/^(\d+\.(\d+\.(\d+\.)?)?|\([a-z]+\)|\([ivxlcdm]+\))\s+/i
```

### Completion Test

- [ ] `src/legal/pleadings.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] Both functions exported

---

## PHASE 18: LEGAL — AFFIDAVIT FORMATTER

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/legal/affidavit.ts. Follow Phase 18 of the plan exactly.`

### Instructions

Create `src/legal/affidavit.ts`.

**`insertJuratBlock(): Promise<{ error?: string }>`**  
Inserts this HTML at the cursor using `selection.insertHtml(html, Word.InsertLocation.after)`:

```html
<p>&nbsp;</p>
<p>SWORN to at ________________________ this _____ day of
   ________________________ 20_____.</p>
<p>&nbsp;</p>
<p>_________________________________</p>
<p>DEPONENT</p>
<p>&nbsp;</p>
<p>Before me:</p>
<p>&nbsp;</p>
<p>_________________________________</p>
<p>COMMISSIONER FOR OATHS / NOTARY PUBLIC</p>
```

**`applyAffidavitNumbering(): Promise<{ error?: string }>`**  
Numbers selected paragraphs starting from 1. Applies `leftIndent = cmToPt(1.27)`, `firstLineIndent = -cmToPt(1.27)`, single line spacing, `spaceBefore = 6`, `spaceAfter = 6`. Prefixes each with `"1."`, `"2."` etc.

### Completion Test

- [ ] `src/legal/affidavit.ts` exists
- [ ] `npx tsc --noEmit` passes

---

## PHASE 19: LEGAL — EXHIBIT MARKER

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/legal/exhibit.ts. Follow Phase 19 of the plan exactly.`

### Instructions

Create `src/legal/exhibit.ts`. Module-level state: `let _lastExhibitLetter = "";`

**`getNextExhibitLetter(): string`**  
Pure function. If `_lastExhibitLetter` is empty, returns `"A"`. Otherwise increments: A→B ... Z→AA, AA→AB ... Updates and returns `_lastExhibitLetter`.

**`insertExhibitLabel(position: "inline" | "standalone"): Promise<{ data?: string; error?: string }>`**  
Gets next letter. If `"standalone"`: inserts a bold paragraph `"EXHIBIT [letter]"` followed by a blank line. If `"inline"`: inserts `(hereinafter referred to as "Exhibit [letter]")` at cursor. Returns the letter used.

**`resetExhibitCounter(): void`**  
Synchronous. Sets `_lastExhibitLetter = ""`.

**`scanDocumentForLastExhibit(): Promise<{ data?: string; error?: string }>`**  
Searches body for `"Exhibit [A-Z]"` and `"EXHIBIT [A-Z]"`. Finds the highest letter. Sets `_lastExhibitLetter`. Returns letter found or `"None found"`. Call this on add-in load so the counter starts from the correct position in existing documents.

### Completion Test

- [ ] `src/legal/exhibit.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 4 functions exported

---

## PHASE 20: LEGAL — RELIEFS FORMATTER

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/legal/reliefs.ts. Follow Phase 20 of the plan exactly.`

### Instructions

Create `src/legal/reliefs.ts`.

**`formatReliefsSection(): Promise<{ error?: string }>`**  
Gets selected paragraphs. Inserts a bold centered header `"THE RELIEFS SOUGHT"` above the selection. Numbers each paragraph `(a)`, `(b)`, `(c)` ... with `leftIndent = cmToPt(1.27)` and `firstLineIndent = -cmToPt(1.27)`.

**`insertClosingRelief(): Promise<{ error?: string }>`**  
Inserts at cursor:
```
AND FOR SUCH FURTHER ORDER(S) as this Honourable Court may deem fit to make
in the circumstances.
```
Formatted with the same indent as the relief paragraphs.

**`insertAnOrderPrefix(): Promise<{ error?: string }>`**  
Inserts `"AN ORDER "` at the beginning of the first selected paragraph.

### Completion Test

- [ ] `src/legal/reliefs.ts` exists
- [ ] `npx tsc --noEmit` passes

---

## PHASE 21: LEGAL — CAUSE HEADING BLOCK

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/legal/causeHeading.ts. Follow Phase 21 of the plan exactly.`

### Instructions

Create `src/legal/causeHeading.ts`.

**`insertCauseHeading(data: CauseHeadingData): Promise<{ error?: string }>`**  
Inserts formatted cause heading HTML at the cursor using `selection.insertHtml(html, Word.InsertLocation.after)`.

Template (all `{{placeholder}}` tokens replaced from `CauseHeadingData`; `division` defaults to `"ABUJA"` if undefined):

```html
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt;font-weight:bold">{{court}}</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt;font-weight:bold">HOLDEN AT {{division}}</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">&nbsp;</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">SUIT NO: {{suitNumber}}</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">&nbsp;</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">BETWEEN</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">&nbsp;</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt;font-weight:bold">{{claimant}}</p>
<p style="text-align:right;font-family:'Times New Roman';font-size:12pt">CLAIMANT/APPLICANT</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">&nbsp;</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">AND</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">&nbsp;</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt;font-weight:bold">{{defendant}}</p>
<p style="text-align:right;font-family:'Times New Roman';font-size:12pt">DEFENDANT/RESPONDENT</p>
<p style="text-align:center;font-family:'Times New Roman';font-size:12pt">&nbsp;</p>
```

### Completion Test

- [ ] `src/legal/causeHeading.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] Manual test: heading block inserts with correct formatting in Word

---

## PHASE 22: LEGAL — CITATION BUILDER

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/legal/citations.ts. Follow Phase 22 of the plan exactly.`

### Instructions

Create `src/legal/citations.ts`.

**`buildCaseCitation(citation: NigerianCaseCitation): string`**  
Pure function. Formats the citation:
- Standard reporters (NWLR, SCNJ, FWLR, NLR, etc.): `[year] volume REPORTER (Pt. part) page`
  - Example: `[2020] 15 NWLR (Pt. 1721) 123`
- LPELR: `[year] LPELR-page` (no volume, no part; page includes court suffix)
  - Example: `[2021] LPELR-55302(CA)`

**`insertCaseCitation(citation: NigerianCaseCitation): Promise<{ error?: string }>`**  
Builds citation string. Inserts at cursor with italic formatting applied after insertion.

**`buildStatutoryCitation(citation: StatutoryCitation): string`**  
Pure function. Format: `Section [section] of the [actName] [year]`.  
Example: `Section 175(1) of the Companies and Allied Matters Act 2020`

**`insertStatutoryCitation(citation: StatutoryCitation): Promise<{ error?: string }>`**  
Builds citation string. Inserts at cursor with no special formatting.

### Completion Test

- [ ] `src/legal/citations.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 4 functions exported

---

## PHASE 23: LEGAL — SCHEDULE FORMATTER

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/legal/schedule.ts. Follow Phase 23 of the plan exactly.`

### Instructions

Create `src/legal/schedule.ts`. Module-level state: `let _lastScheduleNumber = 0;`

**`insertScheduleHeading(label?: string): Promise<{ data?: string; error?: string }>`**  
Increments `_lastScheduleNumber`. If `label` is provided, uses it; otherwise generates `"SCHEDULE " + _lastScheduleNumber`. Inserts a centered bold paragraph. Returns label used.

**`resetScheduleCounter(): void`**  
Sets `_lastScheduleNumber = 0`.

**`scanDocumentForLastSchedule(): Promise<{ data?: number; error?: string }>`**  
Searches body for `"SCHEDULE [number]"`. Finds the highest number. Sets `_lastScheduleNumber`. Returns that number.

### Completion Test

- [ ] `src/legal/schedule.ts` exists
- [ ] `npx tsc --noEmit` passes

---

## PHASE 24: LEGAL PANEL UI

**Prerequisites:** Phases 9, 17–23 complete  
**Resume Prompt:** `Replace the LegalPanel.tsx stub with the full legal workflows panel. Follow Phase 24 of the plan exactly.`

### Instructions

Replace `src/taskpane/components/LegalPanel.tsx`. Props: `{ setError: (e: string | null) => void }`.

**Six Accordion sections:**

**Pleadings:** Scheme `Dropdown` (two options). Apply-to `RadioGroup` (Selection / Document). Buttons: Apply Numbering, Remove Numbering.

**Affidavit:** Buttons: Number Paragraphs, Insert Jurat Block.

**Exhibits:** Label showing next exhibit letter (calls `getNextExhibitLetter()` for display). Buttons: Insert Standalone Label, Insert Inline Reference, Reset Counter. On panel mount: call `scanDocumentForLastExhibit()`.

**Reliefs:** Buttons: Format as Reliefs Section, Insert AN ORDER Prefix, Insert Closing Relief.

**Headings and Citations:**  
Sub-section Cause Heading: `Dropdown` for court name (all `CourtName` options plus "Custom"), inputs for Division, Suit No., Claimant, Defendant. Button: Insert Heading.  
Sub-section Case Citation: inputs for Year, Volume (optional), `Dropdown` for Reporter, Part (optional), Page. Button: Insert Citation.  
Sub-section Statutory Citation: inputs for Section, Act Name, Year. Button: Insert Citation.

**Structure:** Button: Insert Schedule Heading. Button: Reset Schedule Counter. Label showing next schedule number.

**Error handling:** All async calls surface errors via `setError`.

### Completion Test

- [ ] Legal tab renders all six sections
- [ ] Cause heading inserts correctly in Word (manual test)
- [ ] Exhibit counter increments across multiple insertions
- [ ] `npx tsc --noEmit` passes

---

## PHASE 25: FOOTNOTE DETECTION — SCANNER

**Prerequisites:** Phases 7, 8 complete  
**Resume Prompt:** `Create src/footnotes/scanner.ts. This is the most technically involved module in the project. Read the full phase before writing any code. Follow Phase 25 of the plan exactly.`

### Instructions

Create `src/footnotes/scanner.ts`. The scanner runs three detection passes and returns raw markers and blocks for the matcher.

**The three marker patterns:**

| Pattern | Body text example | Block example | Word wildcard |
|---|---|---|---|
| Superscript | `treaty¹ was signed` | `1 See Vienna Convention...` | `[0-9]@` + superscript filter |
| Bracketed | `treaty [1] was signed` | `[1] See Vienna Convention...` | `\[[0-9]@\]` |
| Parenthetical | `treaty (1) was signed` | `1. See Vienna Convention...` | `\([0-9]@\)` |

**Phase 1 — Marker detection (three passes):**

**Pass A — Superscript:**  
```typescript
const digitSearch = context.document.body.search("[0-9]@", { useWildcards: true });
digitSearch.load("font/superscript,text");
await context.sync();
const superscriptHits = digitSearch.items.filter(r => r.font.superscript === true && /^\d+$/.test(r.text));
```
Group consecutive items with the same text digit into single `FootnoteMarker` objects. Record `paragraphIndex` as the index in the full paragraph collection.

**Pass B — Bracketed:**  
```typescript
const bracketSearch = context.document.body.search("\\[[0-9]@\\]", { useWildcards: true });
bracketSearch.load("text");
await context.sync();
```
Parse number from each match text (strip `[` and `]`).

**Pass C — Parenthetical:**  
```typescript
const parenSearch = context.document.body.search("\\([0-9]@\\)", { useWildcards: true });
parenSearch.load("text");
await context.sync();
```
Parse number from each match text (strip `(` and `)`).

**Agent note on wildcard syntax:** The `@` character in Word wildcard syntax means "one or more of the preceding character class." If `[0-9]@` does not match multi-digit numbers on the test machine, fall back to `[0-9][0-9]?[0-9]?` to match 1–3 digit numbers.

**Phase 2 — Footnote block detection:**

Load all body paragraphs: `body.paragraphs.load("text")`.

For each paragraph test against these patterns in order:
- Bracketed block: `/^\s*\[(\d+)\]\s+(.+)$/`
- Decimal block: `/^\s*(\d+)[.)]\s+(.+)$/`
- Bare number block: `/^\s*(\d+)\s{2,}(.+)$/`

For each match, create a `FootnoteBlock` with `number`, `text` (the content after the number), and `paragraphIndex`.

**Consecutive-paragraph merging:** If paragraph N+1 does not start with any number and paragraph N was identified as a block, treat N+1 as a continuation of N. Concatenate with a space.

**Separator detection:**

For each paragraph, test whether `para.text.trim().match(/^[-_]{3,}$/)`. Record the lowest matching `paragraphIndex` as the separator line. If found, only paragraphs with `paragraphIndex > separatorIndex` are treated as block candidates.

**Implement one exported function:**

```typescript
export async function scanDocument(): Promise<{
  data?: {
    markers: FootnoteMarker[];
    blocks: FootnoteBlock[];
    separatorFound: boolean;
    separatorIndex: number;
    scanStartTime: number;
  };
  error?: string;
}>
```

Returns all raw detection results. `scanStartTime` is `Date.now()` recorded at the start of the function, passed to the matcher to compute duration.

### Completion Test

- [ ] `src/footnotes/scanner.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] The function is exported

---

## PHASE 26: FOOTNOTE DETECTION — MATCHER

**Prerequisites:** Phase 25 complete  
**Resume Prompt:** `Create src/footnotes/matcher.ts. Follow Phase 26 of the plan exactly.`

### Instructions

Create `src/footnotes/matcher.ts`. All functions are pure (no Office.js calls).

**`buildMatchTable(markers: FootnoteMarker[], blocks: FootnoteBlock[], separatorFound: boolean, separatorIndex: number): FootnoteMatch[]`**  

For each marker, find the block with the same number. Assign confidence:
- Block found AND `block.paragraphIndex > separatorIndex` (or no separator) → `"high"`
- Block found AND `block.paragraphIndex <= separatorIndex` → `"low"`
- Matched by number but block is ambiguous → `"medium"`

Set `accepted = true` by default. `id = \`match-${marker.number}-${Date.now()}\``

**`extractUnmatchedMarkers(markers: FootnoteMarker[], blocks: FootnoteBlock[]): FootnoteMarker[]`**  
Returns markers with no corresponding block number.

**`extractOrphanedBlocks(markers: FootnoteMarker[], blocks: FootnoteBlock[]): FootnoteBlock[]`**  
Returns blocks with no corresponding marker number.

**`buildScanResult(markers, blocks, separatorFound, separatorIndex, scanStartTime): ScanResult`**  
Assembles the full `ScanResult` object using the three functions above plus `Date.now() - scanStartTime` for `scanDurationMs`.

### Completion Test

- [ ] `src/footnotes/matcher.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 4 functions exported

---

## PHASE 27: FOOTNOTE DETECTION — CONVERTER

**Prerequisites:** Phases 25, 26 complete  
**Resume Prompt:** `Create src/footnotes/converter.ts. Read the full phase before writing any code. The re-scanning approach is critical. Follow Phase 27 of the plan exactly.`

### Instructions

Create `src/footnotes/converter.ts`.

**Critical note on stale range references:**  
Range references stored during the scan session become invalid after `context.sync()` has been called in a later session. The converter must not rely on stored `Word.Range` objects. Instead, it re-locates each marker and block by number at conversion time using fresh searches. This is mandatory.

**Critical note on conversion order:**  
Always process accepted matches in descending `paragraphIndex` order (highest first). This prevents earlier deletions from shifting the indices of later matches.

**`convertAcceptedMatches(scanResult: ScanResult): Promise<{ data?: ConversionResult; error?: string }>`**  

Implementation sequence inside a single `Word.run`:

1. Re-scan for superscript digit ranges matching accepted marker numbers
2. Re-scan for bracketed/parenthetical ranges matching accepted marker numbers
3. Re-scan paragraph texts to re-locate footnote block paragraphs
4. Sort accepted matches by descending `block.paragraphIndex`
5. For each accepted match in that order:
   - Locate the marker range (from step 1 or 2)
   - Call `markerRange.insertFootnote(match.block.text)` — this creates the real Word footnote
   - Clear the marker run: set its text to `""` or call `markerRange.delete()`
   - Record the block's paragraph for deletion in the next step
6. Delete all footnote block paragraphs (also in reverse paragraph order)
7. Return `ConversionResult` with counts

**`convertSingleMatch(match: FootnoteMatch): Promise<{ error?: string }>`**  
Converts one accepted match. Same re-scan approach but for one number only.

**`deleteFootnoteBlocks(blocks: FootnoteBlock[]): Promise<{ data?: number; error?: string }>`**  
Deletes original footnote block paragraphs by re-scanning for their content. Returns count deleted.

### Completion Test

- [ ] `src/footnotes/converter.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 3 functions exported

---

## PHASE 28: FOOTNOTE DETECTION UI

**Prerequisites:** Phases 9, 25, 26, 27 complete  
**Resume Prompt:** `Replace the FootnotePanel.tsx stub with the full footnote detection panel. This is the most complex UI component in the project. Read the full phase before writing any code. Follow Phase 28 of the plan exactly.`

### Instructions

Replace `src/taskpane/components/FootnotePanel.tsx`. Props: `{ setError: (e: string | null) => void }`.

**State:**
```typescript
const [scanResult, setScanResult] = useState<ScanResult | null>(null);
const [scanning, setScanning] = useState(false);
const [converting, setConverting] = useState(false);
const [conversionResult, setConversionResult] = useState<ConversionResult | null>(null);
const [matches, setMatches] = useState<FootnoteMatch[]>([]);
```

**Three-stage layout:**

**Stage 1 — Pre-scan** (when `scanResult === null` and not scanning):  
Brief instruction paragraph. "Scan Document" `Button`. Calls `scanDocument()`, then `buildScanResult()`, sets `scanResult` and `matches`.

**Stage 1 (loading)** (when `scanning === true`):  
Fluent UI `Spinner` with label `"Scanning document..."`.

**Stage 2 — Review table** (when `scanResult !== null` and `conversionResult === null`):  
Summary line: `"Found N matches, M unmatched markers, K orphaned blocks."`  
If `matches.length === 0`: message "No convertible footnotes found." Button: "Scan Again".  
Otherwise, Fluent UI `Table` with columns: **#**, **Marker Context**, **Footnote Text**, **Confidence**, **Include**.
- **Marker Context**: `contextBefore` + bold marker number + `contextAfter`, max 120 chars total
- **Footnote Text**: block text, max 80 chars
- **Confidence**: `Badge` — green for high, yellow for medium, red for low
- **Include**: `Checkbox` bound to `match.accepted`, toggleable per row, updates `matches` state

Below the table:
- Unmatched markers line (if any)
- Orphaned blocks line (if any)
- Button row: "Accept All" | "Reject All" | "Convert Selected" | "Scan Again"

**Stage 2 (converting)** (when `converting === true`):  
Spinner with label `"Converting footnotes..."`.

**Stage 3 — Result** (when `conversionResult !== null`):  
`"Converted N footnote(s) successfully."` If `conversionResult.failed > 0`: error count and list. Button: "Scan Again" (resets all state to Stage 1).

### Completion Test

- [ ] Footnotes tab shows pre-scan stage initially
- [ ] Scan populates the review table (test with a document containing pasted footnotes)
- [ ] Checkboxes toggle individual rows
- [ ] "Convert Selected" converts accepted matches to real Word footnotes
- [ ] Original inline markers are removed from body text after conversion
- [ ] `npx tsc --noEmit` passes

---

## PHASE 29: END-TO-END TESTING

**Prerequisites:** All phases 1–28 complete  
**Resume Prompt:** `Review the chambers-word-toolkit codebase. Run npx tsc --noEmit. Report any TypeScript errors or broken imports. Fix errors found. Do not add new features.`

### Test documents to prepare

**formatting-test.docx:** Mixed fonts, paragraph formats, and styles. Use it to test every formatting engine button.

**pleadings-test.docx:** A Nigerian Statement of Claim with three indent levels. Use it to test pleadings numbering, exhibit marker, reliefs, and cause heading.

**footnotes-test.docx:** A document with pasted content from a PDF. Must contain:
- At least 3 superscript markers in the body
- At least 2 bracketed markers
- Corresponding footnote text blocks at the end, preceded by a line of underscores
- One unmatched marker and one orphaned block for edge case coverage

A Supreme Court or Court of Appeal judgment downloaded from LPELR and pasted into Word (not via Paste Special) will almost always produce the exact pattern the scanner targets.

### Formatting engine tests

- [ ] Bold toggle on mixed selection
- [ ] Font size increase / decrease
- [ ] Apply Times New Roman
- [ ] Clear All Formatting → returns to Normal style
- [ ] Justify alignment applies to selected paragraphs
- [ ] Remove Empty Paragraphs — toast shows count
- [ ] UPPERCASE / Sentence Case conversion

### Legal workflow tests

- [ ] Cause heading block inserts correctly
- [ ] `1. / 1.1 / 1.1.1` numbering applies correctly
- [ ] Jurat block inserts
- [ ] Exhibit A → B → C with counter reset
- [ ] Case citation [2020] 15 NWLR (Pt. 1721) 123 inserts in italics
- [ ] LPELR format: [2021] LPELR-55302(CA) correct
- [ ] Statutory citation Section 175(1) of the CAMA 2020 inserts

### Footnote detection tests

- [ ] Scan completes and shows review table
- [ ] Confidence badges show correct colours
- [ ] Unmatched and orphaned items shown below table
- [ ] Deselect one match → converts without it
- [ ] Conversion creates real Word footnotes (verify in Word's footnote pane)
- [ ] Inline markers removed from body after conversion
- [ ] Block paragraphs removed after conversion

---

## PHASE 30: PACKAGING AND SIDELOADING

**Prerequisites:** Phase 29 complete  
**Resume Prompt:** `Prepare the chambers-word-toolkit project for sideloading and distribution. Follow Phase 30 of the plan exactly.`

### Instructions

#### 30a. Production build

```
npm run build
```

Output goes to `dist/`.

#### 30b. Stable sideload on the development machine

```
npm install -g office-addin-dev-server
office-addin-dev-server start --port 3000 --directory dist
```

The dev server hosts the add-in at `https://localhost:3000`. The manifest already references this URL from `npm start` sessions.

#### 30c. Distribute to another machine (chambers use)

1. Copy `dist/` to a network share accessible from the target machine, e.g., `\\CHAMBERS-SERVER\addins\chambers-word-toolkit\dist\`
2. In `manifest.xml`: replace all `https://localhost:3000` references with the network share URL or a local HTTPS address on the hosting machine
3. On the target machine: Word → File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → add the network share path
4. Restart Word → Insert → My Add-ins → Shared Folder → find PawaTuls

#### 30d. Verify on target machine

- [ ] Add-in loads on the second machine
- [ ] All three panels work
- [ ] Footnote detection converts footnotes correctly

---

---

## APPENDIX A: OFFICE.JS QUICK REFERENCE

### Standard Word.run pattern

```typescript
import { safeWordRun } from "../shared/wordUtils";

export async function myFunction(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text,font");
    await context.sync();
    // operate
    await context.sync();
  });
}
```

### Always load before reading

```typescript
// WRONG — throws "object not loaded"
const text = selection.text;

// CORRECT
selection.load("text");
await context.sync();
const text = selection.text;
```

### Null means mixed selection

```typescript
// In Office.js, a mixed property (some bold, some not) returns null
// In COM interop it returned 9999999 (wdUndefined) — these are different
selection.load("font/bold");
await context.sync();
if (selection.font.bold === null) {
  // mixed — set to true
} else {
  // toggle
}
```

### Iterate paragraphs in reverse for deletions

```typescript
const items = paragraphs.items.slice().reverse();
for (const para of items) {
  if (para.text.trim() === "") {
    para.delete();
  }
}
await context.sync();
```

### Create a footnote (WordApi 1.5)

```typescript
// markerRange is the range of the inline marker in the body
markerRange.insertFootnote("Footnote text here.");
await context.sync();
```

### Search with Word wildcards (not JavaScript regex)

```typescript
// Word wildcard syntax — NOT JavaScript regex
const results = context.document.body.search("[0-9]@", { useWildcards: true });
results.load("font/superscript,text");
await context.sync();
```

### Insert HTML at selection

```typescript
const selection = context.document.getSelection();
selection.insertHtml(htmlString, Word.InsertLocation.after);
await context.sync();
```

---

## APPENDIX B: NIGERIAN LEGAL TEMPLATES

### Standard FCT High Court cause heading

```
IN THE HIGH COURT OF THE FEDERAL CAPITAL TERRITORY
HOLDEN AT ABUJA

SUIT NO: FCT/HC/CV/[NUMBER]/[YEAR]

BETWEEN

[CLAIMANT FULL NAME AND DESCRIPTION]
                                                    CLAIMANT

AND

[DEFENDANT FULL NAME AND DESCRIPTION]
                                                    DEFENDANT
```

### Affidavit structure

```
1.  I, [FULL NAME], of [ADDRESS], do hereby make oath and state as follows:

2.  [Paragraph of facts]

3.  That I depose to this Affidavit in good faith and believe the contents to be
    true and correct to the best of my knowledge, information and belief.

SWORN to at Abuja this _____ day of ________________________ 20_____.

________________________________
DEPONENT

Before me:

________________________________
COMMISSIONER FOR OATHS
```

### Citation formats

**NWLR:** `[2020] 15 NWLR (Pt. 1721) 123`  
**LPELR:** `[2021] LPELR-55302(CA)` or `[2022] LPELR-59001(SC)`  
**SCNJ:** `[2019] 2 SCNJ 45`  
**FWLR:** `[2020] 3 FWLR (Pt. 1050) 200`  
**Statutory:** `Section 175(1) of the Companies and Allied Matters Act 2020`

---

## APPENDIX C: FOOTNOTE DETECTION — ALGORITHM DETAIL

### Why the feature exists

PDF readers and conversion tools flatten footnotes into the document body. The footnote reference number appears inline (as superscript or bracket) and the footnote text appears at the bottom of the page section, separated by a horizontal line. The scanner finds both artefacts and the converter restores them as genuine Word footnotes.

### Five edge cases the converter must handle

**1. Multi-digit numbers (10, 11, 12...).**  
The wildcard `[0-9]@` matches them as a unit. Verify this on the test machine. If consecutive superscript characters `1` and `0` are returned as two separate hits, merge adjacent superscript digit hits with the same paragraph index and adjacent positions.

**2. No separator in the document.**  
Some conversion tools do not insert a separator line. If no separator is found, treat the final cluster of numbered paragraphs as the footnote blocks and assign `"medium"` confidence to all matches.

**3. Multi-paragraph footnote text.**  
If paragraph N starts with a number and paragraph N+1 does not, N+1 is a continuation of N. Concatenate with a space before inserting into the Word footnote.

**4. Conversion must process highest paragraph index first.**  
Deleting a paragraph shifts all subsequent paragraph indices. Process in descending index order. This is enforced in the converter by sorting before processing.

**5. Document edited between scan and convert.**  
The converter does not use stored `Word.Range` objects from the scan session. It re-scans at conversion time. If the user edits the document after scanning, the re-scan finds the current state. This avoids stale reference errors at the cost of a brief re-scan delay on Convert.

---

## APPENDIX D: CROSS-CUTTING PRODUCT RULES (V1 MERGE)

These override silence in earlier phases. They do not change any built module's
contract; they constrain unbuilt work (Phases 16–28, 31–34).

**R1 — Formatting and cleanup stay separate.** Applying a font/paragraph/style
profile must never silently remove blank paragraphs, line breaks, tabs, or
structural content. Cleanup runs only when the user invokes a cleanup command.

**R2 — Preview before destructive scope.** Every whole-document or destructive
panel action shows what it will do first (counts of affected paragraphs/matches)
and applies only on explicit user confirmation. Modules return counts; panels
display them.

**R3 — Idempotence.** Normalizing operations must be safe to run twice: the
second run changes nothing and reports zero. Test this for every cleanup,
conversion, and numbering function.

**R4 — Exclusions: skip and report where safety is unprovable.** Never delete or
rewrite a paragraph the API cannot prove safe. Minimum exclusion set for
deletion-class ops: table-cell paragraphs (done — `parentTableOrNullObject`),
paragraphs carrying bookmarks/fields/section breaks, tracked-change revisions.
When an item is skipped, count it as skipped, not silent.

**R5 — No document-text logging.** Diagnostics record operation IDs, counts,
locations (paragraph indices), and error categories. Never log paragraph text,
footnote content, or names/parties to console, files, or telemetry.

**R6 — General-purpose engine; legal content is data, not defaults.** No Nigerian
legal rule, court name, or citation format may appear as a default inside the
formatting engine or shared helpers. Legal workflows live in `src/legal/` and the
Legal panel only.

**R7 — No silent destructive modification.** Every destructive command has
preconditions (checked), a change plan (counts/preview), and postconditions
(verified). Stop safely when a range cannot be re-resolved or verification fails.
Never claim rollback; report partial state honestly.

---

## APPENDIX E: FORMATTING PROFILES (NEW FEATURE — PHASES 31–32)

Motivation (v1 §2.2): users apply recurring property sets to different document
types; legal filings are demanding instances, not the only instance. A profile
bundles font + paragraph + style selections under a versioned, portable schema.

### Shared profile types (added to `src/shared/types.ts` in the Phase 31 session)

```typescript
export interface FormattingProfile {
  schemaVersion: 1;
  id: string;            // stable, unique, e.g. "court-body-1"
  name: string;          // display name
  description?: string;
  style?: string;        // named style applied first (must exist)
  font?: FontOptions;
  paragraph?: ParagraphOptions;
}
```

### PHASE 31: PROFILE STORE AND APPLY

**Prerequisites:** Phases 10, 11, 12 complete
**Resume Prompt:** `Create src/formatting/profiles.ts. Follow Phase 31 of the plan exactly.`

Create `src/formatting/profiles.ts`. Implement these exported functions
(reuse `applyFont`, `applyParagraphFormat`, `applyStyle` — no duplicated logic):

**`validateProfile(input: unknown): { data?: FormattingProfile; error?: string }`**
Pure. Rejects non-objects, missing/empty `id`/`name`, `schemaVersion !== 1`,
wrong-typed `style`, and any unknown top-level field (list it in the error).
Missing `font`/`paragraph`/`style`/`description` are legal (a profile changes
only what it targets — R1).

**`applyProfile(profile: FormattingProfile, applyTo: "selection" | "document"): Promise<{ data?: number; error?: string }>`**
Validates first (returns error without touching the document on failure).
Order: `applyStyle(profile.style)` if present, then font, then paragraph.
Scope `"document"` operates on body paragraphs directly without moving the
user's cursor (audit lesson: never move selection). Returns paragraphs affected.

**`exportProfile(profile: FormattingProfile): string`**
Pure. `JSON.stringify(profile, null, 2)` after validation.

**`importProfile(json: string): { data?: FormattingProfile; error?: string }`**
Pure. `JSON.parse` in try/catch (parse failure → error, never throw), then
`validateProfile`. Never execute imported content as code.

**Builtins** (exported `const BUILTIN_PROFILES: FormattingProfile[]`):
- `court-body`: Times New Roman, 12, justified, first-line indent 0.5 inch,
  single spacing, space after 6.
- `court-heading`: Times New Roman, 14, bold, centered, space before/after 12.
- `single-space-reset`: single spacing, space before/after 0, no indents.
Builtin values must satisfy R6 (labels only — no legal logic in the engine).

### Completion Test

- [ ] `src/formatting/profiles.ts` exists
- [ ] `npx tsc --noEmit` passes
- [ ] All 5 exports present; invalid JSON/unknown fields rejected with messages
- [ ] Applying twice is a no-op the second time (R3)

---

### PHASE 32: PROFILE SECTION IN FORMATTING PANEL

**Prerequisites:** Phases 16, 31 complete
**Resume Prompt:** `Add the Profiles accordion section to FormattingPanel.tsx. Follow Phase 32 of the plan exactly.`

Add a fourth Accordion section **Profiles** to `src/taskpane/components/FormattingPanel.tsx`:
Profile `Dropdown` (builtins + imported, shows `name`), scope `RadioGroup`
(Selection / Document), **Apply** button (affected-count toast per R2),
**Import** (textarea + button, errors via `setError`), **Export** (read-only
textarea with selected profile JSON for copy). No Office.js calls in the
component. `npx tsc --noEmit` passes.

---

## APPENDIX F: NUMBERING AUDIT AND REPAIR (NEW FEATURE — PHASES 33–34)

Motivation (v1 §2.3): the pleadings formatter (Phase 17) prefixes blindly. A
general-purpose add-in must distinguish **native Word list numbering** from
**manually typed number prefixes**, audit sequences, and repair only through
preview-approved plans. Never auto-renumber a document.

### Shared numbering types (added to `src/shared/types.ts` in the Phase 33 session)

```typescript
export type NumberingKind = "native" | "manual" | "plain";
export interface NumberedParagraph {
  paragraphIndex: number;
  kind: NumberingKind;
  prefix?: string;    // typed prefix text for "manual"
  listId?: number;    // native list id when detectable
  level?: number;     // nesting level when detectable
}
export interface NumberingIssue {
  paragraphIndex: number;
  issue: "gap" | "duplicate" | "restart" | "mixed-style" | "orphan-level";
  detail: string;
}
export interface NumberingRepairPlan {
  edits: { paragraphIndex: number; before: string; after: string }[];
}
```

### PHASE 33: NUMBERING AUDIT MODULE

**Prerequisites:** Phases 7, 8 complete
**Resume Prompt:** `Create src/formatting/numbering.ts. Follow Phase 33 of the plan exactly.`

Create `src/formatting/numbering.ts` with three exported functions:

**`auditNumbering(applyTo: "selection" | "document"): Promise<{ data?: { items: NumberedParagraph[]; issues: NumberingIssue[] }; error?: string }>`**
Read-only, no mutations. Classification per paragraph:
- **native**: paragraph OOXML `<w:pPr>` contains `<w:numPr>` (native list
  membership — the only reliable Office.js signal; verify against the installed
  Word build in the Phase 29 session).
- **manual**: text matches the typed-prefix language
  (same pattern family as pleadings removal).
- **plain**: everything else.
Issue detection on manual sequences per indent level: gap (`1.,3.`), duplicate
(`2.,2.`), restart (reset without level change), mixed-style (`1.` then `(a)` at
one level), orphan-level (level 3 with no level-2 parent). Native items are
reported, never rewritten by the auditor.

**`previewRepair(issues: NumberingIssue[], items: NumberedParagraph[]): NumberingRepairPlan`**
Pure. Proposes minimal prefix rewrites for `gap`/`duplicate` issues only —
restarts and mixed-style are user intent until approved (listed, not rewritten).
`before`/`after` are full-prefix strings for panel display.

**`applyRepairPlan(plan: NumberingRepairPlan): Promise<{ data?: number; error?: string }>`**
Applies only the listed edits, re-resolving each paragraph at apply time in
descending index order and replacing the old prefix with the new one. Any
paragraph whose current text no longer starts with `before` is skipped and
reported in the error text with its index (R4/R7). Returns applied count.

### Completion Test

- [ ] `src/formatting/numbering.ts` exists, `npx tsc --noEmit` passes, 3 fns exported
- [ ] Auditor is read-only (no mutations); native lists never rewritten by repair
- [ ] Repair of a `1.,2.,2.,4.` selection yields `1.,2.,3.,4.`; rerun reports 0 (R3)

---

### PHASE 34: NUMBERING SECTION IN FORMATTING PANEL

**Prerequisites:** Phases 16, 33 complete
**Resume Prompt:** `Add the Numbering accordion section to FormattingPanel.tsx. Follow Phase 34 of the plan exactly.`

Add a fifth Accordion section **Numbering**: scope `RadioGroup`, **Audit**
button → issue list (kind badges native/manual/plain + issue descriptions),
**Preview Repair** (before→after rows), **Apply Repair** (the previewed plan
only). Native items display "managed by Word — not edited".
`npx tsc --noEmit` passes.

---

## APPENDIX G: FOOTNOTE HARDENING DELTAS + TEST FIXTURES (V1 MERGE)

These strengthen Phases 25–29 without changing their shape. Agents running those
phases apply the deltas below alongside the phase text.

### Scanner deltas (Phase 25)

- **Enumerate existing genuine Word footnotes first** via the supported collection
  API (verify the exact member on the installed build). Record
  `existingFootnotes: number` in the scan output (add the field to `ScanResult`
  in the Phase 25 session) and exclude those paragraphs from block candidates.
  Never propose converting an already-correct note.
- Markers that lost superscript are already covered (bracketed/parenthetical
  passes); record formatting-loss as evidence, not as lower confidence.

### Matcher deltas (Phase 26)

- **One-to-one:** a block feeds at most one marker. Duplicate-number blocks match
  at best `"medium"`, surface together as conflicting alternatives, and the UI
  never auto-picks one. Never invent note text; unmatched stays unmatched.
- Review states: `high` → approve-ready, `medium`/`low` → needs-review,
  conflicts → must-pick-one.

### Converter deltas (Phase 27)

- **Idempotent:** re-running converts nothing new (re-scan finds no accepted
  markers left). Test: convert twice, second run converts 0.
- **Verify before delete:** after `insertFootnote`, confirm the note exists with
  matching text before deleting marker/block source; on mismatch keep the source
  and record in `errors[]`.
- **Preserve surroundings:** marker replacement must not move the reference
  position or eat adjacent punctuation/spacing. Save/restore user selection.

### Panel deltas (Phase 28)

- Show the existing-genuine-footnote count; conflicts render as must-pick-one rows.

### Fixture matrix (Phase 29; synthetic or redacted text only)

1. Genuine Word footnotes only (expect: zero proposals, existing count correct).
2. Flattened end-notes, superscript intact.
3. Flattened notes, superscript lost (plain numbers).
4. Notes interleaved between body paragraphs (not end-of-doc).
5. Markers in a different font from body.
6. Numbered pleadings adjacent to note-like paragraphs (expect: no proposals).
7. Law-report citations / bibliography (expect: no proposals).
8. Duplicate numbers, gaps, restarts, missing note text.
9. Mixed genuine + flattened notes.
10. Same number twice in different scopes.
11. Marker with no note; note with no reference.
12. Dates, decimals, section numbers, exponents, superscript ordinals near notes
    (expect: no false proposals).
13. Tables, fields, section breaks, tracked changes around candidates.
14. Long document (100+ notes) for performance sanity.

### Integrity + honesty rules (Phase 29)

- For every destructive test: unrelated text, styles, native lists, existing
  footnote count/text, fields, hyperlinks, tracked changes outside the target
  must be unchanged.
- Raw `.docx` ZIP equality is not the assertion (Word rewrites XML on save) —
  use semantic checks.
- Report unit-only vs Word-run for every test; never claim a Word test passed
  unless it ran in Word. Record Word build/bitness with results.

---

## APPENDIX H: AMENDMENTS TO UNBUILT PHASES

Read before running any listed phase. Amendments override the phase text on
conflict (newer decision wins).

- **Phase 16:** every button surfaces counts (toast/summary per R2);
  whole-document actions confirm first; explain disabled states (audit: silent
  veto); no Office.js calls in the component (unchanged).
- **Phase 17:** `applyPleadingsNumbering` must strip existing number prefixes
  (same regex as removal) before applying — numbering must be idempotent (R3).
  Removal operates on selection/document explicitly, never whole-document by
  surprise.
- **Phase 24:** panel rules as Phase 16; exhibit/schedule counters display next
  values; destructive inserts confirm.
- **Phase 25:** apply scanner deltas (Appendix G) alongside phase text.
- **Phase 26:** apply matcher deltas (Appendix G).
- **Phase 27:** apply converter deltas (Appendix G) plus the four hardened
  converter requirements in `decompilation/AUDIT.md`.
- **Phase 28:** apply panel deltas (Appendix G).
- **Phase 29:** add the Appendix G fixture matrix, integrity checks, and honesty
  rules to the test pass. Do not weaken failing tests to green.
- **Phase 13 (amended, stress-test fix):** case/trim/collapse/tabs rewrite
  `<w:t>` contents via `rewriteTextPreservingRuns` (formatting-preserving);
  plain `insertText(Replace)` only as round-trip-mismatch fallback.
- **Phase 25 (amended):** block starts exclude heading-style, native-list, and
  in-table paragraphs; non-hygienic paras break continuation chains.
- **Phase 26 (amended):** only high-confidence first references default
  accepted; repeats and block-sharers surface unchecked (one-to-one both ways).
- **Phase 33 repair (amended):** prefix-only rewrite via
  `rewriteTextPreservingRuns`, same fallback.

---

## APPENDIX I: KUTOOLS PARITY — REMOVAL, NOTES, CAPTIONS, SELECT, TABLES (PHASES 35–40)

Rule: every KuTools document feature is covered, matched, or explicitly reasoned
about in `docs/KUTOOLS_COVERAGE.md`. The phases below close the remaining gaps.
All follow the global rules (counts returned, confirms in UI, table guards, no
selection churn, `tsc` clean) and Appendix D R1–R7.

### PHASE 33 AMENDMENT (better native-list signal — replaces OOXML `numPr` rule)

Classify via `paragraph.listItemOrNullObject` (verified API): null-object →
not native; otherwise `kind: "native"`, `level: listItem.level`. Keep the OOXML
`<w:numPr>` check only as a Phase 29 cross-validation, not the primary signal.

### PHASE 35: REMOVAL MODULE

**Prerequisites:** Phases 7, 8 complete
**Resume Prompt:** `Create src/formatting/removal.ts. Follow Phase 35 of the plan exactly.`

Create `src/formatting/removal.ts`. Every function takes
`applyTo: "selection" | "document" = "document"`, resolves the paragraph set
(selection paragraphs or body paragraphs), and returns
`{ data?: number; error?: string }` (items removed/changed). Structural
whole-document ops confirm in UI (Phase 36), not here. Verified APIs:
`body.tables/inlinePictures/comments`, `TableRow.delete`, `TableCell.deleteColumn`,
`ContentControl.delete`, `document.deleteBookmark`, `Hyperlink.delete`,
`parentTableOrNullObject`, per-paragraph `getOoxml`/`insertOoxml`.

- `removePageBreaks` / `removeColumnBreaks`: strip `<w:br w:type="page|column"/>`
  from each paragraph's OOXML (batched: queue all, one sync, rewrite changed).
- `removeSectionBreaks`: strip `<w:sectPr>…</w:sectPr>` from `w:pPr` (merges into
  previous section — panel confirms per R2).
- `removeAllBreaks`: calls the three above in one `Word.run`; returns total.
- `removeHiddenText`: strip `<w:r>` runs containing `<w:vanish/>` from OOXML.
- `removeComments`: `body.comments` reverse `delete()`.
- `removeDropCap`: strip `<w:dropCap…/>` from `w:pPr`.
- `removeFrames`: strip `<w:framePr…/>` from `w:pPr`.
- `removePictures`: strip `<w:drawing>` elements containing `<pic:pic/>`
  (covers inline and floating pictures).
- `removeTextBoxes`: strip runs containing `<w:txbxContent/>`.
- `removeTables`: reverse `table.getRange().delete()` (verify at build; fallback
  delete all rows).
- `removeEmptyTableRows`: per table, reverse-delete rows whose cells are all
  whitespace (via `row.values`).
- `removeEmptyTableColumns`: per table, delete columns whose cells are all
  whitespace (via first-cell `deleteColumn()`), reverse column order.
- `removeContentControls(keepText = true)`: `getContentControls()` reverse
  `delete(keepText)` (verify signature at build).
- `removeBookmarks`: names from body-range `getBookmarks()`, `deleteBookmark`
  each (reverse order).
- `removeHyperlinks`: per paragraph `getHyperlinkRanges()` → `hyperlink.delete()`
  (verify display text is kept; else replace range with `textToDisplay`).
- `tabToSpace`: paragraph text `/\t/g → " "`, write back changed only.
- `removeEmptyLines`: paragraph text `/\v{2,}/g → "\v"`, write back changed only.

### Completion Test

- [ ] `src/formatting/removal.ts` exists, `npx tsc --noEmit` passes
- [ ] Rerun on a cleaned document reports 0 for every function (R3)
- [ ] Tables/cells are never corrupted (row/col deletes only via table APIs)

### PHASE 36: REMOVE SECTION IN FORMATTING PANEL

**Prerequisites:** Phases 16, 35 complete
**Resume Prompt:** `Add the Remove accordion section to FormattingPanel.tsx. Follow Phase 36 of the plan exactly.`

Sixth Accordion section **Remove**, grouped buttons (Breaks / Hidden & Marks /
Tables & Objects / References), each surfacing counts; section-break and
remove-tables buttons confirm first (R2). No Office.js calls. `tsc` passes.

### PHASE 37: NOTE MANAGEMENT MODULE + FOOTNOTE PANEL SECTION

**Prerequisites:** Phases 7, 8, 28 complete
**Resume Prompt:** `Create src/footnotes/notes.ts and extend FootnotePanel.tsx. Follow Phase 37 of the plan exactly.`

Create `src/footnotes/notes.ts` (verified APIs: `body.footnotes`/`endnotes`
`NoteItemCollection`, `NoteItem.delete`). Descending-order deletes; counts returned.

- `removeFootnotes(applyTo)` / `removeEndnotes(applyTo)`: reverse-delete the
  collection (selection scope = notes whose reference is inside the selection
  range — verify `NoteItem.reference` at build; else document-scope only with
  panel note).
- `convertEndnoteToText()`: transcribe numbered text (actual reference numbers,
  not loop indices) then backward-delete (KuTools parity minus its bugs).
- `convertFootnoteToEndnote` is DROPPED (unsupported): Office.js has no
  endnote-insert API (`insertFootnote` creates footnotes only; OOXML insert
  cannot add the endnotes part). See `docs/KUTOOLS_COVERAGE.md`.

Extend `FootnotePanel.tsx` with a **Manage Notes** section: four buttons with
counts + confirms; conversion buttons note the plain-text limitation. `tsc` passes.

### Completion Test

- [ ] `src/footnotes/notes.ts` exists (4 fns), panel extended, `tsc` passes
- [ ] Remove twice → second run 0 (R3); conversions verified by recount

### PHASE 38: CAPTIONS MODULE + PANEL BUTTONS

**Prerequisites:** Phases 7, 8, 16 complete
**Resume Prompt:** `Create src/formatting/captions.ts and extend FormattingPanel.tsx. Follow Phase 38 of the plan exactly.`

Create `src/formatting/captions.ts`:

- `insertCaption(kind: "table" | "figure" | "equation", label?: string): Promise<{ error?: string }>`
  Inserts `<label> <SEQ field>` at the cursor via `insertOoxml` (label defaults
  `Table`/`Figure`/`Equation`); numbering is automatic through the SEQ field.
  Verify SEQ-field insertion XML against the installed Word build; update all
  fields after insert so the number renders.
- `refreshCaptions(): Promise<{ data?: number; error?: string }>`
  Updates every SEQ field (`fields.getByTypes` sequence type — verify enum at
  build, else update all fields) and returns the count.

Extend the FormattingPanel **Styles and Tools** section: kind dropdown, label
input, Insert button, Refresh button with count. `tsc` passes.

### PHASE 39: SELECT MODULE + PANEL SECTION

**Prerequisites:** Phases 7, 8, 16 complete
**Resume Prompt:** `Create src/formatting/select.ts and extend FormattingPanel.tsx. Follow Phase 39 of the plan exactly.`

Office.js has a single selection — no multi-select. Contract for every function:
`{ data?: { total: number }; error?: string }`: selects the FIRST match with
`range.select()` and reports how many match (user repeats to step through).

Create `src/formatting/select.ts` (all take no scope — whole document):
`selectTables`, `selectEmptyParagraphs` (same guards as removal),
`selectHeadings` (style name starts with `"Heading"`), `selectShortParagraphs`
(`maxChars = 25` default, KuTools `ShortParaMaxChars` parity), `selectPictures`
(first `inlinePictures` range). New FormattingPanel section **Select** with five
buttons showing match counts. `tsc` passes.

### PHASE 40: TABLES, PICTURES, WIDOW CONTROL + PANEL SECTION

**Prerequisites:** Phases 7, 8, 16 complete
**Resume Prompt:** `Create src/formatting/tables.ts, extend paragraph.ts, extend FormattingPanel.tsx. Follow Phase 40 of the plan exactly.`

Create `src/formatting/tables.ts`:

- `convertTableToText(separator = " "): Promise<{ data?: number; error?: string }>`
  Applies to selected tables (resolve via selection paragraphs'
  `parentTableOrNullObject`, deduplicated): `row.convertToText(separator)` per
  row, reverse order. Returns tables converted. (KuTools `ConvertTableToText`
  separator parity, including its `"-"` default language — default `" "` here.)
- `repeatHeaderRows(enable = true): Promise<{ error?: string }>`
  Sets `table.headerRowCount = enable ? 1 : 0` on selected tables.
- `insertTableDiagonalHeader(type: "down" | "up"): Promise<{ error?: string }>`
  DROPPED (unsupported): true diagonal borders have no Office.js path —
  `BorderLocation` has no diagonal member and `TableCell` exposes no OOXML
  access, so the border cannot be injected at cell precision. See
  `docs/KUTOOLS_COVERAGE.md`.
- `resizeImages(percent: number): Promise<{ data?: number; error?: string }>`
  Scales every inline picture's width/height by `percent/100` (validate 10–400,
  honor `lockAspectRatio`); returns count.
- Extend `src/formatting/paragraph.ts` with `controlWidows(applyTo)`:
  sets `widowControl = true` via OOXML `w:widowControl` (Paragraph exposes no
  direct prop) — the safe, supported replacement for KuTools' font-shrink
  `AdjustParagraphEnd` heuristic (rejected, AUDIT.md #10).

New FormattingPanel section **Tables & Pictures**: separator input + Convert,
Repeat-headers toggle, Diagonal buttons, percent SpinButton + Resize, Widow
control button. `tsc` passes.

### Completion Test (Phases 35–40)

- [ ] All modules exist, `npx tsc --noEmit` passes
- [ ] Every removal/conversion rerun reports 0 (R3)
- [ ] Counts surface in UI; structural ops confirm first (R2)

---

## APPENDIX J: FOOTNOTE WORKFLOW EXTENSIONS (PHASES 41–43)

### PHASE 41: MANUAL FOOTNOTE-BLOCK AREA OVERRIDE

**Status: built.** When auto-detection confidence is low, the user highlights
the footnote block region; `scanSelectionBlocks()` (`src/footnotes/scanner.ts`)
parses selected paragraphs with the same block patterns (no separator needed,
`paragraphIndex: -1`, continuation merging included). The panel rebuilds matches
against the already-scanned markers via `buildScanResult` (medium-grade review
by design — manual areas never claim high confidence). Conversion re-locates by
text, so unknown indices hold through the flow.

### PHASE 42: CITATION-SYSTEM IDENTIFICATION

**Status: built (`src/footnotes/citations.ts`, pure).** Recognizes Nigerian Law
Reports (NWLR/LPELR/SCNJ/FWLR/NLR/ANLR/NSCC/UILR/JELR), statutory citations,
neutral court citations, academic-bracket references, and author–date.
The scanner classifies every paragraph (`detectCitationSystems`), flags
citation-dense blocks (`FootnoteBlock.citationDense`), reports system counts in
`ScanResult.citationSystems`, and the matcher caps citation-dense matches at
medium (bibliographies and law-report lists are never high-confidence footnote
blocks). The panel lists detected systems and warns on mixed usage
(consistency check). Verified by executed assertions (reporters, statutes,
neutral, author–date, plain-text negative, dense indices, cap, conflicts).

### PHASE 43: ORDINAL SUPERSCRIPTING

**Status: built (`superscriptOrdinals` + pure `superscriptOrdinalsInOoxml` in
`src/formatting/text.ts`, button in FormattingPanel).** Finds `\b\d+(st|nd|rd|th)\b`
(1st, 2nd, 3rd … 200th) and superscripts only the suffix via per-paragraph OOXML
run-splitting: original run properties are restored on continuing text,
already-superscript runs are skipped, cross-run matches count as skipped.
Returns `{ changed, skipped }`. Verified by executed assertions (split shape,
format preservation, superscript skip, balanced XML).

---

## APPENDIX J ADDENDUM: EVIDENCE MODEL + RIBBON TAB

### Evidence model (replaces the single-confidence design)

Footnote recovery is three separate decisions with separate scores
(`EVIDENCE_MODEL v1` in `src/footnotes/matcher.ts`; weights provisional,
thresholds high ≥ 70 / medium ≥ 40 — not probabilities):

- **Block classification** (`FootnoteBlock.blockLikelihood` 0–100): sequential-run
  membership +20, note region +10, citation combinations +15/+5, ibid chains +10,
  smaller-than-body font +5, hanging indent +5; contradictions explicit
  (bibliography heading −30, others −15).
- **Citation content** (`citationLikelihood` 0–100): systems ×25, shorthand ×10,
  lone year +5 (Tier D: weak cues never decisive alone).
- **Pairing** (`pairingEvidence[]` + `contradictions[]` on each match):
  number/region base, sequence continuity +10, ambiguity/density/structure caps
  at medium, reference-signal boost low→medium (non-dense only).
- **Boundaries:** long unmarked citation-free prose after a multi-note cluster
  ends the cluster (no fixed cutoff as sole decision); single-footnote pages
  exempt; manual-area selections re-score with the last block flagged
  boundary-uncertain; possible prose merges flagged, never auto-split.
- Citation systems recognized: Nigerian Law Reports, Statutes, Neutral,
  Academic Bracket (covers IEEE/Vancouver/GOST-numeric shapes), Author-Date
  (Harvard-like), OSCOLA, Bluebook, Chicago, MLA, Ibid Chain (chain signal only —
  never density, since a lone "Ibid." is usually a genuine note).
- Review UI shows evidence/contradictions per row (badge tooltip); conflicts
  render as must-pick-one rows; only high-confidence first references start
  checked; restoration stays approval-gated and verified.

### Ribbon tab (Phase 44 — built)

Custom `PawaTuls.Tab` (renders after built-in tabs) with Formatting group
(Format Pane, Bold, Clear All, Empty Paras, Ordinals) and Footnotes group
(Footnotes Pane, Remove Notes, Endnote→Text). Pane buttons deep-link via
`?tab=` (App reads it for the initial tab); quick actions are headless
`ExecuteFunction` calls (failures silent by design — use pane buttons for
reported runs). Manifest 1.1.0.0 — re-sideload (`npm start`) to pick it up.
Resource IDs capped at 23 chars (schema limit 32).

### Review adjudications (Claude input — accepted unless noted)

1. **Thresholds defined:** independent = different evidence tier (EVIDENCE_MODEL
   reviewPolicy); needs-review minimum = 2 supporting cues without Tier A
   linkage (provisional, calibratable); boundary cutoff = max(400 chars, 3×
   median of cluster note lengths).
2. **Global assignment specified:** greedy per-number matching with order
   alignment + human backtracking (duplicates/conflicts surface unchecked).
   Hungarian rejected on record: marker numbers are near-unique keys, so global
   optimization adds nothing once duplicates go to review.
3. **Merge detection specified:** same-paragraph body marker (short numbers
   only — years excluded) + citation tail + 300-char floor → MergeCandidate
   with proposed split offset; never auto-split.
4. **Block-selection UI:** mechanism (a) — select in the document body, confirm
   via task-pane button ("Use Selection as Footnote Area"). Already built (41).
5. **Gap decay:** gap of 1 in runs of 8+ markers → −5; gap of 3+ caps the whole
   run at needs-review with an explicit contradiction.
6. **Mixed native+flattened:** evidence-gated disambiguation — duplicated 1..k
   leading superscript run is dropped as native; otherwise everything stays for
   review. Numbers are matched internally, never against the doc footnote count.
7. **Nigerian Tier B:** LPELR URLs, bare reporter abbreviations, suit-number
   patterns, Per J./JSC quotes as structure-independent signals.
8. **No doc duplication:** EVIDENCE_MODEL + types are the single system.
9. **Undo path:** every conversion runs in one Word.run (single Ctrl+Z unit);
   ConversionResult.insertions logs each insertion + source removed (revert UI
   deferred, log written as specified).

### Conversion contract (identified footnote behavior)

An approved match converts the citation block into a genuine Word footnote and
removes it from the body, exactly: (a) the marker at its position becomes the
footnote reference (marker text deleted, surroundings preserved); (b) the full
block span — starter plus continuations — is re-located by content and deleted
in descending order; (c) any block that cannot be re-located aborts its own
deletion with an explicit error (source left in body, never silently kept or
duplicated); (d) the footnote-count gate runs before any deletion; (e) reruns
convert nothing new (idempotent).
