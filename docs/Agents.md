# AGENTS.MD — MODEL SELECTION GUIDE

**For:** PawaTuls build sessions in OpenCode  
**Purpose:** Helps you choose which model to load with `/model` before each session.  
This file does not prescribe. It describes what each model does well for this project
so you can make the call.

---

## HOW TO SWITCH MODELS IN OPENCODE

At any point in an OpenCode session:
```
/model
```
This opens the model picker. Switch before starting a new phase or when the current
model is producing poor output.

---

## THE MODELS AVAILABLE

### Muse (Meta)
- Lead Engineer Model

### Claude Sonnet (Anthropic)

**Strengths for this project:**
- TypeScript accuracy. It knows the Office.js `Word.*` API hierarchy well and produces
  correct type usage on the first pass more often than the alternatives.
- Complex logic. The footnote scanner (Phase 25) and converter (Phase 27) involve
  multi-pass document searches, range re-scanning, and descending-order processing.
  Claude handles this kind of interconnected logic without losing track of constraints.
- Constraint adherence. When you say "do not use stored range references, re-scan at
  conversion time," Claude follows that constraint even across many lines of code.
  Other models drift on constraints as the function grows long.
- React component state. The FootnotePanel (Phase 28) has three UI stages and five
  state variables. Claude keeps the stage transitions clean.

**Weaknesses:**
- It can over-engineer simple tasks. For a function that just builds an HTML string
  from template variables (the cause heading in Phase 21), Claude may add unnecessary
  abstraction. If you see that happening, tell it to simplify.
- Slower than Gemini Flash for straightforward work.

**Best phases:** 10, 11, 12, 13, 14, 15, 16, 25, 26, 27, 28.  
Also any phase where you are debugging a TypeScript error or a logic fault.

---

### Gemini 3 Pro (Google)

**Strengths for this project:**
- Structured and templated code. The legal workflow modules (Phases 17–24) are
  pattern-based: insert a block of HTML at a cursor position, prefix paragraphs with
  formatted strings, track a counter. Gemini handles this kind of structured,
  repetitive work fast and cleanly.
- Fast iteration. If you want to get through multiple simple phases in one session,
  Gemini is the faster choice.
- HTML template generation. The cause heading block and jurat HTML strings are
  straightforward templating. Gemini writes these accurately.

**Weaknesses:**
- Complex multi-step Office.js logic. The footnote scanner's three detection passes
  and the converter's re-scan approach require sustained constraint tracking. Gemini
  sometimes drops constraints in complex functions.
- TypeScript edge cases. On mixed-selection null handling (`font.bold === null`),
  Gemini occasionally uses the COM interop convention (`=== 9999999`) instead of the
  Office.js convention (`=== null`). Check this specifically.

**Best phases:** 17, 18, 19, 20, 21, 22, 23, 24. Also Phases 4, 5 (scaffold and
manifest) where the task is mechanical.

---

### GPT-4o (OpenAI)

**Strengths for this project:**
- Code review. If Claude or Gemini produced something you are not sure about, paste
  the function into a GPT-4o session and ask it to find bugs or edge cases. It gives
  a useful second perspective.
- Type definition review. Phase 7 (shared types) benefits from a GPT-4o pass to
  check whether the interfaces are internally consistent.
- Debugging. When `npx tsc --noEmit` returns a type error that Claude cannot fix,
  GPT-4o often approaches the resolution differently and finds it.

**Weaknesses:**
- Less reliable on Word API specifics than Claude. It may suggest COM interop
  approaches (`Word.Interop`) instead of Office.js approaches. If it does, correct
  it immediately.
- Tends to suggest installing additional npm packages for problems that do not need them.

**Best phases:** Code review passes. End-to-end debugging (Phase 29).
Not recommended as the primary model for any build phase.

---

### Gemini Flash (Google)

**Strengths for this project:**
- Very fast and cheap for trivial tasks: renaming a function, adjusting an HTML
  template string, fixing a single TypeScript import.
- Good for generating the test documents described in Phase 29 (text content for
  formatting-test.docx, pleadings-test.docx).

**Weaknesses:**
- Not reliable for anything architecturally significant. Do not use it for the
  footnote modules or the formatting engine.

**Best phases:** Microtasks within any phase. Generating test document content.

---

### Local models via Ollama (e.g. Qwen2.5-Coder, DeepSeek-Coder)

**Strengths for this project:**
- Completely offline. If you are working in court or somewhere without reliable
  internet, a local model lets you keep moving.
- No API cost.

**Weaknesses:**
- On the hardware most likely available (8–16 GB RAM), local models produce noticeably
  weaker TypeScript than Claude or GPT-4o. They lose track of type constraints and
  produce incorrect Office.js API calls more often.
- Not recommended for the footnote scanner or converter phases.

**Best phases:** Simple utility functions. HTML template generation. Any phase where
you can closely review every line of output before committing it.

**Recommended models if using locally:**
- `qwen2.5-coder:14b` — best balance of quality and speed on 16 GB RAM
- `deepseek-coder-v2:16b` — strong on TypeScript specifically

---

### Google Antigravity

Antigravity works differently from OpenCode. It is a full IDE with a Planning Mode
and Artifacts feature. Use it for this project in two specific situations:

**Situation 1: Planning a new phase before starting it.**  
Before you open OpenCode for a complex phase (especially Phase 25 — the scanner),
open Antigravity, describe the feature in detail, and run Planning Mode. Antigravity
produces a task breakdown Artifact. Review it, correct anything that contradicts the
plan, then carry the corrected task breakdown into your OpenCode session.

**Situation 2: Architecture review after Phase 6.**  
After the decompilation findings are documented, open Antigravity and ask it to
review `decompilation/integration-notes.md` and identify whether any findings require
changes to the Phase 7 types or Phase 8 utilities. Its structured review output is
easier to work with than a freeform chat response.

**Do not use Antigravity for:**
- Actual code generation for Office.js modules. Antigravity's browser integration is
  not useful for a Word add-in, and it cannot test against the Word desktop app.
- Phases that are straightforward builds (manifests, scaffold, HTML templates).

---

## QUICK REFERENCE

| Task type | Recommended model |
|---|---|
| Complex TypeScript + Office.js API | Claude Sonnet |
| Structured / templated legal modules | Gemini 3 Pro |
| Code review and debugging | GPT-4o |
| Fast trivial edits | Gemini Flash |
| Offline work | Qwen2.5-Coder (local) |
| Pre-phase planning and architecture review | Google Antigravity |

---

## PROMPTING TIPS FOR THIS PROJECT

**Always start with the Master Context Block.** Every new OpenCode session starts
from zero. The Master Context Block (Section 0 of the plan) is what tells the agent
what project it is on, what the rules are, and what has already been built.

**Paste the phase Resume Prompt immediately after the context.** Do not describe
the phase in your own words. The Resume Prompt is written precisely. Use it exactly.

**Tell the agent which file to read first when the phase requires it.** For Phases
10, 11, 13, and 15, the Resume Prompt says "read decompilation/integration-notes.md."
The agent will not do this unless you include that instruction. The Resume Prompts
in the plan already contain it — use them verbatim.

**Do not let an agent modify files outside its phase scope.** If Claude starts
editing `types.ts` while working on Phase 10, stop it. Types must only change in
Phase 7 sessions unless a type is genuinely missing — in which case start a new
Phase 7 session to add it, then return to Phase 10.

**When a model drifts, call it out directly.** "You said you would not use stored
range references. You just did. Fix this." Direct correction on a specific line is
faster than re-prompting from scratch.

**Switch models when:**
- The model has failed twice on the same logical problem
- The output is getting longer but not more correct
- The model is adding abstraction you did not ask for
- You are about to start a different phase

**Save the decompilation outputs before they are needed.** Phases 10–15 all reference
`decompilation/integration-notes.md`. If that file is incomplete when those phases
run, the agent has no decompilation context and will produce generic implementations
that may miss KuTools-specific edge cases.
