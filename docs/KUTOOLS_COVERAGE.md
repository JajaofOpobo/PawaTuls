# KuTools → PawaTuls Coverage Matrix

Source: ribbon button inventory (`--.--.resx`, 156 ids) + forms + string survey.
Rule: every KuTools document feature is covered by a PawaTuls phase, matched,
or listed here with an explicit reason. Nothing is silently skipped.

Status: **covered** (built) · **phased** (Phase # assigned) · **parity-note**
(behavior matched deliberately) · **backlog** (feasible, deferred via scope
decision) · **unsupported** (no Office.js API).

## Formatting engine

| KuTools feature | Status |
|---|---|
| Bold / Italic / Underline / Strike / Super / Sub / AllCaps / SmallCaps / color / highlight / size / name / spacing / hidden, clear char formatting | covered (Phase 10) |
| Case conversion | covered — KuTools has none; we exceed (Phase 13) |
| Alignment, indents, spacing, line spacing, keep-with-next/together, borders, remove spacing/indents/empties | covered (Phase 11, 15) |
| Apply / create / list styles | covered (Phase 12) |
| Collapse spaces, trim, leading tabs | covered (Phase 13) |
| Tab → space | phased (Phase 35) |
| Remove empty manual line breaks | phased (Phase 35) |
| LF ⇄ paragraph marks | covered (Phase 14) |
| List → text | covered (Phase 14) |
| Table → text | phased (Phase 40) |
| Field codes → text | covered (Phase 14, via `unlink`) |
| Remove page / column / section / all breaks | phased (Phase 35) |
| Remove hidden text | phased (Phase 35) |
| Remove paragraph borders | covered (Phase 11 `none` + Phase 15 border branch) |
| Remove drop cap, frames | phased (Phase 35) |
| Remove pictures (inline + floating) | phased (Phase 35) |
| Remove tables, empty rows/cols | phased (Phase 35) |
| Remove content (HTML) controls, keep text | phased (Phase 35) |
| Remove text boxes | phased (Phase 35) |
| Remove comments | phased (Phase 35) |
| Remove hyperlinks, keep text | phased (Phase 35) |
| Remove bookmarks | phased (Phase 35) |
| Remove footnotes / endnotes | phased (Phase 37) |
| Footnote → endnote, endnote → text | endnote → text phased (Phase 37); footnote → endnote unsupported (no endnote-insert API) |
| Adjust paragraph end (font-shrink heuristic) | parity-note: heuristic rejected as buggy (AUDIT.md #10); safe replacement `controlWidows` phased (Phase 40) |
| Clear formatting (individual resets) | covered (Phase 15) |
| Formatting profiles (versioned, import/export, preview) | phased (Phases 31–32, v1 merge) |
| Numbering audit + preview-first repair | phased (Phases 33–34, v1 merge) |
| Resize images 25–200% / to selection | phased (Phase 40) |
| Repeat header rows across pages | phased (Phase 40) |
| Table diagonal header | phased (Phase 40) |
| Insert table / figure / equation captions, refresh, convert-to-caption | phased (Phase 38) |
| Select tables / empty paras / headings / short paras / pictures (select-first + count) | phased (Phase 39) |
| Select pages (current/odd/even/range) | unsupported: Office.js exposes no page model |
| Remove macros (VBA) | unsupported: no Office API touches the VBA project |
| Equation → image / image → equation | unsupported: no OMath API in Office.js |

## Legal workflows (PawaTuls originals — KuTools has no equivalents)

Pleadings, affidavit/jurat, exhibits, reliefs, cause headings, citations,
schedules — covered by plan Phases 17–24. Footnote *detection* (inline pasted
markers → genuine notes) is novel — Phases 25–28.

## Explicitly out of scope (not document features, or app-level)

Window management, file/folder tools, doc merge/split, batch conversion,
export-images-to-file (backlog candidate — feasible via download), barcode /
scanner ActiveX, AutoText/dictionary/autocorrect panes and sync (no API),
bookmark/hyperlink/caption *manager panes* (core ops covered; full panes are
backlog), snapshot/redo-last-command app chrome, register/help/about/options.
Endnote *recovery* stays backlog per v1 (endnote *removal/conversion* is Phase 37).
