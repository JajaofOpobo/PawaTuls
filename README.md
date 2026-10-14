# PawaTuls

Microsoft Word task pane add-in: formatting engine, Nigerian legal workflows, footnote detection.

Stack: Office.js (WordApi 1.5 minimum), React 18, TypeScript 5, Fluent UI v9.
Platform: Windows, Word for Microsoft 365 desktop only.

## Repo layout

This repo root **is** the project root (`chambers-word-toolkit/` in the plan).
Plan doc: `docs/chambers-word-addin-plan.md`. Model guide: `docs/AGENTS.md`.

```
manifest.xml            ← Phase 5 (WordApi 1.5)
package.json / webpack.config.js / tsconfig.json  ← Phase 4 (hand-scaffolded, no Yeoman)
src/shared/types.ts     ← Phase 7
src/shared/wordUtils.ts ← Phase 8
src/taskpane/           ← Phase 9 shell (App + ErrorBanner + stubs)
src/formatting/ src/legal/ src/footnotes/  ← Phases 10–28 (pending)
decompilation/          ← Phases 1–3, 6 outputs (pending)
assets/                 ← add-in icons
```

## Lead-engineer decisions

- **Hand-scaffolded, not Yeoman.** `yo office` (generator-office 3.0.4) install was
  incomplete/broken on Node 24 and Yeoman 7 is fragile there. The scaffold
  reproduces the Yeoman React+TS output (webpack 5 + `taskpane.html`/`commands.html`,
  dev-server on `https://localhost:3000`, `office-addin-debugging` scripts) without
  Yeoman cruft. Phase 4 completion tests still apply (`npm start`, `tsc --noEmit`).
- **Node 24, not 18.** Plan asks for Node 18 LTS; environment has Node 24.15 + npm 11.
  Webpack 5 / TS 5 / React 18 stack is compatible. If `office-addin-dev-certs` misbehaves
  on 24, fall back to Node 18 LTS via nvm.
- **No nested `chambers-word-toolkit/` folder.** Plan paths referencing
  `chambers-word-toolkit/...` map to repo root.

## Dev loop

```powershell
npm install
npm start        # opens Word with task pane (https://localhost:3000/taskpane.html)
npx tsc --noEmit # must pass with zero errors
npm run validate # manifest check
npm run build    # production build → dist/
```

## Status

- [x] Phase 4 — scaffold (hand; `npm run build` compiles, `tsc --noEmit` clean)
- [x] Phase 5 — manifest (WordApi 1.5, `npm run validate` → valid)
- [x] Phase 7 — shared types
- [x] Phase 8 — wordUtils
- [x] Phase 9 — task pane shell + stubs
- [ ] Live Word check — run `npm start`, confirm 3 tabs (manual, Phase 4/9 completion test)
- [x] Phases 1–3, 6 — KuTools decompilation → `decompilation/api-map.md` + `integration-notes.md` (unblocks Phases 10/11/13/15)
- [ ] Phases 10–30 — feature modules, panels, testing, packaging (next: Phase 10 font module)
- Suspended for this version: Phases 17–24 (legal modules + LegalPanel). The Legal
  tab is hidden in `App.tsx`; `LegalPanel.tsx` stub retained. Re-enable by
  restoring the tab and building Phases 17–24 per plan + Appendix H.
- [x] Phases 25–28, 37 — footnote scanner/matcher/converter/panel + note management
- [x] Phases 41–43 (Appendix J) — manual block-area override, citation systems, ordinal superscripting
- [x] Phase 44 + evidence model — ribbon tab with quick actions; tiered evidence scoring with contradictions
- [ ] Phases 10–30 — feature modules, panels, testing, packaging

Note: `tsconfig.json` intentionally has no `noEmit` (webpack `ts-loader` needs emit;
`npx tsc --noEmit` still works via CLI flag).
