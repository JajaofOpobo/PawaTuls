/* tests/smoke.js — real automated tests for pure modules (ChatGPT review item).
 * Transpiles the dependency-safe subset of src/ (no Word runtime executes)
 * and runs node:assert suites. Run: npm test
 */
"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");
const ts = require("typescript");

const ROOT = path.join(__dirname, "..");
const SRC = path.join(ROOT, "src");
const FILES = [
  "shared/types.ts",
  "shared/wordUtils.ts",
  "footnotes/citations.ts",
  "footnotes/matcher.ts",
  "footnotes/converter.ts",
  "formatting/font.ts",
  "formatting/paragraph.ts",
  "formatting/numbering.ts",
  "formatting/profiles.ts",
  "formatting/text.ts",
];

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "pawatuls-tests-"));
const cache = new Map();

function load(rel) {
  if (cache.has(rel)) {
    return cache.get(rel);
  }
  const file = path.join(SRC, rel);
  const source = fs.readFileSync(file, "utf8");
  const { outputText } = ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    fileName: rel,
  });
  const module = { exports: {} };
  const localRequire = (request) => {
    if (request.startsWith(".")) {
      const resolved = path.normalize(path.join(path.dirname(rel), request));
      const withExt = resolved.endsWith(".ts") ? resolved : resolved + ".ts";
      return load(withExt);
    }
    return require(request);
  };
  new Function("module", "exports", "require", outputText)(module, module.exports, localRequire);
  cache.set(rel, module.exports);
  return module.exports;
}

let passed = 0;
function check(name, fn) {
  fn();
  passed++;
  console.log("ok -", name);
}

// ─── citations ─────────────────────────────────────────────────────────────
{
  const C = load("footnotes/citations.ts");
  const has = (text, system) => C.classifyParagraph(text).some((x) => x.system === system);
  check("citations: reporters", () => {
    assert.ok(has("[2020] 15 NWLR (Pt. 1721) 123", "Nigerian Law Reports"));
    assert.ok(has("[2021] LPELR-55302(CA)", "Nigerian Law Reports"));
  });
  check("citations: statutes/neutral/authordate", () => {
    assert.ok(has("Section 175(1) of the Companies and Allied Matters Act 2020", "Statutes"));
    assert.ok(has("[2019] UKSC 42", "Neutral Citations"));
    assert.ok(has("(Adeyemi 2021)", "Author-Date"));
  });
  check("citations: oscola/bluebook/chicago/mla/ibid", () => {
    assert.ok(has("Smith, \u2018Contract Law\u2019 (2021)", "OSCOLA"));
    assert.ok(has("410 U.S. 113 holds.", "Bluebook"));
    assert.ok(has("Adeyemi, Bola. Contract in Nigeria. Lagos: Printwell, 2021.", "Chicago"));
    assert.ok(has('"Things Fall Apart," Research, vol. 12', "MLA"));
    assert.ok(has("Ibid., at 45.", "Ibid Chain"));
  });
  check("citations: plain negative + density", () => {
    assert.strictEqual(C.classifyParagraph("Plain body text, nothing special.").length, 0);
    const sum = C.detectCitationSystems(["Ibid.", "[2020] 15 NWLR (Pt. 1721) 123"]);
    assert.deepStrictEqual(sum.denseParagraphs, [0, 1]);
    assert.deepStrictEqual(sum.listDense, [1]);
  });
  check("citations: signals + shorthand", () => {
    const s = C.detectReferenceSignals("See Smith (2021) and Jones et al. Ibid.");
    assert.ok(s.hasYear && s.hasEtAl && s.hasIbid && s.score >= 3);
    assert.strictEqual(C.detectReferenceSignals("Plain prose.").score, 0);
    assert.deepStrictEqual(
      C.detectAcademicShorthand("See vol. 12, pp. 30-45, ed. Smith, https://x.example"),
      ["vol.", "pp.", "ed.", "URL"]
    );
  });
}

// ─── matcher ───────────────────────────────────────────────────────────────
{
  const { buildMatchTable, EVIDENCE_MODEL } = load("footnotes/matcher.ts");
  assert.strictEqual(EVIDENCE_MODEL.version, "1.0");
  const mk = (n, pi) => ({ number: n, pattern: "bracketed", paragraphIndex: pi, contextBefore: "", contextAfter: "" });
  const blk = (n, pi, extra) => Object.assign({ number: n, text: "Note", paragraphIndex: pi }, extra);
  check("matcher: high default + repeat/conflict unchecked", () => {
    let m = buildMatchTable([mk(1, 1), mk(1, 2)], [blk(1, 10)], false, -1);
    assert.strictEqual(m[0].accepted, true);
    assert.strictEqual(m[1].accepted, false);
    assert.strictEqual(m[1].confidence, "medium");
    m = buildMatchTable([mk(2, 1)], [blk(2, false), blk(2, 11)], false, -1);
  });
  check("matcher: conflict pair surfaced", () => {
    const m = buildMatchTable([mk(2, 1)], [blk(2, 10), { number: 2, text: "Other", paragraphIndex: 11 }], false, -1);
    assert.strictEqual(m.length, 2);
    assert.strictEqual(m[0].accepted, false);
    assert.strictEqual(m[0].confidence, "medium");
    assert.strictEqual(m[1].accepted, false);
  });
  check("matcher: dense cap + low unchecked", () => {
    let m = buildMatchTable([mk(1, 1)], [blk(1, 10, { citationDense: true })], true, 5);
    assert.strictEqual(m[0].confidence, "medium");
    m = buildMatchTable([mk(1, 1)], [blk(1, 2)], true, 5);
    assert.strictEqual(m[0].confidence, "low");
    assert.strictEqual(m[0].accepted, false);
  });
  check("matcher: bibliography contradiction + boundary rule", () => {
    let m = buildMatchTable([mk(1, 1)], [blk(1, 10, { contradictions: ['bibliography: under "References" heading'] })], false, -1);
    assert.strictEqual(m[0].confidence, "medium");
    assert.ok(m[0].contradictions.some((c) => c.indexOf("bibliography") >= 0));
    m = buildMatchTable([mk(1, 1), mk(2, 2)], [blk(1, 10, { followedByLarge: true }), blk(2, 12, {})], false, -1);
    assert.strictEqual(m[0].confidence, "medium");
    m = buildMatchTable([mk(1, 1)], [blk(1, 10, { followedByLarge: true })], false, -1);
    assert.strictEqual(m[0].confidence, "high");
  });
  check("matcher: signals boost + evidence recorded", () => {
    const m = buildMatchTable([mk(3, 1)], [blk(3, 2, { referenceSignals: true })], true, 5);
    assert.strictEqual(m[0].confidence, "medium");
    assert.ok(m[0].pairingEvidence.length > 0);
  });
}

// ─── converter pure helpers ────────────────────────────────────────────────
{
  const { findBlockParagraphIndices, resolveMarkerParagraph } = load("footnotes/converter.ts");
  check("converter: block spans delete fully", () => {
    const texts = ["Body", "1. First note", "continued note", "Body middle", "2. Second"];
    const idx = findBlockParagraphIndices(texts, [
      { number: 1, text: "First note continued note", paragraphIndex: 1, span: 2 },
      { number: 2, text: "Second", paragraphIndex: 4 },
    ]);
    assert.deepStrictEqual(idx, [4, 2, 1]);
  });
  check("converter: resolve by hint + context", () => {
    const texts = ["Ordinary 2024 figures here.", "Treaty history [1] shows agreement here."];
    const match = { marker: { number: 1, pattern: "bracketed", paragraphIndex: 5, contextBefore: "Treaty history ", contextAfter: " shows" }, block: { number: 1, text: "Note", paragraphIndex: 9 } };
    assert.strictEqual(resolveMarkerParagraph(texts, match), 1);
    const nomatch = { marker: { number: 9, pattern: "bracketed", paragraphIndex: 0, contextBefore: "zzz", contextAfter: "zzz" }, block: { number: 9, text: "N", paragraphIndex: 9 } };
    assert.strictEqual(resolveMarkerParagraph(texts, nomatch), -1);
  });
}

// ─── numbering ─────────────────────────────────────────────────────────────
{
  const N = load("formatting/numbering.ts");
  check("numbering: converters", () => {
    assert.strictEqual(N.toAlpha(1), "a");
    assert.strictEqual(N.toAlpha(27), "aa");
    assert.strictEqual(N.toRoman(4), "iv");
    assert.strictEqual(N.toRoman(2026), "mmxxvi");
  });
  check("numbering: gap/dup repair + rerun stability", () => {
    const items = [1, 2, 2, 4].map((n, i) => ({ paragraphIndex: i, kind: "manual", prefix: n + ".", level: 1 }));
    const plan = N.previewRepair(
      [{ paragraphIndex: 2, issue: "duplicate", detail: "" }, { paragraphIndex: 3, issue: "gap", detail: "" }],
      items
    );
    assert.deepStrictEqual(plan.edits[0], { paragraphIndex: 2, before: "2.", after: "3." });
    assert.deepStrictEqual(plan.edits[1], { paragraphIndex: 3, before: "4.", after: "4." });
  });
  check("numbering: (c) is alpha, restart untouched", () => {
    const items = [
      { paragraphIndex: 0, kind: "manual", prefix: "(a)", level: 2 },
      { paragraphIndex: 1, kind: "manual", prefix: "(c)", level: 2 },
      { paragraphIndex: 2, kind: "manual", prefix: "1.", level: 1 },
    ];
    const plan = N.previewRepair(
      [{ paragraphIndex: 1, issue: "gap", detail: "" }, { paragraphIndex: 2, issue: "restart", detail: "" }],
      items
    );
    assert.strictEqual(plan.edits.length, 1);
    assert.deepStrictEqual(plan.edits[0], { paragraphIndex: 1, before: "(c)", after: "(b)" });
  });
}

// ─── profiles ──────────────────────────────────────────────────────────────
{
  const P = load("formatting/profiles.ts");
  check("profiles: validation matrix", () => {
    const good = { schemaVersion: 1, id: "x", name: "X", font: { bold: true, size: 12 } };
    assert.ok(P.validateProfile(good).data);
    assert.ok(P.validateProfile({ schemaVersion: 1, id: "x", name: "X" }).data);
    assert.ok(P.validateProfile({ schemaVersion: 2, id: "x", name: "X" }).error.indexOf("schemaVersion") >= 0);
    assert.ok(P.validateProfile({ schemaVersion: 1, id: "x", name: "X", bogus: 1 }).error.indexOf("bogus") >= 0);
    assert.ok(P.validateProfile({ schemaVersion: 1, id: "x", name: "X", font: { underline: "squiggle" } }).error);
    assert.ok(P.validateProfile("nope").error);
    assert.ok(P.importProfile("{bad").error.indexOf("JSON") >= 0);
  });
  check("profiles: round-trip + builtins", () => {
    const good = { schemaVersion: 1, id: "x", name: "X" };
    assert.strictEqual(P.importProfile(P.exportProfile(good)).data.id, "x");
    assert.throws(() => P.exportProfile({ schemaVersion: 1 }));
    for (const b of P.BUILTIN_PROFILES) {
      assert.ok(P.validateProfile(b).data, "builtin " + b.id);
    }
  });
}

// ─── text helpers ──────────────────────────────────────────────────────────
{
  const T = load("formatting/text.ts");
  check("text: ordinal splits preserve runs", () => {
    const xml = '<w:p><w:pPr/><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">The 1st and 22nd</w:t></w:r></w:p>';
    const r = T.superscriptOrdinalsInOoxml(xml);
    assert.strictEqual(r.changed, 2);
    assert.ok(r.xml.indexOf("<w:b/>") >= 0, "bold kept");
    assert.ok(r.xml.indexOf(" and 22</w:t>") >= 0, "middle text normal");
    const opens = (r.xml.match(/<w:r[\s>]/g) || []).length;
    const closes = (r.xml.match(/<\/w:r>/g) || []).length;
    assert.strictEqual(opens, closes, "balanced runs");
  });
  check("text: rewrite preserves formatting", () => {
    const xml = '<w:p><w:pPr/><w:r><w:rPr><w:b/></w:rPr><w:t>Bold lead</w:t></w:r><w:r><w:t> and plain tail</w:t></w:r></w:p>';
    const out = T.rewriteTextPreservingRuns(xml, "Bold lead and plain tail", "BOLD LEAD AND PLAIN TAIL");
    assert.ok(out && out.indexOf("<w:b/>") >= 0, "bold kept");
    assert.ok(out.indexOf(">BOLD LEAD</w:t>") >= 0, "mapped into bold run");
    const xml2 = '<w:p><w:pPr/><w:r><w:rPr><w:i/></w:rPr><w:t>a   </w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>';
    const out2 = T.rewriteTextPreservingRuns(xml2, "a   b", "a b");
    assert.ok(out2 && out2.indexOf("a </w:t>") >= 0, "collapse exact: " + out2);
    assert.strictEqual(T.rewriteTextPreservingRuns(xml, "different", "DIFFERENT"), null, "mismatch null");
  });
}

console.log(`\nAll ${passed} suites passed.`);
