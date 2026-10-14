/**
 * text.ts — Phase 13: text transform operations.
 * Integration notes §4: KuTools has no case-conversion feature, so there is no
 * decompilation precedent. Office.js exposes no Range.Case equivalent, so all
 * transforms below are JavaScript string processing.
 *
 * Stress-test lesson: Range.insertText(Replace) regenerates runs and WIPES
 * direct character formatting on the rewritten span (same data-loss class as
 * KuTools' .Text assignment). Every function here instead rewrites <w:t>
 * contents inside the paragraph OOXML, preserving all <w:rPr> formatting.
 */
import { CaseTransform } from "../shared/types";
import { getSelectionOrThrow, getSelectionParagraphsOrThrow, safeWordRun } from "../shared/wordUtils";

function applyCaseTransform(text: string, transform: CaseTransform): string {
  switch (transform) {
    case "upper":
      return text.toUpperCase();
    case "lower":
      return text.toLowerCase();
    case "title":
      return text.replace(/\b\w/g, (c) => c.toUpperCase()).replace(/\B\w/g, (c) => c.toLowerCase());
    case "sentence": {
      const lowered = text.toLowerCase();
      return lowered
        .replace(/(^\s*[a-z])/, (c) => c.toUpperCase())
        .replace(/([.!?]\s+[a-z])/g, (c) => c.toUpperCase());
    }
    case "toggle":
      return text
        .split("")
        .map((c) => {
          if (c >= "a" && c <= "z") return c.toUpperCase();
          if (c >= "A" && c <= "Z") return c.toLowerCase();
          return c;
        })
        .join("");
  }
}

export async function transformCase(transform: CaseTransform): Promise<{ error?: string }> {
  return rewriteSelectionText((text) => {
    const next = applyCaseTransform(text, transform);
    return next === text ? null : next;
  });
}

export async function trimSelectionSpaces(): Promise<{ error?: string }> {
  return rewriteSelectionText((text) => {
    const next = text.trim();
    return next === text ? null : next;
  });
}

export async function collapseMultipleSpaces(): Promise<{ error?: string }> {
  return rewriteSelectionText((text) => {
    const next = text.replace(/ {2,}/g, " ");
    return next === text ? null : next;
  });
}

export async function trimParagraphLeadingTabs(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    const ooxmls = paragraphs.items.map((p) => p.getOoxml());
    await context.sync();
    for (let i = 0; i < paragraphs.items.length; i++) {
      const stripped = paragraphs.items[i].text.replace(/^\t+/, "");
      if (stripped === paragraphs.items[i].text) {
        continue;
      }
      const patched = rewriteTextPreservingRuns(ooxmls[i].value, paragraphs.items[i].text, stripped);
      if (patched !== null) {
        paragraphs.items[i].insertOoxml(patched, Word.InsertLocation.replace);
      }
    }
    await context.sync();
  });
}

function decodeXmlEntities(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function encodeXmlEntities(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Rewrite a paragraph's text inside its OOXML, distributing the new text
 * across the existing <w:t> runs so every character keeps its run formatting.
 * Alignment is positional for equal-length transforms (case changes) and
 * deletion-aware for shortening transforms (trim/collapse): kept characters
 * map back to their original runs exactly. Returns null when the OOXML text
 * does not round-trip (caller falls back to plain insertText).
 */
export function rewriteTextPreservingRuns(
  xml: string,
  oldText: string,
  newText: string
): string | null {
  interface Seg {
    start: number;
    end: number;
    attrs: string;
    text: string;
  }
  const segs: Seg[] = [];
  const re = /<w:t([^>]*)>([^<]*)<\/w:t>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    segs.push({ start: m.index, end: m.index + m[0].length, attrs: m[1], text: decodeXmlEntities(m[2]) });
  }
  const flat = segs.map((s) => s.text).join("");
  const source = oldText.endsWith("\r") ? oldText.substring(0, oldText.length - 1) : oldText;
  if (flat !== source) {
    return null;
  }
  // Map each new-text char to the run that owned its source char.
  const runOfChar: number[] = [];
  segs.forEach((seg, runIdx) => {
    for (let k = 0; k < seg.text.length; k++) {
      runOfChar.push(runIdx);
    }
  });
  const owners: number[] = [];
  if (newText.length === flat.length) {
    // Positional (case changes): char i keeps run i's formatting exactly.
    for (let i = 0; i < newText.length; i++) {
      owners.push(runOfChar[i]);
    }
  } else {
    // Deletion-aware (trim/collapse): kept characters map back exactly;
    // insertions (e.g. ß → SS) inherit the current run.
    let oi = 0;
    let ni = 0;
    while (ni < newText.length && oi < flat.length) {
      if (flat[oi] === newText[ni]) {
        owners.push(runOfChar[oi]);
        oi++;
        ni++;
      } else if (newText.length > flat.length) {
        owners.push(runOfChar[Math.min(oi, runOfChar.length - 1)]);
        ni++;
      } else {
        oi++;
      }
    }
    while (ni < newText.length) {
      owners.push(runOfChar.length > 0 ? runOfChar[runOfChar.length - 1] : 0);
      ni++;
    }
  }
  if (owners.length !== newText.length) {
    return null;
  }
  const perRun: string[] = segs.map(() => "");
  newText.split("").forEach((ch, idx) => {
    perRun[owners[idx]] += ch;
  });
  let out = "";
  let cursor = 0;
  segs.forEach((seg, idx) => {
    out += xml.substring(cursor, seg.start);
    out += `<w:t${seg.attrs}>${encodeXmlEntities(perRun[idx])}</w:t>`;
    cursor = seg.end;
  });
  out += xml.substring(cursor);
  return out;
}

async function rewriteSelectionText(
  transform: (text: string) => string | null
): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    const paragraphs = selection.paragraphs;
    paragraphs.load("text");
    const ooxmls = paragraphs.items.map((p) => p.getOoxml());
    await context.sync();
    for (let i = 0; i < paragraphs.items.length; i++) {
      const next = transform(paragraphs.items[i].text);
      if (next === null) {
        continue;
      }
      const patched = rewriteTextPreservingRuns(ooxmls[i].value, paragraphs.items[i].text, next);
      if (patched !== null) {
        paragraphs.items[i].insertOoxml(patched, Word.InsertLocation.replace);
      } else {
        // Round-trip mismatch (e.g. field codes inside runs): legacy path.
        paragraphs.items[i].insertText(next, Word.InsertLocation.replace);
      }
    }
    await context.sync();
  });
}

/**
 * Pure OOXML helper: superscript ordinal suffixes (1st, 2nd, 3rd, 4th … 200th)
 * inside a paragraph's OOXML. Only matches fully inside a single <w:t>;
 * cross-run matches are left for the caller to count as skipped. Runs that are
 * already superscript are left untouched. Returns the patched XML and the
 * number of suffixes rewritten.
 */
export function superscriptOrdinalsInOoxml(xml: string): { xml: string; changed: number } {
  let changed = 0;
  const out = xml.replace(
    /<w:t([^>]*)>([^<]*)<\/w:t>/g,
    (full, attrs: string, text: string, offset: number, whole: string) => {
      if (!/\d(st|nd|rd|th)\b/.test(text)) {
        return full;
      }
      const runStart = Math.max(
        whole.lastIndexOf("<w:r ", offset),
        whole.lastIndexOf("<w:r>", offset)
      );
      if (
        runStart >= 0 &&
        /w:vertAlign\s+w:val="superscript"/.test(whole.substring(runStart, offset))
      ) {
        return full;
      }
      // Preserve the original run properties for the continuing text.
      const rprMatch = whole
        .substring(runStart >= 0 ? runStart : 0, offset)
        .match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
      const rpr = rprMatch ? rprMatch[0] : "";
      const replaced = text.replace(/\b(\d+)(st|nd|rd|th)\b/g, (_, digits: string, suffix: string) => {
        changed++;
        return (
          `${digits}</w:t></w:r>` +
          `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>` +
          `<w:t${attrs}>${suffix}</w:t></w:r>` +
          `<w:r>${rpr}<w:t${attrs}>`
        );
      });
      return replaced === text ? full : `<w:t${attrs}>${replaced}</w:t>`;
    }
  );
  return { xml: out, changed };
}

export async function superscriptOrdinals(
  applyTo: "selection" | "document" = "selection"
): Promise<{ data?: { changed: number; skipped: number }; error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs =
      applyTo === "selection"
        ? await getSelectionParagraphsOrThrow(context)
        : context.document.body.paragraphs;
    paragraphs.load("text");
    const ooxmls = paragraphs.items.map((p) => p.getOoxml());
    await context.sync();
    let changed = 0;
    let skipped = 0;
    for (let i = 0; i < paragraphs.items.length; i++) {
      const candidates = paragraphs.items[i].text.match(/\b\d+(?:st|nd|rd|th)\b/g) ?? [];
      if (candidates.length === 0) {
        continue;
      }
      const patched = superscriptOrdinalsInOoxml(ooxmls[i].value);
      if (patched.changed > 0) {
        paragraphs.items[i].insertOoxml(patched.xml, Word.InsertLocation.replace);
      }
      changed += patched.changed;
      skipped += candidates.length - patched.changed;
    }
    await context.sync();
    return { changed, skipped };
  });
}
