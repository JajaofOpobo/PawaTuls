/**
 * text.ts — Phase 13: text transform operations.
 * Integration notes §4: KuTools has no case-conversion feature, so there is no
 * decompilation precedent. Office.js exposes no Range.Case equivalent, so all
 * transforms below are JavaScript string processing written back with
 * insertText(..., Word.InsertLocation.replace).
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
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.insertText(
      applyCaseTransform(selection.text, transform),
      Word.InsertLocation.replace
    );
    await context.sync();
  });
}

export async function trimSelectionSpaces(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.insertText(selection.text.trim(), Word.InsertLocation.replace);
    await context.sync();
  });
}

export async function collapseMultipleSpaces(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.insertText(selection.text.replace(/ {2,}/g, " "), Word.InsertLocation.replace);
    await context.sync();
  });
}

export async function trimParagraphLeadingTabs(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    for (const para of paragraphs.items) {
      const stripped = para.text.replace(/^\t+/, "");
      if (stripped !== para.text) {
        para.insertText(stripped, Word.InsertLocation.replace);
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
