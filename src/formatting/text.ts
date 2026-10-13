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
