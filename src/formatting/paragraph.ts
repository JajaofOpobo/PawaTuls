/**
 * paragraph.ts — paragraph formatting engine (used by profiles; direct UI
 * removed as native-duplicate — Word's Paragraph dialog covers these).
 * Integration notes §2: deletions iterate in REVERSE order; KuTools' Find-based
 * empty-para removal has no Office.js equivalent (empty paragraphs are unsearchable),
 * so reverse iteration + delete() is the correct workaround.
 */
import { LineSpacingRule, ParagraphOptions } from "../shared/types";
import {
  getSelectionParagraphsOrThrow,
  safeWordRun,
} from "../shared/wordUtils";

export function mapAlignment(
  alignment: NonNullable<ParagraphOptions["alignment"]>
): Word.Alignment {
  switch (alignment) {
    case "left":
      return Word.Alignment.left;
    case "centered":
      return Word.Alignment.centered;
    case "right":
      return Word.Alignment.right;
    case "justified":
      return Word.Alignment.justified;
  }
}

export function mapLineSpacing(rule: LineSpacingRule, value?: number): number {
  switch (rule) {
    case "single":
      return 12;
    case "onePtFive":
      return 18;
    case "double":
      return 24;
    case "atLeast":
    case "exactly":
      if (value === undefined) {
        throw new Error(`Line spacing value (points) is required for "${rule}".`);
      }
      return value;
    case "multiple":
      if (value === undefined) {
        throw new Error('Line spacing multiplier is required for "multiple".');
      }
      return value * 12;
  }
}

export async function applyParagraphFormat(
  options: ParagraphOptions
): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    for (const para of paragraphs.items) {
      if (options.alignment !== undefined) para.alignment = mapAlignment(options.alignment);
      if (options.firstLineIndent !== undefined) para.firstLineIndent = options.firstLineIndent;
      if (options.leftIndent !== undefined) para.leftIndent = options.leftIndent;
      if (options.rightIndent !== undefined) para.rightIndent = options.rightIndent;
      if (options.lineSpacingRule !== undefined) {
        para.lineSpacing = mapLineSpacing(options.lineSpacingRule, options.lineSpacing);
      } else if (options.lineSpacing !== undefined) {
        para.lineSpacing = options.lineSpacing;
      }
      if (options.spaceBefore !== undefined) para.spaceBefore = options.spaceBefore;
      if (options.spaceAfter !== undefined) para.spaceAfter = options.spaceAfter;
      // keepWithNext / keepLinesTogether: Paragraph exposes no direct props and
      // paragraphFormat is unreachable from Paragraph — patch OOXML (see below).
      // pageBreakBefore: no Office.js equivalent — documented no-op (see header).
    }
    if (options.keepWithNext !== undefined || options.keepLinesTogether !== undefined) {
      for (const para of paragraphs.items) {
        await setKeepFlags(context, para, options.keepWithNext, options.keepLinesTogether);
      }
    }
    await context.sync();
  });
}

async function setKeepFlags(
  context: Word.RequestContext,
  para: Word.Paragraph,
  keepNext?: boolean,
  keepLines?: boolean
): Promise<void> {
  const ooxml = para.getOoxml();
  await context.sync();
  let updated = ooxml.value
    .replace(/<w:keepNext[^/]*\/>/g, "")
    .replace(/<w:keepLines[^/]*\/>/g, "");
  let insert = "";
  if (keepNext === true) insert += "<w:keepNext/>";
  if (keepLines === true) insert += "<w:keepLines/>";
  if (insert !== "") {
    if (!/<w:pPr[\s>]/.test(updated)) {
      throw new Error("Could not locate paragraph properties to set keep options.");
    }
    updated = updated.replace(/(<w:pPr[^>]*>)/, `$1${insert}`);
  }
  para.insertOoxml(updated, Word.InsertLocation.replace);
}

export async function removeEmptyParagraphs(): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    // Audit lesson (KuTools ^p-in-tables corruption): never delete paragraph marks
    // that terminate table cells — check table membership before deleting.
    // R4 (plan Appendix D): also skip empty paragraphs carrying bookmarks,
    // fields, or section breaks — their text is empty but they are structural.
    const empties = paragraphs.items.filter((para) => para.text.trim() === "");
    const tableChecks = empties.map((para) => {
      const t = para.parentTableOrNullObject;
      context.load(t);
      return t;
    });
    const ooxmlChecks = empties.map((para) => para.getOoxml());
    await context.sync();
    const PROTECTED = /<w:bookmarkStart[\s>]|<w:fldChar[\s>]|<w:instrText[\s>]|<w:sectPr[\s>]/;
    let deleted = 0;
    for (let i = empties.length - 1; i >= 0; i--) {
      if (tableChecks[i].isNullObject && !PROTECTED.test(ooxmlChecks[i].value)) {
        empties[i].delete();
        deleted++;
      }
    }
    await context.sync();
    return deleted;
  });
}
