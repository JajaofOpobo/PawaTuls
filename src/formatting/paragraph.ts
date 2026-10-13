/**
 * paragraph.ts — Phase 11: paragraph formatting operations.
 * Integration notes §2: deletions iterate in REVERSE order; KuTools' Find-based
 * empty-para removal has no Office.js equivalent (empty paragraphs are unsearchable),
 * so reverse iteration + delete() is the correct workaround.
 * Integration notes §5: no CharacterUnit properties exist in Office.js — point
 * values are the complete equivalent.
 *
 * API notes (verified against @types/office-js):
 * - Word.ParagraphFormat has NO LineSpacingRule enum and NO pageBreakBefore.
 *   lineSpacing is points; COM convention: 12 = single, so single→12,
 *   onePtFive→18, double→24, multiple→value*12, atLeast/exactly→value (points).
 *   pageBreakBefore is accepted in ParagraphOptions but is a documented no-op.
 * - Word.Paragraph has NO getBorder (table-only API). Borders are applied by
 *   injecting <w:pBdr> into the paragraph's OOXML. Width points → sz = pts*8.
 */
import { BorderOptions, LineSpacingRule, ParagraphOptions } from "../shared/types";
import {
  cmToPt,
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

export async function setLineSpacing(
  rule: LineSpacingRule,
  value?: number
): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    const points = mapLineSpacing(rule, value);
    for (const para of paragraphs.items) {
      para.lineSpacing = points;
    }
    await context.sync();
  });
}

export async function removeAllSpacing(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    for (const para of paragraphs.items) {
      para.spaceBefore = 0;
      para.spaceAfter = 0;
    }
    await context.sync();
  });
}

export async function removeAllIndents(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    for (const para of paragraphs.items) {
      para.firstLineIndent = 0;
      para.leftIndent = 0;
      para.rightIndent = 0;
    }
    await context.sync();
  });
}

export async function setFirstLineIndent(cm: number): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    const pt = cmToPt(cm);
    for (const para of paragraphs.items) {
      para.firstLineIndent = pt;
    }
    await context.sync();
  });
}

export async function setHangingIndent(cm: number): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    const pt = cmToPt(cm);
    for (const para of paragraphs.items) {
      para.firstLineIndent = -pt;
      para.leftIndent = pt;
    }
    await context.sync();
  });
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

function buildBorderEdgeXml(edge: string, options: BorderOptions): string {
  if (options.type === "none") return "";
  const val = options.type ?? "single";
  const color = (options.color ?? "000000").replace("#", "");
  const sz = Math.round((options.width ?? 0.75) * 8);
  return `<w:${edge} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
}

export async function setParagraphBorder(
  location: "top" | "bottom" | "left" | "right" | "all",
  options: BorderOptions
): Promise<{ error?: string }> {
  const edges = location === "all" ? ["top", "left", "bottom", "right"] : [location];
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    for (const para of paragraphs.items) {
      const ooxml = para.getOoxml();
      await context.sync();
      const edgesXml = edges.map((e) => buildBorderEdgeXml(e, options)).join("");
      const pBdrXml = `<w:pBdr>${edgesXml}</w:pBdr>`;
      let updated: string;
      if (/<w:pBdr[\s>]/.test(ooxml.value)) {
        updated = ooxml.value.replace(/<w:pBdr[\s\S]*?<\/w:pBdr>/, pBdrXml);
      } else if (/<w:pPr[\s>]/.test(ooxml.value)) {
        updated = ooxml.value.replace(/(<w:pPr[^>]*>)/, `$1${pBdrXml}`);
      } else {
        throw new Error("Could not locate paragraph properties to apply border.");
      }
      para.insertOoxml(updated, Word.InsertLocation.replace);
    }
    await context.sync();
  });
}
