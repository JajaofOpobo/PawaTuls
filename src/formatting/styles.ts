/**
 * styles.ts — Phase 12: Word style operations.
 * Style.font and Style.paragraphFormat are the same object shapes as the
 * selection Range font / ParagraphFormat, so the mappers from font.ts and
 * paragraph.ts are reused here (no duplication).
 */
import { StyleCreationOptions } from "../shared/types";
import { getSelectionParagraphsOrThrow, safeWordRun } from "../shared/wordUtils";
import { mapUnderline } from "./font";
import { mapAlignment, mapLineSpacing } from "./paragraph";

export async function applyStyle(styleName: string): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    for (const para of paragraphs.items) {
      para.style = styleName;
    }
    await context.sync();
  });
}

export async function getAvailableStyles(): Promise<{ data?: string[]; error?: string }> {
  return safeWordRun(async (context) => {
    const styles = context.document.getStyles();
    styles.load("nameLocal, type");
    await context.sync();
    const names = styles.items
      .filter((s) => s.type === Word.StyleType.paragraph)
      .map((s) => s.nameLocal)
      .sort();
    return names;
  });
}

export async function createStyleFromSelection(
  options: StyleCreationOptions
): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const styleType =
      options.styleType === "character" ? Word.StyleType.character : Word.StyleType.paragraph;
    const newStyle = context.document.addStyle(options.name, styleType);
    if (options.basedOn !== undefined) {
      newStyle.baseStyle = options.basedOn;
    }
    const font = options.font;
    if (font !== undefined) {
      if (font.name !== undefined) newStyle.font.name = font.name;
      if (font.size !== undefined) newStyle.font.size = font.size;
      if (font.bold !== undefined) newStyle.font.bold = font.bold;
      if (font.italic !== undefined) newStyle.font.italic = font.italic;
      if (font.underline !== undefined) newStyle.font.underline = mapUnderline(font.underline);
      if (font.strikeThrough !== undefined) newStyle.font.strikeThrough = font.strikeThrough;
      if (font.superscript !== undefined) newStyle.font.superscript = font.superscript;
      if (font.subscript !== undefined) newStyle.font.subscript = font.subscript;
      if (font.color !== undefined) newStyle.font.color = font.color;
      if (font.highlightColor !== undefined) newStyle.font.highlightColor = font.highlightColor;
      if (font.allCaps !== undefined) newStyle.font.allCaps = font.allCaps;
      if (font.smallCaps !== undefined) newStyle.font.smallCaps = font.smallCaps;
      if (font.spacing !== undefined) newStyle.font.spacing = font.spacing;
      if (font.hidden !== undefined) newStyle.font.hidden = font.hidden;
    }
    const paragraph = options.paragraph;
    if (paragraph !== undefined) {
      const pf = newStyle.paragraphFormat;
      if (paragraph.alignment !== undefined) pf.alignment = mapAlignment(paragraph.alignment);
      if (paragraph.firstLineIndent !== undefined) pf.firstLineIndent = paragraph.firstLineIndent;
      if (paragraph.leftIndent !== undefined) pf.leftIndent = paragraph.leftIndent;
      if (paragraph.rightIndent !== undefined) pf.rightIndent = paragraph.rightIndent;
      if (paragraph.lineSpacingRule !== undefined) {
        pf.lineSpacing = mapLineSpacing(paragraph.lineSpacingRule, paragraph.lineSpacing);
      } else if (paragraph.lineSpacing !== undefined) {
        pf.lineSpacing = paragraph.lineSpacing;
      }
      if (paragraph.spaceBefore !== undefined) pf.spaceBefore = paragraph.spaceBefore;
      if (paragraph.spaceAfter !== undefined) pf.spaceAfter = paragraph.spaceAfter;
      if (paragraph.keepWithNext !== undefined) pf.keepWithNext = paragraph.keepWithNext;
      if (paragraph.keepLinesTogether !== undefined) {
        pf.keepTogether = paragraph.keepLinesTogether;
      }
    }
    await context.sync();
  });
}

export async function getSelectionStyle(): Promise<{ data?: string; error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    const first = paragraphs.items[0];
    first.load("style");
    await context.sync();
    return String(first.style);
  });
}
