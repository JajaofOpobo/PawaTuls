/**
 * font.ts — Phase 10: font formatting operations.
 * Integration notes §1: Office.js reports mixed-selection properties as `null`
 * (COM's 9999999 does not apply). Toggles set `true` on null, else invert.
 * There is no toggle constant — always read before inverting.
 */
import { FontOptions, UnderlineStyle } from "../shared/types";
import { getSelectionOrThrow, safeWordRun } from "../shared/wordUtils";

export function mapUnderline(style: UnderlineStyle): Word.UnderlineType {
  switch (style) {
    case "none":
      return Word.UnderlineType.none;
    case "single":
      return Word.UnderlineType.single;
    case "double":
      return Word.UnderlineType.double;
    case "dotted":
      return Word.UnderlineType.dotted;
    case "dashed":
      return Word.UnderlineType.dashLine;
    case "wave":
      return Word.UnderlineType.wave;
    case "thick":
      return Word.UnderlineType.thick;
    case "thickDotted":
      return Word.UnderlineType.dottedHeavy;
    case "thickDashed":
      return Word.UnderlineType.dashLineHeavy;
  }
}

export async function applyFont(options: FontOptions): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    const font = selection.font;
    if (options.name !== undefined) font.name = options.name;
    if (options.size !== undefined) font.size = options.size;
    if (options.bold !== undefined) font.bold = options.bold;
    if (options.italic !== undefined) font.italic = options.italic;
    if (options.underline !== undefined) font.underline = mapUnderline(options.underline);
    if (options.strikeThrough !== undefined) font.strikeThrough = options.strikeThrough;
    if (options.superscript !== undefined) font.superscript = options.superscript;
    if (options.subscript !== undefined) font.subscript = options.subscript;
    if (options.color !== undefined) font.color = options.color;
    if (options.highlightColor !== undefined) font.highlightColor = options.highlightColor;
    if (options.allCaps !== undefined) font.allCaps = options.allCaps;
    if (options.smallCaps !== undefined) font.smallCaps = options.smallCaps;
    if (options.spacing !== undefined) font.spacing = options.spacing;
    if (options.hidden !== undefined) font.hidden = options.hidden;
    await context.sync();
  });
}

export async function increaseFontSize(by: number = 1): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.load("font/size");
    await context.sync();
    const base = selection.font.size ?? 12;
    selection.font.size = Math.max(1, base + by);
    await context.sync();
  });
}

export async function decreaseFontSize(by: number = 1): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.load("font/size");
    await context.sync();
    const base = selection.font.size ?? 12;
    selection.font.size = Math.max(1, base - by);
    await context.sync();
  });
}

export async function clearCharacterFormatting(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    const font = selection.font;
    font.bold = false;
    font.italic = false;
    font.underline = Word.UnderlineType.none;
    font.strikeThrough = false;
    font.superscript = false;
    font.subscript = false;
    font.allCaps = false;
    font.smallCaps = false;
    font.color = "auto";
    font.highlightColor = "none";
    font.spacing = 0;
    font.hidden = false;
    await context.sync();
  });
}

export async function toggleBold(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.load("font/bold");
    await context.sync();
    selection.font.bold = selection.font.bold === null ? true : !selection.font.bold;
    await context.sync();
  });
}

export async function toggleItalic(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.load("font/italic");
    await context.sync();
    selection.font.italic = selection.font.italic === null ? true : !selection.font.italic;
    await context.sync();
  });
}

export async function toggleStrikethrough(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.load("font/strikeThrough");
    await context.sync();
    selection.font.strikeThrough =
      selection.font.strikeThrough === null ? true : !selection.font.strikeThrough;
    await context.sync();
  });
}

export async function setSuperscript(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.font.superscript = true;
    selection.font.subscript = false;
    await context.sync();
  });
}

export async function setSubscript(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const selection = await getSelectionOrThrow(context);
    selection.font.subscript = true;
    selection.font.superscript = false;
    await context.sync();
  });
}
