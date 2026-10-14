/**
 * font.ts — font formatting engine (used by profiles; direct UI removed as
 * native-duplicate — Word's own Font controls cover these one-to-one).
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
