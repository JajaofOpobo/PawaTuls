/**
 * profiles.ts — Phase 31: formatting profile store and apply.
 * A profile changes ONLY what it targets (R1): missing font/paragraph/style
 * sections are legal and left untouched. Reuses applyFont/applyParagraphFormat/
 * applyStyle for selection scope and the exported mappers for document scope
 * (body paragraphs, cursor never moved — audit lesson).
 */
import { FormattingProfile } from "../shared/types";
import { safeWordRun } from "../shared/wordUtils";
import { applyFont, mapUnderline } from "./font";
import { applyParagraphFormat, mapAlignment, mapLineSpacing } from "./paragraph";
import { applyStyle } from "./styles";

const TOP_LEVEL_KEYS = new Set([
  "schemaVersion",
  "id",
  "name",
  "description",
  "style",
  "font",
  "paragraph",
]);

const UNDERLINES = new Set([
  "none",
  "single",
  "double",
  "dotted",
  "dashed",
  "wave",
  "thick",
  "thickDotted",
  "thickDashed",
]);

const ALIGNMENTS = new Set(["left", "centered", "right", "justified"]);

const SPACING_RULES = new Set([
  "single",
  "onePtFive",
  "double",
  "atLeast",
  "exactly",
  "multiple",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function checkNumbers(
  obj: Record<string, unknown>,
  keys: string[],
  path: string,
  problems: string[]
): void {
  for (const key of keys) {
    const value = obj[key];
    if (value !== undefined && typeof value !== "number") {
      problems.push(`${path}.${key} must be a number`);
    }
  }
}

export function validateProfile(input: unknown): {
  data?: FormattingProfile;
  error?: string;
} {
  if (!isRecord(input)) {
    return { error: "Profile must be a JSON object." };
  }
  const problems: string[] = [];
  for (const key of Object.keys(input)) {
    if (!TOP_LEVEL_KEYS.has(key)) {
      problems.push(`Unknown field "${key}"`);
    }
  }
  if (input.schemaVersion !== 1) {
    problems.push("schemaVersion must be 1");
  }
  if (typeof input.id !== "string" || input.id.trim() === "") {
    problems.push("id must be a non-empty string");
  }
  if (typeof input.name !== "string" || input.name.trim() === "") {
    problems.push("name must be a non-empty string");
  }
  if (input.description !== undefined && typeof input.description !== "string") {
    problems.push("description must be a string");
  }
  if (input.style !== undefined && typeof input.style !== "string") {
    problems.push("style must be a string");
  }
  if (input.font !== undefined) {
    if (!isRecord(input.font)) {
      problems.push("font must be an object");
    } else {
      const font = input.font;
      if (font.underline !== undefined && !UNDERLINES.has(String(font.underline))) {
        problems.push(`font.underline is not a known style: ${String(font.underline)}`);
      }
      for (const key of ["bold", "italic", "strikeThrough", "superscript", "subscript", "allCaps", "smallCaps", "hidden"]) {
        if (font[key] !== undefined && typeof font[key] !== "boolean") {
          problems.push(`font.${key} must be boolean`);
        }
      }
      for (const key of ["name", "color", "highlightColor"]) {
        if (font[key] !== undefined && typeof font[key] !== "string") {
          problems.push(`font.${key} must be a string`);
        }
      }
      checkNumbers(font, ["size", "spacing"], "font", problems);
    }
  }
  if (input.paragraph !== undefined) {
    if (!isRecord(input.paragraph)) {
      problems.push("paragraph must be an object");
    } else {
      const paragraph = input.paragraph;
      if (paragraph.alignment !== undefined && !ALIGNMENTS.has(String(paragraph.alignment))) {
        problems.push(`paragraph.alignment is not known: ${String(paragraph.alignment)}`);
      }
      if (
        paragraph.lineSpacingRule !== undefined &&
        !SPACING_RULES.has(String(paragraph.lineSpacingRule))
      ) {
        problems.push(`paragraph.lineSpacingRule is not known: ${String(paragraph.lineSpacingRule)}`);
      }
      for (const key of ["keepWithNext", "keepLinesTogether", "pageBreakBefore"]) {
        if (paragraph[key] !== undefined && typeof paragraph[key] !== "boolean") {
          problems.push(`paragraph.${key} must be boolean`);
        }
      }
      checkNumbers(
        paragraph,
        ["firstLineIndent", "leftIndent", "rightIndent", "lineSpacing", "spaceBefore", "spaceAfter"],
        "paragraph",
        problems
      );
    }
  }
  if (problems.length > 0) {
    return { error: problems.join("; ") };
  }
  return { data: input as unknown as FormattingProfile };
}

async function applyToBodyParagraphs(
  context: Word.RequestContext,
  profile: FormattingProfile
): Promise<number> {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("text");
  await context.sync();
  const count = paragraphs.items.length;
  for (const para of paragraphs.items) {
    if (profile.style !== undefined) {
      para.style = profile.style;
    }
    const font = profile.font;
    if (font !== undefined) {
      if (font.name !== undefined) para.font.name = font.name;
      if (font.size !== undefined) para.font.size = font.size;
      if (font.bold !== undefined) para.font.bold = font.bold;
      if (font.italic !== undefined) para.font.italic = font.italic;
      if (font.underline !== undefined) para.font.underline = mapUnderline(font.underline);
      if (font.strikeThrough !== undefined) para.font.strikeThrough = font.strikeThrough;
      if (font.superscript !== undefined) para.font.superscript = font.superscript;
      if (font.subscript !== undefined) para.font.subscript = font.subscript;
      if (font.color !== undefined) para.font.color = font.color;
      if (font.highlightColor !== undefined) para.font.highlightColor = font.highlightColor;
      if (font.allCaps !== undefined) para.font.allCaps = font.allCaps;
      if (font.smallCaps !== undefined) para.font.smallCaps = font.smallCaps;
      if (font.spacing !== undefined) para.font.spacing = font.spacing;
      if (font.hidden !== undefined) para.font.hidden = font.hidden;
    }
    const paragraph = profile.paragraph;
    if (paragraph !== undefined) {
      if (paragraph.alignment !== undefined) para.alignment = mapAlignment(paragraph.alignment);
      if (paragraph.firstLineIndent !== undefined) para.firstLineIndent = paragraph.firstLineIndent;
      if (paragraph.leftIndent !== undefined) para.leftIndent = paragraph.leftIndent;
      if (paragraph.rightIndent !== undefined) para.rightIndent = paragraph.rightIndent;
      if (paragraph.lineSpacingRule !== undefined) {
        para.lineSpacing = mapLineSpacing(paragraph.lineSpacingRule, paragraph.lineSpacing);
      } else if (paragraph.lineSpacing !== undefined) {
        para.lineSpacing = paragraph.lineSpacing;
      }
      if (paragraph.spaceBefore !== undefined) para.spaceBefore = paragraph.spaceBefore;
      if (paragraph.spaceAfter !== undefined) para.spaceAfter = paragraph.spaceAfter;
    }
  }
  await context.sync();
  return count;
}

export async function applyProfile(
  profile: FormattingProfile,
  applyTo: "selection" | "document"
): Promise<{ data?: number; error?: string }> {
  const check = validateProfile(profile);
  if (!check.data) {
    return { error: check.error };
  }
  const valid = check.data;
  if (applyTo === "document") {
    return safeWordRun(async (context) => applyToBodyParagraphs(context, valid));
  }
  // Selection scope: each module call is its own Word.run (Word.run cannot
  // nest), executed in style → font → paragraph order. Any failure aborts.
  if (valid.style !== undefined) {
    const r = await applyStyle(valid.style);
    if (r.error) {
      return { error: r.error };
    }
  }
  if (valid.font !== undefined) {
    const r = await applyFont(valid.font);
    if (r.error) {
      return { error: r.error };
    }
  }
  if (valid.paragraph !== undefined) {
    const r = await applyParagraphFormat(valid.paragraph);
    if (r.error) {
      return { error: r.error };
    }
  }
  return safeWordRun(async (context) => {
    const paras = context.document.getSelection().paragraphs;
    paras.load("text");
    await context.sync();
    return paras.items.length;
  });
}

export function exportProfile(profile: FormattingProfile): string {
  const check = validateProfile(profile);
  if (!check.data) {
    throw new Error(check.error);
  }
  return JSON.stringify(check.data, null, 2);
}

export function importProfile(json: string): {
  data?: FormattingProfile;
  error?: string;
} {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    return { error: "Invalid JSON: could not parse profile." };
  }
  return validateProfile(parsed);
}

export const BUILTIN_PROFILES: FormattingProfile[] = [
  {
    schemaVersion: 1,
    id: "court-body-1",
    name: "Court Body",
    description: "Justified body text with first-line indent.",
    font: { name: "Times New Roman", size: 12 },
    paragraph: {
      alignment: "justified",
      firstLineIndent: 36,
      lineSpacingRule: "single",
      spaceAfter: 6,
    },
  },
  {
    schemaVersion: 1,
    id: "court-heading-1",
    name: "Court Heading",
    description: "Centered bold heading.",
    font: { name: "Times New Roman", size: 14, bold: true },
    paragraph: { alignment: "centered", spaceBefore: 12, spaceAfter: 12 },
  },
  {
    schemaVersion: 1,
    id: "single-space-reset-1",
    name: "Single Space Reset",
    description: "Single spacing with no indents or gaps.",
    paragraph: {
      lineSpacingRule: "single",
      spaceBefore: 0,
      spaceAfter: 0,
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
    },
  },
];
