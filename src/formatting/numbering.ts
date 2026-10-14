/**
 * numbering.ts — Phase 33: numbering audit and preview-first repair.
 * Appendix I amendment: native membership is classified via
 * paragraph.listItemOrNullObject (verified API), not OOXML numPr.
 * The auditor is strictly read-only; repair applies only caller-approved plans,
 * re-resolves each paragraph at apply time, and skips on mismatch (R4/R7).
 */
import {
  NumberedParagraph,
  NumberingIssue,
  NumberingRepairPlan,
} from "../shared/types";
import { safeWordRun } from "../shared/wordUtils";

type ManualStyle = "decimal" | "alpha" | "roman";

const MANUAL_PREFIX = /^(\d+(?:\.\d+)*\.|\([a-z]+\)|\([ivxlcdm]+\))\s+/i;

function detectManualStyle(prefix: string, hint?: ManualStyle | null): ManualStyle | null {
  if (/^\d/.test(prefix)) {
    return "decimal";
  }
  const m = prefix.match(/^\(([a-z]+)\)$/i);
  if (!m) {
    return null;
  }
  const letters = m[1].toLowerCase();
  if (!/^[ivxlcdm]+$/.test(letters)) {
    return "alpha";
  }
  if (letters.length > 1) {
    // Multi-char: roman unless the running sequence is already alpha.
    return hint === "alpha" ? "alpha" : "roman";
  }
  // Single ambiguous char (c, d, i, l, m, v, x): follow the running sequence;
  // on first sight prefer roman only for the classic i/v/x.
  if (hint === "roman" || hint === "alpha") {
    return hint;
  }
  return letters === "i" || letters === "v" || letters === "x" ? "roman" : "alpha";
}

function parseDecimal(prefix: string): number[] | null {
  const m = prefix.match(/^(\d+(?:\.\d+)*)\./);
  if (!m) {
    return null;
  }
  return m[1].split(".").map((n) => parseInt(n, 10));
}

function parseAlpha(prefix: string): number | null {
  const m = prefix.match(/^\(([a-z]+)\)$/i);
  if (!m) {
    return null;
  }
  const letters = m[1].toLowerCase();
  let value = 0;
  for (const ch of letters) {
    value = value * 26 + (ch.charCodeAt(0) - 96);
  }
  return value;
}

const ROMAN_VALUES: [string, number][] = [
  ["m", 1000], ["cm", 900], ["d", 500], ["cd", 400], ["c", 100],
  ["xc", 90], ["l", 50], ["xl", 40], ["x", 10], ["ix", 9],
  ["v", 5], ["iv", 4], ["i", 1],
];

function parseRoman(prefix: string): number | null {
  const m = prefix.match(/^\(([ivxlcdm]+)\)$/i);
  if (!m) {
    return null;
  }
  let text = m[1].toLowerCase();
  let value = 0;
  for (const [numeral, arabic] of ROMAN_VALUES) {
    while (text.startsWith(numeral)) {
      value += arabic;
      text = text.substring(numeral.length);
    }
  }
  return text === "" ? value : null;
}

export function toAlpha(n: number): string {
  let result = "";
  let value = n;
  while (value > 0) {
    value--;
    result = String.fromCharCode(97 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

export function toRoman(n: number): string {
  let result = "";
  let value = n;
  for (const [numeral, arabic] of ROMAN_VALUES) {
    while (value >= arabic) {
      result += numeral;
      value -= arabic;
    }
  }
  return result;
}

function manualLevel(leftIndent: number): number {
  if (leftIndent < 18) {
    return 1;
  }
  if (leftIndent <= 36) {
    return 2;
  }
  return 3;
}

function scalarOf(prefix: string, style: ManualStyle): number[] | null {  if (style === "decimal") {
    return parseDecimal(prefix);
  }
  if (style === "alpha") {
    const v = parseAlpha(prefix);
    return v === null ? null : [v];
  }
  const v = parseRoman(prefix);
  return v === null ? null : [v];
}

function renderPrefix(style: ManualStyle, values: number[]): string {
  if (style === "decimal") {
    return `${values.join(".")}.`;
  }
  if (style === "alpha") {
    return `(${toAlpha(values[values.length - 1])})`;
  }
  return `(${toRoman(values[values.length - 1])})`;
}

interface Classified {
  style: ManualStyle;
  values: number[];
}

/**
 * Classify manual prefixes in document order, carrying the running style per
 * level as a hint so ambiguous single letters ((c), (d) …) follow their
 * sequence instead of misreading as Roman numerals.
 */
export function classifyManualSequence(
  items: { paragraphIndex: number; prefix: string; level: number }[]
): Map<number, Classified | null> {
  const result = new Map<number, Classified | null>();
  const hintByLevel = new Map<number, ManualStyle>();
  for (const item of items) {
    const style = detectManualStyle(item.prefix, hintByLevel.get(item.level) ?? null);
    if (style === null) {
      result.set(item.paragraphIndex, null);
      continue;
    }
    const values = scalarOf(item.prefix, style);
    if (values === null) {
      result.set(item.paragraphIndex, null);
      continue;
    }
    hintByLevel.set(item.level, style);
    result.set(item.paragraphIndex, { style, values });
  }
  return result;
}

export async function auditNumbering(
  applyTo: "selection" | "document"
): Promise<{
  data?: { items: NumberedParagraph[]; issues: NumberingIssue[] };
  error?: string;
}> {
  return safeWordRun(async (context) => {
    const paragraphs =
      applyTo === "selection"
        ? context.document.getSelection().paragraphs
        : context.document.body.paragraphs;
    paragraphs.load("text, leftIndent");
    await context.sync();

    const items: NumberedParagraph[] = [];
    const listItems = paragraphs.items.map((para) => {
      const listItem = para.listItemOrNullObject;
      context.load(listItem);
      return listItem;
    });
    await context.sync();

    const levels = new Map<number, number>();
    const needLevel: number[] = [];
    listItems.forEach((listItem, i) => {
      if (!listItem.isNullObject) {
        listItem.load("level");
        needLevel.push(i);
      }
    });
    if (needLevel.length > 0) {
      await context.sync();
      for (const i of needLevel) {
        levels.set(i, listItems[i].level);
      }
    }

    for (let i = 0; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];
      if (levels.has(i)) {
        items.push({ paragraphIndex: i, kind: "native", level: levels.get(i) });
        continue;
      }
      const text = para.text;
      const m = text.match(MANUAL_PREFIX);
      if (m) {
        items.push({
          paragraphIndex: i,
          kind: "manual",
          prefix: m[1],
          level: manualLevel(para.leftIndent),
        });
      } else {
        items.push({ paragraphIndex: i, kind: "plain" });
      }
    }

    const issues: NumberingIssue[] = [];
    const seenLevels = new Set<number>();
    const lastByLevel = new Map<number, { values: number[]; style: ManualStyle; index: number }>();
    const manualItems = items.filter(
      (item): item is NumberedParagraph & { prefix: string; level: number } =>
        item.kind === "manual" && item.prefix !== undefined && item.level !== undefined
    );
    const classified = classifyManualSequence(manualItems);
    for (const item of manualItems) {
      if (item.level > 1 && !seenLevels.has(item.level - 1)) {
        issues.push({
          paragraphIndex: item.paragraphIndex,
          issue: "orphan-level",
          detail: `Level ${item.level} item "${item.prefix}" with no level ${item.level - 1} parent seen.`,
        });
      }
      seenLevels.add(item.level);
      const cls = classified.get(item.paragraphIndex);
      if (!cls) {
        continue;
      }
      const { style, values } = cls;
      const last = lastByLevel.get(item.level);
      if (last !== undefined) {
        if (last.style !== style) {
          issues.push({
            paragraphIndex: item.paragraphIndex,
            issue: "mixed-style",
            detail: `Style changed from ${last.style} to ${style} at level ${item.level} ("${item.prefix}"). Review before repair.`,
          });
        } else {
          const prev = last.values[last.values.length - 1];
          const cur = values[values.length - 1];
          const sameStem =
            last.values.slice(0, -1).join(".") === values.slice(0, -1).join(".");
          if (sameStem) {
            if (cur === prev) {
              issues.push({
                paragraphIndex: item.paragraphIndex,
                issue: "duplicate",
                detail: `Duplicate number "${item.prefix}" at level ${item.level}.`,
              });
            } else if (cur < prev) {
              issues.push({
                paragraphIndex: item.paragraphIndex,
                issue: "restart",
                detail: `Sequence resets to "${item.prefix}" at level ${item.level}. Left alone unless approved.`,
              });
            } else if (cur > prev + 1) {
              issues.push({
                paragraphIndex: item.paragraphIndex,
                issue: "gap",
                detail: `Gap at level ${item.level}: "${item.prefix}" follows ${last.values.join(".")}.`,
              });
            }
          }
        }
      }
      lastByLevel.set(item.level, { values, style, index: item.paragraphIndex });
    }

    return { items, issues };
  });
}

export function previewRepair(
  issues: NumberingIssue[],
  items: NumberedParagraph[]
): NumberingRepairPlan {
  // Only gap/duplicate are auto-repaired; restarts and mixed-style are user
  // intent until explicitly approved (listed, never rewritten here).
  const byIndex = new Map(items.map((item) => [item.paragraphIndex, item]));
  const manualItems = items.filter(
    (item): item is NumberedParagraph & { prefix: string; level: number } =>
      item.kind === "manual" && item.prefix !== undefined && item.level !== undefined
  );
  const classified = classifyManualSequence(manualItems);
  const repaired = new Map<number, number[]>();
  const edits: NumberingRepairPlan["edits"] = [];
  const ordered = [...issues].sort((a, b) => a.paragraphIndex - b.paragraphIndex);
  for (const issue of ordered) {
    if (issue.issue !== "gap" && issue.issue !== "duplicate") {
      continue;
    }
    const item = byIndex.get(issue.paragraphIndex);
    if (!item || item.kind !== "manual" || item.prefix === undefined || item.level === undefined) {
      continue;
    }
    const manual = item as NumberedParagraph & { prefix: string; level: number };
    const cls = classified.get(issue.paragraphIndex);
    if (!cls) {
      continue;
    }
    const { style, values } = cls;
    // Walk back to the nearest preceding repaired or original same-level value.
    const stem = values.slice(0, -1);
    let base = 0;
    for (let i = items.indexOf(item) - 1; i >= 0; i--) {
      const prev = items[i];
      if (prev.kind !== "manual" || prev.prefix === undefined || prev.level !== manual.level) {
        continue;
      }
      const prevCls = classified.get(prev.paragraphIndex);
      if (!prevCls || prevCls.style !== style) {
        break;
      }
      const prevValues = prevCls.values;
      if (prevValues.slice(0, -1).join(".") !== stem.join(".")) {
        break;
      }
      const fixed = repaired.get(prev.paragraphIndex);
      base = fixed !== undefined ? fixed[fixed.length - 1] : prevValues[prevValues.length - 1];
      break;
    }
    const next = [...stem, base + 1];
    repaired.set(manual.paragraphIndex, next);
    edits.push({ paragraphIndex: manual.paragraphIndex, before: manual.prefix, after: renderPrefix(style, next) });
  }
  return { edits };
}

export async function applyRepairPlan(
  plan: NumberingRepairPlan,
  applyTo: "selection" | "document" = "document"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs =
      applyTo === "selection"
        ? context.document.getSelection().paragraphs
        : context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    const ordered = [...plan.edits].sort((a, b) => b.paragraphIndex - a.paragraphIndex);
    let applied = 0;
    const skipped: number[] = [];
    for (const edit of ordered) {
      const para = paragraphs.items[edit.paragraphIndex];
      if (!para || !para.text.startsWith(edit.before)) {
        skipped.push(edit.paragraphIndex);
        continue;
      }
      para.insertText(edit.after + para.text.substring(edit.before.length), Word.InsertLocation.replace);
      applied++;
    }
    await context.sync();
    if (skipped.length > 0) {
      throw new Error(
        `Skipped ${skipped.length} paragraph(s) whose text changed since preview (indices ${skipped.join(", ")}).`
      );
    }
    return applied;
  });
}
