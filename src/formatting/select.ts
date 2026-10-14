/**
 * select.ts — Phase 39: jump-to selection helpers.
 * Office.js has a single selection (no multi-select), so every function
 * selects the FIRST match with range.select() and reports the total count;
 * the user repeats the command to step through matches.
 */
import { safeWordRun } from "../shared/wordUtils";

export interface SelectResult {
  total: number;
}

async function selectFirstOf(
  context: Word.RequestContext,
  selectables: { select: () => void }[]
): Promise<SelectResult> {
  if (selectables.length > 0) {
    selectables[0].select();
  }
  await context.sync();
  return { total: selectables.length };
}

export async function selectTables(): Promise<{ data?: SelectResult; error?: string }> {
  return safeWordRun(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    const rows: Word.TableRow[] = [];
    for (const table of tables.items) {
      const tableRows = table.rows;
      tableRows.load("items");
      await context.sync();
      if (tableRows.items.length > 0) {
        rows.push(tableRows.items[0]);
      }
    }
    return selectFirstOf(context, rows);
  });
}

export async function selectEmptyParagraphs(): Promise<{
  data?: SelectResult;
  error?: string;
}> {
  return safeWordRun(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    const empties = paragraphs.items.filter((p) => p.text.trim() === "");
    return selectFirstOf(context, empties);
  });
}

export async function selectHeadings(): Promise<{ data?: SelectResult; error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text, style");
    await context.sync();
    const headings = paragraphs.items.filter((p) => String(p.style).startsWith("Heading"));
    return selectFirstOf(context, headings);
  });
}

export async function selectShortParagraphs(
  maxChars = 25
): Promise<{ data?: SelectResult; error?: string }> {
  // KuTools ShortParaMaxChars parity (default 25, see decompilation config).
  return safeWordRun(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    const shorts = paragraphs.items.filter((p) => {
      const len = p.text.trim().length;
      return len > 0 && len <= maxChars;
    });
    return selectFirstOf(context, shorts);
  });
}

export async function selectPictures(): Promise<{ data?: SelectResult; error?: string }> {
  return safeWordRun(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load("items");
    await context.sync();
    const ranges = pictures.items.map((pic) => pic.getRange());
    return selectFirstOf(context, ranges);
  });
}
