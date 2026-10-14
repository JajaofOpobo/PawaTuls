/**
 * tables.ts — Phase 40: table and picture operations.
 * Table resolution: the selection's parentTable (single table — the 99% case).
 * Multi-table selections convert/apply to the first table found; documented.
 */
import { safeWordRun } from "../shared/wordUtils";

async function getSelectedTables(
  context: Word.RequestContext
): Promise<{ tables: Word.Table[]; error?: string }> {
  const selection = context.document.getSelection();
  const paras = selection.paragraphs;
  paras.load("text");
  await context.sync();
  if (paras.items.length === 0) {
    return { tables: [], error: "No text selected. Select inside a table first." };
  }
  const seen = new Set<string>();
  const tables: Word.Table[] = [];
  const parents = paras.items.map((para) => {
    const table = para.parentTableOrNullObject;
    context.load(table);
    return table;
  });
  await context.sync();
  const inTable = parents.filter((table) => !table.isNullObject);
  for (const table of inTable) {
    table.load("values");
  }
  await context.sync();
  for (const table of inTable) {
    const fingerprint = JSON.stringify(table.values);
    if (!seen.has(fingerprint)) {
      seen.add(fingerprint);
      tables.push(table);
    }
  }
  return { tables };
}

export type TableSeparator = "tab" | "paragraph" | "comma";

function mapSeparator(separator: TableSeparator): Word.TableFieldSeparator {
  switch (separator) {
    case "tab":
      return Word.TableFieldSeparator.tab;
    case "paragraph":
      return Word.TableFieldSeparator.paragraph;
    case "comma":
      return Word.TableFieldSeparator.comma;
  }
}

/**
 * NOTE: Row.convertToText needs the WordApiDesktop 1.4 requirement set
 * (desktop-only). The add-in targets Word desktop exclusively, so this holds;
 * it is unrelated to the WordApi 1.5 minimum in the manifest.
 */
export async function convertTableToText(
  separator: TableSeparator = "tab"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const resolved = await getSelectedTables(context);
    if (resolved.error) {
      throw new Error(resolved.error);
    }
    if (resolved.tables.length === 0) {
      throw new Error("No text selected. Select inside a table first.");
    }
    let converted = 0;
    for (const table of resolved.tables) {
      const rows = table.rows;
      rows.load("items");
      await context.sync();
      for (let i = rows.items.length - 1; i >= 0; i--) {
        rows.items[i].convertToText({ separator: mapSeparator(separator) });
      }
      converted++;
    }
    await context.sync();
    return converted;
  });
}

export async function repeatHeaderRows(enable = true): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const resolved = await getSelectedTables(context);
    if (resolved.error) {
      throw new Error(resolved.error);
    }
    if (resolved.tables.length === 0) {
      throw new Error("No text selected. Select inside a table first.");
    }
    for (const table of resolved.tables) {
      table.headerRowCount = enable ? 1 : 0;
    }
    await context.sync();
  });
}

export async function resizeImages(percent: number): Promise<{ data?: number; error?: string }> {
  if (!(percent >= 10 && percent <= 400)) {
    return { error: "Percent must be between 10 and 400." };
  }
  return safeWordRun(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load("items, width, height");
    await context.sync();
    for (const pic of pictures.items) {
      pic.width = (pic.width * percent) / 100;
      pic.height = (pic.height * percent) / 100;
    }
    await context.sync();
    return pictures.items.length;
  });
}
