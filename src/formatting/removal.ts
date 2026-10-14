/**
 * removal.ts — Phase 35: KuTools-parity removal operations.
 * Every function takes applyTo ("selection" | "document", default "document")
 * and returns { data?: number } (items removed/changed). Structural deletions
 * confirm in UI (Phase 36), never here. Table/cell structure is only ever
 * mutated through table APIs or balanced OOXML surgery — never by deleting
 * cell-terminating paragraphs (audit lesson).
 */
import { getSelectionParagraphsOrThrow, safeWordRun } from "../shared/wordUtils";

export type RemovalScope = "selection" | "document";

function countMatches(xml: string, re: RegExp): number {
  return (xml.match(re) || []).length;
}

/** Batch OOXML rewrite over the resolved paragraphs. Returns changed count. */
async function rewriteParagraphs(
  context: Word.RequestContext,
  applyTo: RemovalScope,
  transform: (xml: string) => { xml: string; count: number }
): Promise<number> {
  // NOTE: collection items must be loaded before use — unloaded access yields
  // an empty array and the op would silently no-op on document scope.
  const paragraphs =
    applyTo === "selection"
      ? await getSelectionParagraphsOrThrow(context)
      : context.document.body.paragraphs;
  paragraphs.load("text");
  await context.sync();
  const ooxmls = paragraphs.items.map((p) => p.getOoxml());
  await context.sync();
  let total = 0;
  for (let i = 0; i < paragraphs.items.length; i++) {
    const result = transform(ooxmls[i].value);
    if (result.count > 0) {
      paragraphs.items[i].insertOoxml(result.xml, Word.InsertLocation.replace);
      total += result.count;
    }
  }
  await context.sync();
  return total;
}

/** Batch plain-text rewrite over the resolved paragraphs. */
async function rewriteParagraphTexts(
  context: Word.RequestContext,
  applyTo: RemovalScope,
  transform: (text: string) => string | null
): Promise<number> {
  const paragraphs =
    applyTo === "selection"
      ? await getSelectionParagraphsOrThrow(context)
      : context.document.body.paragraphs;
  paragraphs.load("text");
  await context.sync();
  let total = 0;
  for (const para of paragraphs.items) {
    const next = transform(para.text);
    if (next !== null && next !== para.text) {
      para.insertText(next, Word.InsertLocation.replace);
      total++;
    }
  }
  await context.sync();
  return total;
}

function stripBreaks(xml: string, type: "page" | "column"): { xml: string; count: number } {
  const re = new RegExp(`<w:br\\b[^>]*\\bw:type="${type}"[^>]*\\/>`, "g");
  return { xml: xml.replace(re, ""), count: countMatches(xml, re) };
}

export async function removePageBreaks(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => rewriteParagraphs(context, applyTo, (xml) => stripBreaks(xml, "page")));
}

export async function removeColumnBreaks(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => rewriteParagraphs(context, applyTo, (xml) => stripBreaks(xml, "column")));
}

export async function removeSectionBreaks(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  // Removing a section break merges the section into the previous one (same as
  // KuTools). Page setup of the trailing section is intentionally dropped.
  return safeWordRun(async (context) =>
    rewriteParagraphs(context, applyTo, (xml) => {
      const re = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/g;
      return { xml: xml.replace(re, ""), count: countMatches(xml, re) };
    })
  );
}

export async function removeAllBreaks(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  // Sequential Word.run calls (Word.run cannot nest); aborts on first error.
  const a = await removePageBreaks(applyTo);
  if (a.error) {
    return a;
  }
  const b = await removeColumnBreaks(applyTo);
  if (b.error) {
    return b;
  }
  const c = await removeSectionBreaks(applyTo);
  if (c.error) {
    return c;
  }
  return { data: (a.data ?? 0) + (b.data ?? 0) + (c.data ?? 0) };
}

export async function removeHiddenText(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) =>
    rewriteParagraphs(context, applyTo, (xml) => {
      const re = /<w:r\b[^>]*>[\s\S]*?<w:vanish\b[^/]*\/>[\s\S]*?<\/w:r>/g;
      return { xml: xml.replace(re, ""), count: countMatches(xml, re) };
    })
  );
}

export async function removeComments(): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();
    const items = comments.items.slice().reverse();
    for (const comment of items) {
      comment.delete();
    }
    await context.sync();
    return items.length;
  });
}

export async function removeDropCap(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) =>
    rewriteParagraphs(context, applyTo, (xml) => {
      const re = /<w:dropCap\b[^>]*\/>/g;
      return { xml: xml.replace(re, ""), count: countMatches(xml, re) };
    })
  );
}

export async function removeFrames(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) =>
    rewriteParagraphs(context, applyTo, (xml) => {
      const re = /<w:framePr\b[^>]*\/>/g;
      return { xml: xml.replace(re, ""), count: countMatches(xml, re) };
    })
  );
}

export async function removePictures(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  // <w:drawing> elements never nest, so non-greedy removal is balanced.
  return safeWordRun(async (context) =>
    rewriteParagraphs(context, applyTo, (xml) => {
      const re = /<w:drawing\b[^>]*>[\s\S]*?<pic:pic\b[\s\S]*?<\/w:drawing>/g;
      return { xml: xml.replace(re, ""), count: countMatches(xml, re) };
    })
  );
}

export async function removeTextBoxes(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  // Textbox content nests runs, so the tempered pattern forbids crossing a
  // nested <w:r> before reaching <w:txbxContent>.
  return safeWordRun(async (context) =>
    rewriteParagraphs(context, applyTo, (xml) => {
      const re =
        /<w:r\b[^>]*>(?:(?!<w:r\b)[\s\S])*?<w:txbxContent\b[\s\S]*?<\/w:txbxContent>(?:(?!<w:r\b)[\s\S])*?<\/w:r>/g;
      return { xml: xml.replace(re, ""), count: countMatches(xml, re) };
    })
  );
}

export async function removeTables(): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    const items = tables.items.slice().reverse();
    for (const table of items) {
      table.delete();
    }
    await context.sync();
    return items.length;
  });
}

export async function removeEmptyTableRows(): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    let deleted = 0;
    for (const table of tables.items) {
      const rows = table.rows;
      rows.load("items, values");
      await context.sync();
      for (let i = rows.items.length - 1; i >= 0; i--) {
        const cells = rows.items[i].values;
        const empty = cells.every((row) => row.every((cell) => cell.trim() === ""));
        if (empty) {
          rows.items[i].delete();
          deleted++;
        }
      }
    }
    await context.sync();
    return deleted;
  });
}

export async function removeEmptyTableColumns(): Promise<{ data?: number; error?: string }> {
  // TableRow.values is a 2D array with one entry per row: values[0][c] is the
  // text of column c. (Treating values.length as the column count only ever
  // evaluated column 0 — fixed.)
  return safeWordRun(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    let deleted = 0;
    for (const table of tables.items) {
      const rows = table.rows;
      rows.load("items, values");
      await context.sync();
      if (rows.items.length === 0) {
        continue;
      }
      const colCount = Math.min(
        ...rows.items.map((r) => (r.values[0] ?? []).length)
      );
      if (colCount === 0) {
        continue;
      }
      const firstCells = rows.items[0].cells;
      firstCells.load("items");
      await context.sync();
      for (let c = Math.min(colCount, firstCells.items.length) - 1; c >= 0; c--) {
        const allEmpty = rows.items.every((row) => {
          const cells = row.values[0] ?? [];
          return (cells[c] ?? "").trim() === "";
        });
        if (allEmpty) {
          firstCells.items[c].deleteColumn();
          deleted++;
        }
      }
    }
    await context.sync();
    return deleted;
  });
}

export async function removeContentControls(
  keepText = true
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const controls = context.document.body.getContentControls();
    controls.load("items");
    await context.sync();
    const items = controls.items.slice().reverse();
    for (const control of items) {
      control.delete(keepText);
    }
    await context.sync();
    return items.length;
  });
}

export async function removeBookmarks(): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const names = context.document.body.getRange().getBookmarks();
    await context.sync();
    const list = names.value.slice();
    for (const name of list) {
      context.document.deleteBookmark(name);
    }
    await context.sync();
    return list.length;
  });
}

export async function removeHyperlinks(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  // Unwrap <w:hyperlink> elements, keeping display content (same outcome as
  // Hyperlink.Delete in COM). Elements never nest.
  return safeWordRun(async (context) =>
    rewriteParagraphs(context, applyTo, (xml) => {
      const re = /<w:hyperlink\b[^>]*>([\s\S]*?)<\/w:hyperlink>/g;
      return { xml: xml.replace(re, "$1"), count: countMatches(xml, re) };
    })
  );
}

export async function removeHorizontalLines(): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    const candidates = paragraphs.items.slice().reverse();
    const ooxmlResults = candidates.map((para) => para.getOoxml());
    const tableChecks = candidates.map((para) => {
      const t = para.parentTableOrNullObject;
      context.load(t);
      return t;
    });
    await context.sync();
    let deleted = 0;
    for (let i = 0; i < candidates.length; i++) {
      // Audit lesson: cell borders are normal table formatting — never delete
      // in-table paragraphs for having borders.
      if (!tableChecks[i].isNullObject) {
        continue;
      }
      const para = candidates[i];
      if (/^[-_=]{3,}$/.test(para.text.trim())) {
        para.delete();
        deleted++;
      } else if (/<w:pBdr[\s\S]*?<w:bottom[\s>/]/.test(ooxmlResults[i].value)) {
        para.delete();
        deleted++;
      }
    }
    await context.sync();
    return deleted;
  });
}

export async function removeExtraDocumentSpacing(): Promise<{
  data?: number;
  error?: string;
}> {
  return safeWordRun(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("spaceBefore, spaceAfter");
    await context.sync();
    let adjusted = 0;
    for (const para of paragraphs.items) {
      if (para.spaceBefore > 12 || para.spaceAfter > 12) {
        para.spaceBefore = 6;
        para.spaceAfter = 6;
        adjusted++;
      }
    }
    await context.sync();
    return adjusted;
  });
}


export async function tabToSpace(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) =>
    rewriteParagraphTexts(context, applyTo, (text) =>
      text.includes("	") ? text.replace(/	/g, " ") : null
    )
  );
}

export async function removeEmptyLines(
  applyTo: RemovalScope = "document"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) =>
    rewriteParagraphTexts(context, applyTo, (text) =>
      /{2,}/.test(text) ? text.replace(/{2,}/g, "") : null
    )
  );
}
