/**
 * convert.ts — Phase 14: conversion tools.
 *
 * API notes (verified against @types/office-js):
 * - Manual line break (Shift+Enter) === "\v" in body text; "\n" inserts a
 *   paragraph mark, "\v" a manual break.
 * - Word.Paragraph has no sub-range delete for the paragraph mark, so
 *   paragraphMarksToLineBreaks merges via selection OOXML: paragraph boundaries
 *   (</w:p><w:p…> + dropped pPr) become <w:br/>. Same end state as the plan's
 *   insert-and-merge description, formatting preserved.
 * - Word.Field has NO result range prop, so fieldCodesToText uses unlink()
 *   (replaces each field with its current result text) instead of the plan's
 *   insert-result-then-delete sequence. Same outcome.
 */
import { getSelectionParagraphsOrThrow, safeWordRun } from "../shared/wordUtils";

export async function lineBreaksToParagraphMarks(): Promise<{
  data?: number;
  error?: string;
}> {
  return safeWordRun(async (context) => {
    const hits = context.document.body.search("\v");
    hits.load("text");
    await context.sync();
    const count = hits.items.length;
    for (const hit of hits.items) {
      hit.insertText("\n", Word.InsertLocation.replace);
    }
    await context.sync();
    return count;
  });
}

export async function paragraphMarksToLineBreaks(): Promise<{
  data?: number;
  error?: string;
}> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    if (paragraphs.items.length < 2) {
      return 0;
    }
    const selection = context.document.getSelection();
    const ooxml = selection.getOoxml();
    await context.sync();
    const xml = ooxml.value;
    if (/<w:tbl[\s>]/.test(xml)) {
      throw new Error("Selection contains a table. Convert paragraphs separately.");
    }
    const merged = xml.replace(
      /<\/w:p>\s*<w:p[^>]*>\s*(<w:pPr>[\s\S]*?<\/w:pPr>)?/g,
      "<w:br/>"
    );
    selection.insertOoxml(merged, Word.InsertLocation.replace);
    await context.sync();
    return paragraphs.items.length - 1;
  });
}

export async function fieldCodesToText(): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const fields = context.document.body.fields;
    fields.load("items");
    await context.sync();
    const count = fields.items.length;
    for (const field of fields.items) {
      field.unlink();
    }
    await context.sync();
    return count;
  });
}

export async function listToText(): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    for (const para of paragraphs.items) {
      para.detachFromList();
      para.leftIndent = 0;
      para.firstLineIndent = 0;
    }
    await context.sync();
  });
}
