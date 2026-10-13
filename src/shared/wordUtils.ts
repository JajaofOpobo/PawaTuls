/**
 * wordUtils.ts
 * Shared helpers for Office.js Word API operations.
 * All Word.run wrappers and unit converters live here.
 * Import these in every formatting, legal, and footnote module.
 */

export async function getSelectionOrThrow(context: Word.RequestContext): Promise<Word.Range> {
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();
  if (selection.text.trim().length === 0) {
    throw new Error("No text selected. Select text before applying this operation.");
  }
  return selection;
}

export async function getSelectionParagraphsOrThrow(
  context: Word.RequestContext
): Promise<Word.ParagraphCollection> {
  const paragraphs = context.document.getSelection().paragraphs;
  paragraphs.load("text");
  await context.sync();
  if (paragraphs.items.length === 0) {
    throw new Error("No text selected. Select text before applying this operation.");
  }
  return paragraphs;
}

export async function safeWordRun<T>(
  fn: (context: Word.RequestContext) => Promise<T>
): Promise<{ data?: T; error?: string }> {
  try {
    const data = await Word.run(fn);
    return { data };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { error: message };
  }
}

export async function getDocumentBodyOrThrow(context: Word.RequestContext): Promise<Word.Body> {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  return body;
}

export function ptToCm(pt: number): number {
  return pt / 28.3465;
}

export function cmToPt(cm: number): number {
  return cm * 28.3465;
}

export function inchToPt(inch: number): number {
  return inch * 72;
}
