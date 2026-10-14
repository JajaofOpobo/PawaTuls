/**
 * captions.ts — Phase 38: caption insertion and refresh.
 * Captions are label + SEQ field paragraphs (Word auto-numbers via the field).
 * SEQ fields render after updateResult, which runs on insert and on refresh.
 */
import { safeWordRun } from "../shared/wordUtils";

export type CaptionKind = "table" | "figure" | "equation";

const KIND_LABELS: Record<CaptionKind, string> = {
  table: "Table",
  figure: "Figure",
  equation: "Equation",
};

function captionOoxml(label: string, seqName: string): string {
  return (
    `<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${label} </w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> SEQ ${seqName} \\* ARABIC</w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>1</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
  );
}

function escapeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

async function updateSeqFields(context: Word.RequestContext): Promise<number> {
  const fields = context.document.body.fields.getByTypes([Word.FieldType.seq]);
  fields.load("items");
  await context.sync();
  for (const field of fields.items) {
    field.updateResult();
  }
  await context.sync();
  return fields.items.length;
}

export async function insertCaption(
  kind: CaptionKind,
  label?: string
): Promise<{ error?: string }> {
  const captionLabel = label !== undefined && label.trim() !== "" ? label.trim() : KIND_LABELS[kind];
  const seqName = KIND_LABELS[kind];
  return safeWordRun(async (context) => {
    const selection = context.document.getSelection();
    selection.insertOoxml(captionOoxml(escapeXml(captionLabel), seqName), Word.InsertLocation.after);
    await context.sync();
    await updateSeqFields(context);
  });
}

export async function refreshCaptions(): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => updateSeqFields(context));
}
