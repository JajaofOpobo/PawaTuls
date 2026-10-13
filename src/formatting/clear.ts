/**
 * clear.ts — Phase 15: clear-formatting operations.
 * Integration notes §3: KuTools has no Clear*AllFormatting calls — individual
 * resets + Normal style, which is exactly what this module does.
 *
 * API note: Word.Paragraph has no getBorder (table-only API), so bottom-border
 * detection in removeHorizontalLines reads each paragraph's OOXML for
 * <w:pBdr>…<w:bottom…/> instead.
 */
import { ClearScope } from "../shared/types";
import { safeWordRun } from "../shared/wordUtils";
import { clearCharacterFormatting } from "./font";
import { applyStyle } from "./styles";
import { removeAllIndents, removeAllSpacing, setLineSpacing } from "./paragraph";

export async function clearFormatting(scope: ClearScope): Promise<{ error?: string }> {
  if (scope === "character") {
    return clearCharacterFormatting();
  }
  if (scope === "paragraph") {
    const r1 = await removeAllIndents();
    if (r1.error) return r1;
    const r2 = await removeAllSpacing();
    if (r2.error) return r2;
    return setLineSpacing("single");
  }
  const r1 = await clearCharacterFormatting();
  if (r1.error) return r1;
  const r2 = await removeAllIndents();
  if (r2.error) return r2;
  const r3 = await removeAllSpacing();
  if (r3.error) return r3;
  const r4 = await setLineSpacing("single");
  if (r4.error) return r4;
  return applyStyle("Normal");
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
