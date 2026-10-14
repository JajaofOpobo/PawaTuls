/**
 * converter.ts — Phase 27: footnote converter.
 *
 * Critical rules (plan + AUDIT.md + Appendix G):
 * - Never use stored Word.Range objects: every marker and block is re-located
 *   by fresh search at conversion time (positions may have shifted since scan).
 * - Descending paragraphIndex order for all deletions.
 * - Placeholder swap: markers are first replaced with unique private-use
 *   placeholders, so newly inserted (superscript, numbered) footnote references
 *   can never be mistaken for remaining markers during relocation.
 * - Verify-before-delete: footnote-count delta is checked before any block
 *   paragraph is deleted; on mismatch sources stay intact and errors report.
 * - Per-match try/catch into errors[]; one bad match never aborts the run.
 * - Idempotent: re-running finds no accepted markers left and converts 0.
 */
import { ConversionResult, FootnoteBlock, FootnoteMatch, ScanResult } from "../shared/types";
import { safeWordRun } from "../shared/wordUtils";

const PH_START = "";
const PH_END = "";

function placeholder(i: number): string {
  return `${PH_START}${i}${PH_END}`;
}

function literalFor(match: FootnoteMatch): string {
  const n = match.marker.number;
  switch (match.marker.pattern) {
    case "bracketed":
      return `[${n}]`;
    case "parenthetical":
      return `(${n})`;
    case "superscript":
      return String(n);
  }
}

/**
 * Pure helper shared by the converter and deleteFootnoteBlocks: locate block
 * paragraphs in a paragraph-text array. Matches by number prefix plus a content
 * probe (first 30 chars) so continuation-merged text still resolves.
 * Returns descending paragraph indices.
 */
export function findBlockParagraphIndices(texts: string[], blocks: FootnoteBlock[]): number[] {
  const starters = [/^\s*\[(\d+)\]/, /^\s*(\d+)[.)]/, /^\s*(\d+)\s{2,}/];
  const found: number[] = [];
  for (const block of blocks) {
    const probe = block.text.substring(0, Math.min(30, block.text.length));
    for (let i = 0; i < texts.length; i++) {
      const t = texts[i];
      let num = -1;
      for (const re of starters) {
        const m = t.match(re);
        if (m) {
          num = parseInt(m[1], 10);
          break;
        }
      }
      if (num === block.number && (probe === "" || t.includes(probe))) {
        found.push(i);
        break;
      }
    }
  }
  return [...new Set(found)].sort((a, b) => b - a);
}

async function convertMatches(
  matches: FootnoteMatch[]
): Promise<{ data?: ConversionResult; error?: string }> {
  const errors: string[] = [];
  if (matches.length === 0) {
    return { data: { converted: 0, failed: 0, errors } };
  }
  return safeWordRun(async (context) => {
    const body = context.document.body;

    // Phase A — locate each accepted marker by fresh literal search and swap
    // in a unique placeholder. All searches are queued first and resolved with
    // one sync (no mutations yet, so ranges stay valid); ordinal consumption
    // per (pattern, number) keeps document order stable. No inserts happen yet
    // so later searches are never polluted by new footnote references.
    const searches: { match: FootnoteMatch; hits: Word.RangeCollection }[] = [];
    for (const match of matches) {
      const hits = body.search(literalFor(match));
      hits.load("text");
      searches.push({ match, hits });
    }
    await context.sync();

    const consumed = new Map<string, number>();
    const swaps: { match: FootnoteMatch; slot: number }[] = [];
    let failed = 0;
    for (const entry of searches) {
      const key = `${entry.match.marker.pattern}:${entry.match.marker.number}`;
      const at = consumed.get(key) ?? 0;
      consumed.set(key, at + 1);
      const candidates = entry.hits.items.filter((h) => h.text === literalFor(entry.match));
      if (at >= candidates.length) {
        failed++;
        errors.push(
          `Marker ${entry.match.marker.number} (${entry.match.marker.pattern}): not found on re-scan; skipped.`
        );
        continue;
      }
      const slot = swaps.length;
      candidates[at].insertText(placeholder(slot), Word.InsertLocation.replace);
      swaps.push({ match: entry.match, slot });
    }
    await context.sync();

    // Phase B — replace placeholders with genuine footnotes, then verify.
    const notes = body.footnotes;
    notes.load("items");
    await context.sync();
    const beforeCount = notes.items.length;

    const inserted: { match: FootnoteMatch }[] = [];
    const phSearches: { swap: { match: FootnoteMatch; slot: number }; hits: Word.RangeCollection }[] = [];
    for (const swap of swaps) {
      const hits = body.search(placeholder(swap.slot));
      hits.load("text");
      phSearches.push({ swap, hits });
    }
    await context.sync();
    for (const entry of phSearches) {
      if (entry.hits.items.length !== 1) {
        failed++;
        errors.push(
          `Marker ${entry.swap.match.marker.number}: placeholder not uniquely re-located; skipped.`
        );
        continue;
      }
      const target = entry.hits.items[0];
      target.delete();
      target.insertFootnote(entry.swap.match.block.text);
      inserted.push({ match: entry.swap.match });
    }
    await context.sync();
    notes.load("items");
    await context.sync();
    const created = notes.items.length - beforeCount;
    if (created !== inserted.length) {
      errors.push(
        `Verification failed: expected ${inserted.length} new footnote(s), found ${created}. ` +
          `Block source paragraphs left intact.`
      );
      return { converted: 0, failed: failed + inserted.length, errors };
    }

    // Phase C — delete original block paragraphs (re-located, descending).
    const paras = body.paragraphs;
    paras.load("text");
    await context.sync();
    const texts = paras.items.map((p) => p.text);
    const indices = findBlockParagraphIndices(
      texts,
      inserted.map((x) => x.match.block)
    );
    for (const idx of indices) {
      paras.items[idx].delete();
    }
    await context.sync();

    return { converted: inserted.length, failed, errors };
  });
}

export async function convertAcceptedMatches(
  scanResult: ScanResult
): Promise<{ data?: ConversionResult; error?: string }> {
  const accepted = [...scanResult.matches]
    .filter((m) => m.accepted)
    .sort((a, b) => b.block.paragraphIndex - a.block.paragraphIndex);
  return convertMatches(accepted);
}

export async function convertSingleMatch(match: FootnoteMatch): Promise<{ error?: string }> {
  const result = await convertMatches([{ ...match, accepted: true }]);
  if (result.error) {
    return { error: result.error };
  }
  const data = result.data;
  if (data && (data.failed > 0 || data.errors.length > 0)) {
    return { error: data.errors.join("; ") };
  }
  return {};
}

export async function deleteFootnoteBlocks(
  blocks: FootnoteBlock[]
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load("text");
    await context.sync();
    const texts = paras.items.map((p) => p.text);
    const indices = findBlockParagraphIndices(texts, blocks);
    for (const idx of indices) {
      paras.items[idx].delete();
    }
    await context.sync();
    return indices.length;
  });
}
