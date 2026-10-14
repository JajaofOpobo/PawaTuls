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
 * Pure: resolve which paragraph holds a marker at conversion time. Prefers the
 * scan-time paragraph index, then ±5 neighbours, then a full sweep — but only
 * accepts paragraphs containing the literal AND whatever context survives
 * (up to 40 chars each side). Returns -1 when unverifiable.
 */
export function resolveMarkerParagraph(texts: string[], match: FootnoteMatch): number {
  const lit = literalFor(match);
  const before = match.marker.contextBefore.slice(-40);
  const after = match.marker.contextAfter.slice(0, 40);
  const fits = (t: string): boolean =>
    t.includes(lit) &&
    (before === "" || t.includes(before)) &&
    (after === "" || t.includes(after));
  const hint = match.marker.paragraphIndex;
  if (hint >= 0 && hint < texts.length && fits(texts[hint])) {
    return hint;
  }
  if (hint >= 0) {
    for (let d = 1; d <= 5; d++) {
      if (hint - d >= 0 && hint - d < texts.length && fits(texts[hint - d])) {
        return hint - d;
      }
      if (hint + d < texts.length && fits(texts[hint + d])) {
        return hint + d;
      }
    }
  }
  for (let i = 0; i < texts.length; i++) {
    if (fits(texts[i])) {
      return i;
    }
  }
  return -1;
}

/**
 * Pure helper shared by the converter and deleteFootnoteBlocks: locate block
 * paragraphs in a paragraph-text array. Matches by number prefix plus a content
 * probe (first 30 chars) so continuation-merged text still resolves. Includes
 * the full span (starter + continuations) so multi-paragraph blocks are fully
 * removed — never just the starter. Returns descending paragraph indices.
 */
export function findBlockParagraphIndices(texts: string[], blocks: FootnoteBlock[]): number[] {
  const starters = [/^\s*\[(\d+)\]/, /^\s*(\d+)[.)]/, /^\s*(\d+)\s{2,}/];
  const stripPrefix = /^\s*(?:\[\d+\]|\d+[.)]|\d+\s{2,})\s*/;
  const found: number[] = [];
  for (const block of blocks) {
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
      if (num !== block.number) {
        continue;
      }
      // The starter paragraph's own content must open the (possibly
      // continuation-merged) block text — never match on merged tail text.
      const content = t.replace(stripPrefix, "");
      if (content !== "" && !block.text.startsWith(content.substring(0, Math.min(20, content.length)))) {
        continue;
      }
      const span = Math.max(1, block.span ?? 1);
      for (let k = 0; k < span && i + k < texts.length; k++) {
        found.push(i + k);
      }
      break;
    }
  }
  return [...new Set(found)].sort((a, b) => b - a);
}

async function convertMatches(
  matches: FootnoteMatch[]
): Promise<{ data?: ConversionResult; error?: string }> {
  const errors: string[] = [];
  if (matches.length === 0) {
    return { data: { converted: 0, failed: 0, errors, insertions: [] } };
  }
  return safeWordRun(async (context) => {
    const body = context.document.body;

    // Relocate markers by paragraph-index hint + context verification — never
    // bare body-wide literal search (that hits ordinary numbers, legal
    // paragraph numbers, citation years, or block numbers before the marker).
    const scanParas = body.paragraphs;
    scanParas.load("text");
    const notes = body.footnotes;
    notes.load("items");
    await context.sync();
    const scanTexts = scanParas.items.map((p) => p.text);
    const beforeCount = notes.items.length;

    // Phase A — paragraph-scoped literal searches (batched), superscript
    // hits verified as superscript, ordinal consumption per paragraph so
    // repeats resolve in document order. No mutations yet.
    interface SearchEntry {
      match: FootnoteMatch;
      paraIndex: number;
      hits: Word.RangeCollection;
    }
    const searches: SearchEntry[] = [];
    const located: { match: FootnoteMatch; paraIndex: number }[] = [];
    let failed = 0;
    for (const match of matches) {
      const paraIndex = resolveMarkerParagraph(scanTexts, match);
      if (paraIndex < 0) {
        failed++;
        errors.push(
          `Marker ${match.marker.number} (${match.marker.pattern}): location unverifiable on re-scan; skipped with source intact.`
        );
        continue;
      }
      located.push({ match, paraIndex });
    }
    for (const entry of located) {
      const hits = scanParas.items[entry.paraIndex].search(literalFor(entry.match));
      hits.load("text, font/superscript");
      searches.push({ match: entry.match, hits: hits, paraIndex: entry.paraIndex });
    }
    await context.sync();

    const consumed = new Map<string, number>();
    const swaps: { match: FootnoteMatch; slot: number }[] = [];
    for (const entry of searches) {
      const lit = literalFor(entry.match);
      let candidates = entry.hits.items.filter((h) => h.text === lit);
      if (entry.match.marker.pattern === "superscript") {
        candidates = candidates.filter((h) => h.font.superscript === true);
      }
      const key = `${entry.paraIndex}:${entry.match.marker.pattern}:${entry.match.marker.number}`;
      const at = consumed.get(key) ?? 0;
      consumed.set(key, at + 1);
      if (at >= candidates.length) {
        failed++;
        errors.push(
          `Marker ${entry.match.marker.number} (${entry.match.marker.pattern}): not found in verified paragraph; skipped with source intact.`
        );
        continue;
      }
      const slot = swaps.length;
      candidates[at].insertText(placeholder(slot), Word.InsertLocation.replace);
      swaps.push({ match: entry.match, slot });
    }
    await context.sync();

    // Phase B — placeholders → genuine footnotes with per-swap rollback: any
    // swap that cannot convert gets its original marker literal restored, so
    // failed conversions never leave private placeholders behind.
    const rollback = async (pending: { match: FootnoteMatch; slot: number }[]): Promise<void> => {
      for (const p of pending) {
        const hits = body.search(placeholder(p.slot));
        hits.load("text");
        await context.sync();
        for (const h of hits.items) {
          h.insertText(literalFor(p.match), Word.InsertLocation.replace);
        }
      }
      await context.sync();
    };

    const inserted: { match: FootnoteMatch }[] = [];
    for (const swap of swaps) {
      const hits = body.search(placeholder(swap.slot));
      hits.load("text");
      await context.sync();
      if (hits.items.length !== 1) {
        await rollback([swap]);
        failed++;
        errors.push(
          `Marker ${swap.match.marker.number}: placeholder not uniquely re-located; marker restored, skipped.`
        );
        continue;
      }
      const target = hits.items[0];
      try {
        target.delete();
        target.insertFootnote(swap.match.block.text);
        await context.sync();
        inserted.push({ match: swap.match });
      } catch (err) {
        await rollback([swap]);
        failed++;
        errors.push(
          `Marker ${swap.match.marker.number}: conversion failed (${
            err instanceof Error ? err.message : String(err)
          }); marker restored, skipped.`
        );
      }
    }
    notes.load("items");
    await context.sync();
    const created = notes.items.length - beforeCount;
    if (created !== inserted.length) {
      errors.push(
        `Verification failed: expected ${inserted.length} new footnote(s), found ${created}. ` +
          `Block source paragraphs left intact; review the document before retrying.`
      );
      return { converted: 0, failed: failed + inserted.length, errors, insertions: [] };
    }

    // Phase C — delete original block paragraphs (re-located, descending).
    const delParas = body.paragraphs;
    delParas.load("text");
    await context.sync();
    const delTexts = delParas.items.map((p) => p.text);
    const indices = findBlockParagraphIndices(
      delTexts,
      inserted.map((x) => x.match.block)
    );
    for (const idx of indices) {
      delParas.items[idx].delete();
    }
    await context.sync();

    return {
      converted: inserted.length,
      failed,
      errors,
      insertions: inserted.map((x) => ({
        markerNumber: x.match.marker.number,
        pattern: x.match.marker.pattern,
        footnoteText: x.match.block.text,
        markerParagraphIndex: x.match.marker.paragraphIndex,
        blockParagraphIndex: x.match.block.paragraphIndex,
      })),
    };
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
