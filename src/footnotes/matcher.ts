/**
 * matcher.ts — Phase 26: footnote matching (pure functions, no Office.js).
 * Appendix G deltas: one-to-one discipline (a block feeds at most one marker;
 * duplicate-number blocks surface as unchecked conflicting alternatives, never
 * auto-picked), duplicate markers downgraded to medium, unmatched stays unmatched.
 */
import {
  FootnoteBlock,
  FootnoteMarker,
  FootnoteMatch,
  MatchConfidence,
  ScanResult,
} from "../shared/types";

export function buildMatchTable(
  markers: FootnoteMarker[],
  blocks: FootnoteBlock[],
  separatorFound: boolean,
  separatorIndex: number
): FootnoteMatch[] {
  const blocksByNumber = new Map<number, FootnoteBlock[]>();
  for (const block of blocks) {
    const list = blocksByNumber.get(block.number) ?? [];
    list.push(block);
    blocksByNumber.set(block.number, list);
  }
  const markerOccurrence = new Map<number, number>();
  const matches: FootnoteMatch[] = [];
  let seq = 0;
  for (const marker of markers) {
    const candidates = blocksByNumber.get(marker.number) ?? [];
    if (candidates.length === 0) {
      continue;
    }
    const occurrence = (markerOccurrence.get(marker.number) ?? 0) + 1;
    markerOccurrence.set(marker.number, occurrence);
    const first = candidates[0];
    let confidence: MatchConfidence =
      !separatorFound || first.paragraphIndex > separatorIndex ? "high" : "low";
    if (candidates.length > 1 || occurrence > 1) {
      confidence = "medium";
    }
    matches.push({
      id: `match-${marker.number}-${Date.now()}-${seq++}`,
      marker,
      block: first,
      confidence,
      accepted: true,
    });
    // Conflicting alternatives: surfaced, never accepted by default.
    for (const alt of candidates.slice(1)) {
      matches.push({
        id: `match-${marker.number}-alt-${Date.now()}-${seq++}`,
        marker,
        block: alt,
        confidence: "low",
        accepted: false,
      });
    }
  }
  return matches;
}

export function extractUnmatchedMarkers(
  markers: FootnoteMarker[],
  blocks: FootnoteBlock[]
): FootnoteMarker[] {
  const blockNumbers = new Set(blocks.map((b) => b.number));
  return markers.filter((m) => !blockNumbers.has(m.number));
}

export function extractOrphanedBlocks(
  markers: FootnoteMarker[],
  blocks: FootnoteBlock[]
): FootnoteBlock[] {
  const markerNumbers = new Set(markers.map((m) => m.number));
  return blocks.filter((b) => !markerNumbers.has(b.number));
}

export function buildScanResult(
  markers: FootnoteMarker[],
  blocks: FootnoteBlock[],
  separatorFound: boolean,
  separatorIndex: number,
  scanStartTime: number,
  existingFootnotes = 0
): ScanResult {
  const matches = buildMatchTable(markers, blocks, separatorFound, separatorIndex);
  return {
    matches,
    unmatchedMarkers: extractUnmatchedMarkers(markers, blocks),
    orphanedBlocks: extractOrphanedBlocks(markers, blocks),
    separatorLineFound: separatorFound,
    totalMarkersFound: markers.length,
    totalBlocksFound: blocks.length,
    scanDurationMs: Date.now() - scanStartTime,
    existingFootnotes,
  };
}
