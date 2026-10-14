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

/**
 * EVIDENCE_MODEL v1 — provisional weights for footnote pairing.
 * Scores map to review tiers (high ≥ 70, medium 40–69, low < 40); they are
 * NOT probabilities. Tune against labelled fixtures (TEST_REPORT matrix);
 * cue patterns live in citations.ts.
 */
export const EVIDENCE_MODEL = {
  version: "1.1",
  thresholds: { high: 70, medium: 40 },
  weights: {
    separatorRegion: 25, // Tier A/C: block in plausible note region (or no separator)
    belowSeparator: -25, // Tier A-negative: block above the separator line
    conflictDuplicate: -15, // Tier A-negative: ambiguity caps at medium
    citationDenseCap: 60, // Tier B-negative ceiling: bibliography-shaped
    followedByLargeCap: 60, // §4 structural ceiling (multi-block only)
    referenceSignalBoost: 15, // Tier B: citation signals in non-dense blocks
    sequenceContinuity: 10, // Tier A: marker run aligns with block run
    bibliographyContradiction: -30, // §3 hard negative
    genericContradiction: -15, // §3 other explicit counter-evidence
    gapSmall: -5, // gap of 1 in a run of 8+ markers: small modifier
    gapLargeCap: 69, // gap of 3+: whole run needs review
  },
  // Review policy: "needs review" = at least two supporting cues without Tier A
  // linkage. "Independent" always means evidence from a different tier —
  // two Tier B signals from one paragraph are one corroboration, not two.
  // High-confidence proposals still require per-match user approval.
  reviewPolicy: {
    needsReviewMinCues: 2,
    independentMeans: "different evidence tier",
  },
};

/**
 * Pure: mixed native+flattened disambiguation. Native footnote references are
 * superscript numbers too, so when genuine footnotes exist AND a 1..k leading
 * run is duplicated later (conflict signature), the leading run is the native
 * one — drop it, keep the flattened run. Without duplication evidence, keep
 * everything for human review.
 */
export function excludeNativeLeadingMarkers(
  markers: FootnoteMarker[],
  existingFootnotes: number
): FootnoteMarker[] {
  if (existingFootnotes <= 0) {
    return markers;
  }
  const byNumber = new Map<number, number[]>();
  markers.forEach((m, i) => {
    if (m.pattern === "superscript") {
      const list = byNumber.get(m.number) ?? [];
      list.push(i);
      byNumber.set(m.number, list);
    }
  });
  const drop = new Set<number>();
  for (let n = 1; n <= existingFootnotes; n++) {
    const idx = byNumber.get(n);
    if (idx !== undefined && idx.length >= 2) {
      drop.add(idx[0]);
    }
  }
  if (drop.size === 0) {
    return markers;
  }
  return markers.filter((_, i) => !drop.has(i));
}

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
  const usedBlocks = new Set<FootnoteBlock>();
  const matches: FootnoteMatch[] = [];
  let seq = 0;
  // Tier A sequence continuity: markers in document order vs block numbers.
  const orderedMarkers = [...markers].sort((a, b) => a.paragraphIndex - b.paragraphIndex);
  // Gap decay: a gap of 1 in a run of 8+ markers is a small modifier; a gap
  // of 3+ makes every assignment in the run needs-review regardless.
  const uniqueNumbers = [...new Set(orderedMarkers.map((m) => m.number))].sort((a, b) => a - b);
  let maxGap = 0;
  for (let i = 1; i < uniqueNumbers.length; i++) {
    maxGap = Math.max(maxGap, uniqueNumbers[i] - uniqueNumbers[i - 1] - 1);
  }
  const gapSmall = maxGap === 1 && uniqueNumbers.length >= 8;
  const gapLarge = maxGap >= 3;
  for (const marker of orderedMarkers) {
    const candidates = blocksByNumber.get(marker.number) ?? [];
    if (candidates.length === 0) {
      continue;
    }
    const occurrence = (markerOccurrence.get(marker.number) ?? 0) + 1;
    markerOccurrence.set(marker.number, occurrence);
    const first = candidates[0];
    const evidence: string[] = [];
    const contradictions: string[] = [...(first.contradictions ?? [])];

    let score = 50;
    if (!separatorFound || first.paragraphIndex > separatorIndex) {
      score += EVIDENCE_MODEL.weights.separatorRegion;
      evidence.push("same-number block in note region");
    } else {
      score += EVIDENCE_MODEL.weights.belowSeparator;
      contradictions.push("block sits above the separator line");
    }
    // Tier A: marker run aligns with an ascending block run.
    const markerRunOk = matches.length === 0 || marker.number >= matches[matches.length - 1].marker.number;
    if (markerRunOk && occurrence === 1) {
      score += EVIDENCE_MODEL.weights.sequenceContinuity;
      evidence.push("marker order continues block sequence");
    }
    let cap = 100;
    if (candidates.length > 1 || occurrence > 1) {
      cap = Math.min(cap, 69); // ambiguity caps at medium; never auto-pick
      contradictions.push(
        candidates.length > 1
          ? `${candidates.length} blocks share number ${marker.number}`
          : `repeat reference to note ${marker.number}`
      );
    }
    if (gapLarge) {
      cap = Math.min(cap, EVIDENCE_MODEL.weights.gapLargeCap);
      contradictions.push(`marker sequence has a gap of ${maxGap} — whole run needs review`);
    } else if (gapSmall) {
      score += EVIDENCE_MODEL.weights.gapSmall;
      evidence.push("single-number gap in a long run");
    }
    if (first.citationDense === true) {
      cap = Math.min(cap, EVIDENCE_MODEL.weights.citationDenseCap);
      contradictions.push("block is citation-list shaped (bibliography?)");
    }
    if (first.followedByLarge === true && blocks.length > 1) {
      cap = Math.min(cap, EVIDENCE_MODEL.weights.followedByLargeCap);
      contradictions.push("block followed by a large body paragraph");
    }
    if (
      first.referenceSignals === true &&
      first.citationDense !== true &&
      score < EVIDENCE_MODEL.thresholds.medium
    ) {
      score += EVIDENCE_MODEL.weights.referenceSignalBoost;
      evidence.push("block carries reference/citation signals");
    }
    for (const c of first.contradictions ?? []) {
      score += c.startsWith("bibliography:")
        ? EVIDENCE_MODEL.weights.bibliographyContradiction
        : EVIDENCE_MODEL.weights.genericContradiction;
    }
    score = Math.max(0, Math.min(cap, score));
    const confidence: MatchConfidence =
      score >= EVIDENCE_MODEL.thresholds.high
        ? "high"
        : score >= EVIDENCE_MODEL.thresholds.medium
          ? "medium"
          : "low";
    // Stress-test lesson: review-first means only high-confidence first
    // references start checked. One-to-one both directions: a block already
    // feeding an accepted match, and any repeat reference to the same number,
    // surface unchecked — the user opts in per row.
    const accepted =
      confidence === "high" && occurrence === 1 && !usedBlocks.has(first);
    if (accepted) {
      usedBlocks.add(first);
    }
    matches.push({
      id: `match-${marker.number}-${Date.now()}-${seq++}`,
      marker,
      block: first,
      confidence,
      accepted,
      pairingEvidence: evidence,
      contradictions,
    });
    // Conflicting alternatives: surfaced, never accepted by default.
    for (const alt of candidates.slice(1)) {
      matches.push({
        id: `match-${marker.number}-alt-${Date.now()}-${seq++}`,
        marker,
        block: alt,
        confidence: "low",
        accepted: false,
        pairingEvidence: ["same number as another block — pick one"],
        contradictions: [...contradictions, "duplicate block number"],
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
  existingFootnotes = 0,
  citationSystems: { system: string; hits: number }[] = []
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
    citationSystems,
  };
}
