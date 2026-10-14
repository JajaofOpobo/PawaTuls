/**
 * scanner.ts — Phase 25: footnote detection scanner (read-only).
 * Returns raw markers + blocks for the matcher. Never mutates the document.
 *
 * Appendix G delta: existing genuine footnotes are counted (body paragraphs
 * exclude note stories, so no paragraph exclusion is needed) and reported.
 *
 * API notes (verified against @types/office-js):
 * - Search option is `matchWildcards` (the plan's `useWildcards` does not exist).
 * - All marker searches run per-paragraph so paragraphIndex is exact; every
 *   search is queued first and resolved with a single sync.
 * - Multi-digit superscript numbers may arrive as single-digit hits; consecutive
 *   single-digit hits in one paragraph are merged only when the joined text is
 *   verified adjacent in the paragraph text (else kept separate). Review-first
 *   design + matcher confidence absorb residual error.
 */
import { FootnoteBlock, FootnoteMarker, MarkerPattern } from "../shared/types";
import { safeWordRun } from "../shared/wordUtils";
import { detectAcademicShorthand, detectCitationSystems, detectReferenceSignals } from "./citations";

export interface CitationSystemCount {
  system: string;
  hits: number;
}

export interface ScanData {
  markers: FootnoteMarker[];
  blocks: FootnoteBlock[];
  separatorFound: boolean;
  separatorIndex: number;
  scanStartTime: number;
  existingFootnotes: number;
  citationSystems: CitationSystemCount[];
}

export interface ScanData {
  markers: FootnoteMarker[];
  blocks: FootnoteBlock[];
  separatorFound: boolean;
  separatorIndex: number;
  scanStartTime: number;
  existingFootnotes: number;
  citationSystems: CitationSystemCount[];
}

const CONTEXT_CHARS = 80;

function pushMarker(
  markers: FootnoteMarker[],
  texts: string[],
  cursors: Map<string, number>,
  paragraphIndex: number,
  number: number,
  pattern: MarkerPattern,
  locatedText: string,
  spanLength: number
): void {
  const text = texts[paragraphIndex];
  const key = `${paragraphIndex}:${pattern}`;
  const cursor = cursors.get(key) ?? 0;
  let idx = text.indexOf(locatedText, cursor);
  if (idx < 0) {
    idx = text.indexOf(locatedText);
  }
  if (idx < 0) {
    return;
  }
  cursors.set(key, idx + spanLength);
  markers.push({
    number,
    pattern,
    paragraphIndex,
    contextBefore: text.substring(Math.max(0, idx - CONTEXT_CHARS), idx),
    contextAfter: text.substring(idx + spanLength, idx + spanLength + CONTEXT_CHARS),
  });
}

export async function scanDocument(): Promise<{ data?: ScanData; error?: string }> {
  const scanStartTime = Date.now();
  return safeWordRun(async (context) => {
    const body = context.document.body;

    const notes = body.footnotes;
    notes.load("items");
    const paras = body.paragraphs;
    paras.load("text, style, firstLineIndent, leftIndent");
    await context.sync();

    const existingFootnotes = notes.items.length;
    const texts = paras.items.map((p) => p.text);
    const count = paras.items.length;

    // Stress-test lesson: body content (numbered headings, native-list items,
    // table cells) must never become footnote blocks. Batch the membership
    // checks with the marker searches below. Font sizes feed the Tier C
    // formatting signal (notes are often set smaller than body text).
    const listItems = paras.items.map((p) => {
      const li = p.listItemOrNullObject;
      context.load(li);
      return li;
    });
    const parentTables = paras.items.map((p) => {
      const t = p.parentTableOrNullObject;
      context.load(t);
      return t;
    });
    const fonts = paras.items.map((p) => {
      const f = p.font;
      f.load("size");
      return f;
    });

    // Appendix J: citation systems feed certainty (dense paragraphs are
    // citation lists/bibliographies, not footnote blocks) and consistency.
    const citation = detectCitationSystems(texts);
    const denseSet = new Set(citation.listDense);

    interface Job {
      index: number;
      pattern: MarkerPattern;
      results: Word.RangeCollection;
    }
    const jobs: Job[] = [];
    for (let i = 0; i < count; i++) {
      const para = paras.items[i];
      const sup = para.search("[0-9]@", { matchWildcards: true });      sup.load("text, font/superscript");
      jobs.push({ index: i, pattern: "superscript", results: sup });
      const br = para.search("\\[[0-9]@\\]", { matchWildcards: true });
      br.load("text");
      jobs.push({ index: i, pattern: "bracketed", results: br });
      const par = para.search("\\([0-9]@\\)", { matchWildcards: true });
      par.load("text");
      jobs.push({ index: i, pattern: "parenthetical", results: par });
    }
    await context.sync();

    const markers: FootnoteMarker[] = [];
    const cursors = new Map<string, number>();

    for (const job of jobs) {
      if (job.pattern === "superscript") {
        let run: Word.Range[] = [];
        const flush = (): void => {
          if (run.length === 0) {
            return;
          }
          const numText = run.map((r) => r.text).join("");
          const first = run[0].text;
          const text = texts[job.index];
          const key = `${job.index}:superscript`;
          const cursor = cursors.get(key) ?? 0;
          let idx = text.indexOf(first, cursor);
          if (idx < 0) {
            idx = text.indexOf(first);
          }
          if (idx >= 0 && text.substring(idx, idx + numText.length) === numText) {
            cursors.set(key, idx + numText.length);
            markers.push({
              number: parseInt(numText, 10),
              pattern: "superscript",
              paragraphIndex: job.index,
              contextBefore: text.substring(Math.max(0, idx - CONTEXT_CHARS), idx),
              contextAfter: text.substring(
                idx + numText.length,
                idx + numText.length + CONTEXT_CHARS
              ),
            });
          } else {
            // Not verifiably adjacent — keep hits as separate markers.
            for (const r of run) {
              pushMarker(markers, texts, cursors, job.index, parseInt(r.text, 10), "superscript", r.text, r.text.length);
            }
          }
          run = [];
        };
        for (const hit of job.results.items) {
          if (hit.font.superscript === true && /^\d+$/.test(hit.text)) {
            if (hit.text.length === 1) {
              run.push(hit);
            } else {
              flush();
              run = [hit];
              flush();
            }
          } else {
            flush();
          }
        }
        flush();
      } else {
        for (const hit of job.results.items) {
          const digits = hit.text.replace(/^\D+/, "").replace(/\D+$/, "");
          if (!/^\d+$/.test(digits)) {
            continue;
          }
          pushMarker(
            markers, texts, cursors, job.index, parseInt(digits, 10),
            job.pattern, hit.text, hit.text.length
          );
        }
      }
    }

    let separatorFound = false;
    let separatorIndex = -1;
    for (let i = 0; i < count; i++) {
      if (/^[-_]{3,}$/.test(texts[i].trim())) {
        separatorFound = true;
        separatorIndex = i;
        break;
      }
    }

    const blocks: FootnoteBlock[] = [];
    const starters = [/^\s*\[(\d+)\]\s+(.+)$/, /^\s*(\d+)[.)]\s+(.+)$/, /^\s*(\d+)\s{2,}(.+)$/];
    // Hygiene gate: headings, native-list items, and table cells are body
    // content, never footnote blocks. Non-hygienic paragraphs also break
    // continuation chains so headings cannot merge into a previous block.
    const hygienic = (i: number): boolean => {
      if (String(paras.items[i].style).startsWith("Heading")) {
        return false;
      }
      if (!listItems[i].isNullObject) {
        return false;
      }
      if (!parentTables[i].isNullObject) {
        return false;
      }
      return true;
    };
    // §4 boundary rule: a long, unmarked, citation-free paragraph after a
    // multi-note cluster usually ENDS the cluster — do not merge it. Length
    // is one input among several, never the sole decision; long genuine
    // notes are still allowed as numbered starts.
    const isLongProse = (t: string): boolean =>
      t.length > 400 && detectReferenceSignals(t).score === 0;
    let lastBlockIdx = -2;
    let numberedStarts = 0;
    for (let i = 0; i < count; i++) {
      if (separatorFound && i <= separatorIndex) {
        continue;
      }
      const t = texts[i];
      let m: RegExpMatchArray | null = null;
      if (hygienic(i)) {
        for (const re of starters) {
          m = t.match(re);
          if (m) {
            break;
          }
        }
      }
      if (m) {
        const blockText = m[2].trim();
        blocks.push({
          number: parseInt(m[1], 10),
          text: blockText,
          paragraphIndex: i,
          citationDense: denseSet.has(i),
          referenceSignals: detectReferenceSignals(blockText).score > 0,
        });
        lastBlockIdx = i;
        numberedStarts++;
      } else if (
        t.trim() !== "" &&
        lastBlockIdx === i - 1 &&
        hygienic(i) &&
        !(numberedStarts >= 2 && isLongProse(t.trim()))
      ) {
        blocks[blocks.length - 1].text += ` ${t.trim()}`;
        lastBlockIdx = i;
      }
    }

    // Structural heuristic: a footnote is rarely followed by a large body
    // paragraph — except on single-footnote pages, where the rule is skipped.
    // Threshold documented: next paragraph over 400 chars.
    if (blocks.length > 1) {
      for (const block of blocks) {
        if (block.paragraphIndex < 0) {
          continue;
        }
        const next = texts[block.paragraphIndex + 1];
        block.followedByLarge = next !== undefined && next.length > 400;
      }
    }
    // Recompute signals on final (continuation-merged) text.
    for (const block of blocks) {
      block.referenceSignals = detectReferenceSignals(block.text).score > 0;
    }

    // Appendix J §1/§6: separate block assessment. Provisional weights
    // (EVIDENCE_MODEL v1, see matcher.ts); scores are review tiers, not
    // probabilities. Formatting signals: notes often run smaller than body
    // text or carry hanging indents.
    const sizes = fonts
      .map((f) => f.size)
      .filter((s): s is number => typeof s === "number");
    const medianSize =
      sizes.length > 0 ? [...sizes].sort((a, b) => a - b)[Math.floor(sizes.length / 2)] : 0;
    const headings: { index: number; text: string }[] = [];
    for (let i = 0; i < count; i++) {
      if (String(paras.items[i].style).startsWith("Heading")) {
        headings.push({ index: i, text: texts[i] });
      }
    }
    const biblioHeadingAbove = (paraIndex: number): string | null => {
      for (const h of headings) {
        if (
          h.index < paraIndex &&
          paraIndex - h.index <= 3 &&
          /references|bibliography|works cited/i.test(h.text)
        ) {
          return h.text.trim().substring(0, 60);
        }
      }
      return null;
    };
    blocks.forEach((block, bi) => {
      const sig = detectReferenceSignals(block.text);
      const shorthand = detectAcademicShorthand(block.text);
      let score = 50;
      const contradictions: string[] = [];
      const prev = bi > 0 ? blocks[bi - 1] : null;
      if (prev !== null && prev.number === block.number - 1) {
        score += 20; // Tier A: sequential run membership
      }
      if (separatorFound && block.paragraphIndex > separatorIndex) {
        score += 10; // Tier C: plausible note region
      }
      if (sig.score >= 2) {
        score += 15; // Tier B: citation combinations (never sufficient alone)
      } else if (sig.score === 1) {
        score += 5;
      }
      if (sig.hasIbid) {
        score += 10; // Tier B: short-form reference chain
      }
      const size = block.paragraphIndex >= 0 ? fonts[block.paragraphIndex].size : null;
      if (typeof size === "number" && medianSize > 0 && size < medianSize - 0.5) {
        score += 5; // Tier C: smaller than body text
      }
      if (block.paragraphIndex >= 0) {
        const para = paras.items[block.paragraphIndex];
        if (para.firstLineIndent < 0 && para.leftIndent > 0) {
          score += 5; // Tier C: hanging indent
        }
      }
      const biblio = block.paragraphIndex >= 0 ? biblioHeadingAbove(block.paragraphIndex) : null;
      if (biblio !== null) {
        contradictions.push(`bibliography: under "${biblio}" heading`);
      }
      block.blockLikelihood = Math.max(0, Math.min(100, score));
      block.citationLikelihood = Math.min(
        100,
        sig.systems.length * 25 + shorthand.length * 10 + (sig.hasYear ? 5 : 0) + (sig.hasIbid ? 10 : 0)
      );
      block.contradictions = contradictions;
      // Long explanatory notes are allowed; flag possible prose-merge for review.
      block.mergedSuspect = block.text.length > 1000;
      block.boundaryUncertain = block.followedByLarge === true;
    });

    return {
      markers,
      blocks,
      separatorFound,
      separatorIndex,
      scanStartTime,
      existingFootnotes,
      citationSystems: citation.systems,
    };
  });
}

/**
 * Manual block-area override (Appendix J): when auto-detection confidence is
 * low, the user highlights the footnote block region and this parses the
 * selected paragraphs as blocks (no separator needed). paragraphIndex is -1
 * (unknown in body order) — matching is by number, conversion re-locates by
 * text, so the flow still holds; confidence stays review-grade by design.
 */
export async function scanSelectionBlocks(): Promise<{
  data?: FootnoteBlock[];
  error?: string;
}> {
  return safeWordRun(async (context) => {
    const paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load("text");
    await context.sync();
    if (paragraphs.items.length === 0) {
      throw new Error("No text selected. Select the footnote block first.");
    }
    const blocks: FootnoteBlock[] = [];
    const starters = [/^\s*\[(\d+)\]\s+(.+)$/, /^\s*(\d+)[.)]\s+(.+)$/, /^\s*(\d+)\s{2,}(.+)$/];
    let prevContinues = false;
    paragraphs.items.forEach((para) => {
      const t = para.text;
      let m: RegExpMatchArray | null = null;
      for (const re of starters) {
        m = t.match(re);
        if (m) {
          break;
        }
      }
      if (m) {
        blocks.push({
          number: parseInt(m[1], 10),
          text: m[2].trim(),
          paragraphIndex: -1,
          citationDense: false,
        });
        prevContinues = true;
      } else if (t.trim() !== "" && blocks.length > 0 && prevContinues) {
        blocks[blocks.length - 1].text += ` ${t.trim()}`;
        prevContinues = true;
      } else {
        prevContinues = false;
      }
    });
    // Manual area narrows the domain but certifies nothing: re-score the
    // selection, propose the last block as the uncertain boundary, and keep
    // review-grade confidence (matcher caps apply normally).
    blocks.forEach((block, bi) => {
      const sig = detectReferenceSignals(block.text);
      block.referenceSignals = sig.score > 0;
      block.blockLikelihood = 50 + Math.min(20, sig.score * 5);
      block.citationLikelihood = Math.min(100, sig.systems.length * 25);
      block.contradictions = [];
      block.mergedSuspect = block.text.length > 1000;
      block.boundaryUncertain = bi === blocks.length - 1;
    });
    if (blocks.length > 0) {
      blocks[blocks.length - 1].contradictions = [
        ...(blocks[blocks.length - 1].contradictions ?? []),
        "manual area: end boundary unverified — adjust selection if the last note runs on",
      ];
    }
    return blocks;
  });
}
