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
import { detectCitationSystems } from "./citations";

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
    paras.load("text");
    await context.sync();

    const existingFootnotes = notes.items.length;
    const texts = paras.items.map((p) => p.text);
    const count = paras.items.length;

    // Appendix J: citation systems feed certainty (dense paragraphs are
    // citation lists/bibliographies, not footnote blocks) and consistency.
    const citation = detectCitationSystems(texts);
    const denseSet = new Set(citation.denseParagraphs);

    interface Job {
      index: number;
      pattern: MarkerPattern;
      results: Word.RangeCollection;
    }
    const jobs: Job[] = [];
    for (let i = 0; i < count; i++) {
      const para = paras.items[i];
      const sup = para.search("[0-9]@", { matchWildcards: true });
      sup.load("text, font/superscript");
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
    for (let i = 0; i < count; i++) {
      if (separatorFound && i <= separatorIndex) {
        continue;
      }
      const t = texts[i];
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
          paragraphIndex: i,
          citationDense: denseSet.has(i),
        });
      } else if (
        t.trim() !== "" &&
        blocks.length > 0 &&
        blocks[blocks.length - 1].paragraphIndex === i - 1
      ) {
        blocks[blocks.length - 1].text += ` ${t.trim()}`;
      }
    }

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
    return blocks;
  });
}
