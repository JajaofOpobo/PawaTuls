/**
 * citations.ts — citation-system identification (pure functions, no Office.js).
 *
 * Purpose: recognize the major standardized citation systems present in a
 * document so the footnote workflow can (a) avoid mistaking bibliographies and
 * law-report lists for footnote blocks (certainty), and (b) report mixed or
 * inconsistent usage (consistency).
 *
 * Systems covered: Nigerian Law Reports (NWLR, LPELR, SCNJ, FWLR, NLR, ANLR,
 * NSCC, UILR, JELR), statutory citations (Section X of the … Act Year),
 * neutral court citations ([Year] COURT Number), academic bracket references
 * ([12] list style), and author–date ((Smith 2020)).
 */
export type CitationSystem =
  | "Nigerian Law Reports"
  | "Statutes"
  | "Neutral Citations"
  | "Academic Bracket"
  | "Author-Date";

export interface CitationHit {
  system: CitationSystem;
  excerpt: string;
}

export interface CitationSummary {
  systems: { system: CitationSystem; hits: number }[];
  denseParagraphs: number[];
}

const PATTERNS: { system: CitationSystem; re: RegExp }[] = [
  {
    system: "Nigerian Law Reports",
    re: /\[\d{4}\]\s*(?:\d+\s+)?(?:NWLR|LPELR|SCNJ|FWLR|NLR|ANLR|NSCC|UILR|JELR)\b|\bLPELR-\d+\([A-Za-z]+\)/gi,
  },
  {
    system: "Statutes",
    re: /\bsection\s+\d[\w()]*\s+of\s+(?:the\s+)?[\w\s&'’-]+?\s*act(?:\s*,?\s*\d{4})?/gi,
  },
  {
    system: "Neutral Citations",
    re: /\[\d{4}\]\s+(?:UKSC|UKPC|EWCA|EWHC|SC|CA|FCA|HCA|ECHR)\s+\d+/g,
  },
  {
    system: "Academic Bracket",
    re: /\[\d{1,3}\]\s*\S/g,
  },
  {
    system: "Author-Date",
    re: /\([A-Z][A-Za-z'’\-]+(?:\s+(?:et\s+al\.?|[A-Z][a-z]+))?,?\s+\d{4}[a-z]?\)/g,
  },
];

export function classifyParagraph(text: string): CitationHit[] {
  const hits: CitationHit[] = [];
  for (const { system, re } of PATTERNS) {
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    // Cap at 5 hits per system per paragraph — presence, not census.
    let guard = 0;
    while ((m = re.exec(text)) !== null && guard < 5) {
      guard++;
      hits.push({ system, excerpt: m[0].substring(0, 80) });
      if (m[0].length === 0) {
        re.lastIndex++;
      }
    }
  }
  return hits;
}

export function detectCitationSystems(texts: string[]): CitationSummary {
  const counts = new Map<CitationSystem, number>();
  const dense: number[] = [];
  texts.forEach((text, i) => {
    const hits = classifyParagraph(text);
    if (hits.length > 0) {
      dense.push(i);
      for (const hit of hits) {
        counts.set(hit.system, (counts.get(hit.system) ?? 0) + 1);
      }
    }
  });
  return {
    systems: [...counts.entries()].map(([system, hits]) => ({ system, hits })),
    denseParagraphs: dense,
  };
}
