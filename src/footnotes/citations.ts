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
  | "Author-Date"
  | "OSCOLA"
  | "Bluebook"
  | "Chicago"
  | "MLA"
  | "Ibid Chain";

// Note: IEEE / Vancouver / GOST-numeric references share the bracketed-number
// shape and are covered by "Academic Bracket" presence detection. Ibid-style
// chains are note CONTENT (a lone "Ibid." is usually a genuine note), so they
// feed certainty signals but never density (see detectReferenceSignals).

export interface CitationHit {
  system: CitationSystem;
  excerpt: string;
}

export interface CitationSummary {
  systems: { system: CitationSystem; hits: number }[];
  denseParagraphs: number[];
  /** Paragraphs matching list-shaped systems (excludes Ibid Chain, which is
   * note content, not bibliography shape). Drives the citationDense flag. */
  listDense: number[];
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
  {
    system: "OSCOLA",
    re: /\b[A-Z][a-z]+,\s*[‘'][^’']{3,}[’']\s*\(\d{4}\)|\([12]\d{3}\)\s*\d*\s*[A-Z]{2,}[A-Za-z&]*(\s+\d+)?/g,
  },
  {
    system: "Bluebook",
    re: /\b\d+\s+(?:U\.S\.|F\.2d|F\.3d|F\.\s*Supp\.?(?:\s*[23]d)?|S\.\s*Ct\.|L\.\s*Ed\.?(?:\s*2d)?|U\.S\.C\.?|N\.W\.2?d?|S\.E\.2?d?|A\.2?d?|P\.2?d?|N\.E\.2?d?|S\.W\.3?d?|So\.\s*3?d?)\s+\d+|\b[A-Z][\w.]*\s+v\.\s+[A-Z][\w.]*/g,
  },
  {
    system: "Chicago",
    re: /\b[A-Z][a-z]+,\s+[A-Z][a-z]+\.\s+[^.]+\.\s+[^:()]+:\s*[^,()]+,\s*\d{4}\./g,
  },
  {
    system: "MLA",
    re: /"[^"]{5,}"\s*,?\s*[A-Z][^,]{2,},\s*(?:vol\.\s*)?\d+/g,
  },
  {
    system: "Ibid Chain",
    re: /\b(ibid\.?|ibidem|id\.|idem|op\.\s*cit\.|loc\.\s*cit\.|supra note|infra note|supra|infra|cf\.|see also)\b/gi,
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
  const listDense: number[] = [];
  texts.forEach((text, i) => {
    const hits = classifyParagraph(text);
    if (hits.length > 0) {
      dense.push(i);
      if (hits.some((h) => h.system !== "Ibid Chain")) {
        listDense.push(i);
      }
      for (const hit of hits) {
        counts.set(hit.system, (counts.get(hit.system) ?? 0) + 1);
      }
    }
  });
  return {
    systems: [...counts.entries()].map(([system, hits]) => ({ system, hits })),
    denseParagraphs: dense,
    listDense,
  };
}

/**
 * Reference-content signals inside a candidate block: years (e.g. 1919),
 * "et al.", ibid-chains, and any recognized citation system. Presence
 * strengthens certainty (a note discussing authorities looks like a note);
 * absence proves nothing. Pure.
 */
export interface ReferenceSignals {
  hasYear: boolean;
  hasEtAl: boolean;
  hasIbid: boolean;
  systems: CitationSystem[];
  score: number;
}

export function detectReferenceSignals(text: string): ReferenceSignals {
  const systems = [...new Set(classifyParagraph(text).map((h) => h.system))];
  const hasYear = /(1[5-9]\d{2}|20\d{2})/.test(text);
  const hasEtAl = /\bet\s+al\.?/i.test(text);
  const hasIbid = /\b(ibid\.?|ibidem|id\.|idem|op\.\s*cit\.|loc\.\s*cit\.|supra|infra|cf\.|see also)\b/i.test(text);
  return {
    hasYear,
    hasEtAl,
    hasIbid,
    systems,
    score: systems.length + (hasYear ? 1 : 0) + (hasEtAl ? 1 : 0) + (hasIbid ? 1 : 0),
  };
}

/**
 * Academic shorthand cues (Tier B/C): et al., vol./no./pp./p./ed./trans./ch./
 * para./DOI/URL/access-date. Returns the distinct cues found. Pure.
 */
export function detectAcademicShorthand(text: string): string[] {
  const cues: string[] = [];
  const tests: [string, RegExp][] = [
    ["et al.", /\bet\s+al\.?/i],
    ["vol.", /\bvols?\.?\s+\d+/i],
    ["no.", /\bnos?\.?\s+\d+/i],
    ["pp.", /\bpp?\.\s*\d+\s*[-–]\s*\d+/],
    ["ed.", /\b(trans\.|eds?\.?)\b/i],
    ["ch.", /\bch\.\s*\d+/i],
    ["para.", /\bparas?\.?\s*\d+/i],
    ["DOI", /\bdoi\s*:\s*\S+|\b10\.\d{4,}\/\S+/i],
    ["URL", /https?:\/\/\S+|www\.\S+/i],
    ["accessed", /\baccessed\s+\d/i],
  ];
  for (const [cue, re] of tests) {
    if (re.test(text)) {
      cues.push(cue);
    }
  }
  return cues;
}
