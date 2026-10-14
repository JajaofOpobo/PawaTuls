// src/shared/types.ts
// Central type definitions for PawaTuls.
// All three feature areas import from here. Do not define types elsewhere.

// ─── FORMATTING ENGINE ────────────────────────────────────────────────────────

export type UnderlineStyle =
  | "none" | "single" | "double" | "dotted"
  | "dashed" | "wave" | "thick" | "thickDotted" | "thickDashed";

export type LineSpacingRule =
  | "single" | "onePtFive" | "double"
  | "atLeast" | "exactly" | "multiple";

export type ClearScope = "character" | "paragraph" | "all";

export type CaseTransform =
  | "upper" | "lower" | "title" | "sentence" | "toggle";

export interface FontOptions {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  strikeThrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  color?: string;           // hex "#RRGGBB" or "auto"
  highlightColor?: string;  // Word English color name or "none" to clear
  allCaps?: boolean;
  smallCaps?: boolean;
  spacing?: number;         // character spacing in points
  hidden?: boolean;
}

export interface ParagraphOptions {
  alignment?: "left" | "centered" | "right" | "justified";
  firstLineIndent?: number;     // points
  leftIndent?: number;          // points
  rightIndent?: number;         // points
  lineSpacing?: number;         // points (used with atLeast, exactly, multiple)
  lineSpacingRule?: LineSpacingRule;
  spaceBefore?: number;         // points
  spaceAfter?: number;          // points
  keepWithNext?: boolean;
  keepLinesTogether?: boolean;
  pageBreakBefore?: boolean;
}

export interface StyleCreationOptions {
  name: string;
  styleType: "paragraph" | "character";
  basedOn?: string;
  font?: FontOptions;
  paragraph?: ParagraphOptions;
}

export interface BorderOptions {
  color?: string;   // hex
  width?: number;   // points
  type?: "single" | "double" | "dotted" | "dashed" | "none";
}

export interface FormattingProfile {
  schemaVersion: 1;
  id: string;            // stable, unique, e.g. "court-body-1"
  name: string;          // display name
  description?: string;
  style?: string;        // named style applied first (must exist)
  font?: FontOptions;
  paragraph?: ParagraphOptions;
}

// ─── LEGAL WORKFLOWS ─────────────────────────────────────────────────────────

export type NumberingScheme =
  | "decimal-decimal-decimal"    // 1. / 1.1 / 1.1.1
  | "decimal-alpha-roman";       // 1. / (a) / (i)

export type LawReporter =
  | "NWLR" | "LPELR" | "SCNJ" | "FWLR"
  | "NLR" | "ANLR" | "NSCC" | "UILR" | "JELR";

export type CourtName =
  | "Supreme Court of Nigeria"
  | "Court of Appeal"
  | "Federal High Court"
  | "High Court of the Federal Capital Territory"
  | "High Court of Justice"
  | "National Industrial Court";

export interface NigerianCaseCitation {
  year: string;
  volume?: string;
  reporter: LawReporter;
  part?: string;
  page: string;
}

export interface CauseHeadingData {
  court: CourtName | string;
  division?: string;
  suitNumber: string;
  claimant: string;
  defendant: string;
  actionType?: string;
}

export interface StatutoryCitation {
  section: string;
  actName: string;
  year: string;
}

// ─── FOOTNOTE DETECTION ──────────────────────────────────────────────────────

export type MarkerPattern = "superscript" | "bracketed" | "parenthetical";

export interface FootnoteMarker {
  number: number;
  pattern: MarkerPattern;
  paragraphIndex: number;
  contextBefore: string;   // up to 80 chars before the marker
  contextAfter: string;    // up to 80 chars after the marker
}

export interface FootnoteBlock {
  number: number;
  text: string;
  paragraphIndex: number;
  /** Paragraph span: 1 + continuation paragraphs merged at scan. Deletion
   * removes the whole span, not just the starter (multi-para fix). */
  span?: number;
  citationDense?: boolean;   // paragraph matches a citation system (Appendix J)
  referenceSignals?: boolean; // text carries reference/citation signals (App. J)
  followedByLarge?: boolean;  // next paragraph is much larger (Appendix J)
  // Appendix J evidence model (separate block assessment):
  blockLikelihood?: number;     // 0–100, provisional (not a probability)
  citationLikelihood?: number;  // 0–100, provisional
  contradictions?: string[];    // explicit counter-evidence
  boundaryUncertain?: boolean;  // start/end boundary needs user eyes
  mergedSuspect?: boolean;      // body prose possibly merged with note text
  mergedSplitAt?: number;       // proposed split offset (review only, never auto-split)
}

export type MatchConfidence = "high" | "medium" | "low";

export interface FootnoteMatch {
  id: string;
  marker: FootnoteMarker;
  block: FootnoteBlock;
  confidence: MatchConfidence;
  accepted: boolean;
  pairingEvidence: string[];   // Appendix J: why this pairing was proposed
  contradictions: string[];    // Appendix J: counter-evidence (explicit)
}

export interface ScanResult {
  matches: FootnoteMatch[];
  unmatchedMarkers: FootnoteMarker[];
  orphanedBlocks: FootnoteBlock[];
  separatorLineFound: boolean;
  totalMarkersFound: number;
  totalBlocksFound: number;
  scanDurationMs: number;
  existingFootnotes: number;   // genuine Word footnotes present (Appendix G)
  citationSystems: { system: string; hits: number }[];   // Appendix J
}

export interface ConversionResult {
  converted: number;
  failed: number;
  errors: string[];
  /** Insertion log: every footnote created + every source removed. The whole
   * conversion runs in one Word.run (single undo unit — Ctrl+Z reverts the
   * batch); this log is the manual-recovery record if granular undo is ever
   * needed (revert UI deferred by design). */
  insertions: {
    markerNumber: number;
    pattern: string;
    footnoteText: string;
    markerParagraphIndex: number;
    blockParagraphIndex: number;
  }[];
}

// ─── NUMBERING AUDIT (Appendix F, Phase 33) ───────────────────────────────────

export type NumberingKind = "native" | "manual" | "plain";

export interface NumberedParagraph {
  paragraphIndex: number;
  kind: NumberingKind;
  prefix?: string;    // typed prefix text for "manual"
  listId?: number;    // native list id when detectable
  level?: number;     // nesting level when detectable
}

export interface NumberingIssue {
  paragraphIndex: number;
  issue: "gap" | "duplicate" | "restart" | "mixed-style" | "orphan-level";
  detail: string;
}

export interface NumberingRepairPlan {
  edits: { paragraphIndex: number; before: string; after: string }[];
}
