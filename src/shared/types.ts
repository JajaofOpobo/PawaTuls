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
}

export type MatchConfidence = "high" | "medium" | "low";

export interface FootnoteMatch {
  id: string;
  marker: FootnoteMarker;
  block: FootnoteBlock;
  confidence: MatchConfidence;
  accepted: boolean;
}

export interface ScanResult {
  matches: FootnoteMatch[];
  unmatchedMarkers: FootnoteMarker[];
  orphanedBlocks: FootnoteBlock[];
  separatorLineFound: boolean;
  totalMarkersFound: number;
  totalBlocksFound: number;
  scanDurationMs: number;
}

export interface ConversionResult {
  converted: number;
  failed: number;
  errors: string[];
}
