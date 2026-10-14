import React, { useState } from "react";
import {
  Badge,
  Button,
  Checkbox,
  Spinner,
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
  Toast,
  ToastTitle,
  Toaster,
  useToastController,
} from "@fluentui/react-components";
import { scanDocument, scanSelectionBlocks } from "../../footnotes/scanner";
import { buildScanResult } from "../../footnotes/matcher";
import { convertAcceptedMatches } from "../../footnotes/converter";
import {
  convertEndnoteToText,
  removeEndnotes,
  removeFootnotes,
} from "../../footnotes/notes";
import { ConversionResult, FootnoteMatch, ScanResult } from "../../shared/types";

export interface FootnotePanelProps {
  setError: (e: string | null) => void;
}

const TOASTER_ID = "footnote-toaster";

function confidenceColor(conf: FootnoteMatch["confidence"]): "success" | "warning" | "danger" {
  switch (conf) {
    case "high":
      return "success";
    case "medium":
      return "warning";
    case "low":
      return "danger";
  }
}

function truncate(text: string, max: number): string {
  return text.length > max ? `${text.substring(0, max)}…` : text;
}

export function FootnotePanel({ setError }: FootnotePanelProps): React.JSX.Element {
  const { dispatchToast } = useToastController(TOASTER_ID);
  const [scanResult, setScanResult] = useState<ScanResult | null>(null);
  const [scanning, setScanning] = useState(false);
  const [converting, setConverting] = useState(false);
  const [conversionResult, setConversionResult] = useState<ConversionResult | null>(null);
  const [matches, setMatches] = useState<FootnoteMatch[]>([]);

  function notify(title: string): void {
    dispatchToast(<Toast><ToastTitle>{title}</ToastTitle></Toast>, { intent: "success" });
  }

  function resetAll(): void {
    setScanResult(null);
    setMatches([]);
    setConversionResult(null);
  }

  async function runCount(
    op: Promise<{ data?: number; error?: string }>,
    noun: string
  ): Promise<void> {
    const result = await op;
    if (result.error) {
      setError(result.error);
    } else {
      notify(`${result.data ?? 0} ${noun}`);
    }
  }

  async function handleScan(): Promise<void> {
    setScanning(true);
    setError(null);
    try {
      const raw = await scanDocument();
      if (raw.error || !raw.data) {
        setError(raw.error ?? "Scan failed.");
      } else {
        const d = raw.data;
        const sr = buildScanResult(
          d.markers, d.blocks, d.separatorFound, d.separatorIndex,
          d.scanStartTime, d.existingFootnotes, d.citationSystems
        );
        setScanResult(sr);
        setMatches(sr.matches);
        setConversionResult(null);
      }
    } finally {
      setScanning(false);
    }
  }

  async function handleConvert(): Promise<void> {
    if (!scanResult) {
      return;
    }
    setConverting(true);
    setError(null);
    try {
      const sr: ScanResult = { ...scanResult, matches };
      const result = await convertAcceptedMatches(sr);
      if (result.error || !result.data) {
        setError(result.error ?? "Conversion failed.");
      } else {
        setConversionResult(result.data);
      }
    } finally {
      setConverting(false);
    }
  }

  // Appendix J manual override: user highlights the footnote block area when
  // auto-detection confidence is low; blocks are re-parsed from the selection
  // and re-matched against the already-scanned markers.
  async function handleUseSelectionAsBlocks(): Promise<void> {
    if (!scanResult) {
      return;
    }
    setError(null);
    const raw = await scanSelectionBlocks();
    if (raw.error || !raw.data) {
      setError(raw.error ?? "Block scan failed.");
      return;
    }
    if (raw.data.length === 0) {
      setError("No numbered footnote blocks found in the selection.");
      return;
    }
    const seen = new Set<string>();
    const markers = [
      ...scanResult.matches.map((m) => m.marker),
      ...scanResult.unmatchedMarkers,
    ].filter((m) => {
      const key = `${m.number}:${m.paragraphIndex}:${m.pattern}`;
      if (seen.has(key)) {
        return false;
      }
      seen.add(key);
      return true;
    });
    const sr = buildScanResult(
      markers, raw.data, false, -1, Date.now(), scanResult.existingFootnotes,
      scanResult.citationSystems
    );
    setScanResult(sr);
    setMatches(sr.matches);
  }

  function setAccepted(id: string, accepted: boolean): void {
    setMatches((prev) => prev.map((m) => (m.id === id ? { ...m, accepted } : m)));
  }

  function setAll(accepted: boolean): void {
    setMatches((prev) => prev.map((m) => ({ ...m, accepted })));
  }

  // Stage 3 — result
  if (conversionResult !== null) {
    return (
      <div>
        <Toaster toasterId={TOASTER_ID} />
        <p>Converted {conversionResult.converted} footnote(s) successfully.</p>
        {conversionResult.failed > 0 && (
          <div>
            <p>{conversionResult.failed} failed:</p>
            <ul>
              {conversionResult.errors.map((e, i) => (
                <li key={i}>{e}</li>
              ))}
            </ul>
          </div>
        )}
        <Button onClick={resetAll}>Scan Again</Button>
      </div>
    );
  }

  // Stage 2 — review table
  if (scanResult !== null) {
    return (
      <div>
        <Toaster toasterId={TOASTER_ID} />
        <p>
          Found {matches.length} matches, {scanResult.unmatchedMarkers.length} unmatched
          markers, {scanResult.orphanedBlocks.length} orphaned blocks.
        </p>
        {scanResult.existingFootnotes > 0 && (
          <p>
            Document already has {scanResult.existingFootnotes} genuine footnote(s) —
            left untouched.
          </p>
        )}
        {scanResult.citationSystems.length > 0 && (
          <p>
            Citation systems detected:{" "}
            {scanResult.citationSystems.map((c) => `${c.system} (${c.hits})`).join(", ")}.
            {scanResult.citationSystems.length > 1 &&
              " Mixed systems — verify citation consistency."}{" "}
            Citation-dense blocks are capped at medium confidence.
          </p>
        )}
        <div>
          <p>
            Low confidence? Highlight the footnote block area in the document, then:
          </p>
          <Button onClick={() => void handleUseSelectionAsBlocks()}>
            Use Selection as Footnote Area
          </Button>
        </div>
        {matches.length === 0 ? (
          <div>
            <p>No convertible footnotes found.</p>
            <Button onClick={resetAll}>Scan Again</Button>
          </div>
        ) : (
          <div>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHeaderCell>#</TableHeaderCell>
                  <TableHeaderCell>Marker Context</TableHeaderCell>
                  <TableHeaderCell>Footnote Text</TableHeaderCell>
                  <TableHeaderCell>Confidence</TableHeaderCell>
                  <TableHeaderCell>Include</TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {matches.map((m) => (
                  <TableRow key={m.id}>
                    <TableCell>{m.marker.number}</TableCell>
                    <TableCell>
                      {truncate(
                        `${m.marker.contextBefore}[${m.marker.number}]${m.marker.contextAfter}`,
                        120
                      )}
                    </TableCell>
                    <TableCell>{truncate(m.block.text, 80)}</TableCell>
                    <TableCell>
                      <Badge
                        appearance="filled"
                        color={confidenceColor(m.confidence)}
                        title={[
                          ...m.pairingEvidence,
                          ...m.contradictions.map((c) => `Against: ${c}`),
                          ...(m.block.boundaryUncertain ? ["Boundary needs eyes"] : []),
                          ...(m.block.mergedSuspect ? ["Possible prose merge — review"] : []),
                        ].join(" | ")}
                      >
                        {m.confidence}
                      </Badge>
                    </TableCell>
                    <TableCell>
                      <Checkbox
                        checked={m.accepted}
                        onChange={(_, data) => setAccepted(m.id, data.checked === true)}
                      />
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
            {scanResult.unmatchedMarkers.length > 0 && (
              <p>
                Unmatched markers:{" "}
                {scanResult.unmatchedMarkers.map((m) => m.number).join(", ")}
              </p>
            )}
            {scanResult.orphanedBlocks.length > 0 && (
              <p>
                Orphaned blocks:{" "}
                {scanResult.orphanedBlocks.map((b) => b.number).join(", ")}
              </p>
            )}
            <div>
              <Button onClick={() => setAll(true)}>Accept All</Button>
              <Button onClick={() => setAll(false)}>Reject All</Button>
              {converting ? (
                <Spinner label="Converting footnotes..." />
              ) : (
                <Button appearance="primary" onClick={() => void handleConvert()}>
                  Convert Selected
                </Button>
              )}
              <Button onClick={resetAll}>Scan Again</Button>
            </div>
          </div>
        )}
        <ManageNotes setError={setError} notify={notify} />
      </div>
    );
  }

  // Stage 1 — pre-scan
  return (
    <div>
      <Toaster toasterId={TOASTER_ID} />
      <p>
        Scans for footnotes pasted inline from PDFs or conversion tools and converts
        approved matches to genuine Word footnotes.
      </p>
      {scanning ? (
        <Spinner label="Scanning document..." />
      ) : (
        <Button appearance="primary" onClick={() => void handleScan()}>
          Scan Document
        </Button>
      )}
      <ManageNotes setError={setError} notify={notify} />
    </div>
  );
}

function ManageNotes({
  setError,
  notify,
}: {
  setError: (e: string | null) => void;
  notify: (title: string) => void;
}): React.JSX.Element {
  async function runCount(
    op: Promise<{ data?: number; error?: string }>,
    noun: string
  ): Promise<void> {
    const result = await op;
    if (result.error) {
      setError(result.error);
    } else {
      notify(`${result.data ?? 0} ${noun}`);
    }
  }
  return (
    <div>
      <h3>Manage Notes</h3>
      <Button onClick={() => void runCount(removeFootnotes(), "footnote(s) removed")}>
        Remove Footnotes
      </Button>
      <Button onClick={() => void runCount(removeEndnotes(), "endnote(s) removed")}>
        Remove Endnotes
      </Button>
      <Button
        onClick={() => void runCount(convertEndnoteToText(), "endnote(s) converted to text")}
      >
        Endnote → Text
      </Button>
    </div>
  );
}
