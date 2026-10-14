/**
 * notes.ts — Phase 37: management of existing genuine Word notes.
 * (Footnote *detection* of pasted inline markers lives in scanner/matcher/
 * converter; this module operates on real NoteItem collections.)
 *
 * API notes (verified against @types/office-js):
 * - body.footnotes / body.endnotes are NoteItemCollections; NoteItem has
 *   body (note content), reference (body mark range), delete().
 * - There is NO endnote-insert API, so footnote→endnote conversion (a KuTools
 *   feature) is infeasible — documented in docs/KUTOOLS_COVERAGE.md.
 * - convertEndnoteToText reads the ACTUAL reference number (KuTools synthesized
 *   it from the loop index and covered only 2 of ~9 number styles — AUDIT.md).
 */
import { safeWordRun } from "../shared/wordUtils";

async function removeNotes(
  kind: "footnote" | "endnote"
): Promise<{ data?: number; error?: string }> {
  return safeWordRun(async (context) => {
    const body = context.document.body;
    const collection = kind === "footnote" ? body.footnotes : body.endnotes;
    collection.load("items");
    await context.sync();
    const items = collection.items.slice().reverse();
    for (const note of items) {
      note.delete();
    }
    await context.sync();
    return items.length;
  });
}

export async function removeFootnotes(): Promise<{ data?: number; error?: string }> {
  return removeNotes("footnote");
}

export async function removeEndnotes(): Promise<{ data?: number; error?: string }> {
  return removeNotes("endnote");
}

export async function convertEndnoteToText(): Promise<{
  data?: number;
  error?: string;
}> {
  return safeWordRun(async (context) => {
    const body = context.document.body;
    const endnotes = body.endnotes;
    endnotes.load("items");
    await context.sync();
    const items = endnotes.items.slice();

    // Snapshot numbers + texts first (forward), so later mutations cannot skew them.
    const snapshots: { number: string; text: string }[] = [];
    for (const note of items) {
      note.reference.load("text");
      note.body.load("text");
    }
    await context.sync();
    for (const note of items) {
      snapshots.push({ number: note.reference.text.trim(), text: note.body.text.trim() });
    }

    // Transcribe at document end (forward).
    for (const snap of snapshots) {
      body.insertText(`\n${snap.number}. ${snap.text}`, Word.InsertLocation.end);
    }
    await context.sync();

    // Replace each reference mark in place, then delete notes (both reverse).
    for (let i = items.length - 1; i >= 0; i--) {
      items[i].reference.insertText(`[${snapshots[i].number}]`, Word.InsertLocation.replace);
    }
    await context.sync();
    for (let i = items.length - 1; i >= 0; i--) {
      items[i].delete();
    }
    await context.sync();
    return items.length;
  });
}
