/* global Office */

import { toggleBold } from "../formatting/font";
import { clearFormatting } from "../formatting/clear";
import { removeEmptyParagraphs } from "../formatting/paragraph";
import { superscriptOrdinals } from "../formatting/text";
import { convertEndnoteToText, removeFootnotes } from "../footnotes/notes";

Office.onReady(() => {
  // Commands function file loaded. Ribbon quick actions below run headless:
  // errors have no task pane to surface in, so failures are silent by design
  // (use the task pane buttons for confirmed, reported runs).
});

async function complete(event: Office.AddinCommands.Event, op: Promise<unknown>): Promise<void> {
  try {
    await op;
  } catch {
    // Headless by design (see above).
  } finally {
    event.completed();
  }
}

function toggleBoldQuick(event: Office.AddinCommands.Event): void {
  void complete(event, toggleBold());
}

function clearAllFormattingQuick(event: Office.AddinCommands.Event): void {
  void complete(event, clearFormatting("all"));
}

function removeEmptyParagraphsQuick(event: Office.AddinCommands.Event): void {
  void complete(event, removeEmptyParagraphs());
}

function superscriptOrdinalsQuick(event: Office.AddinCommands.Event): void {
  void complete(event, superscriptOrdinals("selection"));
}

function removeFootnotesQuick(event: Office.AddinCommands.Event): void {
  void complete(event, removeFootnotes());
}

function convertEndnoteToTextQuick(event: Office.AddinCommands.Event): void {
  void complete(event, convertEndnoteToText());
}

Office.actions.associate("toggleBoldQuick", toggleBoldQuick);
Office.actions.associate("clearAllFormattingQuick", clearAllFormattingQuick);
Office.actions.associate("removeEmptyParagraphsQuick", removeEmptyParagraphsQuick);
Office.actions.associate("superscriptOrdinalsQuick", superscriptOrdinalsQuick);
Office.actions.associate("removeFootnotesQuick", removeFootnotesQuick);
Office.actions.associate("convertEndnoteToTextQuick", convertEndnoteToTextQuick);
