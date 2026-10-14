/* global Office */

/* Ribbon quick actions (headless): formatting, selects, inserts and dialog
 * openers only. Destructive whole-document operations live in the dialogs
 * with confirmation — they are never headless ribbon actions. */

import { insertCaption, refreshCaptions } from "../formatting/captions";
import { clearCharacterFormatting, decreaseFontSize, increaseFontSize, setSubscript, setSuperscript, toggleBold, toggleItalic } from "../formatting/font";
import { applyParagraphFormat, controlWidows, setLineSpacing } from "../formatting/paragraph";
import { BUILTIN_PROFILES, applyProfile } from "../formatting/profiles";
import { selectEmptyParagraphs, selectHeadings, selectPictures, selectShortParagraphs, selectTables } from "../formatting/select";
import { collapseMultipleSpaces, superscriptOrdinals, transformCase } from "../formatting/text";

function boldQuick(event: Office.AddinCommands.Event): void {
  void complete(event, toggleBold());
}

Office.actions.associate("boldQuick", boldQuick);

function italicQuick(event: Office.AddinCommands.Event): void {
  void complete(event, toggleItalic());
}

Office.actions.associate("italicQuick", italicQuick);

function superscriptQuick(event: Office.AddinCommands.Event): void {
  void complete(event, setSuperscript());
}

Office.actions.associate("superscriptQuick", superscriptQuick);

function subscriptQuick(event: Office.AddinCommands.Event): void {
  void complete(event, setSubscript());
}

Office.actions.associate("subscriptQuick", subscriptQuick);

function biggerQuick(event: Office.AddinCommands.Event): void {
  void complete(event, increaseFontSize());
}

Office.actions.associate("biggerQuick", biggerQuick);

function smallerQuick(event: Office.AddinCommands.Event): void {
  void complete(event, decreaseFontSize());
}

Office.actions.associate("smallerQuick", smallerQuick);

function clearCharQuick(event: Office.AddinCommands.Event): void {
  void complete(event, clearCharacterFormatting());
}

Office.actions.associate("clearCharQuick", clearCharQuick);

function alignLeftQuick(event: Office.AddinCommands.Event): void {
  void complete(event, applyParagraphFormat({ alignment: "left" }));
}

Office.actions.associate("alignLeftQuick", alignLeftQuick);

function alignCenterQuick(event: Office.AddinCommands.Event): void {
  void complete(event, applyParagraphFormat({ alignment: "centered" }));
}

Office.actions.associate("alignCenterQuick", alignCenterQuick);

function alignRightQuick(event: Office.AddinCommands.Event): void {
  void complete(event, applyParagraphFormat({ alignment: "right" }));
}

Office.actions.associate("alignRightQuick", alignRightQuick);

function alignJustifyQuick(event: Office.AddinCommands.Event): void {
  void complete(event, applyParagraphFormat({ alignment: "justified" }));
}

Office.actions.associate("alignJustifyQuick", alignJustifyQuick);

function singleQuick(event: Office.AddinCommands.Event): void {
  void complete(event, setLineSpacing("single"));
}

Office.actions.associate("singleQuick", singleQuick);

function oneFiveQuick(event: Office.AddinCommands.Event): void {
  void complete(event, setLineSpacing("onePtFive"));
}

Office.actions.associate("oneFiveQuick", oneFiveQuick);

function doubleQuick(event: Office.AddinCommands.Event): void {
  void complete(event, setLineSpacing("double"));
}

Office.actions.associate("doubleQuick", doubleQuick);

function upperQuick(event: Office.AddinCommands.Event): void {
  void complete(event, transformCase("upper"));
}

Office.actions.associate("upperQuick", upperQuick);

function lowerQuick(event: Office.AddinCommands.Event): void {
  void complete(event, transformCase("lower"));
}

Office.actions.associate("lowerQuick", lowerQuick);

function titleQuick(event: Office.AddinCommands.Event): void {
  void complete(event, transformCase("title"));
}

Office.actions.associate("titleQuick", titleQuick);

function sentenceQuick(event: Office.AddinCommands.Event): void {
  void complete(event, transformCase("sentence"));
}

Office.actions.associate("sentenceQuick", sentenceQuick);

function toggleQuick(event: Office.AddinCommands.Event): void {
  void complete(event, transformCase("toggle"));
}

Office.actions.associate("toggleQuick", toggleQuick);

function collapseQuick(event: Office.AddinCommands.Event): void {
  void complete(event, collapseMultipleSpaces());
}

Office.actions.associate("collapseQuick", collapseQuick);

function ordinalsQuick(event: Office.AddinCommands.Event): void {
  void complete(event, superscriptOrdinals("selection"));
}

Office.actions.associate("ordinalsQuick", ordinalsQuick);

function widowQuick(event: Office.AddinCommands.Event): void {
  void complete(event, controlWidows("selection"));
}

Office.actions.associate("widowQuick", widowQuick);

function tableQuick(event: Office.AddinCommands.Event): void {
  void complete(event, insertCaption("table"));
}

Office.actions.associate("tableQuick", tableQuick);

function figureQuick(event: Office.AddinCommands.Event): void {
  void complete(event, insertCaption("figure"));
}

Office.actions.associate("figureQuick", figureQuick);

function equationQuick(event: Office.AddinCommands.Event): void {
  void complete(event, insertCaption("equation"));
}

Office.actions.associate("equationQuick", equationQuick);

function refreshQuick(event: Office.AddinCommands.Event): void {
  void complete(event, refreshCaptions());
}

Office.actions.associate("refreshQuick", refreshQuick);

function tablesQuick(event: Office.AddinCommands.Event): void {
  void complete(event, selectTables());
}

Office.actions.associate("tablesQuick", tablesQuick);

function emptyParasQuick(event: Office.AddinCommands.Event): void {
  void complete(event, selectEmptyParagraphs());
}

Office.actions.associate("emptyParasQuick", emptyParasQuick);

function headingsQuick(event: Office.AddinCommands.Event): void {
  void complete(event, selectHeadings());
}

Office.actions.associate("headingsQuick", headingsQuick);

function shortParasQuick(event: Office.AddinCommands.Event): void {
  void complete(event, selectShortParagraphs());
}

Office.actions.associate("shortParasQuick", shortParasQuick);

function picturesQuick(event: Office.AddinCommands.Event): void {
  void complete(event, selectPictures());
}

Office.actions.associate("picturesQuick", picturesQuick);

function courtBodyQuick(event: Office.AddinCommands.Event): void {
  void complete(event, applyProfile(BUILTIN_PROFILES.find((p) => p.id === "court-body-1")!, "selection"));
}

Office.actions.associate("courtBodyQuick", courtBodyQuick);

function courtHeadingQuick(event: Office.AddinCommands.Event): void {
  void complete(event, applyProfile(BUILTIN_PROFILES.find((p) => p.id === "court-heading-1")!, "selection"));
}

Office.actions.associate("courtHeadingQuick", courtHeadingQuick);

function singleResetQuick(event: Office.AddinCommands.Event): void {
  void complete(event, applyProfile(BUILTIN_PROFILES.find((p) => p.id === "single-space-reset-1")!, "selection"));
}

Office.actions.associate("singleResetQuick", singleResetQuick);

function openToolsDialog(event: Office.AddinCommands.Event): void {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/tools-dialog.html",
    { height: 70, width: 40 },
    () => event.completed()
  );
}

function openFootnoteDialog(event: Office.AddinCommands.Event): void {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/footnote-dialog.html",
    { height: 70, width: 50 },
    () => event.completed()
  );
}

Office.actions.associate("openToolsDialog", openToolsDialog);
Office.actions.associate("openFootnoteDialog", openFootnoteDialog);

async function complete(event: Office.AddinCommands.Event, op: Promise<unknown>): Promise<void> {
  try {
    await op;
  } catch {
    // Headless by design (see above).
  } finally {
    event.completed();
  }
}

Office.onReady(() => {
  // Commands function file loaded.
});
