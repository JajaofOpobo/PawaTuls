/* global Office */

/* Ribbon commands: quick non-destructive actions run directly;
 * destructive whole-document actions route through a confirm dialog.
 * Headless failures are silent by design — use the dialogs for
 * confirmed, reported runs. */

import { fieldCodesToText, lineBreaksToParagraphMarks, listToText, paragraphMarksToLineBreaks } from "../formatting/convert";
import { removeEmptyParagraphs } from "../formatting/paragraph";
import { BUILTIN_PROFILES, applyProfile } from "../formatting/profiles";
import { removeAllBreaks, removeBookmarks, removeColumnBreaks, removeComments, removeContentControls, removeDropCap, removeEmptyLines, removeEmptyTableColumns, removeEmptyTableRows, removeExtraDocumentSpacing, removeFrames, removeHiddenText, removeHorizontalLines, removeHyperlinks, removePageBreaks, removePictures, removeSectionBreaks, removeTables, removeTextBoxes, tabToSpace } from "../formatting/removal";
import { selectEmptyParagraphs, selectHeadings, selectPictures, selectShortParagraphs, selectTables } from "../formatting/select";
import { resizeImages } from "../formatting/tables";
import { collapseMultipleSpaces, superscriptOrdinals } from "../formatting/text";

function formatCollapseBtn(event: Office.AddinCommands.Event): void {
  void complete(event, collapseMultipleSpaces());
}

Office.actions.associate("formatCollapseBtn", formatCollapseBtn);

function formatOrdinalsBtn(event: Office.AddinCommands.Event): void {
  void complete(event, superscriptOrdinals("selection"));
}

Office.actions.associate("formatOrdinalsBtn", formatOrdinalsBtn);

function formatLfToPmBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Manual line breaks to paragraph marks.", () => lineBreaksToParagraphMarks());
}

Office.actions.associate("formatLfToPmBtn", formatLfToPmBtn);

function formatPmToLfBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Paragraph marks to manual line breaks.", () => paragraphMarksToLineBreaks());
}

Office.actions.associate("formatPmToLfBtn", formatPmToLfBtn);

function formatFieldTextBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Field codes to plain text.", () => fieldCodesToText());
}

Office.actions.associate("formatFieldTextBtn", formatFieldTextBtn);

function formatListTextBtn(event: Office.AddinCommands.Event): void {
  void complete(event, listToText());
}

Office.actions.associate("formatListTextBtn", formatListTextBtn);

function cleanupEmptyParasBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove empty paragraphs.", () => removeEmptyParagraphs());
}

Office.actions.associate("cleanupEmptyParasBtn", cleanupEmptyParasBtn);

function cleanupTabsSpacesBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Convert tabs to spaces.", () => tabToSpace());
}

Office.actions.associate("cleanupTabsSpacesBtn", cleanupTabsSpacesBtn);

function cleanupEmptyLinesBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Collapse repeated manual line breaks.", () => removeEmptyLines());
}

Office.actions.associate("cleanupEmptyLinesBtn", cleanupEmptyLinesBtn);

function cleanupHorizLinesBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove horizontal line paragraphs.", () => removeHorizontalLines());
}

Office.actions.associate("cleanupHorizLinesBtn", cleanupHorizLinesBtn);

function cleanupExtraSpacingBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Clamp spacing above 12pt to 6pt.", () => removeExtraDocumentSpacing());
}

Office.actions.associate("cleanupExtraSpacingBtn", cleanupExtraSpacingBtn);

function cleanupHiddenBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Delete hidden text.", () => removeHiddenText());
}

Office.actions.associate("cleanupHiddenBtn", cleanupHiddenBtn);

function cleanupDropCapsBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove drop caps.", () => removeDropCap());
}

Office.actions.associate("cleanupDropCapsBtn", cleanupDropCapsBtn);

function cleanupFramesBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove frames.", () => removeFrames());
}

Office.actions.associate("cleanupFramesBtn", cleanupFramesBtn);

function breaksPageBreaksBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove all page breaks.", () => removePageBreaks());
}

Office.actions.associate("breaksPageBreaksBtn", breaksPageBreaksBtn);

function breaksColumnBreaksBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove all column breaks.", () => removeColumnBreaks());
}

Office.actions.associate("breaksColumnBreaksBtn", breaksColumnBreaksBtn);

function breaksSectionBreaksBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove section breaks (sections merge).", () => removeSectionBreaks());
}

Office.actions.associate("breaksSectionBreaksBtn", breaksSectionBreaksBtn);

function breaksAllBreaksBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove all breaks.", () => removeAllBreaks());
}

Office.actions.associate("breaksAllBreaksBtn", breaksAllBreaksBtn);

function tablesRemoveTablesBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Delete all tables and content.", () => removeTables());
}

Office.actions.associate("tablesRemoveTablesBtn", tablesRemoveTablesBtn);

function tablesEmptyRowsBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Delete wholly-empty table rows.", () => removeEmptyTableRows());
}

Office.actions.associate("tablesEmptyRowsBtn", tablesEmptyRowsBtn);

function tablesEmptyColsBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Delete wholly-empty table columns.", () => removeEmptyTableColumns());
}

Office.actions.associate("tablesEmptyColsBtn", tablesEmptyColsBtn);

function tablesPicturesBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Delete all pictures.", () => removePictures());
}

Office.actions.associate("tablesPicturesBtn", tablesPicturesBtn);

function tablesTextBoxesBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Delete text boxes and content.", () => removeTextBoxes());
}

Office.actions.associate("tablesTextBoxesBtn", tablesTextBoxesBtn);

function tablesControlsBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove content controls, keep text.", () => removeContentControls());
}

Office.actions.associate("tablesControlsBtn", tablesControlsBtn);

function tablesResize50Btn(event: Office.AddinCommands.Event): void {
  void complete(event, resizeImages(50));
}

Office.actions.associate("tablesResize50Btn", tablesResize50Btn);

function tablesResize100Btn(event: Office.AddinCommands.Event): void {
  void complete(event, resizeImages(100));
}

Office.actions.associate("tablesResize100Btn", tablesResize100Btn);

function tablesResize150Btn(event: Office.AddinCommands.Event): void {
  void complete(event, resizeImages(150));
}

Office.actions.associate("tablesResize150Btn", tablesResize150Btn);

function referencesCommentsBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Delete all comments.", () => removeComments());
}

Office.actions.associate("referencesCommentsBtn", referencesCommentsBtn);

function referencesHyperlinksBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Remove hyperlinks, keep text.", () => removeHyperlinks());
}

Office.actions.associate("referencesHyperlinksBtn", referencesHyperlinksBtn);

function referencesBookmarksBtn(event: Office.AddinCommands.Event): void {
  void confirmAndRun(event, "Delete all bookmarks.", () => removeBookmarks());
}

Office.actions.associate("referencesBookmarksBtn", referencesBookmarksBtn);

function profilesCourtBodyBtn(event: Office.AddinCommands.Event): void {
  void complete(event, applyProfile(BUILTIN_PROFILES.find((p) => p.id === "court-body-1")!, "selection"));
}

Office.actions.associate("profilesCourtBodyBtn", profilesCourtBodyBtn);

function profilesCourtHeadingBtn(event: Office.AddinCommands.Event): void {
  void complete(event, applyProfile(BUILTIN_PROFILES.find((p) => p.id === "court-heading-1")!, "selection"));
}

Office.actions.associate("profilesCourtHeadingBtn", profilesCourtHeadingBtn);

function profilesSingleResetBtn(event: Office.AddinCommands.Event): void {
  void complete(event, applyProfile(BUILTIN_PROFILES.find((p) => p.id === "single-space-reset-1")!, "selection"));
}

Office.actions.associate("profilesSingleResetBtn", profilesSingleResetBtn);

function selectTablesBtn(event: Office.AddinCommands.Event): void {
  void complete(event, selectTables());
}

Office.actions.associate("selectTablesBtn", selectTablesBtn);

function selectEmptyParasBtn(event: Office.AddinCommands.Event): void {
  void complete(event, selectEmptyParagraphs());
}

Office.actions.associate("selectEmptyParasBtn", selectEmptyParasBtn);

function selectHeadingsBtn(event: Office.AddinCommands.Event): void {
  void complete(event, selectHeadings());
}

Office.actions.associate("selectHeadingsBtn", selectHeadingsBtn);

function selectShortParasBtn(event: Office.AddinCommands.Event): void {
  void complete(event, selectShortParagraphs());
}

Office.actions.associate("selectShortParasBtn", selectShortParasBtn);

function selectPicturesBtn(event: Office.AddinCommands.Event): void {
  void complete(event, selectPictures());
}

Office.actions.associate("selectPicturesBtn", selectPicturesBtn);

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

function confirmAndRun(event: Office.AddinCommands.Event, message: string, op: () => Promise<unknown>): void {
  const url = "https://localhost:3000/confirm-dialog.html?message=" + encodeURIComponent(message);
  Office.context.ui.displayDialogAsync(url, { height: 30, width: 30 }, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed();
      return;
    }
    const dialog = result.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: { message?: unknown; error?: unknown }) => {
      dialog.close();
      if (arg.message === "confirm") {
        void complete(event, op());
      } else {
        event.completed();
      }
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
      event.completed();
    });
  });
}

Office.onReady(() => {
  // Commands function file loaded.
});
