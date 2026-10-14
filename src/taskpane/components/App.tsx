import React from "react";
import { Button, FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global Office */

/**
 * Fallback launcher page. PawaTuls lives in the ribbon tab (menus + dialogs);
 * this page remains because the manifest requires a task pane source location.
 */
function openDialog(url: string): void {
  Office.context.ui.displayDialogAsync(url, { height: 70, width: 45 });
}

export function App(): React.JSX.Element {
  return (
    <FluentProvider theme={webLightTheme}>
      <div>
        <h2>PawaTuls</h2>
        <p>Use the PawaTuls tab on the ribbon. You can also open the dialogs here:</p>
        <Button
          appearance="primary"
          onClick={() => openDialog("https://localhost:3000/tools-dialog.html")}
        >
          Open Formatting Tools
        </Button>
        <Button
          onClick={() => openDialog("https://localhost:3000/footnote-dialog.html")}
        >
          Open Footnote Review
        </Button>
      </div>
    </FluentProvider>
  );
}
