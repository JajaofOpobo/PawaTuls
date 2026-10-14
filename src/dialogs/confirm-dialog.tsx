import React from "react";
import { createRoot } from "react-dom/client";
import { Button, FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, window */

function ConfirmDialog(): React.JSX.Element {
  let message = "Apply this action to the document?";
  try {
    message = new URLSearchParams(window.location.search).get("message") ?? message;
  } catch {
    // keep default
  }
  function send(value: string): void {
    try {
      Office.context.ui.messageParent(value);
    } catch {
      // parent already gone; nothing to do
    }
  }
  return (
    <FluentProvider theme={webLightTheme}>
      <div>
        <p>{message}</p>
        <Button appearance="primary" onClick={() => send("confirm")}>
          Apply
        </Button>
        <Button onClick={() => send("cancel")}>Cancel</Button>
      </div>
    </FluentProvider>
  );
}

const container = document.getElementById("container");
if (container) {
  const root = createRoot(container);
  Office.onReady(() => {
    root.render(<ConfirmDialog />);
  });
}
