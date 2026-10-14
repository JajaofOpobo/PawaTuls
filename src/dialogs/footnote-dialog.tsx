import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { FootnotePanel } from "../taskpane/components/FootnotePanel";
import { ErrorBanner } from "../taskpane/components/shared/ErrorBanner";
import "../taskpane/taskpane.css";

/* global document, Office */

function FootnoteDialog(): React.JSX.Element {
  const [error, setError] = useState<string | null>(null);
  return (
    <FluentProvider theme={webLightTheme}>
      <ErrorBanner message={error} onDismiss={() => setError(null)} />
      <FootnotePanel setError={setError} />
    </FluentProvider>
  );
}

const container = document.getElementById("container");
if (container) {
  const root = createRoot(container);
  Office.onReady(() => {
    root.render(<FootnoteDialog />);
  });
}
