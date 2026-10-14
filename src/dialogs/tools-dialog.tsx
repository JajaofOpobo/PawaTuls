import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { FormattingPanel } from "../taskpane/components/FormattingPanel";
import { ErrorBanner } from "../taskpane/components/shared/ErrorBanner";
import "../taskpane/taskpane.css";

/* global document, Office */

function ToolsDialog(): React.JSX.Element {
  const [error, setError] = useState<string | null>(null);
  return (
    <FluentProvider theme={webLightTheme}>
      <ErrorBanner message={error} onDismiss={() => setError(null)} />
      <FormattingPanel setError={setError} />
    </FluentProvider>
  );
}

const container = document.getElementById("container");
if (container) {
  const root = createRoot(container);
  Office.onReady(() => {
    root.render(<ToolsDialog />);
  });
}
