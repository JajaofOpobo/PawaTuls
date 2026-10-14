import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import {
  FluentProvider,
  Toast,
  ToastTitle,
  Toaster,
  useToastController,
  webLightTheme,
} from "@fluentui/react-components";
import { NumberingSection } from "../taskpane/components/NumberingSection";
import { ProfilesSection } from "../taskpane/components/ProfilesSection";
import { ErrorBanner } from "../taskpane/components/shared/ErrorBanner";
import "../taskpane/taskpane.css";

/* global document, Office */

const TOASTER_ID = "tools-toaster";

function ToolsDialog(): React.JSX.Element {
  const [error, setError] = useState<string | null>(null);
  const { dispatchToast } = useToastController(TOASTER_ID);
  function notify(title: string): void {
    dispatchToast(<Toast><ToastTitle>{title}</ToastTitle></Toast>, { intent: "success" });
  }
  return (
    <FluentProvider theme={webLightTheme}>
      <Toaster toasterId={TOASTER_ID} />
      <ErrorBanner message={error} onDismiss={() => setError(null)} />
      <NumberingSection setError={setError} notify={notify} />
      <ProfilesSection setError={setError} notify={notify} />
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
