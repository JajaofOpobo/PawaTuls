import React, { useState } from "react";
import { FluentProvider, Tab, TabList, webLightTheme } from "@fluentui/react-components";
import { ErrorBanner } from "./shared/ErrorBanner";
import { FormattingPanel } from "./FormattingPanel";
import { FootnotePanel } from "./FootnotePanel";

type ActiveTab = "format" | "footnotes";

function initialTab(): ActiveTab {
  try {
    const tab = new URLSearchParams(window.location.search).get("tab");
    return tab === "footnotes" ? "footnotes" : "format";
  } catch {
    return "format";
  }
}

export function App(): React.JSX.Element {
  const [activeTab, setActiveTab] = useState<ActiveTab>(initialTab);
  const [error, setError] = useState<string | null>(null);

  return (
    <FluentProvider theme={webLightTheme}>
      <ErrorBanner message={error} onDismiss={() => setError(null)} />
      <TabList
        selectedValue={activeTab}
        onTabSelect={(_, data) => setActiveTab(data.value as ActiveTab)}
      >
        <Tab value="format">Format</Tab>
        <Tab value="footnotes">Footnotes</Tab>
      </TabList>
      <div className="taskpane-panels">
        {activeTab === "format" && <FormattingPanel setError={setError} />}
        {activeTab === "footnotes" && <FootnotePanel setError={setError} />}
      </div>
    </FluentProvider>
  );
}
