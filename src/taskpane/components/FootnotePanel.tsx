import React from "react";

export interface PanelProps {
  setError: (e: string | null) => void;
}

export function FootnotePanel({ setError }: PanelProps): React.JSX.Element {
  void setError;
  return <div>Footnote Panel — Phase 28</div>;
}
