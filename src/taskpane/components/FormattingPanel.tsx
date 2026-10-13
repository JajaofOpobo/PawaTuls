import React from "react";

export interface PanelProps {
  setError: (e: string | null) => void;
}

export function FormattingPanel({ setError }: PanelProps): React.JSX.Element {
  void setError;
  return <div>Formatting Panel — Phase 16</div>;
}
