import React from "react";

export interface PanelProps {
  setError: (e: string | null) => void;
}

export function LegalPanel({ setError }: PanelProps): React.JSX.Element {
  void setError;
  return <div>Legal Panel — Phase 24</div>;
}
