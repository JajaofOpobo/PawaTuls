import React, { useState } from "react";
import { Badge, Button, Radio, RadioGroup } from "@fluentui/react-components";
import {
  applyRepairPlan,
  auditNumbering,
  previewRepair,
} from "../../formatting/numbering";
import {
  NumberedParagraph,
  NumberingIssue,
  NumberingRepairPlan,
} from "../../shared/types";

export interface SectionProps {
  setError: (e: string | null) => void;
  notify: (title: string) => void;
}

export function NumberingSection({ setError, notify }: SectionProps): React.JSX.Element {
  const [numberScope, setNumberScope] = useState<"selection" | "document">("selection");
  const [numberItems, setNumberItems] = useState<NumberedParagraph[]>([]);
  const [numberIssues, setNumberIssues] = useState<NumberingIssue[]>([]);
  const [repairPlan, setRepairPlan] = useState<NumberingRepairPlan | null>(null);
  const [audited, setAudited] = useState(false);

  async function runCount(
    op: Promise<{ data?: number; error?: string }>,
    noun: string
  ): Promise<void> {
    const result = await op;
    if (result.error) {
      setError(result.error);
    } else {
      notify(`${result.data ?? 0} ${noun}`);
    }
  }

  return (
    <div>
      <h3>Numbering</h3>
      <RadioGroup
        value={numberScope}
        onChange={(_, data) => setNumberScope(data.value as "selection" | "document")}
        layout="horizontal"
      >
        <Radio value="selection" label="Selection" />
        <Radio value="document" label="Document" />
      </RadioGroup>
      <Button
        onClick={() =>
          void (async () => {
            const result = await auditNumbering(numberScope);
            if (result.error || !result.data) {
              setError(result.error ?? "Audit failed.");
              return;
            }
            setNumberItems(result.data.items);
            setNumberIssues(result.data.issues);
            setRepairPlan(null);
            setAudited(true);
          })()
        }
      >
        Audit Numbering
      </Button>
      {audited && (
        <div>
          <p>
            {numberItems.filter((i) => i.kind === "native").length} native,{" "}
            {numberItems.filter((i) => i.kind === "manual").length} manually
            numbered paragraph(s).
          </p>
          {numberIssues.length === 0 ? (
            <p>No numbering issues found.</p>
          ) : (
            <ul>
              {numberIssues.map((issue, i) => (
                <li key={i}>
                  <Badge
                    appearance="filled"
                    color={issue.issue === "mixed-style" || issue.issue === "orphan-level" ? "warning" : "danger"}
                  >
                    {issue.issue}
                  </Badge>{" "}
                  ¶{issue.paragraphIndex}: {issue.detail}
                </li>
              ))}
            </ul>
          )}
          <Button
            onClick={() => {
              setRepairPlan(previewRepair(numberIssues, numberItems));
            }}
          >
            Preview Repair
          </Button>
          {repairPlan !== null && (
            <div>
              {repairPlan.edits.length === 0 ? (
                <p>Nothing to repair (restarts and mixed styles need approval).</p>
              ) : (
                <ul>
                  {repairPlan.edits.map((e, i) => (
                    <li key={i}>
                      ¶{e.paragraphIndex}: {e.before} → {e.after}
                    </li>
                  ))}
                </ul>
              )}
              {repairPlan.edits.length > 0 && (
                <Button
                  appearance="primary"
                  onClick={() =>
                    void runCount(
                      applyRepairPlan(repairPlan, numberScope),
                      "number(s) repaired"
                    )
                  }
                >
                  Apply Repair
                </Button>
              )}
            </div>
          )}
          <p>Native lists are managed by Word — never edited by repair.</p>
        </div>
      )}
    </div>
  );
}
