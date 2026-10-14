import React, { useState } from "react";
import {
  Button,
  Dropdown,
  Option,
  Radio,
  RadioGroup,
  Textarea,
} from "@fluentui/react-components";
import {
  BUILTIN_PROFILES,
  applyProfile,
  exportProfile,
  importProfile,
} from "../../formatting/profiles";
import { FormattingProfile } from "../../shared/types";
import { SectionProps } from "./NumberingSection";

export function ProfilesSection({ setError, notify }: SectionProps): React.JSX.Element {
  const [profiles, setProfiles] = useState<FormattingProfile[]>(BUILTIN_PROFILES);
  const [selectedProfile, setSelectedProfile] = useState<string>(BUILTIN_PROFILES[0].id);
  const [profileScope, setProfileScope] = useState<"selection" | "document">("selection");
  const [importText, setImportText] = useState("");

  const activeProfile = profiles.find((p) => p.id === selectedProfile) ?? profiles[0];

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
      <h3>Profiles</h3>
      <Dropdown
        value={activeProfile.name}
        selectedOptions={[activeProfile.id]}
        onOptionSelect={(_, data) => {
          if (data.optionValue) setSelectedProfile(data.optionValue);
        }}
      >
        {profiles.map((p) => (
          <Option key={p.id} value={p.id}>
            {p.name}
          </Option>
        ))}
      </Dropdown>
      <RadioGroup
        value={profileScope}
        onChange={(_, data) => setProfileScope(data.value as "selection" | "document")}
        layout="horizontal"
      >
        <Radio value="selection" label="Selection" />
        <Radio value="document" label="Document" />
      </RadioGroup>
      <Button
        appearance="primary"
        onClick={() =>
          void runCount(applyProfile(activeProfile, profileScope), "paragraph(s) formatted")
        }
      >
        Apply Profile
      </Button>
      <Textarea
        aria-label="Import profile JSON"
        placeholder="Paste profile JSON to import"
        value={importText}
        onChange={(_, data) => setImportText(data.value)}
      />
      <Button
        onClick={() => {
          const result = importProfile(importText);
          if (result.error || !result.data) {
            setError(result.error ?? "Import failed.");
            return;
          }
          setProfiles((prev) => {
            const without = prev.filter((p) => p.id !== result.data!.id);
            return [...without, result.data!];
          });
          setSelectedProfile(result.data.id);
          setImportText("");
          notify(`Profile "${result.data.name}" imported`);
        }}
      >
        Import Profile
      </Button>
      <Textarea
        aria-label="Export profile JSON"
        readOnly
        value={(() => {
          try {
            return exportProfile(activeProfile);
          } catch {
            return "";
          }
        })()}
      />
    </div>
  );
}
