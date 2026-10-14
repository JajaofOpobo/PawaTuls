import React, { useState } from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Button,
  Dropdown,
  Option,
  Radio,
  RadioGroup,
  SpinButton,
  Textarea,
  Toast,
  ToastBody,
  ToastTitle,
  Toaster,
  ToggleButton,
  useToastController,
} from "@fluentui/react-components";
import { applyFont } from "../../formatting/font";
import {
  applyParagraphFormat,
  removeAllIndents,
  removeAllSpacing,
  removeEmptyParagraphs,
  setFirstLineIndent,
  setHangingIndent,
  setLineSpacing,
} from "../../formatting/paragraph";
import {
  clearCharacterFormatting,
  decreaseFontSize,
  increaseFontSize,
  setSubscript,
  setSuperscript,
  toggleBold,
  toggleItalic,
  toggleStrikethrough,
} from "../../formatting/font";
import { applyStyle, getAvailableStyles } from "../../formatting/styles";
import {
  BUILTIN_PROFILES,
  applyProfile,
  exportProfile,
  importProfile,
} from "../../formatting/profiles";
import {
  applyRepairPlan,
  auditNumbering,
  previewRepair,
} from "../../formatting/numbering";
import {
  removeAllBreaks,
  removeBookmarks,
  removeColumnBreaks,
  removeComments,
  removeContentControls,
  removeDropCap,
  removeEmptyLines,
  removeEmptyTableColumns,
  removeEmptyTableRows,
  removeFrames,
  removeHiddenText,
  removeHyperlinks,
  removePageBreaks,
  removePictures,
  removeSectionBreaks,
  removeTables,
  removeTextBoxes,
  tabToSpace,
} from "../../formatting/removal";
import {
  FormattingProfile,
  NumberingIssue,
  NumberedParagraph,
  NumberingRepairPlan,
} from "../../shared/types";
import { collapseMultipleSpaces, superscriptOrdinals, transformCase } from "../../formatting/text";
import {
  lineBreaksToParagraphMarks,
  paragraphMarksToLineBreaks,
} from "../../formatting/convert";
import { insertCaption, refreshCaptions } from "../../formatting/captions";
import {
  selectEmptyParagraphs,
  selectHeadings,
  selectPictures,
  selectShortParagraphs,
  selectTables,
  SelectResult,
} from "../../formatting/select";
import { controlWidows } from "../../formatting/paragraph";
import {
  convertTableToText,
  repeatHeaderRows,
  resizeImages,
} from "../../formatting/tables";
import { clearFormatting } from "../../formatting/clear";
import { CaseTransform, LineSpacingRule } from "../../shared/types";

export interface FormattingPanelProps {
  setError: (e: string | null) => void;
}

type Alignment = "left" | "centered" | "right" | "justified";

const FONT_FAMILIES = [
  "Times New Roman",
  "Arial",
  "Calibri",
  "Garamond",
  "Book Antiqua",
  "Cambria",
];

const HIGHLIGHT_COLORS = [
  "none",
  "yellow",
  "turquoise",
  "pink",
  "violet",
  "darkBlue",
  "darkRed",
  "darkYellow",
  "darkGreen",
  "teal",
  "darkCyan",
  "black",
  "blue",
  "darkMagenta",
  "darkGray",
  "lightGray",
  "white",
];

const LINE_SPACING_LABELS: { value: LineSpacingRule; label: string }[] = [
  { value: "single", label: "Single" },
  { value: "onePtFive", label: "1.5 Lines" },
  { value: "double", label: "Double" },
  { value: "atLeast", label: "At Least" },
  { value: "exactly", label: "Exactly" },
  { value: "multiple", label: "Multiple" },
];

const CASE_BUTTONS: { value: CaseTransform; label: string }[] = [
  { value: "upper", label: "UPPERCASE" },
  { value: "lower", label: "lowercase" },
  { value: "title", label: "Title Case" },
  { value: "sentence", label: "Sentence case" },
  { value: "toggle", label: "tOGGLE" },
];

const TOASTER_ID = "formatting-toaster";

export function FormattingPanel({ setError }: FormattingPanelProps): React.JSX.Element {
  const { dispatchToast } = useToastController(TOASTER_ID);
  const [underlineOn, setUnderlineOn] = useState(false);
  const [alignment, setAlignment] = useState<Alignment>("left");
  const [spacingRule, setSpacingRule] = useState<LineSpacingRule>("single");
  const [styles, setStyles] = useState<string[]>([]);
  const [profiles, setProfiles] = useState<FormattingProfile[]>(BUILTIN_PROFILES);
  const [selectedProfile, setSelectedProfile] = useState<string>(BUILTIN_PROFILES[0].id);
  const [profileScope, setProfileScope] = useState<"selection" | "document">("selection");
  const [importText, setImportText] = useState("");
  const [captionKind, setCaptionKind] = useState<"table" | "figure" | "equation">("table");
  const [captionLabel, setCaptionLabel] = useState("");
  const [tableSeparator, setTableSeparator] = useState<"tab" | "paragraph" | "comma">("tab");
  const [repeatHeaders, setRepeatHeaders] = useState(true);
  const [resizePercent, setResizePercent] = useState(100);
  const [numberScope, setNumberScope] = useState<"selection" | "document">("selection");
  const [numberItems, setNumberItems] = useState<NumberedParagraph[]>([]);
  const [numberIssues, setNumberIssues] = useState<NumberingIssue[]>([]);
  const [repairPlan, setRepairPlan] = useState<NumberingRepairPlan | null>(null);
  const [audited, setAudited] = useState(false);

  const activeProfile = profiles.find((p) => p.id === selectedProfile) ?? profiles[0];

  function notify(title: string, body?: string): void {
    dispatchToast(
      <Toast>
        <ToastTitle>{title}</ToastTitle>
        {body !== undefined && <ToastBody>{body}</ToastBody>}
      </Toast>,
      { intent: "success" }
    );
  }

  async function run(op: Promise<{ error?: string }>): Promise<void> {
    const result = await op;
    if (result.error) {
      setError(result.error);
    }
  }

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

  async function runSelect(
    op: Promise<{ data?: SelectResult; error?: string }>,
    noun: string
  ): Promise<void> {
    const result = await op;
    if (result.error) {
      setError(result.error);
    } else if ((result.data?.total ?? 0) === 0) {
      notify(`No ${noun} found`);
    } else {
      notify(`${result.data?.total} ${noun} — first selected`);
    }
  }

  async function loadStyles(): Promise<void> {
    const result = await getAvailableStyles();
    if (result.error) {
      setError(result.error);
    } else {
      setStyles(result.data ?? []);
    }
  }

  const needsSpacingValue =
    spacingRule === "atLeast" || spacingRule === "exactly" || spacingRule === "multiple";

  return (
    <div>
      <Toaster toasterId={TOASTER_ID} />
      <Accordion multiple collapsible>
        <AccordionItem value="font">
          <AccordionHeader>Font</AccordionHeader>
          <AccordionPanel>
            <Dropdown
              placeholder="Font family"
              onOptionSelect={(_, data) => {
                if (data.optionValue) void run(applyFont({ name: data.optionValue }));
              }}
            >
              {FONT_FAMILIES.map((f) => (
                <Option key={f} value={f}>
                  {f}
                </Option>
              ))}
            </Dropdown>
            <SpinButton
              defaultValue={12}
              min={6}
              max={72}
              onChange={(_, data) => {
                if (data.value != null) void run(applyFont({ size: data.value }));
              }}
            />
            <div>
              <ToggleButton onClick={() => void run(toggleBold())}>B</ToggleButton>
              <ToggleButton onClick={() => void run(toggleItalic())}>I</ToggleButton>
              <ToggleButton
                checked={underlineOn}
                onClick={() => {
                  const next = !underlineOn;
                  setUnderlineOn(next);
                  void run(applyFont({ underline: next ? "single" : "none" }));
                }}
              >
                U
              </ToggleButton>
              <ToggleButton onClick={() => void run(toggleStrikethrough())}>S</ToggleButton>
            </div>
            <div>
              <Button onClick={() => void run(setSuperscript())}>x²</Button>
              <Button onClick={() => void run(setSubscript())}>x₂</Button>
              <Button onClick={() => void run(increaseFontSize())}>A+</Button>
              <Button onClick={() => void run(decreaseFontSize())}>A−</Button>
            </div>
            <div>
              <input
                type="color"
                aria-label="Text color"
                onChange={(e) => {
                  if (e.target.value) void run(applyFont({ color: e.target.value }));
                }}
              />
              <Dropdown
                placeholder="Highlight"
                onOptionSelect={(_, data) => {
                  if (data.optionValue)
                    void run(applyFont({ highlightColor: data.optionValue }));
                }}
              >
                {HIGHLIGHT_COLORS.map((c) => (
                  <Option key={c} value={c}>
                    {c}
                  </Option>
                ))}
              </Dropdown>
            </div>
            <SpinButton
              defaultValue={0}
              min={-20}
              max={200}
              step={0.5}
              aria-label="Character spacing (points)"
              onChange={(_, data) => {
                if (data.value != null) void run(applyFont({ spacing: data.value }));
              }}
            />
            <Button onClick={() => void run(clearCharacterFormatting())}>
              Clear character formatting
            </Button>
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="paragraph">
          <AccordionHeader>Paragraph</AccordionHeader>
          <AccordionPanel>
            <div>
              {(["left", "centered", "right", "justified"] as Alignment[]).map((a) => (
                <ToggleButton
                  key={a}
                  checked={alignment === a}
                  onClick={() => {
                    setAlignment(a);
                    void run(applyParagraphFormat({ alignment: a }));
                  }}
                >
                  {a[0].toUpperCase()}
                </ToggleButton>
              ))}
            </div>
            <Dropdown
              value={LINE_SPACING_LABELS.find((l) => l.value === spacingRule)?.label}
              selectedOptions={[spacingRule]}
              onOptionSelect={(_, data) => {
                const rule = data.optionValue as LineSpacingRule;
                setSpacingRule(rule);
                if (rule === "single" || rule === "onePtFive" || rule === "double") {
                  void run(setLineSpacing(rule));
                }
              }}
            >
              {LINE_SPACING_LABELS.map((l) => (
                <Option key={l.value} value={l.value}>
                  {l.label}
                </Option>
              ))}
            </Dropdown>
            {needsSpacingValue && (
              <SpinButton
                defaultValue={12}
                min={1}
                aria-label="Line spacing value"
                onChange={(_, data) => {
                  if (data.value != null) void run(setLineSpacing(spacingRule, data.value));
                }}
              />
            )}
            <SpinButton
              defaultValue={0}
              min={0}
              aria-label="Space before (points)"
              onChange={(_, data) => {
                if (data.value != null)
                  void run(applyParagraphFormat({ spaceBefore: data.value }));
              }}
            />
            <SpinButton
              defaultValue={0}
              min={0}
              aria-label="Space after (points)"
              onChange={(_, data) => {
                if (data.value != null)
                  void run(applyParagraphFormat({ spaceAfter: data.value }));
              }}
            />
            <SpinButton
              defaultValue={0}
              min={0}
              aria-label="First-line indent (cm)"
              onChange={(_, data) => {
                if (data.value != null) void run(setFirstLineIndent(data.value));
              }}
            />
            <SpinButton
              defaultValue={0}
              min={0}
              aria-label="Hanging indent (cm)"
              onChange={(_, data) => {
                if (data.value != null) void run(setHangingIndent(data.value));
              }}
            />
            <Button onClick={() => void run(removeAllSpacing())}>Remove All Spacing</Button>
            <Button onClick={() => void run(removeAllIndents())}>Remove All Indents</Button>
            <Button
              onClick={() => void runCount(removeEmptyParagraphs(), "empty paragraph(s) removed")}
            >
              Remove Empty Paragraphs
            </Button>
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="styles-tools">
          <AccordionHeader>Styles and Tools</AccordionHeader>
          <AccordionPanel>
            <Dropdown
              placeholder="Apply style"
              onOptionSelect={(_, data) => {
                if (data.optionValue) void run(applyStyle(data.optionValue));
              }}
              onClick={() => void loadStyles()}
            >
              {styles.map((s) => (
                <Option key={s} value={s}>
                  {s}
                </Option>
              ))}
            </Dropdown>
            <div>
              {CASE_BUTTONS.map((c) => (
                <Button key={c.value} onClick={() => void run(transformCase(c.value))}>
                  {c.label}
                </Button>
              ))}
            </div>
            <Button onClick={() => void run(collapseMultipleSpaces())}>Collapse spaces</Button>
            <Button
              onClick={() =>
                void (async () => {
                  const r = await superscriptOrdinals("selection");
                  if (r.error) {
                    setError(r.error);
                  } else {
                    notify(
                      `${r.data?.changed ?? 0} ordinal(s) superscripted` +
                        ((r.data?.skipped ?? 0) > 0 ? ` (${r.data?.skipped} skipped)` : "")
                    );
                  }
                })()
              }
            >
              Superscript ordinals (1st, 2nd…)
            </Button>
            <Button
              onClick={() =>
                void runCount(lineBreaksToParagraphMarks(), "line break(s) converted")
              }
            >
              Line Breaks → Paragraph Marks
            </Button>
            <Button
              onClick={() =>
                void runCount(paragraphMarksToLineBreaks(), "paragraph mark(s) converted")
              }
            >
              Paragraph Marks → Line Breaks
            </Button>
            <Button onClick={() => void run(clearFormatting("all"))}>
              Clear All Formatting
            </Button>
            <Dropdown
              value={captionKind}
              selectedOptions={[captionKind]}
              aria-label="Caption kind"
              onOptionSelect={(_, data) => {
                if (data.optionValue) {
                  setCaptionKind(data.optionValue as "table" | "figure" | "equation");
                }
              }}
            >
              <Option value="table">Table</Option>
              <Option value="figure">Figure</Option>
              <Option value="equation">Equation</Option>
            </Dropdown>
            <input
              aria-label="Caption label (optional)"
              placeholder="Label (optional)"
              value={captionLabel}
              onChange={(e) => setCaptionLabel(e.target.value)}
            />
            <Button onClick={() => void run(insertCaption(captionKind, captionLabel))}>
              Insert Caption
            </Button>
            <Button onClick={() => void runCount(refreshCaptions(), "caption(s) refreshed")}>
              Refresh Captions
            </Button>
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="profiles">
          <AccordionHeader>Profiles</AccordionHeader>
          <AccordionPanel>
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
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="numbering">
          <AccordionHeader>Numbering</AccordionHeader>
          <AccordionPanel>
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
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="remove">
          <AccordionHeader>Remove</AccordionHeader>
          <AccordionPanel>
            <h4>Breaks</h4>
            <div>
              <Button onClick={() => void runCount(removePageBreaks(), "page break(s) removed")}>
                Page Breaks
              </Button>
              <Button onClick={() => void runCount(removeColumnBreaks(), "column break(s) removed")}>
                Column Breaks
              </Button>
              <Button onClick={() => void runCount(removeSectionBreaks(), "section break(s) removed")}>
                Section Breaks
              </Button>
              <Button onClick={() => void runCount(removeAllBreaks(), "break(s) removed")}>
                All Breaks
              </Button>
            </div>
            <h4>Hidden &amp; Marks</h4>
            <div>
              <Button onClick={() => void runCount(removeHiddenText(), "hidden run(s) removed")}>
                Hidden Text
              </Button>
              <Button onClick={() => void runCount(removeDropCap(), "drop cap(s) removed")}>
                Drop Caps
              </Button>
              <Button onClick={() => void runCount(removeFrames(), "frame(s) removed")}>
                Frames
              </Button>
              <Button onClick={() => void runCount(removeEmptyLines(), "line run(s) collapsed")}>
                Empty Lines
              </Button>
              <Button onClick={() => void runCount(tabToSpace(), "paragraph(s) detabbed")}>
                Tabs → Spaces
              </Button>
            </div>
            <h4>Tables &amp; Objects</h4>
            <div>
              <Button onClick={() => void runCount(removeTables(), "table(s) removed")}>
                Tables
              </Button>
              <Button onClick={() => void runCount(removeEmptyTableRows(), "empty row(s) removed")}>
                Empty Rows
              </Button>
              <Button onClick={() => void runCount(removeEmptyTableColumns(), "empty column(s) removed")}>
                Empty Cols
              </Button>
              <Button onClick={() => void runCount(removePictures(), "picture(s) removed")}>
                Pictures
              </Button>
              <Button onClick={() => void runCount(removeTextBoxes(), "text box(es) removed")}>
                Text Boxes
              </Button>
              <Button onClick={() => void runCount(removeContentControls(), "control(s) removed")}>
                Controls
              </Button>
            </div>
            <h4>References</h4>
            <div>
              <Button onClick={() => void runCount(removeComments(), "comment(s) removed")}>
                Comments
              </Button>
              <Button onClick={() => void runCount(removeHyperlinks(), "hyperlink(s) removed")}>
                Hyperlinks
              </Button>
              <Button onClick={() => void runCount(removeBookmarks(), "bookmark(s) removed")}>
                Bookmarks
              </Button>
            </div>
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="select">
          <AccordionHeader>Select</AccordionHeader>
          <AccordionPanel>
            <Button onClick={() => void runSelect(selectTables(), "table(s)")}>
              Select Tables
            </Button>
            <Button onClick={() => void runSelect(selectEmptyParagraphs(), "empty paragraph(s)")}>
              Empty Paragraphs
            </Button>
            <Button onClick={() => void runSelect(selectHeadings(), "heading(s)")}>
              Headings
            </Button>
            <Button onClick={() => void runSelect(selectShortParagraphs(), "short paragraph(s)")}>
              Short Paragraphs
            </Button>
            <Button onClick={() => void runSelect(selectPictures(), "picture(s)")}>
              Pictures
            </Button>
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="tables-pictures">
          <AccordionHeader>Tables &amp; Pictures</AccordionHeader>
          <AccordionPanel>
            <Dropdown
              value={tableSeparator}
              selectedOptions={[tableSeparator]}
              aria-label="Table to text separator"
              onOptionSelect={(_, data) => {
                if (data.optionValue) {
                  setTableSeparator(data.optionValue as "tab" | "paragraph" | "comma");
                }
              }}
            >
              <Option value="tab">Tab</Option>
              <Option value="paragraph">Paragraph</Option>
              <Option value="comma">Comma</Option>
            </Dropdown>
            <Button
              onClick={() => void runCount(convertTableToText(tableSeparator), "table(s) converted")}
            >
              Table → Text
            </Button>
            <ToggleButton
              checked={repeatHeaders}
              onClick={() => {
                const next = !repeatHeaders;
                setRepeatHeaders(next);
                void run(repeatHeaderRows(next));
              }}
            >
              Repeat Header Rows
            </ToggleButton>
            <SpinButton
              defaultValue={100}
              min={10}
              max={400}
              aria-label="Image size percent"
              onChange={(_, data) => {
                if (data.value != null) setResizePercent(data.value);
              }}
            />
            <Button
              onClick={() => void runCount(resizeImages(resizePercent), "image(s) resized")}
            >
              Resize Images
            </Button>
            <Button onClick={() => void run(controlWidows("selection"))}>
              Widow Control
            </Button>
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </div>
  );
}
