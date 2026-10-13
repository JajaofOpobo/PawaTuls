import React, { useState } from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  Dropdown,
  Option,
  SpinButton,
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
import { collapseMultipleSpaces, transformCase } from "../../formatting/text";
import {
  lineBreaksToParagraphMarks,
  paragraphMarksToLineBreaks,
} from "../../formatting/convert";
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
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </div>
  );
}
