# make-fixtures.ps1 — Phase 29: generate Word-run fixture documents.
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File .\decompilation\make-fixtures.ps1
$ErrorActionPreference = "Stop"
$root = Get-Location
$fixDir = Join-Path $root "decompilation\fixtures"
New-Item -ItemType Directory -Force -Path $fixDir | Out-Null

$word = New-Object -ComObject Word.Application
try {
  $word.Visible = $false
  $word.DisplayAlerts = 0

  # ── formatting-test.docx ──────────────────────────────────────────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Heading 1")
  $sel.TypeText("Formatting Fixture")
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.TypeText("Plain body paragraph one.")
  $sel.TypeParagraph()
  $sel.Font.Bold = $true
  $sel.TypeText("Bold lead")
  $sel.Font.Bold = $false
  $sel.TypeText(" and ")
  $sel.Font.Italic = $true
  $sel.TypeText("italic tail")
  $sel.Font.Italic = $false
  $sel.TypeText(" mixed.")
  $sel.TypeParagraph()
  $sel.Font.Size = 20
  $sel.Font.Color = 255
  $sel.TypeText("Big red line.")
  $sel.Font.Size = 11
  $sel.Font.Color = 0
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 2  # center
  $sel.TypeText("Centered line")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 3  # right
  $sel.TypeText("Right line")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 3
  $sel.TypeText("justify  me  doubLES  and	tabs")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 0
  $sel.TypeText("lowercase line for UPPERCASE test")
  $sel.TypeParagraph()
  $sel.TypeText("SENTENCE CASE SOURCE. second sentence here! third one?")
  $sel.TypeParagraph()
  $sel.TypeText("First manual line")
  $sel.TypeText([char]11)  # manual line break
  $sel.TypeText("second manual line")
  $sel.TypeParagraph()
  $sel.TypeText("1st 2nd 3rd 4th 21st 200th ordinals line")
  $sel.TypeParagraph()
  $sel.TypeText("--------------------------------------------------")
  $sel.TypeParagraph()
  $sel.TypeText("")
  $sel.TypeParagraph()  # truly empty
  $sel.TypeText("   ")
  $sel.TypeParagraph()  # whitespace-only
  # native numbered list
  $sel.Range.ListFormat.ApplyNumberDefault() | Out-Null
  $sel.TypeText("native one")
  $sel.TypeParagraph()
  $sel.TypeText("native two")
  $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers() | Out-Null
  # table with an empty row
  $rng = $sel.Range
  $table = $doc.Tables.Add($rng, 3, 2)
  $table.Cell(1, 1).Range.Text = "A1"
  $table.Cell(1, 2).Range.Text = "B1"
  $table.Cell(2, 1).Range.Text = ""
  $table.Cell(2, 2).Range.Text = ""
  $table.Cell(3, 1).Range.Text = "A3"
  $table.Cell(3, 2).Range.Text = "B3"
  # genuine footnote + endnote (must be left alone by scanner)
  $sel.EndKey(6) | Out-Null  # wdStory
  $sel.TypeParagraph()
  $sel.TypeText("Body text with a genuine footnote")
  $fnRange = $sel.Range
  $fn = $doc.Footnotes.Add($fnRange)
  $fn.Range.Text = "Genuine footnote text."
  $word.ActiveWindow.View.SeekView = 0  # back to main story
  $sel.EndKey(6) | Out-Null
  $sel.TypeText(" and genuine endnote")
  $enRange = $sel.Range
  $en = $doc.Endnotes.Add($enRange)
  $en.Range.Text = "Genuine endnote text."
  $word.ActiveWindow.View.SeekView = 0
  $sel.EndKey(6) | Out-Null
  $sel.TypeText(".")
  # hyperlink + comment + bookmark
  $sel.TypeParagraph()
  $sel.TypeText("See example site ")
  $hlStart = $sel.Range.Start
  $sel.TypeText("link here")
  $hlRange = $doc.Range($hlStart, $sel.Range.Start)
  $doc.Hyperlinks.Add($hlRange, "https://example.com") | Out-Null
  $sel.TypeParagraph()
  $cRange = $sel.Range
  $sel.TypeText("Commented text.")
  $doc.Comments.Add($cRange, "Reviewer note.") | Out-Null
  $bmRange = $sel.Range
  $doc.Bookmarks.Add("FixtureMark", $bmRange) | Out-Null
  # page break + hidden text
  $sel.TypeParagraph()
  $sel.InsertBreak(7) | Out-Null  # wdPageBreak
  $sel.Font.Hidden = $true
  $sel.TypeText("Hidden fixture text.")
  $sel.Font.Hidden = $false
  $sel.TypeParagraph()
  $sel.TypeText("After page break.")
  $doc.SaveAs([string](Join-Path $fixDir "formatting-test.docx"))
  $doc.Close()
  Write-Host "saved formatting-test.docx"

  # ── footnotes-test.docx ───────────────────────────────────────────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.TypeText("Treaty history")
  $sel.Font.Superscript = $true
  $sel.TypeText("1")
  $sel.Font.Superscript = $false
  $sel.TypeText(" shows early agreement")
  $sel.Font.Superscript = $true
  $sel.TypeText("12")
  $sel.Font.Superscript = $false
  $sel.TypeText(" and later text [4] with brackets (5) in parens.")
  $sel.TypeParagraph()
  $sel.TypeText("Unmatched marker")
  $sel.Font.Superscript = $true
  $sel.TypeText("7")
  $sel.Font.Superscript = $false
  $sel.TypeText(" has no block.")
  $sel.TypeParagraph()
  $sel.TypeText("________________________________________")
  $sel.TypeParagraph()
  $sel.TypeText("1 See Vienna Convention on the Law of Treaties.")
  $sel.TypeParagraph()
  $sel.TypeText("12 Continued note text for the twelfth marker spanning lines.")
  $sel.TypeParagraph()
  $sel.TypeText("still part of note twelve.")
  $sel.TypeParagraph()
  $sel.TypeText("[4] See British Yearbook of International Law.")
  $sel.TypeParagraph()
  $sel.TypeText("(5) See American Journal of International Law.")
  $sel.TypeParagraph()
  $sel.TypeText("9 Orphan block with no marker in body.")
  $sel.TypeParagraph()
  $sel.TypeText("[2020] 15 NWLR (Pt. 1721) 123 is a citation, not a footnote.")
  $sel.TypeParagraph()
  $doc.SaveAs([string](Join-Path $fixDir "footnotes-test.docx"))
  $doc.Close()
  Write-Host "saved footnotes-test.docx"

  # ── numbering-test.docx ───────────────────────────────────────────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.Range.ListFormat.ApplyNumberDefault() | Out-Null
  $sel.TypeText("native alpha")
  $sel.TypeParagraph()
  $sel.TypeText("native beta")
  $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers() | Out-Null
  foreach ($t in @("1. manual one", "2. manual two", "2. manual dup", "4. manual gap")) {
    $sel.TypeText($t)
    $sel.TypeParagraph()
  }
  $sel.TypeText("Ordinary 2026 prose with 3.14 decimals, not a list.")
  $sel.TypeParagraph()
  $doc.SaveAs([string](Join-Path $fixDir "numbering-test.docx"))
  $doc.Close()
  Write-Host "saved numbering-test.docx"

  # ── pleadings-test.docx (fixture for later legal re-enable) ───────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.TypeText("The Claimant states as follows:")
  $sel.TypeParagraph()
  $sel.TypeText("First averment paragraph.")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.LeftIndent = 28.35
  $sel.TypeText("Sub-particular under first averment.")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.LeftIndent = 0
  $sel.TypeText("Second averment paragraph.")
  $sel.TypeParagraph()
  $doc.SaveAs([string](Join-Path $fixDir "pleadings-test.docx"))
  $doc.Close()
  Write-Host "saved pleadings-test.docx"

  # ── footnotes-edge-test.docx (hard negatives + boundaries) ────────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Normal")
  # 1. numbered substantive prose with years, no citation structure
  $sel.TypeText("1. In 1919 the committee met and resolved nothing of note.")
  $sel.TypeParagraph()
  $sel.TypeText("2. The 2024 review repeated the exercise at length.")
  $sel.TypeParagraph()
  # 2. bibliography under References heading (must NOT be footnotes)
  $sel.Style = $doc.Styles.Item("Heading 1")
  $sel.TypeText("References")
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.TypeText("1. Adeyemi, Bola. Contract in Nigeria. Lagos: Printwell, 2021.")
  $sel.TypeParagraph()
  $sel.TypeText("2. Smith, John. “Treaties.” Journal of Law, vol. 3, 1919.")
  $sel.TypeParagraph()
  # 3. numbered legal argument, no coherent note block
  $sel.TypeText("1. Reliance is placed on s. 191 and section 2 of the Act.")
  $sel.TypeParagraph()
  $sel.TypeText("2. See the 2020 ruling where “et al.” appears in quoted text.")
  $sel.TypeParagraph()
  $sel.TypeText("3. Full particulars at https://example.invalid/report.")
  $sel.TypeParagraph()
  # 4. ibid chain + later long body paragraph (boundary test)
  $sel.TypeText("Body claim one")
  $sel.Font.Superscript = $true
  $sel.TypeText("1")
  $sel.Font.Superscript = $false
  $sel.TypeText(" and claim two")
  $sel.Font.Superscript = $true
  $sel.TypeText("2")
  $sel.Font.Superscript = $false
  $sel.TypeText(".")
  $sel.TypeParagraph()
  $sel.TypeText("1. Smith, ‘Contract Law’ (2021).")
  $sel.TypeParagraph()
  $sel.TypeText("2. Ibid., at 45.")
  $sel.TypeParagraph()
  $sel.TypeText("This is a long ordinary body paragraph with no note marker and no citation structure whatsoever; it discusses the weather, the harvest, and other matters at considerable length so that any block-boundary rule must stop the footnote cluster before absorbing it into note two.")
  $sel.TypeParagraph()
  # 5. single note + long following paragraph (exception test)
  $sel.TypeText("A lone assertion")
  $sel.Font.Superscript = $true
  $sel.TypeText("3")
  $sel.Font.Superscript = $false
  $sel.TypeText(" stands here.")
  $sel.TypeParagraph()
  $sel.TypeText("3. Okafor (2019) explains the doctrine in a single compact note.")
  $sel.TypeParagraph()
  $sel.TypeText("Another long body paragraph follows the single note; it is ordinary prose about procedure and timelines, running well past four hundred characters in total length so that the single-footnote exception is exercised and the candidate is not rejected merely because a long paragraph follows it.")
  $sel.TypeParagraph()
  # 6. long explanatory footnote + 7. duplicate/missing markers
  $sel.TypeText("Analysis continues")
  $sel.Font.Superscript = $true
  $sel.TypeText("4")
  $sel.Font.Superscript = $false
  $sel.TypeText(" with a repeated marker")
  $sel.Font.Superscript = $true
  $sel.TypeText("4")
  $sel.Font.Superscript = $false
  $sel.TypeText(" and a gap to")
  $sel.Font.Superscript = $true
  $sel.TypeText("6")
  $sel.Font.Superscript = $false
  $sel.TypeText(".")
  $sel.TypeParagraph()
  $sel.TypeText("4. This is a genuinely long explanatory footnote that discusses the history of the doctrine from 1919 through 2024 with multiple citations including Smith (2021), Jones et al., 410 U.S. 113, and commentary in several journals across many lines of detailed analysis that far exceeds ordinary note length.")
  $sel.TypeParagraph()
  $sel.TypeText("6. Short note after the gap.")
  $sel.TypeParagraph()
  # 8. merged body prose + citation in one long paragraph
  $sel.TypeText("The committee sat late into the evening and the minutes record a wide-ranging debate about funding before noting, almost in passing, see M. Placeholder, Integrated Note, 9 Journal 101 (2024), after which discussion returned to scheduling and the paragraph continues as ordinary prose.")
  $sel.TypeParagraph()
  # 9. mixed styles cluster
  $sel.TypeText("Style check")
  $sel.Font.Superscript = $true
  $sel.TypeText("7")
  $sel.Font.Superscript = $false
  $sel.TypeText(" and")
  $sel.Font.Superscript = $true
  $sel.TypeText("8")
  $sel.Font.Superscript = $false
  $sel.TypeText(".")
  $sel.TypeParagraph()
  $sel.TypeText("7. Smith, ‘Contract Law’ (2021) is OSCOLA-like in form.")
  $sel.TypeParagraph()
  $sel.TypeText("8. 410 U.S. 113 is Bluebook-like in form.")
  $sel.TypeParagraph()
  $doc.SaveAs([string](Join-Path $fixDir "footnotes-edge-test.docx"))
  $doc.Close()
  Write-Host "saved footnotes-edge-test.docx"
}
finally {
  $word.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
Get-ChildItem $fixDir | Select-Object Name, @{n='KB';e={[int]($_.Length/1KB)}}
