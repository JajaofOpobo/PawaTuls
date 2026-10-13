# probe-kutools.ps1 — check for an automation API (COMAddIn.Object) and CommandBar controls.
$ErrorActionPreference = "Continue"
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  $ka = $w.COMAddIns.Item("WordKutools.Connect")
  Write-Host ("Connect=" + $ka.Connect)
  try {
    $o = $ka.Object
    if ($null -eq $o) { Write-Host "Object: null (no automation API exposed)" }
    else { $o | Get-Member -MemberType Method | Select-Object -First 10 Name }
  } catch { Write-Host ("Object probe failed: " + $_.Exception.Message) }
  Write-Host "---COMMAND BARS---"
  $found = 0
  foreach ($bar in $w.CommandBars) {
    try {
      $controls = $bar.Controls
    } catch { continue }
    foreach ($c in $controls) {
      try { $cap = $c.Caption } catch { continue }
      if ($cap -like "*utool*" -or $bar.Name -like "*utool*") {
        Write-Host ($bar.Name + " :: " + $cap)
        $found++
      }
    }
  }
  Write-Host ("matches=" + $found)
}
finally {
  $w.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($w) | Out-Null
}
