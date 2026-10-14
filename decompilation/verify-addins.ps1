# verify-addins.ps1 — list Word COM add-ins (post-KuTools-removal check).
$ErrorActionPreference = "Continue"
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  foreach ($a in $w.COMAddIns) {
    Write-Host ($a.ProgId + "  Connect=" + $a.Connect)
  }
}
finally {
  $w.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($w) | Out-Null
}
