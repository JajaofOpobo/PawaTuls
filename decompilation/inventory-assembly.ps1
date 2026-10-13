# inventory-assembly.ps1 — list types/methods in Kutools.dll via reflection (no code execution).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File .\decompilation\inventory-assembly.ps1 [dllPath] [outFile]
param(
  [string]$dllPath = "C:\Program Files (x86)\ExtendOffice\WordKutools\Kutools\Kutools.dll",
  [string]$outFile = ""
)
$ErrorActionPreference = "Continue"
$libDir = Split-Path $dllPath
$resolver = {
  param($sender, $args)
  $name = (New-Object System.Reflection.AssemblyName($args.Name)).Name + ".dll"
  $candidate = Join-Path $libDir $name
  if (Test-Path $candidate) {
    return [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($candidate)
  }
  return $null
}
$handler = [System.ResolveEventHandler]{
  param($sender, $args)
  $name = (New-Object System.Reflection.AssemblyName($args.Name)).Name + ".dll"
  $candidate = Join-Path $libDir $name
  if (Test-Path $candidate) {
    return [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($candidate)
  }
  return $null
}
[System.AppDomain]::CurrentDomain.add_ReflectionOnlyAssemblyResolve($handler)
try {
  $asm = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($dllPath)
} catch {
  Write-Host ("LOAD FAILED: " + $_.Exception.Message)
  exit 1
}
Write-Host ("Assembly: " + $asm.FullName)
$types = @()
try { $types = $asm.GetTypes() }
catch {
  Write-Host ("GetTypes threw: " + $_.Exception.GetType().FullName + " :: " + $_.Exception.Message)
  Write-Host ("LoaderExceptions: " + ($_.Exception.LoaderExceptions | Select-Object -First 5 | ForEach-Object { $_.Message } | Out-String))
  try { $types = $_.Exception.Types | Where-Object { $_ -ne $null } } catch { $types = @() }
}
Write-Host ("Types found: " + $types.Count)
$interesting = $types | Where-Object {
  $_.Name -match "Font|Paragraph|Footnote|Format|Style|Case|Indent|Spacing|Clear|Delete|Find|Convert|Character|Number|Caption|Table|Text"
} | Select-Object FullName
Write-Host "---INTERESTING TYPES---"
$interesting | ForEach-Object { Write-Host $_.FullName }
Write-Host ("interesting count: " + @($interesting).Count)
# Obfuscation gauge: single-char or unreadable type names in top-level namespaces
$weird = $types | Where-Object { $_.Name -match "^[a-zA-Z0-9_]{1,2}$" }
Write-Host ("single/double-char type names (obfuscation gauge): " + @($weird).Count)
if ($outFile -ne "") {
  $types | ForEach-Object { $_.FullName } | Set-Content $outFile
  Write-Host ("full list saved: " + $outFile)
}
