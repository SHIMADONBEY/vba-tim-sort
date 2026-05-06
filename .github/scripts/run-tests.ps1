param(
  [string]$Workbook = "test-runner.xlsb",
  [string]$Macro = "TestTimSort.RunAll_Headless",
  [string]$OutPath = "test-result.json",
  [int]$TimeoutSeconds = 30,
  [switch]$Export,
  [switch]$AllowAutomation
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$child = Join-Path $scriptDir "run-tests-core.ps1"

if (-not (Test-Path $child)) {
  Write-Host "Child script not found: $child"
  exit 2
}

# Forward args to child in a fresh PowerShell process to isolate COM and host effects
$argList = @(
  '-NoProfile','-ExecutionPolicy','Bypass','-File', $child,
  '-Workbook', $Workbook,
  '-Macro', $Macro,
  '-OutPath', $OutPath,
  '-TimeoutSeconds', $TimeoutSeconds.ToString()
)
if ($Export) { $argList += '-Export' }
if ($AllowAutomation) { $argList += '-AllowAutomation' }

Write-Host "Launching child process for test run..."
powershell @argList
$ec = $LASTEXITCODE
Write-Host "Child exit code: $ec"
# propagate child's exit code
exit $ec
