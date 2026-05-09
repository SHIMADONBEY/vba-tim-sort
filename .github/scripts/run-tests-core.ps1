param(
  [string]$Workbook = "test-runner.xlsb",
  [string]$Macro = "TestTimSort.RunAll_Headless",
  [string]$OutPath = "test-result.json",
  [int]$TimeoutSeconds = 30,
  [switch]$Export,
  [switch]$AllowAutomation
)

# Internal script for running tests with COM interaction, called by run-tests.ps1 in a child process to isolate COM and host effects.
$global:__ResultWritten = $false
$global:__ResultExitCode = 0

# pre-run setup: generate a unique RunId and timestamp for this test run, and determine the per-run output path with a timestamp and GUID to avoid conflicts between runs
$global:__RunId = [System.Guid]::NewGuid().ToString()
$global:__StartTime = Get-Date
$timestamp = $global:__StartTime.ToString("yyyyMMddTHHmmssfff")
$baseName = ($OutPath -replace '\.json$','')
$RunOutPath = Join-Path (Get-Location) ("{0}.{1}.{2}.json" -f $baseName, $timestamp, $global:__RunId)

# debug log for tracing the test run steps and timing, especially around file writing and COM interactions
"{0} RunId={1} PID={2} Start={3}" -f ((Get-Date).ToString("o"), $global:__RunId, $PID, $global:__StartTime.ToString("o")) | Out-File -FilePath .\test-result.write-debug.log -Append -Encoding utf8

function Write-ResultAndExit($obj, $code) {
  if ($global:__ResultWritten) { return }
  $time = (Get-Date).ToString("o")
  $debug = ".\test-result.write-debug.log"

  "$time ENTER Write-ResultAndExit code=$code runId=$global:__RunId PID=$PID" | Out-File -FilePath $debug -Append -Encoding utf8

  try {
    # This section does not perform JSON conversion; it assumes the caller has already serialized the object to a string if needed. This allows for more flexible output formats and avoids issues with ConvertTo-Json depth and formatting.
    if ($obj -is [String]) {
      $outText = $obj
    } else {
      # if $obj is not a string, convert it to a string representation using Out-String. This is a fallback for cases where the caller did not serialize the object to JSON, allowing for more flexible output formats.
      $outText = $obj | Out-String
    }

    # Write to a temp file first to avoid issues with concurrent writes or file locks, then move to the final destination. This is more robust especially when multiple runs might be writing to the same output path pattern.
    $tmp = "$RunOutPath.tmp.$([System.Guid]::NewGuid().ToString())"
    "$time Writing tmp -> $tmp (size: $($outText.Length))" | Out-File -FilePath $debug -Append -Encoding utf8
    $outText | Out-File -FilePath $tmp -Encoding utf8 -Force
    "$time After Out-File Test-Path tmp: $([bool](Test-Path $tmp)); tmp-size: $((Get-Item $tmp).Length)" | Out-File -FilePath $debug -Append -Encoding utf8

    try {
      Move-Item -Path $tmp -Destination $RunOutPath -Force
      "$time Move-Item succeeded: $tmp -> $RunOutPath" | Out-File -FilePath $debug -Append -Encoding utf8
    } catch {
      "$time Move-Item failed: $($_.Exception.Message) -- trying Copy+Remove" | Out-File -FilePath $debug -Append -Encoding utf8
      Copy-Item -Path $tmp -Destination $RunOutPath -Force
      Remove-Item -Path $tmp -Force
      "$time Copy+Remove fallback succeeded" | Out-File -FilePath $debug -Append -Encoding utf8
    }

    "$time Wrote $RunOutPath with code $code" | Out-File -FilePath .\test-result.log -Append -Encoding utf8
    $global:__ResultWritten = $true
    $global:__ResultExitCode = $code

    return
  } catch {
    "$time Write-ResultAndExit CATCH: $($_.Exception.Message)" | Out-File -FilePath $debug -Append -Encoding utf8
    $global:__ResultWritten = $true
    $global:__ResultExitCode = 3
    try {
      $dump = @{ error="write_fatal"; message=$_.Exception.Message; raw = $_ | Out-String; timestamp=$time }
      $dump | ConvertTo-Json -Depth 5 | Out-File -FilePath ($RunOutPath + ".writerror.json") -Encoding utf8 -Force
    } catch {
      "$time Failed to write writerror.json: $($_.Exception.Message)" | Out-File -FilePath $debug -Append -Encoding utf8
    }
    return
  }
}

try {
  # --- optional: import vba-files into workbook ---
  if ($Export) {
    Write-Host "Importing vba-files -> $Workbook using COM importer script..."
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $importScript = Join-Path $scriptDir "import-vba-into-workbook.ps1"
    if (-not (Test-Path $importScript)) {
      Write-ResultAndExit @{ error = "importer_missing"; path = $importScript; timestamp = (Get-Date).ToString("o") } 4
      Exit $global:__ResultExitCode
    }
    & $importScript -Workbook $Workbook -VbaFolder "vba-files" -RemoveExisting
    $ec = $LASTEXITCODE
    if ($ec -ne 0) {
      Write-ResultAndExit @{ error = "import_failed"; exit = $ec; timestamp = (Get-Date).ToString("o") } 4
      Exit $global:__ResultExitCode
    }
  }

  $fullPath = Resolve-Path -Path $Workbook -ErrorAction SilentlyContinue
  if (-not $fullPath) {
    Write-ResultAndExit @{ error = "workbook_not_found"; workbook = $Workbook; timestamp = (Get-Date).ToString("o") } 5
    Exit $global:__ResultExitCode
  }

  # COM interaction to open Excel, run the specified macro, and handle the output. The macro is expected to write its results to $RunOutPath, which this script will monitor and read. The script also includes robust error handling and cleanup to ensure that COM objects are released properly.
  $xl = New-Object -ComObject Excel.Application
  if ($AllowAutomation) { $xl.AutomationSecurity = 1 }
  $xl.Visible = $false
  $xl.DisplayAlerts = $false

  $wb = $xl.Workbooks.Open($fullPath.Path)

  try {
    $wb.Activate()
    $macroCall = $Macro
    Write-Host "Running macro: $macroCall with outPath: $RunOutPath"
    $result = $xl.Run($macroCall, $RunOutPath)
    Write-Host ("Macro returned: {0}" -f $result)

    # Wait for the per-run file written by the macro ($RunOutPath)
    $start = Get-Date
    $deadline = $start.AddSeconds($TimeoutSeconds)
    $json = $null
    while ((Get-Date) -lt $deadline) {
      if (-not (Test-Path $RunOutPath)) { Start-Sleep -Milliseconds 200; continue }
      try {
        $raw = Get-Content -Raw -Path $RunOutPath -ErrorAction Stop
        $json = $raw | ConvertFrom-Json -ErrorAction Stop
        break
      } catch {
        Start-Sleep -Milliseconds 200
      }
    }

    if ($null -eq $json) {
      Write-ResultAndExit @{ error = "invalid_json_or_truncated"; message = "Could not parse JSON or file truncated"; timestamp = (Get-Date).ToString("o") } 7
      Exit $global:__ResultExitCode
    }

    # Determine pass/fail from JSON content, with fallbacks to the macro return value or default to false. This allows the VBA macro to have control over the test result while still providing a fallback mechanism in case the JSON output is not as expected.
    $passed = $false
    if ($null -ne $json.passed) { $passed = [bool]$json.passed } elseif ($null -ne $result) { $passed = [bool]$result } else { $passed = $false }

    Write-Host ("Test run completed. Passed: {0}" -f $passed)

    if ($null -ne $json) {
      # Log the test result and the decision made based on the JSON content.
      # This is useful for debugging and tracing the test outcomes, especially when multiple runs are involved.
      $time = (Get-Date).ToString("o")

      if ($passed) {
        $logMessage = "$time Test passed; JSON indicates success. Keeping VBA-authored result file as authoritative: $RunOutPath"
        $global:__ResultExitCode = 0
      } else {
        $logMessage = "$time Test failed; JSON indicates failure. Keeping VBA-authored result file as authoritative: $RunOutPath"
        $global:__ResultExitCode = 2
      }

      Out-File -FilePath .\test-result.write-debug.log -Append -Encoding utf8 -InputObject $logMessage
      $global:__ResultWritten = $true
      Exit $global:__ResultExitCode
    }

    # If we reach this point, it means the JSON was not valid or did not contain the expected 'passed' property.
    # We will treat this as a failure but still keep the VBA-authored result file for debugging.
    Write-ResultAndExit @{ error = "invalid_json"; message = "JSON does not contain expected 'passed' property"; timestamp = (Get-Date).ToString("o") } 7
    Exit $global:__ResultExitCode
  } finally {
    # COM cleanup (this runs even after Write-ResultAndExit)
    Write-Host "Cleaning up COM objects..."
    if ($null -ne $wb) {
      Write-Host "Closing workbook..."
      try { $wb.Close($false) } catch {}
      try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
      $wb = $null
    }
    if ($null -ne $xl) {
      Write-Host "Quitting Excel application..."
      try { $xl.Quit() } catch {}
      try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
      $xl = $null
      Write-Host "Forcing garbage collection to clean up COM references..."
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Host "Cleanup complete."
  }
} catch {
  # If a result has already been written, exit with that code
  if ($global:__ResultWritten) { Exit $global:__ResultExitCode }

  $err = @{ error = "unexpected"; message = $_.Exception.Message; stack = $_.Exception.StackTrace; timestamp = (Get-Date).ToString("o") }
  Write-ResultAndExit $err 3
  Exit $global:__ResultExitCode
}
