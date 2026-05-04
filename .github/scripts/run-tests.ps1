param(
  [string]$Workbook = "test-runner.xlsb",
  [string]$Macro = "TestModule.RunAllTests",
  [switch]$Export   # 指定すると npx xvba export を先に実行する
)

# Exit codes:
# 0 = pass
# 2 = tests failed (macro returned false)
# 3 = exception during test run
# 4 = export failed
# 5 = workbook open failed / not found

function Write-ResultAndExit($obj, $code) {
  $obj | ConvertTo-Json -Depth 5 | Out-File -FilePath "test-result.json" -Encoding utf8
  exit $code
}

try {
  if ($Export) {
    Write-Host "Exporting vba-files -> $Workbook using npx @localsmart/xvba-cli..."
    & npx @localsmart/xvba-cli export --src vba-files --target $Workbook
    $ec = $LASTEXITCODE
    if ($ec -ne 0) {
      Write-Host "Export failed (exit $ec)"
      Write-ResultAndExit @{ error = "export_failed"; exit = $ec; timestamp = (Get-Date).ToString("o") } 4
    }
  }

  $fullPath = Resolve-Path -Path $Workbook -ErrorAction SilentlyContinue
  if (-not $fullPath) {
    Write-Host "Workbook not found: $Workbook"
    Write-ResultAndExit @{ error = "workbook_not_found"; workbook = $Workbook } 5
  }

  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false

  $wb = $xl.Workbooks.Open($fullPath.Path)

  try {
    Write-Host "Running macro: $Macro"
    $result = $xl.Run($Macro)
    $passed = [bool]$result
    Write-Host ("Macro returned: {0}" -f $result)

    $out = @{
      passed = $passed
      macro = $Macro
      workbook = $fullPath.Path
      timestamp = (Get-Date).ToString("o")
    }
    Write-ResultAndExit $out (if ($passed) { 0 } else { 2 })
  } catch {
    $err = @{
      error = "macro_exception"
      message = $_.Exception.Message
      stack = $_.Exception.StackTrace
      macro = $Macro
      timestamp = (Get-Date).ToString("o")
    }
    Write-ResultAndExit $err 3
  } finally {
    if ($wb -ne $null) {
      $wb.Close($false)
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
    if ($xl -ne $null) {
      $xl.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
} catch {
  $err = @{
    error = "unexpected"
    message = $_.Exception.Message
    stack = $_.Exception.StackTrace
    timestamp = (Get-Date).ToString("o")
  }
  Write-ResultAndExit $err 3
}
