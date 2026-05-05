param(
  [string]$Workbook = "test-runner.xlsb",
  [string]$VbaFolder = "vba-files",
  [switch]$RemoveExisting
)

try {
  $wbPath = (Resolve-Path -Path $Workbook -ErrorAction Stop).ProviderPath
} catch {
  Write-Error "Workbook not found: $Workbook"
  exit 1
}

if (-not (Test-Path -Path $VbaFolder)) {
  Write-Error "VBA folder not found: $VbaFolder"
  exit 2
}

Write-Host "Importing VBA modules from '$VbaFolder' into workbook '$wbPath'..."

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($wbPath)
  try {
    $vbProj = $wb.VBProject
    $files = Get-ChildItem -Path $VbaFolder -Recurse -Include *.bas,*.cls,*.frm -File | Sort-Object FullName
    foreach ($f in $files) {
      $modName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
      if ($RemoveExisting) {
        try {
          $existing = $vbProj.VBComponents.Item($modName)
          if ($existing -ne $null) {
            $vbProj.VBComponents.Remove($existing)
            Write-Host "Removed existing module: $modName"
          }
        } catch { }
      }
      try {
        # Create a temp copy with CRLF and system encoding to avoid VBProject.Import issues
        $ext = [System.IO.Path]::GetExtension($f.FullName)
        $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName() + $ext)
        $content = Get-Content -Raw -Path $f.FullName -ErrorAction Stop
        $content = $content -replace "(\r?\n)", "`r`n"
        $content | Out-File -FilePath $tmp -Encoding Default
        $vbProj.VBComponents.Import($tmp)
        Write-Host "Imported: $($f.FullName)"
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
      } catch {
        Write-Warning "Import failed for $($f.FullName): $_"
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
      }
    }
    $wb.Save()
    Write-Host "Workbook saved: $wbPath"
  } finally {
    $wb.Close($false)
  }
} finally {
  $xl.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

exit 0
