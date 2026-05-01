param(
  [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
$project = Join-Path $root "CategoryDockVsto.csproj"

& $msbuild $project /p:Configuration=$Configuration /v:minimal
if ($LASTEXITCODE -ne 0) {
  throw "Build failed."
}

$vsto = Join-Path $root "bin\$Configuration\CategoryDockVsto.vsto"
if (-not (Test-Path $vsto)) {
  throw "VSTO manifest not found: $vsto"
}

$manifest = ((New-Object System.Uri($vsto)).AbsoluteUri + "|vstolocal")
$addin = "HKCU\Software\Microsoft\Office\Outlook\Addins\CategoryDockVsto"

& cmd.exe /c "reg delete ""HKCU\Software\Microsoft\Office\Outlook\Addins\CategoryDockClassic.Connect"" /f 2>nul" | Out-Null
& cmd.exe /c "reg delete ""HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\CategoryDockClassic.Connect"" /f 2>nul" | Out-Null
& cmd.exe /c "reg delete ""HKCU\Software\Microsoft\Office\Outlook\AddinsData\CategoryDockClassic.Connect"" /f 2>nul" | Out-Null

& reg.exe add $addin /v FriendlyName /t REG_SZ /d "Category Dock VSTO" /f | Out-Null
& reg.exe add $addin /v Description /t REG_SZ /d "Batch category manager for Outlook classic." /f | Out-Null
& reg.exe add $addin /v LoadBehavior /t REG_DWORD /d 3 /f | Out-Null
& reg.exe add $addin /v Manifest /t REG_SZ /d $manifest /f | Out-Null
& cmd.exe /c "reg delete ""HKCU\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"" /v CategoryDockVsto.Microsoft.Outlook.Explorer /f 2>nul" | Out-Null

Write-Host "Installed Category Dock VSTO for current user."
Write-Host "Manifest: $manifest"
