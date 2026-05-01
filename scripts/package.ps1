param(
  [string]$Version = "1.1"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
$project = Join-Path $root "CategoryDockVsto.csproj"
$release = Join-Path $root "bin\Release"
$dist = Join-Path $root "dist"
$packageName = "CategoryDockVsto-$Version"
$packageRoot = Join-Path $dist $packageName
$scriptsRoot = Join-Path $packageRoot "scripts"
$zip = Join-Path $dist "$packageName.zip"

if (Test-Path $packageRoot) {
  Remove-Item -LiteralPath $packageRoot -Recurse -Force
}

if (Test-Path $zip) {
  Remove-Item -LiteralPath $zip -Force
}

& $msbuild $project /p:Configuration=Release /v:minimal
if ($LASTEXITCODE -ne 0) {
  throw "Release build failed."
}

New-Item -ItemType Directory -Path $packageRoot | Out-Null
New-Item -ItemType Directory -Path $scriptsRoot | Out-Null

Get-ChildItem -Path $release -File | Where-Object Extension -ne ".pdb" | Copy-Item -Destination $packageRoot -Force
Copy-Item -Path (Join-Path $PSScriptRoot "install-package.ps1") -Destination $scriptsRoot -Force
Copy-Item -Path (Join-Path $PSScriptRoot "Install Category Dock.cmd") -Destination $packageRoot -Force
Copy-Item -Path (Join-Path $PSScriptRoot "uninstall.ps1") -Destination $scriptsRoot -Force
Copy-Item -Path (Join-Path $root "README.md") -Destination $packageRoot -Force

Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zip -Force

Write-Host "Created package: $zip"
