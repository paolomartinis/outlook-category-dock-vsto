$ErrorActionPreference = "Stop"

& reg.exe delete "HKCU\Software\Microsoft\Office\Outlook\Addins\CategoryDockVsto" /f 2>$null | Out-Null

Write-Host "Uninstalled Category Dock VSTO for current user."
