# Installation

## Package install

1. Download the release zip.
2. Extract it to a local folder.
3. Close Outlook classic.
4. Double-click `Install Category Dock.cmd`.
5. Reopen Outlook classic.

The add-in should appear as **Category Dock VSTO** in Outlook COM Add-ins, and the dock should open automatically.

## PowerShell install

If the command file is blocked, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-package.ps1
```

## Uninstall

Close Outlook and run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1
```
