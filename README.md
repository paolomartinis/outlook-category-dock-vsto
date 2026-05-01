# Category Dock VSTO

Outlook classic VSTO add-in for category assignment, category management, and category-based search.

## Requirements

- Outlook classic for Windows.
- Microsoft Visual Studio Tools for Office Runtime.
- For development builds: Visual Studio 2022 with Office/SharePoint workload.

## Development Install

Close Outlook, then run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

Reopen Outlook classic and look for **Category Dock VSTO** in COM add-ins and **Category Dock** in the Home ribbon.

## Package

Create a distributable zip:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\package.ps1 -Version 1.0
```

The package is created under `dist\CategoryDockVsto-1.0.zip`.

## Package Install

Extract the zip, close Outlook, then run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-package.ps1
```

To uninstall:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1
```
