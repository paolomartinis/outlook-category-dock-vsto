# Category Dock VSTO

Outlook classic VSTO add-in for category assignment, category management, and category-based search.

## Version

Current development package: **1.1**.

## Documentation

See the [GitHub Wiki](https://github.com/paolomartinis/outlook-category-dock-vsto/wiki) for installation, usage, and release notes.

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
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\package.ps1 -Version 1.1
```

The package is created under `dist\CategoryDockVsto-1.1.zip`.

## Package Install

Extract the zip, close Outlook, then double-click `Install Category Dock.cmd`.

Alternatively, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-package.ps1
```

To uninstall:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1
```

## Changelog

See [CHANGELOG.md](CHANGELOG.md).
