# AD User Export & Audit

AD User Export & Audit is a PowerShell WPF GUI application that helps sysadmins export Active Directory users (disabled, active, or both) within a chosen date range. It uses `Get-ADUser` with `whenChanged` to build CSV reports, with options for OU scoping, name filters, column selection, and live preview. The tool is designed to make leaver audits and account reviews faster and more consistent, without needing to run raw PowerShell commands.

## Features

- WPF GUI with dark theme.
- Date range selection (based on `whenChanged`).
- Filter by user status: disabled / active / both.
- Optional OU scope and DisplayName wildcard.
- Select which columns to export.
- Results preview grid.
- Export to CSV and Out-GridView.

## Requirements

- Windows with PowerShell 5.1.
- RSAT / ActiveDirectory module installed.
- AD read access.

## Usage

```powershell
.\ADUserExportAudit.ps1
```

Optionally, compile to EXE using PS2EXE:

```powershell
Invoke-PS2EXE -InputFile .\ADUserExportAudit.ps1 -OutputFile .\ADUserExportAudit.exe -NoConsole
```
