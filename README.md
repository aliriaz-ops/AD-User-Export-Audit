# AD User Export & Audit

AD User Export & Audit is a PowerShell WPF GUI application that helps sysadmins export Active Directory users (disabled, active, or both) over a selected date range. It uses `Get-ADUser` with `whenChanged` to build CSV reports, with options for OU scoping, name filters, column selection, and live preview. The tool is designed to make leaver audits and account reviews faster and more consistent, without needing to run raw PowerShell commands.
