# outlookmAnAger

Local PowerShell WPF tool to help users manage Outlook signatures and view/set mailbox/calendar permissions (scaffold/MVP).

Prerequisites
- Windows with PowerShell 5.1
- Outlook 2016/2019/365 installed (this tool uses the Outlook client COM for account enumeration)
- Script execution policy allowing running unsigned scripts (see below)

Running
1. Open PowerShell (recommended: PowerShell 5.1)
2. Navigate to the project folder and run the convenience script:

```powershell
cd C:\Users\nony\Documents\vsco\outlookmAnAger
.\build-and-run.ps1
```

Notes
- This is an unsigned script bundle intended for local use by end users. If your org blocks unsigned scripts, you will need to set ExecutionPolicy accordingly for the CurrentUser scope.
- Signature operations copy the signature files in the current user's %APPDATA%\Microsoft\Signatures folder. Export/Import uses ZIP packages in the `artifacts` folder.
- Permission setting is currently a placeholder (not implemented) — changing mailbox/calendar permissions from a non-admin, client-side tool depends on server-side capabilities. The UI scaffold is present and will call into the permissions module when implemented.

Security
- The tool runs client-side and does not transmit data externally. Review scripts before running in your environment.

Next steps (planned)
- Implement Set-CalendarPermission using Outlook OOM or Exchange remote PowerShell where appropriate.
- Add tests (Pester) and packaging instructions.
