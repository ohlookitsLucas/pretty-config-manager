# Pretty Configuration Manager
### Mail Utilities for Microsoft Outlook

A lightweight Windows desktop tool that makes managing Outlook signatures and mailbox permissions simple — even if you've never done it before.

---

## What it does

### Signatures
Create and manage your Outlook email signatures in one place.

- **Create, edit, and preview** signatures with a built-in HTML editor and live preview
- **Assign signatures** to individual mailboxes — choose a different signature for new emails vs. replies
- **Export and import** signatures as ZIP packages, so you can back them up or move them to another machine
- **Manage multiple accounts** — if you have more than one email account in Outlook, each one can have its own signature

### Permissions
Control who can see and interact with your mailbox folders.

- **Browse your mailboxes** — all accounts in your current Outlook profile are listed on the left
- **Pick a folder** — select Inbox, Calendar, Contacts, or any subfolder to manage its sharing
- **See who has access** — a clear table shows every person who has been granted access and what they can do
- **Add or change access** — search your organisation's directory by name or email, pick an access level in plain English, and save with one click
- **Automatic parent folder access** — when you grant access to a subfolder, the tool automatically grants "Can view" on parent folders so the user can navigate to it in Outlook
- **Remove access** — select someone from the list and click Remove to revoke their access

### Settings
- **Themes** — choose between Dark (default), Light, or Retro to match your preference
- **Language** — English and German

### Extras
There are a few surprises hidden in here. You'll know how to find them.

---

## Requirements

- **Windows** with PowerShell 5.1 or later
- **Microsoft Outlook 2016, 2019, or Microsoft 365** installed and configured
- Script execution policy that allows running local scripts (see below)

---

## Getting started

1. Open **PowerShell** (version 5.1 recommended)
2. Navigate to the project folder and run the launcher:

```powershell
cd C:\path\to\outlookmAnAger
.\src\PrettyConfMan.ps1
```

> **Execution policy:** If your machine blocks unsigned scripts, run this once first:
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
> ```

---

## Development & testing

The project includes a full offline test harness — no Outlook or Active Directory required.

### Run all tests

```powershell
.\tests\Run-Tests.ps1
```

### Sandbox mode

Launch the app with mock data (3 mailboxes, 8 AD users, realistic folder permissions):

```powershell
.\tests\Start-Sandbox.ps1
```

Use `-NoTeardown` to keep the sandbox environment after closing for inspection.

### Test structure

| File | What it tests |
|---|---|
| `SignatureManager.Tests.ps1` | Signature CRUD, registry operations, export/import |
| `PermissionsManager.Tests.ps1` | Account enumeration, folder permissions, AD search, ancestor auto-grant |
| `Ui.Tests.ps1` | WPF UI automation — create/rename/delete signatures, assign/unassign, preview |

Tests use mock COM objects and a sandbox registry at `HKCU:\Software\outlookmAnAger-TEST` so nothing touches your real Outlook data.

---

## Project structure

```
src/
  PrettyConfMan.ps1        # App launcher
  MainWindow.xaml           # WPF layout
  Ui.ps1                    # UI code-behind
  SignatureManager.psm1     # Signature file/registry operations
  PermissionsManager.psm1   # Outlook COM permissions + AD search
  Theme.ps1                 # Theme definitions
  Language.ps1              # EN/DE translations
  extras/                   # Easter egg scripts
tests/
  Run-Tests.ps1             # Run full Pester suite
  Start-Sandbox.ps1         # Launch app with mock data
  UiTestHelper.psm1         # WPF UI test infrastructure
  Mocks/                    # Mock Outlook COM + AD objects
  fixtures/                 # Test signature files + registry data
```

---

## Security & privacy

- All operations run **locally on your machine** — nothing is sent over the network
- Signature files are stored in your `%APPDATA%\Microsoft\Signatures` folder as normal
- Mailbox permission changes are applied directly through the Outlook application on your computer
- Review the source scripts before running if required by your organisation's policy
