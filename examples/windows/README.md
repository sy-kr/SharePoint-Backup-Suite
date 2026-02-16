# Windows Task Scheduler Examples

Scheduled task XML definitions for running SharePoint backups automatically.

## Setup

### 1. Set environment variables (machine-level)

Open an **elevated** PowerShell prompt:

```powershell
[Environment]::SetEnvironmentVariable('TENANT_ID',     'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx', 'Machine')
[Environment]::SetEnvironmentVariable('CLIENT_ID',     'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx', 'Machine')
[Environment]::SetEnvironmentVariable('CLIENT_SECRET', 'your-secret-value',                    'Machine')
```

> **Tip:** For tighter security, store the secret in Windows Credential Manager or Azure Key Vault and retrieve it in a wrapper script instead of a plain environment variable.

### 2. Edit the XML files

Update each XML file with your actual values:

| Placeholder | Replace with |
|-------------|-------------|
| `C:\spbackup\` | Path to your cloned repo |
| `C:\backups\lists` / `loop` / `docs` | Desired output directories |
| `https://contoso.sharepoint.com/sites/team` | Your SharePoint site URL |
| `Tasks` | Your list display name (list backup only) |
| `Documents` | Your library display name (library backup only) |
| `https://loop.cloud.microsoft/p/...` | Your Loop workspace URL (loop backup only) |
| `C:\Program Files\PowerShell\7\pwsh.exe` | Path to `pwsh.exe` if non-default |

### 3. Import the tasks

```powershell
# From an elevated prompt:
schtasks /Create /TN "SPBackup-List"    /XML "examples\windows\SPBackup-List.xml"
schtasks /Create /TN "SPBackup-Loop"    /XML "examples\windows\SPBackup-Loop.xml"
schtasks /Create /TN "SPBackup-Library" /XML "examples\windows\SPBackup-Library.xml"
```

Or import via **Task Scheduler GUI**: Action → Import Task → select the `.xml` file.

### 4. Test

```powershell
schtasks /Run /TN "SPBackup-List"
schtasks /Run /TN "SPBackup-Library"
```

Check output in the configured backup directories and review logs under `<out>/logs/`.

## Schedule

All three tasks default to **every 6 hours**, staggered by 15 minutes:

| Task | Start Time |
|------|-----------|
| SPBackup-List | 02:00 |
| SPBackup-Loop | 02:15 |
| SPBackup-Library | 02:30 |

Edit `<StartBoundary>` and `<Interval>` in the XML to change the schedule.

## Managing tasks

```powershell
# List status
schtasks /Query /TN "SPBackup-List"
schtasks /Query /TN "SPBackup-Library"

# Disable
schtasks /Change /TN "SPBackup-Library" /Disable

# Delete
schtasks /Delete /TN "SPBackup-Library" /F
```
