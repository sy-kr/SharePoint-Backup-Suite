# Windows Scheduling & Credential Management

## Credential Setup (one-time)

Credentials are stored in **Windows Credential Manager** — encrypted with DPAPI,
tied to your user account, never visible in plain text or process lists.

### 1. Store credentials

Run **once**, interactively, as the **same user account** that will run the
scheduled task:

```powershell
pwsh .\Setup-Credentials.ps1
```

You'll be prompted for:

| Value | Required? | Notes |
|-------|-----------|-------|
| Tenant ID | Yes | Azure AD / Entra tenant GUID |
| Client ID | Yes | App registration client GUID |
| Client Secret | **No** | Press Enter to skip if using certificate-only auth |

If you skip the client secret, `Get-GraphToken` will automatically fall back to
certificate authentication (JWT client assertion). Place your `.pfx` at
`<project-root>/certs/spbackup.pfx` or set `CERT_PATH`.

### 2. Verify stored credentials

```cmd
cmdkey /list:SPBackup:*
```

### 3. Remove credentials

```cmd
cmdkey /delete:SPBackup:TenantId
cmdkey /delete:SPBackup:ClientId
cmdkey /delete:SPBackup:ClientSecret
```

---

## How it works

```
┌──────────────────┐     ┌──────────────────┐     ┌──────────────────────────┐
│  Task Scheduler   │────▶│  Run-Backups.cmd  │────▶│    Run-Backups.ps1       │
│  (triggers)       │     │  (thin launcher)  │     │  ┌────────────────────┐  │
└──────────────────┘     └──────────────────┘     │  │Load-Credentials.ps1│  │
                                                   │  │ Cred Manager → env │  │
                                                   │  └────────────────────┘  │
                                                   │         │                │
                                                   │         ▼                │
                                                   │  spbackup.ps1 list …    │
                                                   │  spbackup.ps1 loop …    │
                                                   │  spbackup.ps1 library … │
                                                   │         │                │
                                                   │         ▼                │
                                                   │  Email report (SMTP)     │
                                                   └──────────────────────────┘
```

**Auth flow:**
1. `Load-Credentials.ps1` reads `TENANT_ID` + `CLIENT_ID` from Credential Manager (always)
2. If `CLIENT_SECRET` is stored → sets `$env:CLIENT_SECRET` → client-secret auth
3. If no secret → `$env:CLIENT_SECRET` stays unset → `Get-GraphToken` falls back to certificate auth
4. SharePoint REST API (list attachments) always uses certificate auth regardless

Credentials exist **only** in:
- Windows Credential Manager (encrypted at rest)
- Environment variables (current process only, never inherited by child shells)

---

## Task Scheduler Setup

### Option A: Single orchestrated task (recommended)

Edit `Run-Backups.ps1` with your job list, then create one task:

```powershell
$action   = New-ScheduledTaskAction -Execute "C:\Backup\examples\windows\Run-Backups.cmd"
$trigger  = New-ScheduledTaskTrigger -Daily -At 2:00AM
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 8) `
    -StartWhenAvailable `
    -DontStopOnIdleEnd

Register-ScheduledTask -TaskName "SPBackup-All" `
    -Action $action -Trigger $trigger -Settings $settings `
    -User "DOMAIN\svc-backup" -Password (Read-Host -AsSecureString "Password")
```

> **Important:** The `-User` must be the same account that ran `Setup-Credentials.ps1`.

### Option B: Import individual XML tasks

Update each XML file with your paths, then import:

```powershell
schtasks /Create /TN "SPBackup-List"    /XML "examples\windows\SPBackup-List.xml"
schtasks /Create /TN "SPBackup-Loop"    /XML "examples\windows\SPBackup-Loop.xml"
schtasks /Create /TN "SPBackup-Library" /XML "examples\windows\SPBackup-Library.xml"
```

### Option C: Command-line (schtasks.exe)

```cmd
schtasks /create /tn "SPBackup-All" ^
    /tr "C:\Backup\examples\windows\Run-Backups.cmd" ^
    /sc daily /st 02:00 ^
    /ru DOMAIN\svc-backup /rp * ^
    /rl HIGHEST
```

---

## Schedule

Default staggered schedule for individual tasks:

| Task | Start Time |
|------|-----------|
| SPBackup-List | 02:00 |
| SPBackup-Loop | 02:15 |
| SPBackup-Library | 02:30 |

Edit `<StartBoundary>` and `<Interval>` in the XML to change.

---

## Managing tasks

```powershell
# Status
schtasks /Query /TN "SPBackup-All"

# Disable
schtasks /Change /TN "SPBackup-All" /Disable

# Delete
schtasks /Delete /TN "SPBackup-All" /F
```

---

## Files

| File | Purpose |
|------|---------|
| `Setup-Credentials.ps1` | One-time: stores credentials in Credential Manager |
| `Load-Credentials.ps1` | Runtime: reads credentials → `$env:` variables |
| `Run-Backups.ps1` | Orchestrator: runs jobs, collects results, emails report |
| `Run-Backups.cmd` | Thin launcher for Task Scheduler / double-click |
| `SPBackup-*.xml` | Individual Task Scheduler XML templates |
