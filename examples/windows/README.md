# Windows Scheduling & Credential Management

## Quick start

1. **Copy files to the project root** (same directory as `spbackup.ps1`):
   ```powershell
   Copy-Item examples\Run-Backups.ps1          .
   Copy-Item examples\windows\Load-Credentials.ps1  .
   Copy-Item examples\windows\Setup-Credentials.ps1 .
   Copy-Item examples\windows\Run-Backups.cmd       .
   ```
2. **Store credentials** (one-time) — `pwsh .\Setup-Credentials.ps1`
3. **Edit jobs & SMTP settings** in `Run-Backups.ps1`
4. **Test interactively** — `pwsh .\Run-Backups.ps1`
5. **Schedule** — create a Task Scheduler task that runs `Run-Backups.cmd`

---

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

## Configuring backup jobs

Open `Run-Backups.ps1` and edit the `$Jobs` array. Each job is a hashtable with
three keys:

| Key | Description |
|-----|-------------|
| `Label` | Friendly name shown in logs and email reports |
| `Tool` | Which backup tool to use: `library`, `list`, or `loop` |
| `Args` | Array of arguments passed to `spbackup.ps1 <tool>` |

### Document library backup

Backs up all files and folders from a SharePoint document library.

```powershell
@{
    Label = 'Team Site - Documents'
    Tool  = 'library'
    Args  = @('backup',
               '--url',     'https://contoso.sharepoint.com/sites/TeamSite',
               '--library', 'Documents',
               '--out',     (Join-Path $BackupBase 'TeamSite'))
}
```

Required arguments: `--url` (site URL), `--library` (library display name),
`--out` (output directory).

> **Tip:** Run `spbackup.ps1 library enumerate --url <site-url>` to list all
> document libraries on a site if you're not sure of the library name.

### Microsoft List backup

Backs up list items to CSV and downloads any attachments.

```powershell
@{
    Label = 'Project Tasks'
    Tool  = 'list'
    Args  = @('backup',
               '--url',  'https://contoso.sharepoint.com/sites/Projects/Lists/Tasks/AllItems.aspx',
               '--out',  (Join-Path $BackupBase 'Lists\ProjectTasks'))
}
```

Required arguments: `--url` (list URL — the list name is extracted
automatically), `--out` (output directory).

You can also specify `--list "Tasks"` or `--list-id <GUID>` explicitly instead
of relying on URL parsing.

> **Tip:** Run `spbackup.ps1 list enumerate --url <site-url>` to list all
> lists on a site.

### Microsoft Loop workspace backup

Backs up all pages from a Loop workspace as `.loop` (raw), HTML, and Markdown.

```powershell
@{
    Label = 'Team Wiki'
    Tool  = 'loop'
    Args  = @('backup',
               '--url',  'https://loop.cloud.microsoft/p/<your-workspace-id>',
               '--out',  (Join-Path $BackupBase 'TeamWiki'))
}
```

Required arguments: `--url` (Loop workspace URL from your browser), `--out`
(output directory).

Optional flags: `--no-html` (skip HTML export), `--no-md` (skip Markdown
export).

> **Tip:** Run `spbackup.ps1 loop resolve --url <loop-url>` to verify the URL
> resolves to a valid SharePoint storage location.

### Common optional flags

These can be appended to any job's `Args` array:

| Flag | Effect |
|------|--------|
| `--verbose` | Detailed per-file logging (visible interactively; suppressed in headless mode) |
| `--verify` | Run a verify pass after backup completes |
| `--concurrency <n>` | Max parallel downloads (default varies by tool) |

> **Tip:** `--verbose` is safe for both interactive and scheduled runs.
> In interactive mode the output streams to the console in real time.
> In headless mode (Task Scheduler) the output only goes to per-job log
> files — there is no pipeline backpressure.

### Disabling a job

Comment out the entire block:

```powershell
# @{
#     Label = 'Disabled Job'
#     Tool  = 'library'
#     Args  = @('backup', '--url', '...', '--library', '...', '--out', '...')
# }
```

---

## How it works

All scripts should be in the **project root** alongside `spbackup.ps1`:

```
Task Scheduler --> Run-Backups.cmd --> Run-Backups.ps1
                   (thin launcher)     |
                                       +--> Load-Credentials.ps1
                                       |    (Credential Manager -> env vars)
                                       |
                                       +--> spbackup.ps1 library backup ...
                                       +--> spbackup.ps1 list backup ...
                                       +--> spbackup.ps1 loop backup ...
                                       |
                                       +--> Email report (SMTP)
```

**Auth flow:**
1. `Load-Credentials.ps1` reads `TENANT_ID` + `CLIENT_ID` from Credential Manager (always)
2. If `CLIENT_SECRET` is stored -> sets `$env:CLIENT_SECRET` -> client-secret auth
3. If no secret -> `$env:CLIENT_SECRET` stays unset -> `Get-GraphToken` falls back to certificate auth
4. SharePoint REST API (list attachments) always uses certificate auth regardless

**Output capture:**
Each job's output is written to a per-job log file under
`<BackupBase>/orchestrator-logs/job-<label>.log`.

- **Interactive mode** (default): `Tee-Object` streams output to both the
  console and the log file, so `--verbose` logging is visible in real time.
- **Headless mode** (`-Headless`): `Out-File` sends everything to the log
  only — zero console output. Progress bars are also suppressed.

The `.cmd` launcher and Task Scheduler should always use `-Headless`.

Credentials exist **only** in:
- Windows Credential Manager (encrypted at rest)
- Environment variables (current process only, never inherited by child shells)

---

## Task Scheduler Setup

### 1. Edit Run-Backups.ps1

Update the configuration section at the top:

- `$SpBackupRoot` — path to the folder containing `spbackup.ps1`
- `$BackupBase` — root output directory
- `$SmtpServer` / `$SmtpPort` / `$SmtpFrom` / `$SmtpTo` — email settings
- `$Jobs` — your backup job definitions (see above)

### 2. Test interactively

Run from a PowerShell terminal to make sure everything works:

```powershell
pwsh .\Run-Backups.ps1
```

This runs in **interactive mode** — you'll see all job output (including
`--verbose` logging) in real time, as well as the summary report.

> The `.cmd` launcher and Task Scheduler use `-Headless`, which suppresses
> all console output. The report is still saved to disk and emailed.

### 3. Create a scheduled task

Open **Task Scheduler** and create a new task:

| Setting | Value |
|---------|-------|
| **Run as** | The same user account that ran `Setup-Credentials.ps1` |
| **Run whether user is logged on or not** | Yes |
| **Trigger** | Daily, e.g. 02:00 AM |
| **Action** | Start a program |
| **Program/script** | `C:\spbackup\Run-Backups.cmd` |
| **Start in** | `C:\spbackup` |

Under **Settings**:
- Set "Stop the task if it runs longer than" to a reasonable limit (e.g. 12 hours)
- Enable "If the task fails, restart every" 1 hour, up to 2 times
- Enable "Start the task only if the computer is on AC power" if applicable

> **Important:** The "Run as" user **must** match the account that ran
> `Setup-Credentials.ps1`. Credential Manager entries are per-user.

---

## Managing the scheduled task

```powershell
# Check status
schtasks /Query /TN "SPBackup-All"

# Run immediately
schtasks /Run /TN "SPBackup-All"

# Disable
schtasks /Change /TN "SPBackup-All" /Disable

# Delete
schtasks /Delete /TN "SPBackup-All" /F
```

---

## Log files

All logs are written to `<BackupBase>/orchestrator-logs/`:

| File | Contents |
|------|----------|
| `backup-report-<timestamp>.txt` | Overall summary (also emailed) |
| `job-<label>.log` | Full output for each individual job |

Failed-job output (last 100 lines) is included in the email report automatically.

---

## Files

Copy these to the project root (same directory as `spbackup.ps1`):

| File | Source | Purpose |
|------|--------|---------|
| `Run-Backups.ps1` | `examples/` | Cross-platform orchestrator (shared with Linux) |
| `Load-Credentials.ps1` | `examples/windows/` | Runtime: reads credentials into `$env:` variables |
| `Setup-Credentials.ps1` | `examples/windows/` | One-time: stores credentials in Credential Manager |
| `Run-Backups.cmd` | `examples/windows/` | Thin launcher for Task Scheduler / double-click |
