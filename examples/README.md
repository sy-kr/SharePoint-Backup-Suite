# Examples

Example scripts and configuration for automating SharePoint backups.

## Overview

| File / Folder | Description |
|---------------|-------------|
| [`Run-Backups.ps1`](Run-Backups.ps1) | Cross-platform orchestrator — runs multiple backup jobs and sends an email report. Supports `-Headless` for scheduled (silent) runs. |
| [`windows/`](windows/) | Windows-specific: Credential Manager integration, `.cmd` launcher, Task Scheduler setup |
| [`linux/`](linux/) | Linux-specific: thin shell wrapper, cron / systemd timer setup |

## Getting started

### 1. Copy files to the project root

Copy the files you need to the **same directory as `spbackup.ps1`**:

```bash
# Always needed
cp examples/Run-Backups.ps1  .

# Windows only
cp examples/windows/Load-Credentials.ps1  .
cp examples/windows/Setup-Credentials.ps1 .
cp examples/windows/Run-Backups.cmd        .

# Linux only
cp examples/linux/run-backups.sh           .
```

> **Why copy?** The scripts use `$PSScriptRoot` to locate `spbackup.ps1` in
> the same directory. Keeping everything together avoids path configuration.

### 2. Configure backup jobs

Open `Run-Backups.ps1` and edit the **CONFIGURATION** section:

- `$BackupBase` — root output directory
- `$SmtpServer` / `$SmtpTo` — email report settings (set `$SmtpServer = ''` to disable)
- `$Jobs` — your backup job definitions

Each job specifies a `Label`, a `Tool` (`library`, `list`, or `loop`), and an
`Args` array passed to `spbackup.ps1`:

```powershell
@{
    Label = 'Team Documents'
    Tool  = 'library'
    Args  = @('backup',
               '--url',     'https://contoso.sharepoint.com/sites/TeamSite',
               '--library', 'Documents',
               '--out',     (Join-Path $BackupBase 'TeamSite'))
}
```

### 3. Set up credentials & scheduling

- **Windows** — see [windows/README.md](windows/README.md)
- **Linux** — see [linux/README.md](linux/README.md)

---

## Do I need Run-Backups.ps1?

**No.** You can call `spbackup.ps1` directly from cron, Task Scheduler, or the
command line:

```bash
pwsh ./spbackup.ps1 library backup --url "https://..." --library "Documents" --out ./backup
```

`Run-Backups.ps1` is useful when you want to:

- Run **multiple** backup jobs in sequence from a single scheduled task
- Get a **consolidated email report** with per-job status and timing
- Capture **per-job log files** automatically
- Load Windows credentials from **Credential Manager** (via `Load-Credentials.ps1`)

If you only have one backup job and don't need email reports, calling
`spbackup.ps1` directly is simpler.
