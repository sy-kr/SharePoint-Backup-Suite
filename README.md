# SharePoint Backup Suite

> [!WARNING]
> **Vibe-coded.** This project was built almost entirely through AI-assisted development (GitHub Copilot / Claude). It has been tested and is currently used in a production environment for performing backups of Microsoft Loop workspaces and Microsoft Lists. Review the code before trusting it with anything important, and please [open an issue](../../issues) if you find bugs.

A unified PowerShell backup utility for Microsoft 365 — export **Microsoft Lists** to CSV (with attachments) and **Microsoft Loop** pages to HTML / Markdown.

```bash
# Back up a Microsoft List (CSV + attachments)
pwsh ./spbackup.ps1 list backup --url "https://contoso.sharepoint.com/sites/team" --list "Tasks" --out ./backup/lists

# Back up a Loop workspace (HTML + Markdown)
pwsh ./spbackup.ps1 loop backup --url "https://loop.cloud.microsoft/p/..." --out ./backup/loop
```

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Azure App Registration Setup](#azure-app-registration-setup)
- [Command Reference](#command-reference)
- [Environment Variables](#environment-variables)
- [Incremental Backups](#incremental-backups)
- [Backup Verification](#backup-verification)
- [Diagnostics](#diagnostics)
- [Scheduling](#scheduling)
- [Output Structure](#output-structure)
- [Troubleshooting](#troubleshooting)
- [Project Structure](#project-structure)
- [License](#license)

---

## Features

| Feature | List Backup | Loop Backup |
|---------|:-----------:|:-----------:|
| Incremental sync (skip unchanged items) | ✅ | ✅ |
| Atomic writes (no partial files) | ✅ | ✅ |
| SHA-256 integrity verification | ✅ | ✅ |
| JSONL structured logging | ✅ | ✅ |
| Concurrency control (semaphore) | ✅ | ✅ |
| Retry with exponential backoff | ✅ | ✅ |
| Dry-run mode | ✅ | ✅ |
| CSV export | ✅ | — |
| List attachment download | ✅ | — |
| HTML export | — | ✅ |
| Markdown export (Python / Pandoc) | — | ✅ |
| Raw `.loop` file download | — | ✅ |
| Loop URL resolution (6 strategies) | — | ✅ |
| Certificate-based SharePoint auth | ✅ | — |

---

## Requirements

- **PowerShell 7.0+** (`pwsh`) — cross-platform (Linux, macOS, Windows)
- **Microsoft Entra ID app registration** with the [permissions below](#azure-app-registration-setup)
- **Python 3.8+** *(optional)* — for Loop HTML → Markdown conversion (venv auto-created on first run)
- **Pandoc** *(optional)* — fallback Markdown converter

---

## Quick Start

### 1. Clone & navigate

```bash
git clone <repo-url> spbackup
cd spbackup
```

### 2. Set environment variables

```bash
export TENANT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export CLIENT_SECRET="your-client-secret"
```

### 3. Run a backup

```bash
# Back up a Microsoft List
pwsh ./spbackup.ps1 list backup \
  --url "https://contoso.sharepoint.com/sites/team" \
  --list "Tasks" \
  --out ./backup/lists \
  --verbose

# Back up a Loop workspace
pwsh ./spbackup.ps1 loop backup \
  --url "https://loop.cloud.microsoft/p/..." \
  --out ./backup/loop \
  --verbose
```

### 4. (Optional) Python for Markdown conversion

If Python 3.8+ is on your PATH, the Loop backup will **automatically create a venv** and install
dependencies on first run. No manual setup needed.

To force-recreate the venv (e.g. after updating `requirements.txt`):

```bash
pwsh ./spbackup.ps1 loop setup-venv
```

---

## Azure App Registration Setup

### Microsoft Graph API Permissions

Grant these **Application** permissions and click **Grant admin consent**:

| Permission | Type | Required For |
|------------|------|-------------|
| `Sites.Read.All` | Application | Site / list enumeration, Loop resolution |
| `Files.Read.All` | Application | Loop page content download |

### SharePoint REST API (for list attachments only)

List attachments require **certificate-based** authentication via the SharePoint REST API:

1. Generate a certificate:
   ```bash
   openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 730 -nodes -subj "/CN=spbackup"
   openssl pkcs12 -export -out certs/spbackup.pfx -inkey key.pem -in cert.pem -passout pass:
   ```

2. Upload `cert.pem` to your app registration → **Certificates & secrets** → **Certificates** → **Upload certificate**

3. Grant the app **SharePoint** → `Sites.Read.All` (Application) permission + admin consent

4. Place `spbackup.pfx` in the `certs/` directory (auto-discovered) or set `CERT_PATH`

### SharePoint Embedded (for Loop workspaces)

Loop workspaces use SharePoint Embedded (SPE) containers. Even with `Sites.Read.All`, your app needs a **guest app permission** on the Loop container type:

```powershell
# Run in Windows PowerShell 5.1 (not pwsh) with SharePoint Online Management Shell
Connect-SPOService -Url "https://contoso-admin.sharepoint.com"
Set-SPOApplicationPermission `
  -OwningApplicationId "a187e399-0c36-4b98-8f04-1edc167a0996" `
  -GuestApplicationId "<YOUR_CLIENT_ID>" `
  -PermissionAppOnly Read, ReadContent
```

> `a187e399-0c36-4b98-8f04-1edc167a0996` is Microsoft's Loop Web Application owning app ID (fixed).

---

## Command Reference

### List Backup

```
pwsh ./spbackup.ps1 list <command> [options]
```

| Command | Description |
|---------|-------------|
| `backup` | Export a Microsoft List to CSV and download attachments |
| `enumerate` | List all Microsoft Lists in a SharePoint site |
| `verify` | Verify backup integrity against manifest |
| `diagnose` | Check auth, decode JWT, test Graph API access |

<details>
<summary><strong>Backup options</strong></summary>

| Option | Description |
|--------|-------------|
| `--url <URL>` | SharePoint site or list URL (required) |
| `--site-url <URL>` | Alias for `--url` (backward-compatible) |
| `--list <name>` | List display name (required unless `--list-id` used) |
| `--list-id <GUID>` | List GUID (alternative to `--list`) |
| `--out <dir>` | Output directory (required) |
| `--concurrency <N>` | Max parallel downloads (default: 4) |
| `--since <ISO date>` | Only items modified after this date |
| `--state <path>` | State file path (default: `<out>/.state.json`) |
| `--force` | Re-export everything, ignoring state |
| `--dry-run` | Enumerate only, no downloads |
| `--skip-attachments` | CSV only, skip attachment download |
| `--verbose` | Human-readable console output |

</details>

**Examples:**

```bash
# List all Microsoft Lists in a site
pwsh ./spbackup.ps1 list enumerate --url "https://contoso.sharepoint.com/sites/team"

# Back up a list by ID, skipping attachments
pwsh ./spbackup.ps1 list backup --url "https://contoso.sharepoint.com/sites/team" \
  --list-id "a1b2c3d4-..." --out ./backup/lists --skip-attachments

# Back up only items modified in the last 7 days
pwsh ./spbackup.ps1 list backup --url "https://contoso.sharepoint.com/sites/team" \
  --list "Tasks" --out ./backup/lists --since "2026-02-09T00:00:00Z"
```

### Loop Backup

```
pwsh ./spbackup.ps1 loop <command> [options]
```

| Command | Description |
|---------|-------------|
| `backup` | Export Loop pages (HTML, Markdown, raw `.loop`) |
| `resolve` | Resolve a URL to Graph resource(s) and print JSON |
| `verify` | Verify backup integrity against manifest |
| `diagnose` | Check auth, decode JWT, test Graph / SPE access |
| `setup-venv` | Force-(re)create Python venv (auto-created on first backup) |

<details>
<summary><strong>Backup options</strong></summary>

| Option | Description |
|--------|-------------|
| `--url <URL>` | Loop URL, SharePoint sharing link, or `loop.cloud.microsoft` URL (required) |
| `--out <dir>` | Output directory (required) |
| `--mode <mode>` | `page` / `workspace` / `auto` (default: auto) |
| `--raw-loop` | Also download raw `.loop` file bytes |
| `--html` | Export HTML (default: on) |
| `--md` | Export Markdown (default: on) |
| `--concurrency <N>` | Max parallel downloads (default: 4) |
| `--since <ISO date>` | Only items modified after this date |
| `--state <path>` | State file path (default: `<out>/.state.json`) |
| `--force` | Re-export all items, ignoring state |
| `--dry-run` | Resolve and enumerate only, no downloads |
| `--verbose` | Human-readable console output |

</details>

**Examples:**

```bash
# Back up a single Loop page
pwsh ./spbackup.ps1 loop backup --url "https://loop.cloud.microsoft/p/..." \
  --out ./backup/loop --mode page

# Back up with raw .loop files and verbose output
pwsh ./spbackup.ps1 loop backup --url "https://loop.cloud.microsoft/p/..." \
  --out ./backup/loop --raw-loop --verbose

# Resolve a Loop URL to its Graph resource (useful for debugging)
pwsh ./spbackup.ps1 loop resolve --url "https://loop.cloud.microsoft/p/..."
```

---

## Environment Variables

| Variable | Required | Description |
|----------|:--------:|-------------|
| `TENANT_ID` | ✅ | Azure AD / Entra tenant ID |
| `CLIENT_ID` | ✅ | App registration client ID |
| `CLIENT_SECRET` | ✅ | App registration client secret |
| `CERT_PATH` | — | Path to `.pfx` certificate (auto-discovered from `certs/`) |
| `CERT_PASSWORD` | — | Certificate password (if any) |
| `PYTHON` | — | Path to Python executable |
| `PANDOC` | — | Path to Pandoc executable |
| `SEARCH_REGION` | — | SharePoint region for search (e.g. `NAM`, `EUR`, `APC`) |

---

## Incremental Backups

Both tools track state in a `.state.json` file in the output directory:

- **List backup** — tracks `lastModifiedDateTime` per list; skips re-export if unchanged
- **Loop backup** — tracks `eTag` per item; skips re-export if unchanged, falls back to SHA-256 hash comparison

Use `--force` to re-export everything. Use `--since <ISO date>` to only process recently modified items.

---

## Backup Verification

```bash
pwsh ./spbackup.ps1 list verify --out ./backup/lists
pwsh ./spbackup.ps1 loop verify --out ./backup/loop
```

Reads `manifest.json` and checks all files against their stored SHA-256 hashes.

| Exit Code | Meaning |
|:---------:|---------|
| `0` | All files verified OK |
| `2` | Missing or mismatched files |

---

## Diagnostics

```bash
pwsh ./spbackup.ps1 list diagnose --url "https://..."
pwsh ./spbackup.ps1 loop diagnose --url "https://..."
```

Checks environment variables, acquires a token, decodes the JWT to show permissions, and tests Graph API connectivity. Pass a URL for site-specific or SPE container tests.

---

## Scheduling

### systemd (Linux)

Copy and edit the example service and timer files:

```bash
sudo cp examples/systemd/spbackup-list.service /etc/systemd/system/
sudo cp examples/systemd/spbackup-list.timer /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now spbackup-list.timer
```

Create the credential file:

```bash
sudo mkdir -p /etc/spbackup
sudo tee /etc/spbackup/env << 'EOF'
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=your-secret-value
EOF
sudo chmod 600 /etc/spbackup/env
```

### cron

```cron
# Every 6 hours
0 */6 * * * cd /opt/spbackup && /usr/bin/pwsh ./spbackup.ps1 list backup --url "https://..." --list "Tasks" --out /var/lib/spbackup/lists --verbose >> /var/log/spbackup-list.log 2>&1
0 */6 * * * cd /opt/spbackup && /usr/bin/pwsh ./spbackup.ps1 loop backup --url "https://..." --out /var/lib/spbackup/loop --verbose >> /var/log/spbackup-loop.log 2>&1
```

### Windows Task Scheduler

```powershell
schtasks /Create /TN "SPBackup-List" `
  /TR "pwsh -File C:\spbackup\spbackup.ps1 list backup --url 'https://...' --list 'Tasks' --out C:\backups\lists --verbose" `
  /SC DAILY /ST 02:00 /RU SYSTEM
```

---

## Output Structure

### List Backup

```
<out>/
├── <ListName>.csv                  CSV export of all list items
├── <ListName>_columns.json         Column definitions (name, type, display name)
├── attachments/
│   └── <ItemTitle>__<id>/
│       ├── attachment1.pdf
│       └── attachment2.docx
├── manifest.json                   Backup manifest with hashes
├── .state.json                     Incremental sync state
└── logs/
    └── run-<timestamp>.log.jsonl   Structured log
```

### Loop Backup

```
<out>/
├── items/
│   └── <PageName>__<id>/
│       ├── meta.json               Drive item metadata
│       ├── page.html               HTML export
│       ├── page.md                 Markdown export
│       └── page.loop               Raw .loop file (if --raw-loop)
├── manifest.json                   Backup manifest with hashes
├── .state.json                     Incremental sync state
└── logs/
    └── run-<timestamp>.log.jsonl   Structured log
```

---

## Direct Script Invocation

You can also invoke the backup scripts directly (without the dispatcher):

```bash
pwsh ./backup-lists.ps1 backup --url "https://..." --list "Tasks" --out ./backup
pwsh ./backup-loop.ps1 backup --url "https://..." --out ./backup
```

---

## Troubleshooting

<details>
<summary><strong>"Token has NO application permissions"</strong></summary>

Admin consent has not been granted. In Azure Portal → App registrations → API permissions → click **Grant admin consent**.
</details>

<details>
<summary><strong>List attachments not downloading</strong></summary>

Attachments require certificate-based SharePoint REST API auth. Run `diagnose` to check. See [SharePoint REST API setup](#sharepoint-rest-api-for-list-attachments-only).
</details>

<details>
<summary><strong>Loop backup returns "Could not resolve URL"</strong></summary>

1. Run `diagnose --url "..."` for detailed analysis
2. Check that `Files.Read.All` and `Sites.Read.All` are granted with admin consent
3. For SPE containers, grant guest app permission (see [SharePoint Embedded](#sharepoint-embedded-for-loop-workspaces))
4. Personal Loop workspaces ("My workspace") in OneDrive **cannot** be accessed with app-only credentials
</details>

<details>
<summary><strong>"Region is required" error in search</strong></summary>

Set `SEARCH_REGION` env var (e.g., `SEARCH_REGION=NAM`). Required for app-only search.
</details>

---

## Project Structure

```
spbackup.ps1              Main entry point (routes to list / loop)
backup-lists.ps1           Microsoft List backup script
backup-loop.ps1            Microsoft Loop backup script
lib/
├── Common.ps1             Shared helpers (logging, filesystem, state, CLI)
├── GraphAuth.ps1          Graph & SharePoint token acquisition
├── GraphApi.ps1           HTTP wrappers with retry / throttle
└── SiteResolver.ps1       SharePoint site URL resolution
tools/
├── html_to_md.py          HTML → Markdown converter (Loop-specific)
└── requirements.txt       Python dependencies
certs/
└── README.md              Certificate setup instructions
examples/
└── systemd/               systemd service & timer units
```

---

## License

MIT
