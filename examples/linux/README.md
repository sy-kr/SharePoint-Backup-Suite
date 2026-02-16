# Linux Scheduling

Run backups automatically on Linux using `run-backups.sh` + cron (or systemd
timers).  The shell script loads credentials from an env file, then hands off
to the cross-platform `Run-Backups.ps1` orchestrator.

## Quick start

1. Copy files to the project root (same directory as `spbackup.ps1`):

   ```bash
   cp examples/Run-Backups.ps1  .
   cp examples/linux/run-backups.sh .
   ```

2. Create the credential env file
3. Edit `Run-Backups.ps1` — configure jobs, SMTP, and paths
4. Test interactively: `./run-backups.sh`
5. Add a cron entry (or systemd timer)

---

## Credential env file

Store credentials in a file readable only by the backup user.
`run-backups.sh` sources this file automatically.

```bash
sudo mkdir -p /etc/spbackup
sudo tee /etc/spbackup/env << 'EOF'
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=your-client-secret
EOF
sudo chmod 600 /etc/spbackup/env
sudo chown root:root /etc/spbackup/env
```

> **Note:** Values should **not** be quoted or exported. The wrapper script
> sources the file and exports the variables itself.

If your env file is somewhere other than `/etc/spbackup/env`, override it:

```bash
ENV_FILE=/home/backups/.spbackup-env ./run-backups.sh
```

---

## Cron

### Using the orchestrator (recommended)

The orchestrator runs all jobs, collects results, and sends a single email
report:

```cron
# Daily at 02:00
0 2 * * * cd /opt/spbackup && /opt/spbackup/run-backups.sh >> /var/log/spbackup.log 2>&1
```

### Running individual jobs directly

If you don't need the orchestrator, call `spbackup.ps1` directly:

```cron
# Every 6 hours
0 */6 * * * . /etc/spbackup/env && cd /opt/spbackup && /usr/bin/pwsh ./spbackup.ps1 list backup --url "https://contoso.sharepoint.com/sites/team" --list "Tasks" --out /var/lib/spbackup/lists >> /var/log/spbackup-list.log 2>&1
0 */6 * * * . /etc/spbackup/env && cd /opt/spbackup && /usr/bin/pwsh ./spbackup.ps1 library backup --url "https://contoso.sharepoint.com/sites/team" --library "Documents" --out /var/lib/spbackup/docs >> /var/log/spbackup-library.log 2>&1
```

---

## systemd timer (alternative to cron)

If you prefer systemd timers, create two unit files:

**`/etc/systemd/system/spbackup.service`**

```ini
[Unit]
Description=SharePoint Backup Suite
After=network-online.target
Wants=network-online.target

[Service]
Type=oneshot
User=spbackup
WorkingDirectory=/opt/spbackup
ExecStart=/opt/spbackup/run-backups.sh
TimeoutStartSec=14400
EnvironmentFile=/etc/spbackup/env

# Hardening
ProtectSystem=full
PrivateTmp=true
NoNewPrivileges=true
```

**`/etc/systemd/system/spbackup.timer`**

```ini
[Unit]
Description=Run SharePoint backups daily

[Timer]
OnCalendar=*-*-* 02:00:00
Persistent=true
RandomizedDelaySec=300

[Install]
WantedBy=timers.target
```

Enable:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now spbackup.timer
```

Check status:

```bash
systemctl list-timers spbackup*
journalctl -u spbackup.service --since today
```

---

## How it works

```
cron / systemd
  └── run-backups.sh
        ├── source /etc/spbackup/env   (sets TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        └── exec pwsh Run-Backups.ps1
              ├── spbackup.ps1 library backup ...
              ├── spbackup.ps1 list backup ...
              ├── spbackup.ps1 loop backup ...
              └── Email report (SMTP)
```

---

## Log files

All logs are written to `<BackupBase>/orchestrator-logs/`:

| File | Contents |
|------|----------|
| `backup-report-<timestamp>.txt` | Overall summary (also emailed) |
| `job-<label>.log` | Full output for each individual job |

---

## Files

| File | Purpose |
|------|---------|
| `run-backups.sh` | Thin bash wrapper — loads env, calls `Run-Backups.ps1` |
| `Run-Backups.ps1` | Cross-platform orchestrator (shared — see `examples/`) |
