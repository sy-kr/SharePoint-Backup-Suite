#!/usr/bin/env bash
# ---------------------------------------------------------------------------
#  run-backups.sh - Thin wrapper for Run-Backups.ps1 on Linux / macOS.
#
#  Loads credentials from an env file, then calls the PowerShell
#  orchestrator.  Use this as the target for a cron job.
#
#  Usage:
#    ./run-backups.sh                         # uses /etc/spbackup/env
#    ENV_FILE=/path/to/env ./run-backups.sh   # custom env file
#
#  First-time setup:
#    1. Copy Run-Backups.ps1 to the same directory as spbackup.ps1
#    2. Edit the CONFIGURATION section in Run-Backups.ps1
#    3. Create an env file (see below)
#    4. Add a cron entry (see README.md)
# ---------------------------------------------------------------------------
set -euo pipefail

# Directory where spbackup.ps1 and Run-Backups.ps1 live
# (assumes this script has been copied to the project root)
SPBACKUP_DIR="$(cd "$(dirname "$0")" && pwd)"

# Credential / env file (override with ENV_FILE environment variable)
ENV_FILE="${ENV_FILE:-/etc/spbackup/env}"

if [[ -f "$ENV_FILE" ]]; then
    # shellcheck source=/dev/null
    source "$ENV_FILE"
    export TENANT_ID CLIENT_ID CLIENT_SECRET 2>/dev/null || true
else
    echo "WARNING: env file not found at $ENV_FILE" >&2
    echo "Set TENANT_ID, CLIENT_ID, CLIENT_SECRET before running." >&2
fi

exec pwsh -NoProfile -NonInteractive -File "${SPBACKUP_DIR}/Run-Backups.ps1"
