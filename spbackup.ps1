#!/usr/bin/env pwsh
#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Backup Suite — unified entry point for list and loop backup tools.

.DESCRIPTION
    Routes commands to the appropriate backup tool:
      spbackup.ps1 list    <command> [options]  → backup-lists.ps1
      spbackup.ps1 loop    <command> [options]  → backup-loop.ps1
      spbackup.ps1 library <command> [options]  → backup-library.ps1

.EXAMPLE
    pwsh ./spbackup.ps1 list backup --url "https://..." --list "Tasks" --out "./backup"
    pwsh ./spbackup.ps1 loop backup --url "https://..." --out "./backup"
    pwsh ./spbackup.ps1 library backup --url "https://..." --library "Documents" --out "./backup"
    pwsh ./spbackup.ps1 list enumerate --url "https://..."
    pwsh ./spbackup.ps1 loop resolve --url "https://..."
    pwsh ./spbackup.ps1 --version
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SUITE_VERSION = '2.0.0'

function Show-MainUsage {
    $usage = @"
SharePoint Backup Suite v$($script:SUITE_VERSION)

USAGE:
  pwsh ./spbackup.ps1 <tool> <command> [options]

TOOLS:
  list        Microsoft List backup (CSV export + attachments)
  loop        Microsoft Loop backup (HTML / Markdown / raw .loop)
  library     SharePoint Document Library backup (file download)

GLOBAL COMMANDS:
  --version   Show version
  --help      Show this help

LIST COMMANDS:
  pwsh ./spbackup.ps1 list backup    --url <URL> --list <name> --out <dir> [options]
  pwsh ./spbackup.ps1 list enumerate --url <URL>
  pwsh ./spbackup.ps1 list verify    --out <dir>
  pwsh ./spbackup.ps1 list diagnose  [--url <URL>]

LOOP COMMANDS:
  pwsh ./spbackup.ps1 loop backup    --url <URL> --out <dir> [options]
  pwsh ./spbackup.ps1 loop resolve   --url <URL>
  pwsh ./spbackup.ps1 loop verify    --out <dir>
  pwsh ./spbackup.ps1 loop diagnose  [--url <URL>]
  pwsh ./spbackup.ps1 loop setup-venv

LIBRARY COMMANDS:
  pwsh ./spbackup.ps1 library backup    --url <URL> --library <name> --out <dir> [options]
  pwsh ./spbackup.ps1 library enumerate --url <URL>
  pwsh ./spbackup.ps1 library verify    --out <dir>
  pwsh ./spbackup.ps1 library diagnose  [--url <URL>]

ENVIRONMENT VARIABLES (required):
  TENANT_ID             Azure AD / Entra tenant ID
  CLIENT_ID             App registration client ID
  CLIENT_SECRET         App registration client secret

See the individual tool help for full options:
  pwsh ./spbackup.ps1 list help
  pwsh ./spbackup.ps1 loop help
  pwsh ./spbackup.ps1 library help
"@
    Write-Host $usage
}

if ($args.Count -eq 0) {
    Show-MainUsage
    exit 1
}

$tool = $args[0].ToLower()
$remaining = @()
if ($args.Count -gt 1) {
    $remaining = @($args[1..($args.Count - 1)])
}

switch ($tool) {
    '--version' {
        Write-Host "SharePoint Backup Suite v$($script:SUITE_VERSION)"
        exit 0
    }
    '--help' {
        Show-MainUsage
        exit 0
    }
    'help' {
        Show-MainUsage
        exit 0
    }
    'list' {
        $scriptPath = Join-Path $PSScriptRoot 'backup-lists.ps1'
        if (-not (Test-Path $scriptPath)) {
            Write-Error "backup-lists.ps1 not found at $scriptPath"
            exit 1
        }
        & $scriptPath @remaining
        exit $LASTEXITCODE
    }
    'loop' {
        $scriptPath = Join-Path $PSScriptRoot 'backup-loop.ps1'
        if (-not (Test-Path $scriptPath)) {
            Write-Error "backup-loop.ps1 not found at $scriptPath"
            exit 1
        }
        & $scriptPath @remaining
        exit $LASTEXITCODE
    }
    'library' {
        $scriptPath = Join-Path $PSScriptRoot 'backup-library.ps1'
        if (-not (Test-Path $scriptPath)) {
            Write-Error "backup-library.ps1 not found at $scriptPath"
            exit 1
        }
        & $scriptPath @remaining
        exit $LASTEXITCODE
    }
    default {
        Write-Error "Unknown tool: '$tool'. Use 'list', 'loop', or 'library'."
        Write-Host ''
        Show-MainUsage
        exit 1
    }
}
