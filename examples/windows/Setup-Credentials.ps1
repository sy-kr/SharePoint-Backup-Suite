<#
.SYNOPSIS
    One-time setup: stores SharePoint backup credentials in Windows Credential Manager.
    Run this interactively as the same user account that will execute the scheduled task.

.DESCRIPTION
    Uses cmdkey.exe (built into Windows) so there are no module dependencies.
    Credentials are encrypted with the user's DPAPI key — only this Windows
    account on this machine can retrieve them.

    If you are using certificate-only auth (no client secret), just press Enter
    when prompted for the secret and it will be skipped.

.EXAMPLE
    pwsh .\Setup-Credentials.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '=== SPBackup Credential Setup ===' -ForegroundColor Cyan
Write-Host ''
Write-Host 'Credentials will be stored in Windows Credential Manager (DPAPI encrypted).'
Write-Host 'Only this Windows user account on this machine can read them.'
Write-Host ''

# ── Prompt for values ────────────────────────────────────────────────────────

$TenantId = Read-Host 'Tenant ID  (GUID, required)'
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    Write-Host 'ERROR: Tenant ID is required.' -ForegroundColor Red
    exit 1
}

$ClientId = Read-Host 'Client ID  (GUID, required)'
if ([string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Host 'ERROR: Client ID is required.' -ForegroundColor Red
    exit 1
}

Write-Host ''
Write-Host 'Client Secret is optional if you have a certificate (.pfx) for auth.'
Write-Host 'Press Enter to skip if using certificate-only authentication.'
$SecureSecret = Read-Host 'Client Secret' -AsSecureString

# Convert SecureString back to plain text for cmdkey (it stores it encrypted)
$BSTR   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureSecret)
$Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

# ── Store in Credential Manager ──────────────────────────────────────────────

Write-Host ''
cmdkey /generic:"SPBackup:TenantId" /user:"spbackup" /pass:"$TenantId"
cmdkey /generic:"SPBackup:ClientId" /user:"spbackup" /pass:"$ClientId"

if (-not [string]::IsNullOrWhiteSpace($Secret)) {
    cmdkey /generic:"SPBackup:ClientSecret" /user:"spbackup" /pass:"$Secret"
    $authMode = 'client_secret + certificate fallback'
} else {
    # Remove any stale secret entry
    cmdkey /delete:"SPBackup:ClientSecret" 2>$null
    $authMode = 'certificate only'
}

# Zero out the plain-text copy
$Secret = $null
[System.GC]::Collect()

Write-Host ''
Write-Host 'Credentials stored in Windows Credential Manager.' -ForegroundColor Green
Write-Host "   Auth mode: $authMode"
Write-Host '   Targets:   SPBackup:TenantId, SPBackup:ClientId'
if ($authMode -like 'client*') {
    Write-Host '              SPBackup:ClientSecret'
}
Write-Host ''
Write-Host 'To verify:   cmdkey /list:SPBackup:*'
Write-Host ''
Write-Host 'To remove all:'
Write-Host '   cmdkey /delete:SPBackup:TenantId'
Write-Host '   cmdkey /delete:SPBackup:ClientId'
Write-Host '   cmdkey /delete:SPBackup:ClientSecret'
Write-Host ''

if ($authMode -eq 'certificate only') {
    Write-Host 'REMINDER: Place your .pfx certificate at:' -ForegroundColor Yellow
    Write-Host '   <project-root>\certs\spbackup.pfx' -ForegroundColor Yellow
    Write-Host '   or set the CERT_PATH environment variable.' -ForegroundColor Yellow
    Write-Host ''
}
