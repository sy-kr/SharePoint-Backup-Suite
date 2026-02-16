<#
.SYNOPSIS
    Retrieves SPBackup credentials from Windows Credential Manager and sets
    them as environment variables for the current process.

.DESCRIPTION
    Uses the Win32 CredRead API (P/Invoke, no external modules) to read
    Generic credentials stored by Setup-Credentials.ps1.

    Dot-source this from Run-Backups.ps1 before calling spbackup.ps1:
        . "$PSScriptRoot\Load-Credentials.ps1"

    After execution, the following are set (if stored):
        $env:TENANT_ID
        $env:CLIENT_ID
        $env:CLIENT_SECRET   (empty string if cert-only mode)

.EXAMPLE
    . .\Load-Credentials.ps1
#>

# ── P/Invoke for CredRead (no external modules needed) ───────────────────────

if (-not ([System.Management.Automation.PSTypeName]'CredManager').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class CredManager {
    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CredReadW(string target, int type, int flags, out IntPtr credential);

    [DllImport("advapi32.dll")]
    private static extern void CredFree(IntPtr credential);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct CREDENTIAL {
        public int    Flags;
        public int    Type;
        public string TargetName;
        public string Comment;
        public long   LastWritten;
        public int    CredentialBlobSize;
        public IntPtr CredentialBlob;
        public int    Persist;
        public int    AttributeCount;
        public IntPtr Attributes;
        public string TargetAlias;
        public string UserName;
    }

    /// <summary>Read a Generic credential's password by target name. Returns null if not found.</summary>
    public static string Read(string target) {
        IntPtr credPtr;
        // Type 1 = CRED_TYPE_GENERIC
        if (!CredReadW(target, 1, 0, out credPtr))
            return null;

        try {
            CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(credPtr, typeof(CREDENTIAL));
            if (cred.CredentialBlobSize > 0)
                return Marshal.PtrToStringUni(cred.CredentialBlob, cred.CredentialBlobSize / 2);
            return null;
        } finally {
            CredFree(credPtr);
        }
    }
}
"@
}

# ── Read and set ─────────────────────────────────────────────────────────────

$_tenantId     = [CredManager]::Read('SPBackup:TenantId')
$_clientId     = [CredManager]::Read('SPBackup:ClientId')
$_clientSecret = [CredManager]::Read('SPBackup:ClientSecret')   # may be $null for cert-only

$_missing = @()
if (-not $_tenantId) { $_missing += 'SPBackup:TenantId' }
if (-not $_clientId) { $_missing += 'SPBackup:ClientId' }

if ($_missing.Count -gt 0) {
    Write-Error ("Missing credentials in Credential Manager: {0}`nRun Setup-Credentials.ps1 first." -f ($_missing -join ', '))
    exit 1
}

$env:TENANT_ID = $_tenantId
$env:CLIENT_ID = $_clientId

if ($_clientSecret) {
    $env:CLIENT_SECRET = $_clientSecret
    $_authMode = 'client_secret'
} else {
    # Leave CLIENT_SECRET unset — GraphAuth.ps1 will fall back to certificate
    $_authMode = 'certificate'
}

# Clean up variables — don't leak to caller's scope
Remove-Variable -Name _tenantId, _clientId, _clientSecret, _missing -ErrorAction SilentlyContinue

Write-Verbose "Loaded SPBackup credentials from Credential Manager (auth: $_authMode)"
Remove-Variable -Name _authMode -ErrorAction SilentlyContinue
