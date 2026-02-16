# ─────────────────────────────────────────────────────────────────────────────
# Common.ps1 — Shared helpers for the SharePoint Backup suite
# ─────────────────────────────────────────────────────────────────────────────
# Dot-source this file from any backup script:
#   . (Join-Path $PSScriptRoot 'lib' 'Common.ps1')
# ─────────────────────────────────────────────────────────────────────────────

# Constants
$script:SUITE_VERSION    = '2.0.0'
$script:GRAPH_BASE       = 'https://graph.microsoft.com/v1.0'
$script:TOKEN_URL_TEMPLATE = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token'
$script:MAX_RETRIES      = 8
$script:SAFE_NAME_MAX    = 80

# Derive project root from the lib/ directory this file lives in.
# Only set if the caller hasn't already defined it.
if (-not $script:ProjectRoot) {
    $script:ProjectRoot = Split-Path -Parent $PSScriptRoot
}

# Global state (populated at runtime by each backup script).
# Initialised unconditionally — callers override these *after* dot-sourcing.
$script:CachedToken       = $null
$script:CachedTokenExpiry = [datetime]::MinValue
$script:CachedSpToken       = $null
$script:CachedSpTokenExpiry = [datetime]::MinValue
$script:SharePointHostname  = $null
$script:LogFilePath       = $null
$script:VerboseOutput     = $false
$script:Semaphore         = $null
$script:RunTimestamp      = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
$script:ExitCode          = 0

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
function Write-LogJsonl {
    [CmdletBinding()]
    param(
        [string]$Level,
        [string]$Event,
        [string]$Url         = '',
        [string]$SiteId      = '',
        [string]$ListId      = '',
        [string]$DriveId     = '',
        [string]$ItemId      = '',
        [int]$Attempt        = 0,
        [int]$StatusCode     = 0,
        [string]$Message     = ''
    )
    $entry = [ordered]@{
        timestamp  = (Get-Date).ToUniversalTime().ToString('o')
        level      = $Level
        event      = $Event
    }
    if ($Url)        { $entry['url']        = $Url }
    if ($SiteId)     { $entry['siteId']     = $SiteId }
    if ($ListId)     { $entry['listId']     = $ListId }
    if ($DriveId)    { $entry['driveId']    = $DriveId }
    if ($ItemId)     { $entry['itemId']     = $ItemId }
    if ($Attempt)    { $entry['attempt']    = $Attempt }
    if ($StatusCode) { $entry['statusCode'] = $StatusCode }
    $entry['message'] = $Message

    $json = $entry | ConvertTo-Json -Compress -Depth 4
    if ($script:LogFilePath) {
        $json | Out-File -FilePath $script:LogFilePath -Append -Encoding utf8
    }
    if ($script:VerboseOutput) {
        $color = switch ($Level) {
            'ERROR'   { 'Red' }
            'WARN'    { 'Yellow' }
            'INFO'    { 'Cyan' }
            'DEBUG'   { 'DarkGray' }
            default   { 'White' }
        }
        Write-Host "[$Level] $Event - $Message" -ForegroundColor $color
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Safe property access (strict mode compatible)
# ─────────────────────────────────────────────────────────────────────────────
function Get-SafeProp {
    [CmdletBinding()]
    param([object]$Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [hashtable]) { return $Obj[$Name] }
    if ($Obj.PSObject.Properties[$Name]) { return $Obj.$Name }
    return $null
}

function Test-SafeProp {
    [CmdletBinding()]
    param([object]$Obj, [string]$Name)
    if ($null -eq $Obj) { return $false }
    if ($Obj -is [hashtable]) { return $Obj.ContainsKey($Name) }
    return [bool]$Obj.PSObject.Properties[$Name]
}

# ─────────────────────────────────────────────────────────────────────────────
# Filesystem helpers
# ─────────────────────────────────────────────────────────────────────────────
function New-SafeName {
    <#
    .SYNOPSIS
        Create a filesystem-safe name from a display name, optionally suffixed
        with an ID for uniqueness.
    .PARAMETER StripExtension
        Strip the file extension before sanitising. Use for Loop .loop files
        where the extension is not meaningful. Do NOT use for list attachments
        where the original extension must be preserved.
    #>
    [CmdletBinding()]
    param(
        [string]$Name,
        [string]$Suffix = '',
        [switch]$StripExtension
    )
    $base = if ($StripExtension) { [System.IO.Path]::GetFileNameWithoutExtension($Name) } else { $Name }
    $safe = $base -replace '[<>:"/\\|?*\x00-\x1F]', '_'
    $safe = $safe -replace '[_\s]+', '_'
    $safe = $safe.Trim('_', ' ', '.')
    if ($safe.Length -gt $script:SAFE_NAME_MAX) {
        $safe = $safe.Substring(0, $script:SAFE_NAME_MAX)
    }
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'unnamed' }
    if ($Suffix) { return "${safe}__${Suffix}" }
    return $safe
}

function Write-AtomicFile {
    [CmdletBinding()]
    param(
        [string]$Path,
        [object]$Content,
        [switch]$Binary
    )
    $dir = Split-Path $Path -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = "$Path.tmp.$PID"
    try {
        if ($Binary) {
            [System.IO.File]::WriteAllBytes($tmp, [byte[]]$Content)
        } else {
            [System.IO.File]::WriteAllText($tmp, [string]$Content, [System.Text.Encoding]::UTF8)
        }
        Move-Item -Path $tmp -Destination $Path -Force
    } catch {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        throw
    }
}

function Get-FileSHA256 {
    [CmdletBinding()]
    param([string]$Path)
    if (-not (Test-Path $Path)) { return '' }
    $hash = Get-FileHash -Path $Path -Algorithm SHA256
    return $hash.Hash.ToLowerInvariant()
}

function Get-ConciseErrorMessage {
    <#
    .SYNOPSIS
        Strip HTTP response headers and ResponsePreview trailers from
        exception messages to keep logs clean.
    #>
    [CmdletBinding()]
    param([string]$Msg)
    if ([string]::IsNullOrEmpty($Msg)) { return $Msg }
    foreach ($marker in @(' Headers: ', "`nHeaders: ", "`n Headers: ", "`rHeaders: ", "`r`nHeaders: ", "`r`n Headers: ")) {
        $idx = $Msg.IndexOf($marker)
        if ($idx -gt 0) { return $Msg.Substring(0, $idx) }
    }
    foreach ($marker in @(' ResponsePreview: ', "`nResponsePreview: ", "`n ResponsePreview: ")) {
        $idx = $Msg.IndexOf($marker)
        if ($idx -gt 0) { return $Msg.Substring(0, $idx) }
    }
    return $Msg
}

# ─────────────────────────────────────────────────────────────────────────────
# State management (incremental backup)
# ─────────────────────────────────────────────────────────────────────────────
function Read-State {
    [CmdletBinding()]
    param([string]$Path)
    if (-not (Test-Path $Path)) { return @{} }
    $raw = Get-Content $Path -Raw -Encoding utf8
    $obj = $raw | ConvertFrom-Json -AsHashtable
    if ($null -eq $obj) { return @{} }
    return $obj
}

function Save-State {
    [CmdletBinding()]
    param([string]$Path, [hashtable]$State)
    $json = $State | ConvertTo-Json -Depth 10
    Write-AtomicFile -Path $Path -Content $json
}

# ─────────────────────────────────────────────────────────────────────────────
# CSV helper
# ─────────────────────────────────────────────────────────────────────────────
function Format-CsvField {
    [CmdletBinding()]
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    if ($Value.Contains(',') -or $Value.Contains('"') -or $Value.Contains("`n") -or $Value.Contains("`r")) {
        $escaped = $Value.Replace('"', '""')
        return "`"$escaped`""
    }
    return $Value
}

# ─────────────────────────────────────────────────────────────────────────────
# CLI argument parser (shared between all scripts)
# ─────────────────────────────────────────────────────────────────────────────
function Parse-Arguments {
    [CmdletBinding()]
    param([string[]]$RawArgs)

    if ($RawArgs.Count -eq 0) {
        return @{ Command = 'help'; Options = @{} }
    }

    $command = $RawArgs[0].ToLower()
    $opts = @{}
    $i = 1

    # Superset of all boolean switches across list and loop scripts
    $boolSwitches = @(
        'dry-run', 'verbose', 'force', 'include-hidden', 'skip-attachments',
        'raw-loop', 'html', 'md', 'yes-large'
    )

    while ($i -lt $RawArgs.Count) {
        $arg = $RawArgs[$i]
        if ($arg.StartsWith('--')) {
            $key = $arg.Substring(2).ToLower()

            if ($key -in $boolSwitches) {
                $opts[$key] = $true
                $i++
                continue
            }

            # Key-value parameters
            if (($i + 1) -lt $RawArgs.Count -and -not $RawArgs[$i + 1].StartsWith('--')) {
                $opts[$key] = $RawArgs[$i + 1]
                $i += 2
            } else {
                $opts[$key] = $true
                $i++
            }
        } else {
            $i++
        }
    }

    return @{
        Command = $command
        Options = $opts
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Shared diagnostic helpers
# ─────────────────────────────────────────────────────────────────────────────
function Show-TokenDiagnostic {
    <#
    .SYNOPSIS
        Decode a JWT token and display its claims. Returns the decoded claims
        object or $null on failure.
    #>
    [CmdletBinding()]
    param([string]$Token)

    try {
        $parts = $Token.Split('.')
        if ($parts.Count -lt 2) { return $null }
        $payload = $parts[1]
        $padLen = (4 - ($payload.Length % 4)) % 4
        $payload = $payload.Replace('-', '+').Replace('_', '/') + ('=' * $padLen)
        $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
        $claims = $decoded | ConvertFrom-Json

        $aud   = Get-SafeProp $claims 'aud'
        $iss   = Get-SafeProp $claims 'iss'
        $appId = Get-SafeProp $claims 'appid'
        $tid   = Get-SafeProp $claims 'tid'
        $roles = if ($claims.PSObject.Properties['roles']) { @($claims.roles) } else { @() }

        Write-Host "   Audience (aud):  $aud"
        Write-Host "   Issuer (iss):    $iss"
        Write-Host "   App ID (appid):  $appId"
        Write-Host "   Tenant ID (tid): $tid"
        Write-Host ''
        Write-Host '   Application roles (permissions) in token:' -ForegroundColor Yellow
        if ($roles.Count -eq 0) {
            Write-Host '     (NONE — this is the problem!)' -ForegroundColor Red
            Write-Host '     Your token has NO application permissions.' -ForegroundColor Red
            Write-Host '     Admin consent was NOT granted, or permissions were added but not consented to.' -ForegroundColor Red
        } else {
            foreach ($r in ($roles | Sort-Object)) {
                $color = if ($r -in @('Files.Read.All', 'Sites.Read.All', 'Sites.ReadWrite.All', 'Sites.FullControl.All')) { 'Green' } else { 'White' }
                Write-Host "     - $r" -ForegroundColor $color
            }
        }
        Write-Host ''

        # Explicit permission check table
        $hasFilesRead = 'Files.Read.All' -in $roles
        $hasSitesRead = 'Sites.Read.All' -in $roles
        $hasSitesFull = 'Sites.FullControl.All' -in $roles
        $hasSitesRW   = 'Sites.ReadWrite.All' -in $roles

        Write-Host '   Permission check:' -ForegroundColor Yellow
        Write-Host "     Files.Read.All:        $(if ($hasFilesRead) { 'YES' } else { 'MISSING' })" -ForegroundColor $(if ($hasFilesRead) { 'Green' } else { 'Red' })
        Write-Host "     Sites.Read.All:        $(if ($hasSitesRead) { 'YES' } else { 'MISSING' })" -ForegroundColor $(if ($hasSitesRead) { 'Green' } else { 'Red' })
        Write-Host "     Sites.ReadWrite.All:   $(if ($hasSitesRW) { 'YES' } else { 'not set (optional)' })" -ForegroundColor $(if ($hasSitesRW) { 'Green' } else { 'DarkGray' })
        Write-Host "     Sites.FullControl.All: $(if ($hasSitesFull) { 'YES' } else { 'not set (optional)' })" -ForegroundColor $(if ($hasSitesFull) { 'Green' } else { 'DarkGray' })

        if (-not $hasFilesRead -and -not $hasSitesRead -and -not $hasSitesFull -and -not $hasSitesRW) {
            Write-Host ''
            Write-Host '   DIAGNOSIS: Token has NONE of the required permissions.' -ForegroundColor Red
            Write-Host '   Admin consent has not been granted. In Azure Portal:' -ForegroundColor Red
            Write-Host '     1. Go to App registrations > [your app] > API permissions' -ForegroundColor Red
            Write-Host '     2. Click "Grant admin consent for [tenant]"' -ForegroundColor Red
            Write-Host '     3. Confirm all permissions show green checkmarks' -ForegroundColor Red
        }

        return [PSCustomObject]@{
            Claims = $claims
            Roles  = $roles
        }
    } catch {
        Write-Host "   Could not decode JWT: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Show-GraphConnectivityTests {
    <#
    .SYNOPSIS
        Run basic Graph API connectivity tests (/organization, /sites).
    #>
    [CmdletBinding()]
    param()

    try {
        $org = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/organization?`$select=id,displayName"
        $orgValue = if ($org.PSObject.Properties['value']) { @($org.value) } else { @($org) }
        if ($orgValue.Count -gt 0) {
            $orgName = Get-SafeProp $orgValue[0] 'displayName'
            Write-Host "   /organization:  OK — $orgName" -ForegroundColor Green
        }
    } catch {
        Write-Host "   /organization:  FAIL — $($_.Exception.Message)" -ForegroundColor Red
    }

    try {
        $sites = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/sites?search=*&`$top=3&`$select=id,displayName,webUrl"
        $siteValues = if ($sites.PSObject.Properties['value']) { @($sites.value) } else { @() }
        Write-Host "   /sites?search=*: OK — found $($siteValues.Count) site(s)" -ForegroundColor Green
        foreach ($s in $siteValues) {
            $sName = Get-SafeProp $s 'displayName'
            $sUrl  = Get-SafeProp $s 'webUrl'
            Write-Host "     - $sName ($sUrl)" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "   /sites?search=*: FAIL — $($_.Exception.Message)" -ForegroundColor Red
    }
}
