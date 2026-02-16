#!/usr/bin/env pwsh
#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Document Library Backup — download all files preserving folder structure.

.DESCRIPTION
    Part of the SharePoint Backup suite. Can be invoked directly or via spbackup.ps1.
    Recursively downloads every file in a SharePoint document library using the
    Microsoft Graph API. Designed for very large libraries (hundreds of thousands of
    files) with paginated enumeration, per-file retry, and incremental sync.

.EXAMPLE
    pwsh ./backup-library.ps1 backup --url "https://contoso.sharepoint.com/sites/team" --library "Documents" --out "./backup"
    pwsh ./backup-library.ps1 enumerate --url "https://contoso.sharepoint.com/sites/team"
    pwsh ./backup-library.ps1 verify --out "./backup"
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# Load shared libraries
# ─────────────────────────────────────────────────────────────────────────────
$script:ProjectRoot = $PSScriptRoot
. (Join-Path $PSScriptRoot 'lib' 'Common.ps1')
. (Join-Path $PSScriptRoot 'lib' 'GraphAuth.ps1')
. (Join-Path $PSScriptRoot 'lib' 'GraphApi.ps1')
. (Join-Path $PSScriptRoot 'lib' 'SiteResolver.ps1')

$script:TOOL_NAME         = 'spbackup library'
$script:FILE_MAX_RETRIES  = 5
$script:FILE_RETRY_BASE   = 5      # seconds — base for exponential backoff (5, 10, 20, 40…)

# ─────────────────────────────────────────────────────────────────────────────
# Drive / Library Discovery
# ─────────────────────────────────────────────────────────────────────────────
function Get-SiteDocumentLibraries {
    <#
    .SYNOPSIS
        Return all document-library drives for a site (paginated).
    #>
    [CmdletBinding()]
    param(
        [string]$SiteId,
        [switch]$IncludeHidden
    )

    $drives = [System.Collections.Generic.List[object]]::new()
    $uri = "$($script:GRAPH_BASE)/sites/$SiteId/drives?`$select=id,name,driveType,webUrl,description,lastModifiedDateTime&`$top=200"

    while ($uri) {
        Write-LogJsonl -Level 'DEBUG' -Event 'drives_enumerate' -SiteId $SiteId -Url $uri -Message 'Listing drives'
        $response = Invoke-GraphRequest -Uri $uri -SiteId $SiteId

        $items = @()
        if (Test-SafeProp $response 'value') { $items = @($response.value) }

        foreach ($item in $items) {
            $driveType = Get-SafeProp $item 'driveType'
            # Skip personal OneDrive drives that occasionally appear
            if ($driveType -eq 'personal') { continue }
            $drives.Add($item)
        }

        $uri = $null
        if (Test-SafeProp $response '@odata.nextLink') {
            $uri = Get-SafeProp $response '@odata.nextLink'
        }
    }

    return $drives
}

function Find-DriveByName {
    <#
    .SYNOPSIS
        Resolve a library by display name (case-insensitive) or ID.
    #>
    [CmdletBinding()]
    param(
        [string]$SiteId,
        [string]$LibraryName = '',
        [string]$LibraryId   = ''
    )

    if ($LibraryId) {
        Write-LogJsonl -Level 'INFO' -Event 'drive_lookup_by_id' -SiteId $SiteId -DriveId $LibraryId -Message 'Looking up drive by ID'
        try {
            return Invoke-GraphRequest `
                -Uri "$($script:GRAPH_BASE)/drives/${LibraryId}?`$select=id,name,driveType,webUrl,description,lastModifiedDateTime" `
                -DriveId $LibraryId
        } catch {
            throw "Could not find document library with ID '$LibraryId': $($_.Exception.Message)"
        }
    }

    $allDrives = Get-SiteDocumentLibraries -SiteId $SiteId
    # Exact match
    foreach ($d in $allDrives) {
        if ((Get-SafeProp $d 'name') -eq $LibraryName) { return $d }
    }
    # Case-insensitive
    foreach ($d in $allDrives) {
        $n = Get-SafeProp $d 'name'
        if ($n -and $n.Equals($LibraryName, [System.StringComparison]::OrdinalIgnoreCase)) { return $d }
    }
    # URL-decoded match for "Shared Documents" → "Documents" etc.
    foreach ($d in $allDrives) {
        $wUrl = Get-SafeProp $d 'webUrl'
        if ($wUrl -and $wUrl -match "/$([regex]::Escape($LibraryName))(/|$)") { return $d }
    }

    $available = ($allDrives | ForEach-Object { "'$(Get-SafeProp $_ 'name')'" }) -join ', '
    throw "Could not find a document library named '$LibraryName' in site. Available: $available"
}

# ─────────────────────────────────────────────────────────────────────────────
# Recursive File Enumeration (paginated, large-library safe)
# ─────────────────────────────────────────────────────────────────────────────
function Get-DriveItemsRecursive {
    <#
    .SYNOPSIS
        Enumerate every file in a drive using delta or recursive children.
        Uses pagination ($top + @odata.nextLink) for libraries with hundreds
        of thousands of items.
    .PARAMETER DriveId
        The Graph drive ID.
    .PARAMETER SinceDate
        Optional ISO-8601 date filter — only items modified after this.
    .PARAMETER DeltaLink
        Optional delta link from a previous sync to get only changes.
    .OUTPUTS
        A hashtable with 'files' (list of drive items) and 'deltaLink'
        (for next incremental sync).
    #>
    [CmdletBinding()]
    param(
        [string]$DriveId,
        [string]$SinceDate = '',
        [string]$DeltaLink = ''
    )

    $files = [System.Collections.Generic.List[object]]::new()
    $newDeltaLink = ''
    $totalEnumerated = 0

    # Prefer delta query for incremental sync (returns all items on first call)
    if ($DeltaLink) {
        $uri = $DeltaLink
    } else {
        $uri = "$($script:GRAPH_BASE)/drives/$DriveId/root/delta?`$select=id,name,size,file,folder,parentReference,lastModifiedDateTime,eTag,webUrl&`$top=500"
    }

    while ($uri) {
        Write-LogJsonl -Level 'DEBUG' -Event 'drive_enumerate_page' -DriveId $DriveId -Url $uri `
            -Message "Enumerating drive items (so far: $totalEnumerated)"

        try {
            $response = Invoke-GraphRequest -Uri $uri -DriveId $DriveId
        } catch {
            Write-LogJsonl -Level 'WARN' -Event 'drive_enumerate_fail' -DriveId $DriveId `
                -Message "Pagination failed after $totalEnumerated items: $($_.Exception.Message)"
            break
        }

        $items = @()
        if (Test-SafeProp $response 'value') { $items = @($response.value) }

        foreach ($item in $items) {
            $totalEnumerated++

            # Skip folders (we only want files; folder structure comes from parentReference)
            if (Test-SafeProp $item 'folder') { continue }

            # Skip deleted items in delta responses
            if (Test-SafeProp $item 'deleted') { continue }

            # Optional date filter
            if ($SinceDate) {
                $lastMod = Get-SafeProp $item 'lastModifiedDateTime'
                if ($lastMod -and $lastMod -lt $SinceDate) { continue }
            }

            $files.Add($item)
        }

        # Check for next page or delta link
        $uri = $null
        if (Test-SafeProp $response '@odata.nextLink') {
            $uri = Get-SafeProp $response '@odata.nextLink'
        }
        if (Test-SafeProp $response '@odata.deltaLink') {
            $newDeltaLink = Get-SafeProp $response '@odata.deltaLink'
        }

        # Progress heartbeat every 5000 items
        if ($totalEnumerated % 5000 -eq 0 -and $totalEnumerated -gt 0) {
            Write-LogJsonl -Level 'INFO' -Event 'enumerate_progress' -DriveId $DriveId `
                -Message "Enumerated $totalEnumerated items so far ($($files.Count) files)"
        }
    }

    Write-LogJsonl -Level 'INFO' -Event 'enumerate_complete' -DriveId $DriveId `
        -Message "Enumeration complete: $totalEnumerated total items, $($files.Count) files"

    return @{
        files     = $files
        deltaLink = $newDeltaLink
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# File Download with Per-File Retry
# ─────────────────────────────────────────────────────────────────────────────
function Get-DriveItemRelativePath {
    <#
    .SYNOPSIS
        Build a relative filesystem path from a drive item's parentReference.
        E.g. "General/Subfolder/report.docx"
    #>
    [CmdletBinding()]
    param([object]$DriveItem)

    $parentRef = Get-SafeProp $DriveItem 'parentReference'
    $parentPath = Get-SafeProp $parentRef 'path'
    $itemName  = Get-SafeProp $DriveItem 'name'

    if (-not $parentPath -or -not $itemName) {
        return $itemName ?? 'unknown'
    }

    # parentReference.path looks like: /drives/<id>/root:/General/Subfolder
    # We want the part after "root:" (or "root:/" )
    $relFolder = ''
    $rootIdx = $parentPath.IndexOf('root:')
    if ($rootIdx -ge 0) {
        $relFolder = $parentPath.Substring($rootIdx + 5).TrimStart('/')
    }

    # URL-decode folder names (SharePoint encodes spaces, special chars)
    $relFolder = [System.Uri]::UnescapeDataString($relFolder)

    if ($relFolder) {
        return "$relFolder/$itemName"
    }
    return $itemName
}

function Download-DriveFile {
    <#
    .SYNOPSIS
        Download a single file from a drive, with up to $FILE_MAX_RETRIES
        attempts and exponential backoff on failure.
    .OUTPUTS
        Hashtable with 'success', 'path', 'sha256', 'size', and optionally 'error'.
    #>
    [CmdletBinding()]
    param(
        [object]$DriveItem,
        [string]$OutputDir,
        [string]$DriveId
    )

    $itemId   = Get-SafeProp $DriveItem 'id'
    $itemName = Get-SafeProp $DriveItem 'name'
    $relPath  = Get-DriveItemRelativePath -DriveItem $DriveItem

    # Sanitise each path segment individually to preserve folder structure
    $segments = $relPath -split '/'
    $safeSegments = $segments | ForEach-Object {
        # Preserve original file/folder names as much as possible but make them filesystem-safe
        # Include \ so backslashes aren't misinterpreted as path separators on Windows
        $seg = $_ -replace '[<>:"\\/|?*\x00-\x1F]', '_'
        $seg = $seg.Trim('.', ' ')
        if ([string]::IsNullOrWhiteSpace($seg)) { $seg = '_' }
        $seg
    }
    $safeRelPath = $safeSegments -join [System.IO.Path]::DirectorySeparatorChar
    $outPath = Join-Path $OutputDir $safeRelPath

    $outDir = Split-Path $outPath -Parent
    if (-not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $tmpPath = "$outPath.tmp.$PID"
    $downloadUri = "$($script:GRAPH_BASE)/drives/$DriveId/items/$itemId/content"

    for ($attempt = 1; $attempt -le $script:FILE_MAX_RETRIES; $attempt++) {
        try {
            Invoke-GraphRequest -Uri $downloadUri -Raw -OutFile $tmpPath `
                -DriveId $DriveId -ItemId $itemId | Out-Null

            Move-Item -LiteralPath $tmpPath -Destination $outPath -Force
            $hash = Get-FileSHA256 -Path $outPath
            $fileSize = (Get-Item -LiteralPath $outPath).Length

            Write-LogJsonl -Level 'DEBUG' -Event 'file_downloaded' -DriveId $DriveId -ItemId $itemId `
                -Attempt $attempt -Message "Downloaded: $relPath ($fileSize bytes)"

            return @{
                success  = $true
                relPath  = $relPath
                path     = $outPath
                sha256   = $hash
                size     = $fileSize
                itemId   = $itemId
                itemName = $itemName
            }
        } catch {
            if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue }

            $conciseErr = Get-ConciseErrorMessage $_.Exception.Message

            if ($attempt -lt $script:FILE_MAX_RETRIES) {
                $delay = $script:FILE_RETRY_BASE * [math]::Pow(2, $attempt - 1) + (Get-Random -Minimum 0.0 -Maximum 1.0)
                Write-LogJsonl -Level 'WARN' -Event 'file_download_retry' -DriveId $DriveId -ItemId $itemId `
                    -Attempt $attempt -Message "Failed: $relPath — retrying in ${delay}s — $conciseErr"
                Start-Sleep -Seconds $delay
            } else {
                Write-LogJsonl -Level 'ERROR' -Event 'file_download_fail' -DriveId $DriveId -ItemId $itemId `
                    -Attempt $attempt -Message "Failed after $($script:FILE_MAX_RETRIES) attempts: $relPath — $conciseErr"
                return @{
                    success  = $false
                    relPath  = $relPath
                    itemId   = $itemId
                    itemName = $itemName
                    error    = $conciseErr
                }
            }
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Commands
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-LibraryEnumerateCommand {
    <#
    .SYNOPSIS
        List all document libraries (drives) in a SharePoint site.
    #>
    [CmdletBinding()]
    param([hashtable]$Opts)

    $siteUrl = $Opts['url']
    if ([string]::IsNullOrWhiteSpace($siteUrl)) { throw '--url is required for the enumerate command.' }

    $site    = Resolve-SiteFromUrl -Url $siteUrl
    $siteId  = Get-SafeProp $site 'id'
    $siteName = Get-SafeProp $site 'displayName'

    $drives = Get-SiteDocumentLibraries -SiteId $siteId

    Write-Host ''
    Write-Host "Document Libraries in '$siteName':" -ForegroundColor Cyan
    Write-Host ''

    if ($drives.Count -eq 0) {
        Write-Host '  (no document libraries found)' -ForegroundColor Yellow
        return
    }

    foreach ($drive in $drives) {
        $dName    = Get-SafeProp $drive 'name'
        $dId      = Get-SafeProp $drive 'id'
        $dType    = Get-SafeProp $drive 'driveType'
        $dWebUrl  = Get-SafeProp $drive 'webUrl'
        $dLastMod = Get-SafeProp $drive 'lastModifiedDateTime'

        Write-Host "  - $dName" -ForegroundColor White
        Write-Host "    ID:        $dId" -ForegroundColor DarkGray
        Write-Host "    Type:      $dType" -ForegroundColor DarkGray
        Write-Host "    URL:       $dWebUrl" -ForegroundColor DarkGray
        Write-Host "    Modified:  $dLastMod" -ForegroundColor DarkGray
        Write-Host ''
    }

    Write-LogJsonl -Level 'INFO' -Event 'drives_enumerated' -SiteId $siteId `
        -Message "Found $($drives.Count) document library/libraries"
}

function Invoke-LibraryBackupCommand {
    <#
    .SYNOPSIS
        Download all files from a document library, preserving folder structure.
    #>
    [CmdletBinding()]
    param([hashtable]$Opts)

    $siteUrl     = $Opts['url']
    $libName     = $Opts['library']
    $libId       = $Opts['library-id']
    $outDir      = $Opts['out']
    $concurrency = [int]($Opts['concurrency'] ?? 4)
    $sinceDate   = $Opts['since']       ?? ''
    $statePath   = $Opts['state']       ?? ''
    $dryRun      = [bool]$Opts['dry-run']
    $force       = [bool]$Opts['force']

    # Auto-extract library name from URL if it contains a known doc-lib path segment
    if ($siteUrl -and -not $libName -and -not $libId) {
        if ($siteUrl -match '(?:/(?:Shared%20Documents|Documents|[^/]+))\s*$') {
            $lastSeg = ([uri]$siteUrl).Segments[-1].TrimEnd('/')
            $libName = [System.Uri]::UnescapeDataString($lastSeg)
            Write-LogJsonl -Level 'INFO' -Event 'url_library_extracted' `
                -Message "Extracted library name from URL: '$libName'"
        }
    }

    if ([string]::IsNullOrWhiteSpace($siteUrl))                              { throw '--url is required for the backup command.' }
    if ([string]::IsNullOrWhiteSpace($libName) -and [string]::IsNullOrWhiteSpace($libId)) {
        throw '--library or --library-id is required for the backup command.'
    }
    if ([string]::IsNullOrWhiteSpace($outDir))                               { throw '--out is required for the backup command.' }

    if (-not $statePath) { $statePath = Join-Path $outDir '.state.json' }

    $script:Semaphore = [System.Threading.Semaphore]::new($concurrency, $concurrency)

    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    # Set up logging
    $logsDir = Join-Path $outDir 'logs'
    if (-not (Test-Path $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }
    $script:LogFilePath = Join-Path $logsDir "run-$($script:RunTimestamp).log.jsonl"

    Write-LogJsonl -Level 'INFO' -Event 'backup_start' -Url $siteUrl `
        -Message "Library backup starting: library=$($libName ?? $libId), concurrency=$concurrency"

    # 1) Resolve site
    $site      = Resolve-SiteFromUrl -Url $siteUrl
    $siteId    = Get-SafeProp $site 'id'
    $siteName  = Get-SafeProp $site 'displayName'

    Write-LogJsonl -Level 'INFO' -Event 'site_resolved' -SiteId $siteId -Message "Site: $siteName"

    # 2) Find the drive
    $drive = Find-DriveByName -SiteId $siteId -LibraryName $libName -LibraryId $libId
    $driveId      = Get-SafeProp $drive 'id'
    $driveName    = Get-SafeProp $drive 'name'
    $driveLastMod = Get-SafeProp $drive 'lastModifiedDateTime'

    Write-LogJsonl -Level 'INFO' -Event 'drive_found' -SiteId $siteId -DriveId $driveId `
        -Message "Found library: $driveName (lastModified=$driveLastMod)"

    # 3) Load state (for delta link and per-file eTags)
    $state = @{}
    if (-not $force) { $state = Read-State -Path $statePath }

    $stateKey  = "drive:${driveId}"
    $deltaLink = ''
    $fileEtags = @{}

    if ($state.ContainsKey($stateKey)) {
        $driveState = $state[$stateKey]
        $deltaLink = $driveState['deltaLink'] ?? ''
        $fileEtags = $driveState['fileEtags'] ?? @{}
    }

    # 4) Enumerate files (using delta for incremental)
    Write-LogJsonl -Level 'INFO' -Event 'enumerate_start' -DriveId $driveId `
        -Message "Enumerating files$(if ($deltaLink) { ' (incremental delta)' } else { ' (full scan)' })"

    $enumResult = Get-DriveItemsRecursive -DriveId $driveId -SinceDate $sinceDate -DeltaLink $(if (-not $force) { $deltaLink } else { '' })
    $allFiles   = $enumResult['files']
    $newDelta   = $enumResult['deltaLink']

    Write-Host "Found $($allFiles.Count) file(s) in '$driveName'" -ForegroundColor Cyan

    # 5) Filter out unchanged files (by eTag)
    $filesToDownload = [System.Collections.Generic.List[object]]::new()
    $skippedCount = 0

    foreach ($file in $allFiles) {
        $fId   = Get-SafeProp $file 'id'
        $fEtag = Get-SafeProp $file 'eTag'

        if (-not $force -and $fEtag -and $fileEtags.ContainsKey($fId) -and $fileEtags[$fId] -eq $fEtag) {
            $skippedCount++
            continue
        }

        $filesToDownload.Add($file)
    }

    if ($skippedCount -gt 0) {
        Write-LogJsonl -Level 'INFO' -Event 'skip_unchanged' -DriveId $driveId `
            -Message "Skipping $skippedCount unchanged file(s)"
        Write-Host "  Skipping $skippedCount unchanged file(s)" -ForegroundColor DarkGray
    }

    Write-Host "  $($filesToDownload.Count) file(s) to download" -ForegroundColor Cyan

    if ($dryRun) {
        Write-Host ''
        Write-Host "Dry run summary for '$driveName':" -ForegroundColor Cyan
        Write-Host "  Total files in library: $($allFiles.Count)" -ForegroundColor DarkGray
        Write-Host "  Files to download:      $($filesToDownload.Count)" -ForegroundColor DarkGray
        Write-Host "  Skipped (unchanged):    $skippedCount" -ForegroundColor DarkGray

        if ($filesToDownload.Count -le 50) {
            Write-Host ''
            foreach ($f in $filesToDownload) {
                $relPath = Get-DriveItemRelativePath -DriveItem $f
                $size = Get-SafeProp $f 'size'
                Write-Host "  $relPath ($size bytes)" -ForegroundColor DarkGray
            }
        }
        return
    }

    if ($filesToDownload.Count -eq 0) {
        Write-Host 'Nothing to download.' -ForegroundColor Green
        # Still save state (delta link may have advanced)
        $state[$stateKey] = @{
            deltaLink        = $newDelta
            fileEtags        = $fileEtags
            lastBackupTime   = (Get-Date).ToUniversalTime().ToString('o')
            totalFiles       = $allFiles.Count
        }
        Save-State -Path $statePath -State $state
        return
    }

    # 6) Download files
    $filesDir = Join-Path $outDir 'files'
    if (-not (Test-Path $filesDir)) {
        New-Item -ItemType Directory -Path $filesDir -Force | Out-Null
    }

    $downloadedFiles = [System.Collections.Generic.List[object]]::new()
    $failedFiles     = [System.Collections.Generic.List[object]]::new()
    $totalBytes      = [long]0
    $downloadCount   = 0

    foreach ($file in $filesToDownload) {
        $downloadCount++
        $relPath = Get-DriveItemRelativePath -DriveItem $file

        if ($downloadCount % 100 -eq 0 -or $downloadCount -eq 1) {
            Write-LogJsonl -Level 'INFO' -Event 'download_progress' -DriveId $driveId `
                -Message "Downloading $downloadCount / $($filesToDownload.Count): $relPath"
            if ($script:VerboseOutput) {
                Write-Host "  [$downloadCount / $($filesToDownload.Count)] $relPath" -ForegroundColor DarkGray
            }
        }

        $result = Download-DriveFile -DriveItem $file -OutputDir $filesDir -DriveId $driveId

        if ($result['success']) {
            $downloadedFiles.Add([ordered]@{
                relPath  = $result['relPath']
                path     = $result['path']
                sha256   = $result['sha256']
                size     = $result['size']
                itemId   = $result['itemId']
            })
            $totalBytes += $result['size']

            # Update eTag for this file
            $fId   = Get-SafeProp $file 'id'
            $fEtag = Get-SafeProp $file 'eTag'
            if ($fId -and $fEtag) { $fileEtags[$fId] = $fEtag }
        } else {
            $failedFiles.Add([ordered]@{
                relPath  = $result['relPath']
                itemId   = $result['itemId']
                itemName = $result['itemName']
                error    = $result['error']
            })
        }
    }

    # 7) Write manifest
    $manifest = [ordered]@{
        toolVersion    = $script:SUITE_VERSION
        toolName       = 'spbackup-library'
        runTimestamp    = $script:RunTimestamp
        site           = [ordered]@{
            id          = $siteId
            displayName = $siteName
            webUrl      = Get-SafeProp $site 'webUrl'
        }
        library        = [ordered]@{
            id                   = $driveId
            name                 = $driveName
            lastModifiedDateTime = $driveLastMod
        }
        totalFiles     = $allFiles.Count
        downloaded     = $downloadedFiles.Count
        failed         = $failedFiles.Count
        skipped        = $skippedCount
        totalBytes     = $totalBytes
        files          = $downloadedFiles
        failures       = $failedFiles
    }
    $manifestPath = Join-Path $outDir 'manifest.json'
    $manifestJson = $manifest | ConvertTo-Json -Depth 20
    Write-AtomicFile -Path $manifestPath -Content $manifestJson

    # 8) Update state
    $state[$stateKey] = @{
        deltaLink        = $newDelta
        fileEtags        = $fileEtags
        lastBackupTime   = (Get-Date).ToUniversalTime().ToString('o')
        totalFiles       = $allFiles.Count
        driveName        = $driveName
    }
    Save-State -Path $statePath -State $state

    # Summary
    $sizeDisplay = if ($totalBytes -ge 1GB) {
        '{0:N2} GB' -f ($totalBytes / 1GB)
    } elseif ($totalBytes -ge 1MB) {
        '{0:N1} MB' -f ($totalBytes / 1MB)
    } else {
        '{0:N0} KB' -f ($totalBytes / 1KB)
    }

    $summaryMsg = "Backup complete: $($downloadedFiles.Count) file(s) downloaded ($sizeDisplay)"
    if ($skippedCount -gt 0) {
        $summaryMsg += ", $skippedCount skipped (unchanged)"
    }
    if ($failedFiles.Count -gt 0) {
        $summaryMsg += ", $($failedFiles.Count) FAILED"
        $script:ExitCode = 1
    }

    Write-LogJsonl -Level 'INFO' -Event 'backup_complete' -DriveId $driveId -Message $summaryMsg
    Write-Host ''
    Write-Host $summaryMsg -ForegroundColor $(if ($failedFiles.Count -gt 0) { 'Yellow' } else { 'Green' })

    if ($failedFiles.Count -gt 0) {
        Write-Host ''
        Write-Host 'Failed files:' -ForegroundColor Yellow
        foreach ($f in $failedFiles) {
            Write-Host "  - $($f.relPath): $($f.error)" -ForegroundColor Yellow
        }
    }
}

function Invoke-LibraryVerifyCommand {
    [CmdletBinding()]
    param([hashtable]$Opts)

    $outDir = $Opts['out']
    if ([string]::IsNullOrWhiteSpace($outDir)) { throw '--out is required for the verify command.' }

    $manifestPath = Join-Path $outDir 'manifest.json'
    if (-not (Test-Path $manifestPath)) {
        Write-Error "manifest.json not found in $outDir"
        $script:ExitCode = 1
        return
    }

    $manifest = Get-Content $manifestPath -Raw -Encoding utf8 | ConvertFrom-Json

    $missingFiles = 0; $hashMismatch = 0; $checkedFiles = 0; $okFiles = 0

    $files = Get-SafeProp $manifest 'files'
    if ($files) {
        foreach ($f in @($files)) {
            $fPath   = Get-SafeProp $f 'path'
            $fSha256 = Get-SafeProp $f 'sha256'
            $fRel    = Get-SafeProp $f 'relPath'
            if (-not $fPath) { continue }
            $checkedFiles++
            if (-not (Test-Path -LiteralPath $fPath)) {
                Write-Host "MISSING: $fRel ($fPath)" -ForegroundColor Red
                $missingFiles++
            } elseif ($fSha256) {
                $actualHash = Get-FileSHA256 -Path $fPath
                if ($actualHash -ne $fSha256) {
                    Write-Host "HASH MISMATCH: $fRel (expected=$fSha256, actual=$actualHash)" -ForegroundColor Yellow
                    $hashMismatch++
                } else { $okFiles++ }
            } else { $okFiles++ }
        }
    }

    Write-Host ''
    Write-Host 'Verification complete:' -ForegroundColor Cyan
    Write-Host "  Checked:  $checkedFiles files"
    Write-Host "  OK:       $okFiles"
    Write-Host "  Missing:  $missingFiles"
    Write-Host "  Mismatch: $hashMismatch"

    $script:ExitCode = if ($missingFiles -gt 0 -or $hashMismatch -gt 0) { 2 } else { 0 }
}

function Invoke-LibraryDiagnoseCommand {
    [CmdletBinding()]
    param([hashtable]$Opts)

    Write-Host ''
    Write-Host '=== Document Library Backup Diagnostic ===' -ForegroundColor Cyan
    Write-Host ''

    # 1) Env vars & auth method
    Write-Host '1. Environment variables' -ForegroundColor Yellow
    $tenantId     = $env:TENANT_ID
    $clientId     = $env:CLIENT_ID
    $clientSecret = $env:CLIENT_SECRET
    Write-Host "   TENANT_ID:     $(if ($tenantId) { $tenantId } else { '(NOT SET)' })"
    Write-Host "   CLIENT_ID:     $(if ($clientId) { $clientId } else { '(NOT SET)' })"
    Write-Host "   CLIENT_SECRET: $(if ($clientSecret) { $clientSecret.Substring(0, [math]::Min(4, $clientSecret.Length)) + '***' + ' (' + $clientSecret.Length + ' chars)' } else { '(not set — will use certificate auth)' })"

    $cert = Find-Certificate
    if ($cert) {
        Write-Host "   CERT:          $($cert.Subject) (thumbprint=$($cert.Thumbprint))" -ForegroundColor Green
        $cert.Dispose()
    } elseif (-not $clientSecret) {
        Write-Host '   CERT:          (not found)' -ForegroundColor Red
    }

    if (-not $clientSecret -and -not $cert) {
        $authMethod = 'certificate'
    } elseif ($clientSecret) {
        $authMethod = 'client_secret'
    } else {
        $authMethod = 'certificate'
    }
    Write-Host "   Auth method:   $authMethod"
    Write-Host ''

    if (-not $tenantId -or -not $clientId) {
        Write-Host '   FAIL: TENANT_ID and CLIENT_ID are required.' -ForegroundColor Red
        return
    }
    if (-not $clientSecret -and -not (Find-Certificate)) {
        Write-Host '   FAIL: Neither CLIENT_SECRET nor a certificate is available.' -ForegroundColor Red
        return
    }

    # 2) Token
    Write-Host '2. Token acquisition' -ForegroundColor Yellow
    try {
        $token = Get-GraphToken
        Write-Host '   OK: Token acquired successfully' -ForegroundColor Green
    } catch {
        Write-Host "   FAIL: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    Write-Host ''

    # 3) JWT
    Write-Host '3. JWT token claims' -ForegroundColor Yellow
    $diag = Show-TokenDiagnostic -Token $token
    Write-Host ''

    # 4) Graph API
    Write-Host '4. Graph API connectivity tests' -ForegroundColor Yellow
    Show-GraphConnectivityTests
    Write-Host ''

    # 5) Site-specific tests
    $siteUrl = $Opts['url']
    if ($siteUrl) {
        Write-Host '5. Site-specific tests' -ForegroundColor Yellow
        try {
            $site = Resolve-SiteFromUrl -Url $siteUrl
            $siteId = Get-SafeProp $site 'id'
            $siteName = Get-SafeProp $site 'displayName'
            Write-Host "   Site resolved: $siteName (id=$siteId)" -ForegroundColor Green

            try {
                $drives = Get-SiteDocumentLibraries -SiteId $siteId
                Write-Host "   Document libraries found: $($drives.Count)" -ForegroundColor Green
                foreach ($d in $drives) {
                    $dName = Get-SafeProp $d 'name'
                    $dId   = Get-SafeProp $d 'id'
                    $dType = Get-SafeProp $d 'driveType'
                    Write-Host "     - $dName (id=$dId, type=$dType)" -ForegroundColor DarkGray
                }
            } catch {
                Write-Host "   Drive enumeration: FAIL — $($_.Exception.Message)" -ForegroundColor Red
            }

            # Try listing root of first drive
            if ($drives.Count -gt 0) {
                $testDrive = $drives[0]
                $testDriveId = Get-SafeProp $testDrive 'id'
                $testDriveName = Get-SafeProp $testDrive 'name'
                Write-Host ''
                Write-Host "   Testing file enumeration on '$testDriveName'..." -ForegroundColor Yellow
                try {
                    $testUri = "$($script:GRAPH_BASE)/drives/$testDriveId/root/children?`$select=id,name,size,file,folder&`$top=5"
                    $testResp = Invoke-GraphRequest -Uri $testUri -DriveId $testDriveId
                    $testItems = @()
                    if (Test-SafeProp $testResp 'value') { $testItems = @($testResp.value) }
                    Write-Host "   Root children (first 5): $($testItems.Count) item(s)" -ForegroundColor Green
                    foreach ($ti in $testItems) {
                        $tiName = Get-SafeProp $ti 'name'
                        $isFolder = Test-SafeProp $ti 'folder'
                        $tiSize = Get-SafeProp $ti 'size'
                        $type = if ($isFolder) { 'folder' } else { 'file' }
                        Write-Host "     - $tiName ($type$(if (-not $isFolder) { ", $tiSize bytes" }))" -ForegroundColor DarkGray
                    }
                } catch {
                    Write-Host "   File enumeration: FAIL — $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        } catch {
            Write-Host "   Site resolution: FAIL — $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host ''
    Write-Host '=== Diagnostic complete ===' -ForegroundColor Cyan
    Write-Host ''
}

# ─────────────────────────────────────────────────────────────────────────────
# Usage & Entry Point
# ─────────────────────────────────────────────────────────────────────────────
function Show-LibraryUsage {
    $usage = @"
$($script:TOOL_NAME) — SharePoint Document Library Backup

USAGE:
  pwsh ./backup-library.ps1 backup --url "<URL>" --library "<name>" --out "<dir>" [OPTIONS]
  pwsh ./backup-library.ps1 enumerate --url "<URL>"
  pwsh ./backup-library.ps1 verify --out "<dir>"
  pwsh ./backup-library.ps1 diagnose [--url "<URL>"] [--verbose]

  Or via the unified entry point:
  pwsh ./spbackup.ps1 library backup --url "<URL>" --library "<name>" --out "<dir>"

COMMANDS:
  backup      Download all files from a document library
  enumerate   List all document libraries in a SharePoint site
  verify      Verify backup integrity against manifest
  diagnose    Check auth, decode JWT, test Graph API access
  help        Show this usage information

BACKUP OPTIONS:
  --url <URL>           SharePoint site URL (required)
  --library <name>      Document library display name (required unless --library-id)
  --library-id <id>     Drive GUID (alternative to --library)
  --out <dir>           Output directory (required)
  --concurrency <N>     Max parallel downloads (default: 4)
  --since <ISO date>    Only files modified after this date
  --state <path>        State file path (default: <out>/.state.json)
  --force               Re-download everything, ignoring state / delta
  --dry-run             Enumerate files only; no downloads
  --verbose             Human-readable console output

ENVIRONMENT VARIABLES (required):
  TENANT_ID             Azure AD / Entra tenant ID
  CLIENT_ID             App registration client ID

AUTHENTICATION (one of the following):
  CLIENT_SECRET         App registration client secret
  CERT_PATH             Path to .pfx certificate (auto-discovered from ./certs/)
  CERT_PASSWORD         Certificate password (if any)
"@
    Write-Host $usage
}

function Main {
    [CmdletBinding()]
    param([string[]]$RawArgs)

    $parsed  = Parse-Arguments -RawArgs $RawArgs
    $command = $parsed.Command
    $opts    = $parsed.Options

    $script:VerboseOutput = [bool]$opts['verbose']

    switch ($command) {
        'enumerate' { Invoke-LibraryEnumerateCommand -Opts $opts }
        'libraries' { Invoke-LibraryEnumerateCommand -Opts $opts }   # alias
        'backup'    { Invoke-LibraryBackupCommand -Opts $opts }
        'verify'    { Invoke-LibraryVerifyCommand -Opts $opts }
        'diagnose'  { Invoke-LibraryDiagnoseCommand -Opts $opts }
        'help'      { Show-LibraryUsage }
        default {
            Write-Error "Unknown command: $command"
            Show-LibraryUsage
            exit 1
        }
    }
}

# Run when invoked directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.') {
    Main -RawArgs $args
    exit $script:ExitCode
}
