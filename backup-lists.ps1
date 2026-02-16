#!/usr/bin/env pwsh
#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft List Backup — export lists to CSV and download attachments.

.DESCRIPTION
    Part of the SharePoint Backup suite. Can be invoked directly or via spbackup.ps1.
    Exports Microsoft Lists to CSV and downloads all attachments using Microsoft Graph
    and SharePoint REST APIs.

.EXAMPLE
    pwsh ./backup-lists.ps1 backup --url "https://contoso.sharepoint.com/sites/team" --list "Tasks" --out "./backup"
    pwsh ./backup-lists.ps1 enumerate --url "https://contoso.sharepoint.com/sites/team"
    pwsh ./backup-lists.ps1 verify --out "./backup"
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

$script:TOOL_NAME = 'spbackup list'

# ─────────────────────────────────────────────────────────────────────────────
# List Discovery & Enumeration
# ─────────────────────────────────────────────────────────────────────────────
function Get-SiteLists {
    [CmdletBinding()]
    param(
        [string]$SiteId,
        [switch]$IncludeHidden
    )

    $lists = [System.Collections.Generic.List[object]]::new()
    $uri = "$($script:GRAPH_BASE)/sites/$SiteId/lists?`$select=id,displayName,description,webUrl,list,lastModifiedDateTime&`$top=200"

    while ($uri) {
        Write-LogJsonl -Level 'DEBUG' -Event 'list_enumerate' -SiteId $SiteId -Url $uri -Message 'Listing site lists'
        $response = Invoke-GraphRequest -Uri $uri -SiteId $SiteId

        $items = @()
        if (Test-SafeProp $response 'value') { $items = @($response.value) }

        foreach ($item in $items) {
            $listInfo = Get-SafeProp $item 'list'
            $isHidden = Get-SafeProp $listInfo 'hidden'
            $template = Get-SafeProp $listInfo 'template'

            if ($isHidden -and -not $IncludeHidden) { continue }
            if ($template -eq 'documentLibrary') { continue }

            $lists.Add($item)
        }

        $uri = $null
        if (Test-SafeProp $response '@odata.nextLink') {
            $uri = Get-SafeProp $response '@odata.nextLink'
        }
    }

    return $lists
}

function Get-ListColumns {
    [CmdletBinding()]
    param(
        [string]$SiteId,
        [string]$ListId
    )

    $columns = [System.Collections.Generic.List[object]]::new()
    $uri = "$($script:GRAPH_BASE)/sites/$SiteId/lists/$ListId/columns?`$top=200"

    while ($uri) {
        Write-LogJsonl -Level 'DEBUG' -Event 'columns_enumerate' -SiteId $SiteId -ListId $ListId -Url $uri -Message 'Listing columns'
        $response = Invoke-GraphRequest -Uri $uri -SiteId $SiteId -ListId $ListId

        $items = @()
        if (Test-SafeProp $response 'value') { $items = @($response.value) }

        foreach ($col in $items) {
            $columns.Add($col)
        }

        $uri = $null
        if (Test-SafeProp $response '@odata.nextLink') {
            $uri = Get-SafeProp $response '@odata.nextLink'
        }
    }

    return $columns
}

function Get-ListItems {
    [CmdletBinding()]
    param(
        [string]$SiteId,
        [string]$ListId,
        [string]$SinceDate = '',
        [string[]]$ExpandFields = @('fields')
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $expand = ($ExpandFields -join ',')
    $uri = "$($script:GRAPH_BASE)/sites/$SiteId/lists/$ListId/items?`$expand=${expand}&`$top=200"

    if ($SinceDate -ne '') {
        $uri += "&`$filter=lastModifiedDateTime ge '$SinceDate'"
    }

    while ($uri) {
        Write-LogJsonl -Level 'DEBUG' -Event 'items_enumerate' -SiteId $SiteId -ListId $ListId -Url $uri -Message 'Listing items'

        try {
            $response = Invoke-GraphRequest -Uri $uri -SiteId $SiteId -ListId $ListId
        } catch {
            Write-LogJsonl -Level 'WARN' -Event 'items_enumerate_fail' -SiteId $SiteId -ListId $ListId `
                -Url $uri -Message "Pagination failed (collected $($items.Count) items so far): $($_.Exception.Message)"
            break
        }

        $values = @()
        if (Test-SafeProp $response 'value') { $values = @($response.value) }

        Write-LogJsonl -Level 'DEBUG' -Event 'items_page' -SiteId $SiteId -ListId $ListId `
            -Message "Got $($values.Count) items in this page"

        foreach ($item in $values) {
            $items.Add($item)
        }

        $uri = $null
        if (Test-SafeProp $response '@odata.nextLink') {
            $uri = Get-SafeProp $response '@odata.nextLink'
        }
    }

    return $items
}

# ─────────────────────────────────────────────────────────────────────────────
# Attachments (via SharePoint REST API — requires certificate auth)
# ─────────────────────────────────────────────────────────────────────────────
function Get-ListItemAttachments {
    [CmdletBinding()]
    param(
        [string]$SiteWebUrl,
        [string]$ListId,
        [string]$ItemId
    )

    $attachments = [System.Collections.Generic.List[object]]::new()

    $spUri = "$SiteWebUrl/_api/web/lists(guid%27$ListId%27)/items($ItemId)/AttachmentFiles"
    $resp = Invoke-SharePointRequest -Uri $spUri -ListId $ListId -ItemId $ItemId

    $files = @()
    if ($resp -and (Test-SafeProp $resp 'value')) {
        $files = @($resp.value)
    } elseif ($resp -is [array]) {
        $files = @($resp)
    }

    foreach ($att in $files) {
        $fileName     = Get-SafeProp $att 'FileName'
        $serverRelUrl = Get-SafeProp $att 'ServerRelativeUrl'

        $downloadUrl = $null
        if ($serverRelUrl) {
            $downloadUrl = "https://$($script:SharePointHostname)$serverRelUrl"
        }

        $attachments.Add(@{
            name         = $fileName
            downloadUrl  = $downloadUrl
            serverRelUrl = $serverRelUrl
        })
    }

    return $attachments
}

function Download-Attachment {
    [CmdletBinding()]
    param(
        [hashtable]$Attachment,
        [string]$OutputDir,
        [string]$SiteWebUrl,
        [string]$SiteId,
        [string]$ListId,
        [string]$ItemId
    )

    $attName  = $Attachment['name']
    $safeName = New-SafeName -Name $attName
    $outPath  = Join-Path $OutputDir $safeName
    $tmpPath  = "$outPath.tmp.$PID"

    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    Write-LogJsonl -Level 'INFO' -Event 'download_attachment' -SiteId $SiteId -ListId $ListId -ItemId $ItemId `
        -Message "Downloading attachment: $attName"

    try {
        $serverRelUrl = $Attachment['serverRelUrl']
        if ($serverRelUrl) {
            $encodedPath = [uri]::EscapeDataString($serverRelUrl)
            $fileUri = "$SiteWebUrl/_api/web/GetFileByServerRelativeUrl(%27$encodedPath%27)/`$value"
            Invoke-SharePointRequest -Uri $fileUri -OutFile $tmpPath -ListId $ListId -ItemId $ItemId | Out-Null
        } else {
            throw 'No server-relative URL available for attachment.'
        }

        Move-Item -Path $tmpPath -Destination $outPath -Force
        Write-LogJsonl -Level 'INFO' -Event 'download_attachment_success' -SiteId $SiteId -ListId $ListId -ItemId $ItemId `
            -Message "Downloaded: $attName"
        return @{
            success  = $true
            name     = $attName
            path     = $outPath
            safeName = $safeName
        }
    } catch {
        if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force -ErrorAction SilentlyContinue }
        $conciseErr = Get-ConciseErrorMessage $_.Exception.Message
        Write-LogJsonl -Level 'ERROR' -Event 'download_attachment_fail' -SiteId $SiteId -ListId $ListId -ItemId $ItemId `
            -Message "Failed to download $attName : $conciseErr"
        return @{
            success = $false
            name    = $attName
            error   = $conciseErr
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CSV Export
# ─────────────────────────────────────────────────────────────────────────────
function Export-ListToCsv {
    [CmdletBinding()]
    param(
        [object[]]$Items,
        [object[]]$Columns,
        [string]$OutputPath
    )

    # Build column name mapping — filter to user-facing columns
    $userColumns = [System.Collections.Generic.List[object]]::new()
    foreach ($col in $Columns) {
        $colName        = Get-SafeProp $col 'name'
        $colDisplayName = Get-SafeProp $col 'displayName'
        $hidden         = Get-SafeProp $col 'hidden'

        $systemKeep = @('Title', 'Created', 'Modified', 'Author', 'Editor', 'id', 'Attachments', '_UIVersionString')
        if ($hidden -and $colName -notin $systemKeep) { continue }

        $userColumns.Add(@{
            name        = $colName
            displayName = $colDisplayName ?? $colName
        })
    }

    $fieldNames = [System.Collections.Generic.List[string]]::new()
    foreach ($col in $userColumns) {
        $fieldNames.Add($col['name'])
    }

    # Build CSV rows
    $csvRows = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($item in $Items) {
        $fields = Get-SafeProp $item 'fields'
        if (-not $fields) { continue }

        $row = @{}
        foreach ($colInfo in $userColumns) {
            $colName = $colInfo['name']
            $val = Get-SafeProp $fields $colName

            if ($null -eq $val) {
                $row[$colName] = ''
            } elseif ($val -is [array] -or $val -is [System.Collections.IList]) {
                $row[$colName] = ($val | ForEach-Object {
                    if ($_ -is [PSCustomObject] -or $_ -is [hashtable]) {
                        $lv    = Get-SafeProp $_ 'LookupValue'
                        $email = Get-SafeProp $_ 'Email'
                        if ($lv) { $lv } elseif ($email) { $email } else { ($_ | ConvertTo-Json -Compress -Depth 2) }
                    } else {
                        [string]$_
                    }
                }) -join '; '
            } elseif ($val -is [PSCustomObject]) {
                $lv    = Get-SafeProp $val 'LookupValue'
                $email = Get-SafeProp $val 'Email'
                if ($lv) { $row[$colName] = $lv }
                elseif ($email) { $row[$colName] = $email }
                else { $row[$colName] = ($val | ConvertTo-Json -Compress -Depth 2) }
            } else {
                $row[$colName] = [string]$val
            }
        }

        $csvRows.Add($row)
    }

    # Build header row with deduplication
    $headerMap = @{}
    $displayNameCounts = @{}
    foreach ($col in $userColumns) {
        $dn = $col['displayName']
        if ($displayNameCounts.ContainsKey($dn)) { $displayNameCounts[$dn]++ }
        else { $displayNameCounts[$dn] = 1 }
    }
    foreach ($col in $userColumns) {
        $dn = $col['displayName']
        if ($displayNameCounts[$dn] -gt 1) {
            $headerMap[$col['name']] = "$dn ($($col['name']))"
        } else {
            $headerMap[$col['name']] = $dn
        }
    }

    $orderedFieldNames = @($fieldNames)
    $sb = [System.Text.StringBuilder]::new()

    # Header
    $headerLine = ($orderedFieldNames | ForEach-Object { Format-CsvField ($headerMap[$_] ?? $_) }) -join ','
    $sb.AppendLine($headerLine) | Out-Null

    # Data rows
    foreach ($row in $csvRows) {
        $line = ($orderedFieldNames | ForEach-Object { Format-CsvField ($row[$_] ?? '') }) -join ','
        $sb.AppendLine($line) | Out-Null
    }

    Write-AtomicFile -Path $OutputPath -Content $sb.ToString()
    Write-LogJsonl -Level 'INFO' -Event 'csv_export' -Message "Exported $($csvRows.Count) rows to CSV"
}

# ─────────────────────────────────────────────────────────────────────────────
# Commands
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-ListEnumerateCommand {
    [CmdletBinding()]
    param([hashtable]$Opts)

    $siteUrl = $Opts['url'] ?? $Opts['site-url']
    if ([string]::IsNullOrWhiteSpace($siteUrl)) { throw '--url is required for the enumerate command.' }

    $includeHidden = [bool]$Opts['include-hidden']

    $site    = Resolve-SiteFromUrl -Url $siteUrl
    $siteId  = Get-SafeProp $site 'id'
    $siteName = Get-SafeProp $site 'displayName'

    $lists = Get-SiteLists -SiteId $siteId -IncludeHidden:$includeHidden

    Write-Host ''
    Write-Host "Lists in '$siteName':" -ForegroundColor Cyan
    Write-Host ''

    if ($lists.Count -eq 0) {
        Write-Host '  (no lists found)' -ForegroundColor Yellow
        return
    }

    foreach ($list in $lists) {
        $listName = Get-SafeProp $list 'displayName'
        $listId   = Get-SafeProp $list 'id'
        $listInfo = Get-SafeProp $list 'list'
        $template = Get-SafeProp $listInfo 'template'
        $lastMod  = Get-SafeProp $list 'lastModifiedDateTime'

        Write-Host "  - $listName" -ForegroundColor White
        Write-Host "    ID: $listId" -ForegroundColor DarkGray
        Write-Host "    Template: $template" -ForegroundColor DarkGray
        Write-Host "    Last modified: $lastMod" -ForegroundColor DarkGray
        Write-Host ''
    }

    Write-LogJsonl -Level 'INFO' -Event 'lists_enumerated' -SiteId $siteId -Message "Found $($lists.Count) list(s)"
}

function Invoke-ListBackupCommand {
    [CmdletBinding()]
    param([hashtable]$Opts)

    $siteUrl      = $Opts['url'] ?? $Opts['site-url']
    $listName     = $Opts['list']
    $listId       = $Opts['list-id']
    $outDir       = $Opts['out']
    $concurrency  = [int]($Opts['concurrency'] ?? 4)
    $sinceDate    = $Opts['since']       ?? ''
    $statePath    = $Opts['state']       ?? ''
    $dryRun       = [bool]$Opts['dry-run']
    $force        = [bool]$Opts['force']
    $skipAttach   = [bool]$Opts['skip-attachments']

    # Auto-extract list name from URL if it contains /Lists/<name>
    if ($siteUrl -and -not $listName -and -not $listId) {
        if ($siteUrl -match '/Lists/([^/]+)') {
            $listName = [System.Uri]::UnescapeDataString($Matches[1])
            Write-LogJsonl -Level 'INFO' -Event 'url_list_extracted' -Message "Extracted list name from URL: '$listName'"
        }
    }

    if ([string]::IsNullOrWhiteSpace($siteUrl))   { throw '--url is required for the backup command.' }
    if ([string]::IsNullOrWhiteSpace($listName) -and [string]::IsNullOrWhiteSpace($listId)) {
        throw '--list or --list-id is required for the backup command (or use --url with a /Lists/<name> URL).'
    }
    if ([string]::IsNullOrWhiteSpace($outDir))    { throw '--out is required for the backup command.' }

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
        -Message "List backup starting: list=$($listName ?? $listId), concurrency=$concurrency, skipAttachments=$skipAttach"

    # 1) Resolve site
    $site      = Resolve-SiteFromUrl -Url $siteUrl
    $siteId    = Get-SafeProp $site 'id'
    $siteName  = Get-SafeProp $site 'displayName'
    $siteWebUrl = Get-SafeProp $site 'webUrl'

    Write-LogJsonl -Level 'INFO' -Event 'site_resolved' -SiteId $siteId -Message "Site: $siteName"

    # 2) Find the list
    $targetList = $null

    if ($listId) {
        Write-LogJsonl -Level 'INFO' -Event 'list_lookup_by_id' -SiteId $siteId -ListId $listId -Message "Looking up list by ID"
        try {
            $targetList = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/sites/$siteId/lists/${listId}?`$select=id,displayName,description,webUrl,list,lastModifiedDateTime" `
                -SiteId $siteId -ListId $listId
        } catch {
            throw "Could not find list with ID '$listId': $($_.Exception.Message)"
        }
    } else {
        Write-LogJsonl -Level 'INFO' -Event 'list_lookup_by_name' -SiteId $siteId -Message "Looking for list named '$listName'"
        $allLists = Get-SiteLists -SiteId $siteId -IncludeHidden
        foreach ($l in $allLists) {
            $lName = Get-SafeProp $l 'displayName'
            if ($lName -eq $listName) { $targetList = $l; break }
        }
        if (-not $targetList) {
            foreach ($l in $allLists) {
                $lName = Get-SafeProp $l 'displayName'
                if ($lName -and $lName.Equals($listName, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $targetList = $l; break
                }
            }
        }
        if (-not $targetList) {
            $urlEncodedName = [System.Uri]::EscapeDataString($listName)
            foreach ($l in $allLists) {
                $lWebUrl = Get-SafeProp $l 'webUrl'
                if ($lWebUrl -and ($lWebUrl -match "/Lists/$([regex]::Escape($listName))(/|$)" -or
                    $lWebUrl -match "/Lists/$([regex]::Escape($urlEncodedName))(/|$)")) {
                    $targetList = $l
                    Write-LogJsonl -Level 'INFO' -Event 'list_matched_by_url' -SiteId $siteId `
                        -Message "Matched list by URL path: displayName='$(Get-SafeProp $l 'displayName')', searched='$listName'"
                    break
                }
            }
        }
        if (-not $targetList) {
            $availableNames = ($allLists | ForEach-Object { "'$(Get-SafeProp $_ 'displayName')'" }) -join ', '
            throw "Could not find a list named '$listName' in site '$siteName'. Available lists: $availableNames"
        }
    }

    $targetListId      = Get-SafeProp $targetList 'id'
    $targetListName    = Get-SafeProp $targetList 'displayName'
    $targetListLastMod = Get-SafeProp $targetList 'lastModifiedDateTime'

    Write-LogJsonl -Level 'INFO' -Event 'list_found' -SiteId $siteId -ListId $targetListId `
        -Message "Found list: $targetListName (lastModified=$targetListLastMod)"

    # 3) Get column definitions
    Write-LogJsonl -Level 'INFO' -Event 'columns_fetch' -SiteId $siteId -ListId $targetListId -Message 'Fetching column definitions'
    $columns = @(Get-ListColumns -SiteId $siteId -ListId $targetListId)
    Write-LogJsonl -Level 'INFO' -Event 'columns_fetched' -SiteId $siteId -ListId $targetListId `
        -Message "Found $($columns.Count) columns"

    $hasAttachmentsCol = $false
    foreach ($col in $columns) {
        if ((Get-SafeProp $col 'name') -eq 'Attachments') { $hasAttachmentsCol = $true; break }
    }

    # 4) Incremental sync check
    $state = @{}
    if (-not $force) { $state = Read-State -Path $statePath }

    $stateListKey  = "list:${targetListId}"
    $lastBackupMod = $null
    if ($state.ContainsKey($stateListKey)) {
        $lastBackupMod = $state[$stateListKey]['lastModifiedDateTime']
    }

    if (-not $force -and $lastBackupMod -and $targetListLastMod -eq $lastBackupMod -and -not $sinceDate) {
        Write-LogJsonl -Level 'INFO' -Event 'skip_unchanged_list' -SiteId $siteId -ListId $targetListId `
            -Message "List unchanged since last backup (lastModified=$targetListLastMod)"
        Write-Host "List '$targetListName' has not changed since the last backup. Use --force to re-export." -ForegroundColor Yellow
        return
    }

    # 5) Fetch all list items
    Write-LogJsonl -Level 'INFO' -Event 'items_fetch' -SiteId $siteId -ListId $targetListId -Message 'Fetching list items'
    $listItems = @(Get-ListItems -SiteId $siteId -ListId $targetListId -SinceDate $sinceDate)
    Write-LogJsonl -Level 'INFO' -Event 'items_fetched' -SiteId $siteId -ListId $targetListId `
        -Message "Found $($listItems.Count) items"

    if ($listItems.Count -eq 0) {
        Write-LogJsonl -Level 'WARN' -Event 'no_items' -Message 'No items found in the list.'
        Write-Warning 'No items found in the list.'
    }

    if ($dryRun) {
        Write-Host "Dry run: found $($listItems.Count) items in '$targetListName'" -ForegroundColor Cyan
        Write-Host "  Columns: $($columns.Count)" -ForegroundColor DarkGray
        Write-Host "  Has Attachments column: $hasAttachmentsCol" -ForegroundColor DarkGray
        return
    }

    # 6) Export to CSV
    $safeListName = New-SafeName -Name $targetListName
    $csvPath = Join-Path $outDir "${safeListName}.csv"

    Write-LogJsonl -Level 'INFO' -Event 'csv_export_start' -SiteId $siteId -ListId $targetListId `
        -Message "Exporting $($listItems.Count) items to CSV"

    Export-ListToCsv -Items $listItems -Columns $columns -OutputPath $csvPath

    # Save column definitions as JSON
    $columnsPath = Join-Path $outDir "${safeListName}_columns.json"
    $columnsJson = ($columns | ForEach-Object {
        $colType = 'unknown'
        foreach ($t in @('text', 'number', 'dateTime', 'choice', 'lookup', 'boolean',
                         'calculated', 'personOrGroup', 'currency', 'hyperlink',
                         'thumbnail', 'contentApprovalStatus', 'geolocation', 'term')) {
            if (Test-SafeProp $_ $t) { $colType = $t; break }
        }
        [ordered]@{
            name        = Get-SafeProp $_ 'name'
            displayName = Get-SafeProp $_ 'displayName'
            type        = $colType
            readOnly    = Get-SafeProp $_ 'readOnly'
            hidden      = Get-SafeProp $_ 'hidden'
        }
    }) | ConvertTo-Json -Depth 5
    Write-AtomicFile -Path $columnsPath -Content $columnsJson

    # 7) Download attachments
    $attachmentResults  = [System.Collections.Generic.List[object]]::new()
    $attachmentFailures = [System.Collections.Generic.List[object]]::new()
    $totalAttachments   = 0
    $downloadedAttachments = 0

    if (-not $skipAttach -and $hasAttachmentsCol) {
        $spAccessOk = Test-SharePointAccess -SiteWebUrl $siteWebUrl
        if (-not $spAccessOk) {
            Write-LogJsonl -Level 'WARN' -Event 'attachments_sp_no_access' -SiteId $siteId -ListId $targetListId `
                -Message 'SharePoint REST API not accessible — cannot download attachments. Cert-based auth required.'
            Write-Host ''
            Write-Host '  ⚠ Attachment download skipped — SharePoint REST API access not available.' -ForegroundColor Yellow
            Write-Host '    SharePoint requires certificate-based auth for app-only REST API access.' -ForegroundColor DarkYellow
            Write-Host '    To enable attachments:' -ForegroundColor DarkYellow
            Write-Host '      1. Place a .pfx certificate at ./certs/spbackup.pfx (or set CERT_PATH)' -ForegroundColor White
            Write-Host '      2. Upload the .cer public key to your app registration in Azure Portal' -ForegroundColor White
            Write-Host '      3. Grant SharePoint > Sites.Read.All (Application) permission + admin consent' -ForegroundColor White
            Write-Host ''
        } else {
            $attachDir = Join-Path $outDir 'attachments'
            if (-not (Test-Path $attachDir)) {
                New-Item -ItemType Directory -Path $attachDir -Force | Out-Null
            }

            Write-LogJsonl -Level 'INFO' -Event 'attachments_start' -SiteId $siteId -ListId $targetListId `
                -Message 'Starting attachment download'

            foreach ($item in $listItems) {
                $itemId = Get-SafeProp $item 'id'
                $fields = Get-SafeProp $item 'fields'
                $hasAttachments = Get-SafeProp $fields 'Attachments'

                if (-not $hasAttachments) { continue }

                $itemTitle = Get-SafeProp $fields 'Title'
                if (-not $itemTitle) { $itemTitle = "item_$itemId" }
                $safeItemName  = New-SafeName -Name $itemTitle
                $itemAttachDir = Join-Path $attachDir "${safeItemName}__${itemId}"

                Write-LogJsonl -Level 'INFO' -Event 'attachments_item' -SiteId $siteId -ListId $targetListId -ItemId $itemId `
                    -Message "Checking attachments for: $itemTitle"

                try {
                    $attachments = @(Get-ListItemAttachments -SiteWebUrl $siteWebUrl -ListId $targetListId -ItemId $itemId)

                    if ($attachments.Count -eq 0) {
                        Write-LogJsonl -Level 'DEBUG' -Event 'attachments_none' -SiteId $siteId -ListId $targetListId -ItemId $itemId `
                            -Message "No attachments found (Attachments flag was true but no files)"
                        continue
                    }

                    $totalAttachments += $attachments.Count
                    Write-LogJsonl -Level 'INFO' -Event 'attachments_found' -SiteId $siteId -ListId $targetListId -ItemId $itemId `
                        -Message "Found $($attachments.Count) attachment(s)"

                    foreach ($att in $attachments) {
                        $attResult = Download-Attachment -Attachment $att -OutputDir $itemAttachDir `
                            -SiteWebUrl $siteWebUrl -SiteId $siteId -ListId $targetListId -ItemId $itemId

                        if ($attResult['success']) {
                            $downloadedAttachments++
                            $attachmentResults.Add([ordered]@{
                                itemId    = $itemId
                                itemTitle = $itemTitle
                                fileName  = $attResult['name']
                                safeName  = $attResult['safeName']
                                path      = $attResult['path']
                                sha256    = Get-FileSHA256 -Path $attResult['path']
                            })
                        } else {
                            $attachmentFailures.Add([ordered]@{
                                itemId    = $itemId
                                itemTitle = $itemTitle
                                fileName  = $attResult['name']
                                error     = $attResult['error']
                            })
                        }
                    }
                } catch {
                    $conciseErr = Get-ConciseErrorMessage $_.Exception.Message
                    Write-LogJsonl -Level 'ERROR' -Event 'attachments_item_fail' -SiteId $siteId -ListId $targetListId -ItemId $itemId `
                        -Message "Failed to get attachments for item $itemId : $conciseErr"
                    $attachmentFailures.Add([ordered]@{
                        itemId    = $itemId
                        itemTitle = $itemTitle
                        fileName  = '(enumeration failed)'
                        error     = $conciseErr
                    })
                }
            }

            Write-LogJsonl -Level 'INFO' -Event 'attachments_complete' -SiteId $siteId -ListId $targetListId `
                -Message "Attachments: $downloadedAttachments downloaded, $($attachmentFailures.Count) failed (of $totalAttachments total)"
        }
    } elseif (-not $skipAttach -and -not $hasAttachmentsCol) {
        Write-LogJsonl -Level 'INFO' -Event 'attachments_skip_no_column' -SiteId $siteId -ListId $targetListId `
            -Message 'List does not have an Attachments column — skipping attachment download'
    } else {
        Write-LogJsonl -Level 'INFO' -Event 'attachments_skip_flag' -SiteId $siteId -ListId $targetListId `
            -Message 'Attachment download skipped (--skip-attachments)'
    }

    # 8) Write manifest
    $manifest = [ordered]@{
        toolVersion    = $script:SUITE_VERSION
        toolName       = 'spbackup-list'
        runTimestamp    = $script:RunTimestamp
        site           = [ordered]@{
            id          = $siteId
            displayName = $siteName
            webUrl      = Get-SafeProp $site 'webUrl'
        }
        list           = [ordered]@{
            id                   = $targetListId
            displayName          = $targetListName
            lastModifiedDateTime = $targetListLastMod
        }
        totalItems     = $listItems.Count
        csvFile        = "${safeListName}.csv"
        csvSha256      = Get-FileSHA256 -Path $csvPath
        columnsFile    = "${safeListName}_columns.json"
        attachments    = [ordered]@{
            total      = $totalAttachments
            downloaded = $downloadedAttachments
            failed     = $attachmentFailures.Count
            files      = $attachmentResults
            failures   = $attachmentFailures
        }
    }
    $manifestPath = Join-Path $outDir 'manifest.json'
    $manifestJson = $manifest | ConvertTo-Json -Depth 20
    Write-AtomicFile -Path $manifestPath -Content $manifestJson

    # 9) Update state
    $state[$stateListKey] = @{
        lastModifiedDateTime = $targetListLastMod
        lastBackupTime       = (Get-Date).ToUniversalTime().ToString('o')
        itemCount            = $listItems.Count
        csvSha256            = Get-FileSHA256 -Path $csvPath
    }
    Save-State -Path $statePath -State $state

    # Summary
    $summaryMsg = "Backup complete: $($listItems.Count) items exported to CSV"
    if (-not $skipAttach -and $hasAttachmentsCol) {
        $summaryMsg += ", $downloadedAttachments attachments downloaded"
        if ($attachmentFailures.Count -gt 0) {
            $summaryMsg += ", $($attachmentFailures.Count) attachment(s) failed"
            $script:ExitCode = 1
        }
    }
    Write-LogJsonl -Level 'INFO' -Event 'backup_complete' -Message $summaryMsg
    Write-Host $summaryMsg -ForegroundColor $(if ($attachmentFailures.Count -gt 0) { 'Yellow' } else { 'Green' })

    if ($attachmentFailures.Count -gt 0) {
        Write-Host "Attachment failures:" -ForegroundColor Yellow
        foreach ($f in $attachmentFailures) {
            Write-Host "  - Item '$($f.itemTitle)' / $($f.fileName): $($f.error)" -ForegroundColor Yellow
        }
    }
}

function Invoke-ListVerifyCommand {
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

    # Verify CSV file
    $csvFile   = Get-SafeProp $manifest 'csvFile'
    $csvSha256 = Get-SafeProp $manifest 'csvSha256'
    if ($csvFile) {
        $csvPath = Join-Path $outDir $csvFile
        $checkedFiles++
        if (-not (Test-Path $csvPath)) {
            Write-Host "MISSING: $csvFile" -ForegroundColor Red
            $missingFiles++
        } elseif ($csvSha256) {
            $actualHash = Get-FileSHA256 -Path $csvPath
            if ($actualHash -ne $csvSha256) {
                Write-Host "HASH MISMATCH: $csvFile (expected=$csvSha256, actual=$actualHash)" -ForegroundColor Yellow
                $hashMismatch++
            } else { $okFiles++ }
        } else { $okFiles++ }
    }

    # Verify attachment files
    $attachments = Get-SafeProp $manifest 'attachments'
    $attFiles = if ($attachments) { Get-SafeProp $attachments 'files' } else { $null }
    if ($attFiles) {
        foreach ($att in @($attFiles)) {
            $attPath   = Get-SafeProp $att 'path'
            $attSha256 = Get-SafeProp $att 'sha256'
            $attName   = Get-SafeProp $att 'fileName'
            if (-not $attPath) { continue }
            $checkedFiles++
            if (-not (Test-Path $attPath)) {
                Write-Host "MISSING: $attName ($attPath)" -ForegroundColor Red
                $missingFiles++
            } elseif ($attSha256) {
                $actualHash = Get-FileSHA256 -Path $attPath
                if ($actualHash -ne $attSha256) {
                    Write-Host "HASH MISMATCH: $attName (expected=$attSha256, actual=$actualHash)" -ForegroundColor Yellow
                    $hashMismatch++
                } else { $okFiles++ }
            } else { $okFiles++ }
        }
    }

    Write-Host ""
    Write-Host "Verification complete:" -ForegroundColor Cyan
    Write-Host "  Checked:  $checkedFiles files"
    Write-Host "  OK:       $okFiles"
    Write-Host "  Missing:  $missingFiles"
    Write-Host "  Mismatch: $hashMismatch"

    $script:ExitCode = if ($missingFiles -gt 0 -or $hashMismatch -gt 0) { 2 } else { 0 }
}

function Invoke-ListDiagnoseCommand {
    [CmdletBinding()]
    param([hashtable]$Opts)

    Write-Host ''
    Write-Host '=== List Backup Diagnostic ===' -ForegroundColor Cyan
    Write-Host ''

    # 1) Env vars
    Write-Host '1. Environment variables' -ForegroundColor Yellow
    $tenantId     = $env:TENANT_ID
    $clientId     = $env:CLIENT_ID
    $clientSecret = $env:CLIENT_SECRET
    Write-Host "   TENANT_ID:     $(if ($tenantId) { $tenantId } else { '(NOT SET)' })"
    Write-Host "   CLIENT_ID:     $(if ($clientId) { $clientId } else { '(NOT SET)' })"
    Write-Host "   CLIENT_SECRET: $(if ($clientSecret) { $clientSecret.Substring(0, [math]::Min(4, $clientSecret.Length)) + '***' + ' (' + $clientSecret.Length + ' chars)' } else { '(NOT SET)' })"

    $certPath = $env:CERT_PATH
    if (-not $certPath) {
        $autoCert = Join-Path $script:ProjectRoot 'certs' 'spbackup.pfx'
        if (Test-Path $autoCert) { $certPath = $autoCert }
    }
    if ($certPath -and (Test-Path $certPath)) {
        Write-Host "   CERT_PATH:     $certPath (found)" -ForegroundColor Green
    } else {
        Write-Host '   CERT_PATH:     (not found — attachments will be unavailable)' -ForegroundColor DarkYellow
    }
    Write-Host ''

    if (-not $tenantId -or -not $clientId -or -not $clientSecret) {
        Write-Host '   FAIL: Required environment variables are missing.' -ForegroundColor Red
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
    $siteUrl = $Opts['url'] ?? $Opts['site-url']
    if ($siteUrl) {
        Write-Host '5. Site-specific tests' -ForegroundColor Yellow
        try {
            $site = Resolve-SiteFromUrl -Url $siteUrl
            $siteId = Get-SafeProp $site 'id'
            $siteName = Get-SafeProp $site 'displayName'
            Write-Host "   Site resolved: $siteName (id=$siteId)" -ForegroundColor Green
            try {
                $lists = Get-SiteLists -SiteId $siteId
                Write-Host "   Lists found: $($lists.Count)" -ForegroundColor Green
                foreach ($l in $lists) {
                    Write-Host "     - $(Get-SafeProp $l 'displayName') (id=$(Get-SafeProp $l 'id'))" -ForegroundColor DarkGray
                }
            } catch {
                Write-Host "   List enumeration: FAIL — $($_.Exception.Message)" -ForegroundColor Red
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
function Show-ListUsage {
    $usage = @"
$($script:TOOL_NAME) — Microsoft List Backup

USAGE:
  pwsh ./backup-lists.ps1 backup --url "<URL>" --list "<name>" --out "<dir>" [OPTIONS]
  pwsh ./backup-lists.ps1 enumerate --url "<URL>"
  pwsh ./backup-lists.ps1 verify --out "<dir>"
  pwsh ./backup-lists.ps1 diagnose [--url "<URL>"] [--verbose]

  Or via the unified entry point:
  pwsh ./spbackup.ps1 list backup --url "<URL>" --list "<name>" --out "<dir>"

COMMANDS:
  backup      Export a Microsoft List to CSV and download attachments
  enumerate   List all Microsoft Lists in a SharePoint site
  verify      Verify backup integrity against manifest
  diagnose    Check auth, decode JWT, test Graph API access
  help        Show this usage information

BACKUP OPTIONS:
  --url <URL>           SharePoint site or list URL (required)
  --site-url <URL>      Alias for --url (backward-compatible)
  --list <name>         List display name (required unless --list-id is given)
  --list-id <id>        List GUID (alternative to --list)
  --out <dir>           Output directory (required)
  --concurrency <N>     Max parallel downloads (default: 4)
  --since <ISO date>    Only items modified after this date
  --state <path>        State file path (default: <out>/.state.json)
  --force               Re-export everything, ignoring state
  --dry-run             Enumerate only; no downloads or CSV export
  --skip-attachments    Skip attachment download (CSV only)
  --verbose             Human-readable console output

ENVIRONMENT VARIABLES (required):
  TENANT_ID             Azure AD / Entra tenant ID
  CLIENT_ID             App registration client ID
  CLIENT_SECRET         App registration client secret

ENVIRONMENT VARIABLES (for attachments):
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
        'enumerate' { Invoke-ListEnumerateCommand -Opts $opts }
        'lists'     { Invoke-ListEnumerateCommand -Opts $opts }   # alias
        'backup'    { Invoke-ListBackupCommand -Opts $opts }
        'verify'    { Invoke-ListVerifyCommand -Opts $opts }
        'diagnose'  { Invoke-ListDiagnoseCommand -Opts $opts }
        'help'      { Show-ListUsage }
        default {
            Write-Error "Unknown command: $command"
            Show-ListUsage
            exit 1
        }
    }
}

# Run when invoked directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.') {
    Main -RawArgs $args
    exit $script:ExitCode
}
