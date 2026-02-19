#!/usr/bin/env pwsh
#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Loop Backup — export Loop pages to HTML, Markdown, and raw .loop files.

.DESCRIPTION
    Part of the SharePoint Backup suite. Can be invoked directly or via spbackup.ps1.
    Resolves Loop URLs (including loop.cloud.microsoft links), enumerates .loop pages,
    and exports them as HTML, Markdown, and/or raw .loop files.

.EXAMPLE
    pwsh ./backup-loop.ps1 backup --url "https://..." --out "./backup"
    pwsh ./backup-loop.ps1 resolve --url "https://..."
    pwsh ./backup-loop.ps1 verify --out "./backup"
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

$script:TOOL_NAME        = 'spbackup loop'
$script:LARGE_ITEM_THRESHOLD = 1000
$script:PythonExe        = $null
$script:PandocExe        = $null

# ─────────────────────────────────────────────────────────────────────────────
# Loop URL Decoding & Resolution
# ─────────────────────────────────────────────────────────────────────────────
function Expand-LoopAppUrl {
    <#
    .SYNOPSIS
        Detect and decode Loop app URLs (loop.cloud.microsoft, loop.microsoft.com).
        Returns an ordered list of candidate SharePoint/OneDrive URLs to try,
        plus any extracted item GUIDs. Returns $null if the URL is not a Loop app URL.
    #>
    [CmdletBinding()]
    param([string]$Url)

    if ($Url -notmatch 'loop\.(cloud\.)?microsoft(\.com)?(/|$)') {
        return $null
    }

    Write-LogJsonl -Level 'INFO' -Event 'loop_app_url' -Url $Url -Message 'Detected Loop app URL; decoding payload...'

    $payload = $null
    if ($Url -match '/p/([A-Za-z0-9_\-+/=%]+)$') {
        $payload = $Matches[1]
    } elseif ($Url -match '/p/([A-Za-z0-9_\-+/=%]+)\?') {
        $payload = $Matches[1]
    }

    if (-not $payload) {
        Write-LogJsonl -Level 'WARN' -Event 'loop_app_decode_fail' -Url $Url -Message 'Could not extract base64 payload from Loop URL'
        return $null
    }

    $payload = [System.Uri]::UnescapeDataString($payload)
    $padLen = (4 - ($payload.Length % 4)) % 4
    $payload = $payload + ('=' * $padLen)

    try {
        $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
        $data = $decoded | ConvertFrom-Json
    } catch {
        Write-LogJsonl -Level 'WARN' -Event 'loop_app_decode_fail' -Url $Url -Message "Base64/JSON decode failed: $($_.Exception.Message)"
        return $null
    }

    $result = @{
        candidateUrls   = [System.Collections.Generic.List[string]]::new()
        itemId          = $null
        workspaceUrl    = $null
        pageUrl         = $null
        driveIdentifier = $null
        folderItemId    = $null
        pageItemId      = $null
    }

    # Extract item GUID (i.i)
    $dataI  = Get-SafeProp $data 'i'
    $dataIi = Get-SafeProp $dataI 'i'
    if ($dataIi) {
        $result.itemId = $dataIi
        Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_item_id' -ItemId $dataIi -Message "Extracted item ID"
    }

    # Extract page URL (p.u)
    $dataP  = Get-SafeProp $data 'p'
    $dataPu = Get-SafeProp $dataP 'u'
    if ($dataPu) {
        $pageUrl = [string]$dataPu
        $result.pageUrl = $pageUrl

        $cleanPage = $pageUrl -replace '[&?]nav=[^&]*', ''
        $result.candidateUrls.Add($cleanPage)
        Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_page_url' -Url $cleanPage -Message "Extracted page URL"

        if ($pageUrl -match 'nav=([A-Za-z0-9_\-+/=%]+)') {
            $navB64 = [System.Uri]::UnescapeDataString($Matches[1])
            $navPad = (4 - ($navB64.Length % 4)) % 4
            $navB64 = $navB64 + ('=' * $navPad)
            try {
                $navDecoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($navB64))
                if ($navDecoded -match 'd=([^&]+)') {
                    $result['driveIdentifier'] = $Matches[1]
                    Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_page_drive_id' -Message "Page nav drive identifier: $($Matches[1])"
                }
                if ($navDecoded -match 'f=([^&]+)') {
                    $result['pageItemId'] = $Matches[1]
                    Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_page_item_id' -Message "Page nav item ID: $($Matches[1])"
                }
            } catch {
                Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_page_nav_decode_fail' -Message $_.Exception.Message
            }
        }
    }

    # Extract workspace URL (w.u)
    $dataW  = Get-SafeProp $data 'w'
    $dataWu = Get-SafeProp $dataW 'u'
    if ($dataWu) {
        $wsUrl = [string]$dataWu
        $result.workspaceUrl = $wsUrl

        $cleanWs = $wsUrl -replace '[&?]nav=[^&]*', ''
        if ($cleanWs -ne ($result.candidateUrls | Select-Object -First 1)) {
            $result.candidateUrls.Add($cleanWs)
        }
        Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_ws_url' -Url $cleanWs -Message "Extracted workspace URL"

        if ($wsUrl -match 'nav=([A-Za-z0-9_\-+/=%]+)') {
            $navB64 = [System.Uri]::UnescapeDataString($Matches[1])
            $navPad = (4 - ($navB64.Length % 4)) % 4
            $navB64 = $navB64 + ('=' * $navPad)
            try {
                $navDecoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($navB64))
                if ($navDecoded -match 'd=([^&]+)') {
                    if (-not $result['driveIdentifier']) {
                        $result['driveIdentifier'] = $Matches[1]
                    }
                    Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_drive_id' -Message "Extracted drive identifier: $($Matches[1])"
                }
                if ($navDecoded -match 'f=([^&]+)') {
                    $result['folderItemId'] = $Matches[1]
                    Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_folder_id' -Message "Extracted folder item ID: $($Matches[1])"
                }
            } catch {
                Write-LogJsonl -Level 'DEBUG' -Event 'loop_app_nav_decode_fail' -Message $_.Exception.Message
            }
        }
    }

    if ($result.candidateUrls.Count -eq 0) {
        Write-LogJsonl -Level 'WARN' -Event 'loop_app_no_urls' -Url $Url -Message 'No inner URLs extracted from Loop app payload'
        return $null
    }

    Write-LogJsonl -Level 'INFO' -Event 'loop_app_decoded' -Message "Extracted $($result.candidateUrls.Count) candidate URL(s) and itemId=$($result.itemId)"
    return $result
}

function ConvertTo-ShareId {
    [CmdletBinding()]
    param([string]$Url)
    $bytes  = [System.Text.Encoding]::UTF8.GetBytes($Url)
    $b64    = [Convert]::ToBase64String($bytes)
    $b64url = $b64.TrimEnd('=').Replace('+', '-').Replace('/', '_')
    return "u!$b64url"
}

function Resolve-ViaSharesApi {
    [CmdletBinding()]
    param([string]$Url)
    $shareId = ConvertTo-ShareId -Url $Url
    $uri = "$($script:GRAPH_BASE)/shares/$shareId/driveItem"
    Write-LogJsonl -Level 'INFO' -Event 'resolve_shares' -Url $uri -Message "Trying Shares API with shareId"
    try {
        $item = Invoke-GraphRequest -Uri "${uri}?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference"
        return @{ success = $true; driveItem = $item; method = 'shares_api' }
    } catch {
        $ex = $_.Exception
        $code = 0
        if ($ex.PSObject.Properties['Response'] -and $ex.Response) { $code = [int]$ex.Response.StatusCode }
        Write-LogJsonl -Level 'WARN' -Event 'resolve_shares_fail' -Url $uri -StatusCode $code -Message $ex.Message
        return @{ success = $false; method = 'shares_api'; error = $ex.Message }
    }
}

function Extract-UrlTokens {
    [CmdletBinding()]
    param([string]$Url)
    $tokens = @()
    $cleanUrl = $Url -replace '[&?]nav=[^&]*', ''

    $guidPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
    $matches_ = [regex]::Matches($cleanUrl, $guidPattern)
    foreach ($m in $matches_) { $tokens += $m.Value }

    if ($cleanUrl -match '(CSP_[0-9a-fA-F\-]+)') { $tokens += $Matches[1] }

    try {
        $parsed = [uri]$cleanUrl
        $segments = $parsed.AbsolutePath -split '/' | Where-Object { $_ -ne '' }
        if ($segments.Count -gt 0) {
            $last = [System.Uri]::UnescapeDataString($segments[-1])
            if ($last -and $last.Length -gt 3 -and $last -notmatch '^\:') {
                $tokens += $last
            }
        }
    } catch {}
    return @($tokens | Select-Object -Unique)
}

function Resolve-ViaSearch {
    [CmdletBinding()]
    param([string]$Url)
    $tokens = @(Extract-UrlTokens -Url $Url)
    if ($tokens.Count -eq 0) {
        return @{ success = $false; method = 'search'; error = 'No tokens extracted from URL' }
    }

    $tokenClauses = ($tokens | ForEach-Object { "`"$_`"" }) -join ' OR '
    $queryString = "filetype:loop AND ($tokenClauses)"

    Write-LogJsonl -Level 'INFO' -Event 'resolve_search' -Url $Url -Message "Search query: $queryString"

    $searchReq = @{
        entityTypes = @('listItem')
        query       = @{ queryString = $queryString }
        from        = 0
        size        = 50
    }
    $region = $env:SEARCH_REGION
    if ($region) { $searchReq['region'] = $region }
    $searchBody = @{ requests = @($searchReq) }

    try {
        $result = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/search/query" -Method POST -Body $searchBody
    } catch {
        $errMsg = $_.Exception.Message
        if (-not $region -and $errMsg -match 'Region is required') {
            $errMsg += ' (Hint: set SEARCH_REGION env var, e.g. SEARCH_REGION=NAM)'
        }
        Write-LogJsonl -Level 'WARN' -Event 'resolve_search_fail' -Url $Url -Message $errMsg
        return @{ success = $false; method = 'search'; error = $errMsg }
    }

    $hits = [System.Collections.Generic.List[hashtable]]::new()
    $valueItems = if ($result -and $result.PSObject.Properties['value']) { @($result.value) } else { @() }
    foreach ($response in $valueItems) {
        if (-not $response -or -not $response.PSObject.Properties['hitsContainers']) { continue }
        foreach ($hitContainer in @($response.hitsContainers)) {
            if (-not $hitContainer -or -not $hitContainer.PSObject.Properties['hits']) { continue }
            foreach ($hit in @($hitContainer.hits)) {
                if (-not $hit -or -not $hit.PSObject.Properties['resource']) { continue }
                $resource = $hit.resource
                $confidence = 0
                foreach ($t in $tokens) {
                    if ($resource.PSObject.Properties['name'] -and $resource.name.Contains($t)) { $confidence += 40 }
                    if ($resource.PSObject.Properties['webUrl'] -and $resource.webUrl.Contains($t)) { $confidence += 30 }
                    if ($resource.PSObject.Properties['id'] -and $resource.id -eq $t) { $confidence += 50 }
                }
                $confidence = [math]::Min($confidence, 100)
                $hits.Add(@{
                    driveItem  = $resource
                    confidence = $confidence
                    hitId      = if ($hit.PSObject.Properties['hitId']) { $hit.hitId } else { '' }
                })
            }
        }
    }

    if ($hits.Count -eq 0) {
        return @{ success = $false; method = 'search'; error = 'No search results matched' }
    }

    $hits = @($hits | Sort-Object { -$_.confidence })
    $best = $hits[0]
    $ambiguous = $hits.Count -gt 1 -and $best.confidence -lt 80

    if ($ambiguous) {
        Write-LogJsonl -Level 'WARN' -Event 'resolve_ambiguous' -Url $Url `
            -Message "Ambiguous: $($hits.Count) candidates; best confidence=$($best.confidence)"
    }

    $driveId = $null
    $itemId  = $null
    $di = $best.driveItem
    if ($di.PSObject.Properties['parentReference'] -and $di.parentReference -and $di.parentReference.PSObject.Properties['driveId']) {
        $driveId = $di.parentReference.driveId
    }
    if ($di.PSObject.Properties['id']) { $itemId = $di.id }
    if (-not $driveId -and $di.PSObject.Properties['driveId']) { $driveId = $di.driveId }

    $fullItem = $best.driveItem
    if ($driveId -and $itemId) {
        try {
            $fullItem = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$driveId/items/${itemId}?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference"
        } catch {
            Write-LogJsonl -Level 'WARN' -Event 'resolve_metadata_fail' -DriveId $driveId -ItemId $itemId -Message $_.Exception.Message
        }
    }

    return @{
        success    = $true
        driveItem  = $fullItem
        method     = 'search'
        candidates = $hits
        ambiguous  = $ambiguous
    }
}

function Resolve-LoopUrl {
    [CmdletBinding()]
    param([string]$Url)

    Write-LogJsonl -Level 'INFO' -Event 'resolve_start' -Url $Url -Message 'Beginning URL resolution'

    # 0) Pre-process Loop app URLs
    $loopAppData = Expand-LoopAppUrl -Url $Url
    $urlsToTry = [System.Collections.Generic.List[string]]::new()

    if ($loopAppData) {
        foreach ($u in $loopAppData['candidateUrls']) { $urlsToTry.Add($u) }
        $urlsToTry.Add($Url)
    } else {
        $urlsToTry.Add($Url)
    }

    # 1) Try Shares API with each candidate URL
    foreach ($candidateUrl in $urlsToTry) {
        if ($candidateUrl -match 'loop\.(cloud\.)?microsoft(\.com)?(/|$)') {
            Write-LogJsonl -Level 'DEBUG' -Event 'resolve_skip_shares' -Url $candidateUrl -Message 'Skipping Shares API for Loop app URL'
            continue
        }
        if ($candidateUrl -match '/contentstorage/' -and $candidateUrl -notmatch '/:[a-z]+:/') {
            Write-LogJsonl -Level 'DEBUG' -Event 'resolve_skip_shares' -Url $candidateUrl -Message 'Skipping Shares API for contentstorage URL'
            continue
        }
        $result = Resolve-ViaSharesApi -Url $candidateUrl
        if ($result['success']) {
            Write-LogJsonl -Level 'INFO' -Event 'resolve_success' -Url $candidateUrl -Message "Resolved via Shares API"
            if ($loopAppData) { $result['loopAppData'] = $loopAppData }
            return $result
        }
    }

    # 1b) If no loopAppData, try to extract nav params from the raw URL
    if (-not $loopAppData -and $Url -match 'nav=([A-Za-z0-9_\-+/=%]+)') {
        $navB64 = [System.Uri]::UnescapeDataString($Matches[1])
        $navPad = (4 - ($navB64.Length % 4)) % 4
        $navB64 = $navB64 + ('=' * $navPad)
        try {
            $navDecoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($navB64))
            $loopAppData = @{
                candidateUrls   = [System.Collections.Generic.List[string]]::new()
                itemId          = $null
                workspaceUrl    = $null
                pageUrl         = $Url
                driveIdentifier = $null
                folderItemId    = $null
                pageItemId      = $null
            }
            if ($navDecoded -match 'd=([^&]+)') {
                $loopAppData['driveIdentifier'] = $Matches[1]
                Write-LogJsonl -Level 'DEBUG' -Event 'url_nav_drive_id' -Message "Nav drive identifier: $($Matches[1])"
            }
            if ($navDecoded -match 'f=([^&]+)') {
                $loopAppData['pageItemId'] = $Matches[1]
                Write-LogJsonl -Level 'DEBUG' -Event 'url_nav_item_id' -Message "Nav item ID: $($Matches[1])"
            }
        } catch {
            Write-LogJsonl -Level 'DEBUG' -Event 'url_nav_decode_fail' -Message $_.Exception.Message
        }
    }

    # 2) Try Direct Drive Access using decoded nav parameters
    if ($loopAppData -and $loopAppData['driveIdentifier']) {
        $navDriveId = $loopAppData['driveIdentifier']
        $navItemIds = [System.Collections.Generic.List[string]]::new()
        if ($loopAppData['pageItemId'])  { $navItemIds.Add($loopAppData['pageItemId']) }
        if ($loopAppData['folderItemId'] -and $loopAppData['folderItemId'] -ne $loopAppData['pageItemId']) {
            $navItemIds.Add($loopAppData['folderItemId'])
        }

        foreach ($navItemId in $navItemIds) {
            Write-LogJsonl -Level 'INFO' -Event 'resolve_direct_drive' -DriveId $navDriveId -ItemId $navItemId `
                -Message "Trying direct drive access: drives/$navDriveId/items/$navItemId"
            try {
                $item = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$navDriveId/items/${navItemId}?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference" `
                    -DriveId $navDriveId -ItemId $navItemId
                Write-LogJsonl -Level 'INFO' -Event 'resolve_success' -DriveId $navDriveId -ItemId $navItemId `
                    -Message "Resolved via direct drive access: $(Get-SafeProp $item 'name')"
                return @{
                    success     = $true
                    driveItem   = $item
                    method      = 'direct_drive'
                    loopAppData = $loopAppData
                }
            } catch {
                $ex = $_.Exception
                $code = 0
                if ($ex.PSObject.Properties['Response'] -and $ex.Response) { $code = [int]$ex.Response.StatusCode }
                Write-LogJsonl -Level 'WARN' -Event 'resolve_direct_drive_fail' -DriveId $navDriveId -ItemId $navItemId `
                    -StatusCode $code -Message $ex.Message
            }
        }

        # Try drive root
        Write-LogJsonl -Level 'INFO' -Event 'resolve_direct_drive_root' -DriveId $navDriveId -Message "Trying drive root"
        try {
            $driveRoot = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$navDriveId/root?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference" -DriveId $navDriveId
            Write-LogJsonl -Level 'INFO' -Event 'resolve_success' -DriveId $navDriveId `
                -Message "Resolved drive root: $(Get-SafeProp $driveRoot 'name')"
            return @{
                success     = $true
                driveItem   = $driveRoot
                method      = 'direct_drive_root'
                loopAppData = $loopAppData
            }
        } catch {
            $ex = $_.Exception
            $code = 0
            if ($ex.PSObject.Properties['Response'] -and $ex.Response) { $code = [int]$ex.Response.StatusCode }
            Write-LogJsonl -Level 'WARN' -Event 'resolve_direct_drive_root_fail' -DriveId $navDriveId `
                -StatusCode $code -Message $ex.Message
        }
    }

    # 3) Try Storage Containers API (SPE containers — CSP_*)
    if ($loopAppData) {
        $containerIds = [System.Collections.Generic.List[string]]::new()
        if ($loopAppData['driveIdentifier']) { $containerIds.Add($loopAppData['driveIdentifier']) }
        $pageUrl = $loopAppData['pageUrl']
        if ($pageUrl -and $pageUrl -match 'CSP_([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
            $rawGuid = $Matches[1]
            if ($rawGuid -notin $containerIds) { $containerIds.Add($rawGuid) }
        }
        foreach ($containerId in $containerIds) {
            Write-LogJsonl -Level 'INFO' -Event 'resolve_container_api' -Message "Trying Storage Containers API with containerId=$containerId"
            try {
                $containerDrive = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/storage/fileStorage/containers/$containerId/drive?`$select=id,name,webUrl,description"
                $containerDriveId = Get-SafeProp $containerDrive 'id'
                if ($containerDriveId) {
                    Write-LogJsonl -Level 'INFO' -Event 'resolve_container_drive' -DriveId $containerDriveId `
                        -Message "Got drive from container: $(Get-SafeProp $containerDrive 'name')"

                    $pageItemId = $loopAppData['pageItemId']
                    if ($pageItemId) {
                        try {
                            $item = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$containerDriveId/items/${pageItemId}?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference" `
                                -DriveId $containerDriveId -ItemId $pageItemId
                            Write-LogJsonl -Level 'INFO' -Event 'resolve_success' -DriveId $containerDriveId -ItemId $pageItemId `
                                -Message "Resolved via container API + page item: $(Get-SafeProp $item 'name')"
                            return @{
                                success     = $true
                                driveItem   = $item
                                method      = 'container_api'
                                loopAppData = $loopAppData
                            }
                        } catch {
                            Write-LogJsonl -Level 'WARN' -Event 'resolve_container_item_fail' -DriveId $containerDriveId -ItemId $pageItemId -Message $_.Exception.Message
                        }
                    }

                    # Fall back to drive root
                    try {
                        $driveRoot = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$containerDriveId/root?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference" -DriveId $containerDriveId
                        Write-LogJsonl -Level 'INFO' -Event 'resolve_success' -DriveId $containerDriveId `
                            -Message "Resolved container drive root: $(Get-SafeProp $driveRoot 'name')"
                        return @{
                            success     = $true
                            driveItem   = $driveRoot
                            method      = 'container_api_root'
                            loopAppData = $loopAppData
                        }
                    } catch {
                        Write-LogJsonl -Level 'WARN' -Event 'resolve_container_root_fail' -DriveId $containerDriveId -Message $_.Exception.Message
                    }
                }
            } catch {
                $ex = $_.Exception
                $code = 0
                if ($ex.PSObject.Properties['Response'] -and $ex.Response) { $code = [int]$ex.Response.StatusCode }
                Write-LogJsonl -Level 'WARN' -Event 'resolve_container_api_fail' -StatusCode $code -Message "containerId=$containerId $($ex.Message)"
            }
        }
    }

    # 4) Try Microsoft Search
    $searchUrl = $Url
    if ($loopAppData -and $loopAppData['candidateUrls'].Count -gt 0) {
        $searchUrl = $loopAppData['candidateUrls'][0]
    }
    $result = Resolve-ViaSearch -Url $searchUrl
    if ($result['success']) {
        Write-LogJsonl -Level 'INFO' -Event 'resolve_success' -Url $searchUrl -Message "Resolved via Search API (ambiguous=$($result['ambiguous']))"
        if ($loopAppData) { $result['loopAppData'] = $loopAppData }
        return $result
    }

    # 5) Search by item ID
    if ($loopAppData -and $loopAppData['itemId']) {
        Write-LogJsonl -Level 'INFO' -Event 'resolve_search_by_id' -ItemId $loopAppData['itemId'] -Message "Trying search by item GUID"
        $idSearchReq = @{
            entityTypes = @('listItem')
            query       = @{ queryString = "filetype:loop AND `"$($loopAppData['itemId'])`"" }
            from        = 0
            size        = 10
        }
        if ($env:SEARCH_REGION) { $idSearchReq['region'] = $env:SEARCH_REGION }
        $idSearchBody = @{ requests = @($idSearchReq) }
        try {
            $searchResult = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/search/query" -Method POST -Body $idSearchBody
            $srValues = if ($searchResult -and $searchResult.PSObject.Properties['value']) { @($searchResult.value) } else { @() }
            foreach ($resp in $srValues) {
                if (-not $resp -or -not $resp.PSObject.Properties['hitsContainers']) { continue }
                foreach ($hc in @($resp.hitsContainers)) {
                    if (-not $hc -or -not $hc.PSObject.Properties['hits']) { continue }
                    $hcHits = @($hc.hits)
                    if ($hcHits.Count -gt 0) {
                        $firstHit = $hcHits[0]
                        if (-not $firstHit -or -not $firstHit.PSObject.Properties['resource']) { continue }
                        $resource = $firstHit.resource
                        $driveId = $null
                        if ($resource.PSObject.Properties['parentReference'] -and $resource.parentReference -and $resource.parentReference.PSObject.Properties['driveId']) {
                            $driveId = $resource.parentReference.driveId
                        }
                        if (-not $driveId -and $resource.PSObject.Properties['driveId']) { $driveId = $resource.driveId }
                        $rItemId = if ($resource.PSObject.Properties['id']) { $resource.id } else { $null }
                        $fullItem = $resource
                        if ($driveId -and $rItemId) {
                            try {
                                $fullItem = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$driveId/items/${rItemId}?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference"
                            } catch {}
                        }
                        Write-LogJsonl -Level 'INFO' -Event 'resolve_success' -ItemId $loopAppData['itemId'] -Message "Resolved via item GUID search"
                        return @{
                            success     = $true
                            driveItem   = $fullItem
                            method      = 'search_by_id'
                            loopAppData = $loopAppData
                        }
                    }
                }
            }
        } catch {
            Write-LogJsonl -Level 'WARN' -Event 'resolve_search_by_id_fail' -Message $_.Exception.Message
        }
    }

    # 6) Fail
    Write-LogJsonl -Level 'ERROR' -Event 'resolve_fail' -Url $Url -Message 'Could not resolve URL via any method'

    $errorMsg = "Could not resolve the provided URL. Tried Shares API, Direct Drive Access, Storage Containers API, and Microsoft Search."
    if ($loopAppData) {
        $containerId = $loopAppData['pageUrl'] -replace '.*CSP_([0-9a-fA-F-]+).*','$1'
        $errorMsg += " The Loop app URL was decoded successfully — inner page URL: $($loopAppData['pageUrl']) | Container ID: $containerId"

        Write-Host ''
        Write-Host '  This is a SharePoint Embedded container (used by Loop).' -ForegroundColor Yellow
        Write-Host '  Your Entra ID app needs these permissions:' -ForegroundColor Yellow
        Write-Host "    1. 'Files.Read.All' (Application) — with admin consent"
        Write-Host "    2. 'Sites.Read.All' (Application) — with admin consent"
        Write-Host "    3. Optionally 'FileStorageContainer.Selected' (Application) — for container-specific access"
        Write-Host ''
        Write-Host '  Common fixes:' -ForegroundColor Yellow
        Write-Host '    - Verify admin consent is granted (green checkmarks in Azure Portal > API permissions)'
        Write-Host '    - Wait 5-10 minutes after granting consent for propagation'
        Write-Host '    - Check that the app is in the SAME tenant as the Loop workspace'
        Write-Host ''
    } else {
        $errorMsg += " Suggestions: (1) Verify the URL is correct. (2) Ensure the app has Files.Read.All and Sites.Read.All permissions. (3) Try a direct SharePoint/OneDrive sharing link. (4) Set SEARCH_REGION env var (e.g. SEARCH_REGION=NAM)."
    }
    return @{ success = $false; method = 'none'; error = $errorMsg }
}

# ─────────────────────────────────────────────────────────────────────────────
# Enumerate Loop items
# ─────────────────────────────────────────────────────────────────────────────
function Get-LoopItemsRecursive {
    [CmdletBinding()]
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$SinceDate = ''
    )
    $items = [System.Collections.Generic.List[object]]::new()
    $uri = "$($script:GRAPH_BASE)/drives/$DriveId/items/$ItemId/children?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference&`$top=200"

    while ($uri) {
        Write-LogJsonl -Level 'DEBUG' -Event 'enumerate' -DriveId $DriveId -ItemId $ItemId -Url $uri -Message 'Listing children'

        $response = $null
        try {
            $response = Invoke-GraphRequest -Uri $uri -DriveId $DriveId -ItemId $ItemId
        } catch {
            Write-LogJsonl -Level 'WARN' -Event 'enumerate_page_fail' -DriveId $DriveId -ItemId $ItemId `
                -Url $uri -Message "Pagination failed (collected $($items.Count) items so far): $($_.Exception.Message)"
            break
        }

        $children = @()
        if (Test-SafeProp $response 'value') { $children = @($response.value) }

        foreach ($child in $children) {
            $childName = Get-SafeProp $child 'name'
            $childId   = Get-SafeProp $child 'id'

            # Normalize parentReference.driveId
            $childParent = Get-SafeProp $child 'parentReference'
            if ($childParent) {
                $childDriveId = Get-SafeProp $childParent 'driveId'
                if (-not $childDriveId -or $childDriveId -ne $DriveId) {
                    $childParent.driveId = $DriveId
                }
            } else {
                $child | Add-Member -NotePropertyName 'parentReference' -NotePropertyValue ([PSCustomObject]@{ driveId = $DriveId }) -Force
            }

            if (Test-SafeProp $child 'folder') {
                Write-LogJsonl -Level 'DEBUG' -Event 'enumerate_folder' -DriveId $DriveId -ItemId $childId `
                    -Message "Recursing into folder: $childName"
                try {
                    $subItems = Get-LoopItemsRecursive -DriveId $DriveId -ItemId $childId -SinceDate $SinceDate
                    foreach ($si in $subItems) { $items.Add($si) }
                } catch {
                    Write-LogJsonl -Level 'WARN' -Event 'enumerate_folder_fail' -DriveId $DriveId -ItemId $childId `
                        -Message "Failed to enumerate folder '$childName': $($_.Exception.Message)"
                }
            }
            elseif ($childName -and $childName.EndsWith('.loop', [System.StringComparison]::OrdinalIgnoreCase)) {
                if ($childName -eq 'Untitled.loop') {
                    Write-LogJsonl -Level 'DEBUG' -Event 'enumerate_skip_root' -DriveId $DriveId -ItemId $childId `
                        -Message "Skipping workspace root container: $childName"
                    continue
                }
                $childModified = Get-SafeProp $child 'lastModifiedDateTime'
                if ($SinceDate -ne '' -and $childModified) {
                    $modified = [datetime]::Parse($childModified)
                    $since    = [datetime]::Parse($SinceDate)
                    if ($modified -lt $since) { continue }
                }
                Write-LogJsonl -Level 'DEBUG' -Event 'enumerate_loop_file' -DriveId $DriveId -ItemId $childId `
                    -Message "Found .loop file: $childName"
                $items.Add($child)
            }
            else {
                Write-LogJsonl -Level 'DEBUG' -Event 'enumerate_skip' -DriveId $DriveId -ItemId $childId `
                    -Message "Skipping non-.loop file: $childName"
            }
        }

        $uri = $null
        if (Test-SafeProp $response '@odata.nextLink') {
            $uri = Get-SafeProp $response '@odata.nextLink'
        }
    }

    return $items
}

# ─────────────────────────────────────────────────────────────────────────────
# Export Pipeline — Python / Pandoc
# ─────────────────────────────────────────────────────────────────────────────
function Find-PythonExe {
    [CmdletBinding()]
    param()

    $toolsDir = Join-Path $script:ProjectRoot 'tools'
    $venvDir  = Join-Path $toolsDir '.venv'
    $venvPy   = if ($IsWindows) {
        Join-Path $venvDir 'Scripts' 'python.exe'
    } else {
        Join-Path $venvDir 'bin' 'python'
    }

    if (Test-Path $venvPy) {
        # Venv exists — check if deps need updating (requirements.txt changed)
        $stampFile = Join-Path $venvDir '.stamp'
        $reqFile   = Join-Path $toolsDir 'requirements.txt'
        $needsUpdate = $true
        if ((Test-Path $stampFile) -and (Test-Path $reqFile)) {
            $stampHash = (Get-Content $stampFile -Raw).Trim()
            $reqHash   = Get-FileSHA256 -Path $reqFile
            $needsUpdate = ($stampHash -ne $reqHash)
        }
        if ($needsUpdate) {
            Initialize-PythonVenv -Auto | Out-Null
        }
        Write-LogJsonl -Level 'DEBUG' -Event 'python_venv' -Message "Using venv python: $venvPy"
        return $venvPy
    }

    # No venv yet — try to auto-bootstrap one
    $created = Initialize-PythonVenv -Auto
    if ($created -and (Test-Path $venvPy)) {
        Write-LogJsonl -Level 'DEBUG' -Event 'python_venv' -Message "Using newly created venv python: $venvPy"
        return $venvPy
    }

    # Fallback: bare system python (deps may not be installed)
    if ($env:PYTHON -and (Get-Command $env:PYTHON -ErrorAction SilentlyContinue)) {
        return $env:PYTHON
    }

    $candidates = if ($IsWindows) { @('python', 'python3') } else { @('python3', 'python') }
    foreach ($candidate in $candidates) {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($cmd) { return $candidate }
    }
    return $null
}

function Initialize-PythonVenv {
    <#
    .SYNOPSIS
        Create (or recreate) the Python venv in tools/.venv and install dependencies.
    .PARAMETER Auto
        When set, the function runs silently as part of auto-bootstrap during backup.
        If Python is not on PATH it returns $false instead of throwing.
    .PARAMETER Force
        When set, delete any existing venv and recreate from scratch.
    #>
    [CmdletBinding()]
    param(
        [switch]$Auto,
        [switch]$Force
    )
    $toolsDir = Join-Path $script:ProjectRoot 'tools'
    $venvDir  = Join-Path $toolsDir '.venv'
    $reqFile  = Join-Path $toolsDir 'requirements.txt'
    $stampFile = Join-Path $venvDir '.stamp'

    if (-not (Test-Path $reqFile)) {
        if ($Auto) { return $false }
        throw "requirements.txt not found at $reqFile"
    }

    # If stamp exists and matches current requirements.txt, skip (unless -Force)
    if (-not $Force -and (Test-Path $stampFile)) {
        $stampHash = (Get-Content $stampFile -Raw).Trim()
        $reqHash   = (Get-FileSHA256 -Path $reqFile)
        if ($stampHash -eq $reqHash) {
            if (-not $Auto) { Write-Host 'Python venv is up to date.' -ForegroundColor Green }
            return $true
        }
    }

    # Find a base Python interpreter
    # On Windows, prefer 'python' first — 'python3' often resolves to the
    # Microsoft Store app-execution alias (a stub that opens the Store instead
    # of running Python). On Unix, prefer 'python3' to avoid Python 2.
    $basePython = $null
    if ($env:PYTHON -and (Get-Command $env:PYTHON -ErrorAction SilentlyContinue)) {
        $basePython = $env:PYTHON
    } else {
        $candidates = if ($IsWindows) { @('python', 'python3') } else { @('python3', 'python') }
        foreach ($candidate in $candidates) {
            $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
            if ($cmd) { $basePython = $candidate; break }
        }
    }
    if (-not $basePython) {
        if ($Auto) {
            Write-Host '[auto-venv] Python not found on PATH — skipping venv setup. Install Python 3.8+ for Markdown conversion.' -ForegroundColor Yellow
            return $false
        }
        throw 'No Python interpreter found. Install Python 3.8+ and ensure it is on PATH.'
    }

    # Remove existing venv on Force or if it looks broken
    if ($Force -and (Test-Path $venvDir)) {
        Write-Host 'Removing existing venv...' -ForegroundColor Yellow
        Remove-Item -Recurse -Force $venvDir
    }

    $label = if ($Auto) { '[auto-venv]' } else { '' }

    if (-not (Test-Path $venvDir)) {
        Write-Host "$label Creating Python venv in $venvDir ..." -ForegroundColor Cyan
        & $basePython -m venv $venvDir
        if ($LASTEXITCODE -ne 0) {
            if ($Auto) {
                Write-Host "$label Failed to create venv (exit $LASTEXITCODE). Markdown conversion may not work." -ForegroundColor Yellow
                return $false
            }
            throw "Failed to create venv (exit code $LASTEXITCODE)"
        }
    }

    $venvPip = if ($IsWindows) { Join-Path $venvDir 'Scripts' 'pip' } else { Join-Path $venvDir 'bin' 'pip' }

    Write-Host "$label Installing dependencies..." -ForegroundColor Cyan
    & $venvPip install --upgrade pip 2>&1 | Out-Null
    & $venvPip install -r $reqFile
    if ($LASTEXITCODE -ne 0) {
        if ($Auto) {
            Write-Host "$label pip install failed (exit $LASTEXITCODE). Markdown conversion may not work." -ForegroundColor Yellow
            return $false
        }
        throw "pip install failed (exit code $LASTEXITCODE)"
    }

    # Write stamp with requirements hash so we skip next time
    $reqHash = Get-FileSHA256 -Path $reqFile
    Set-Content -Path $stampFile -Value $reqHash -NoNewline

    Write-Host "$label Python venv ready." -ForegroundColor Green
    Write-LogJsonl -Level 'INFO' -Event 'venv_initialized' -Message "Python venv created/updated in $venvDir"
    return $true
}

function Find-PandocExe {
    [CmdletBinding()]
    param()
    if ($env:PANDOC -and (Get-Command $env:PANDOC -ErrorAction SilentlyContinue)) { return $env:PANDOC }
    $cmd = Get-Command 'pandoc' -ErrorAction SilentlyContinue
    if ($cmd) { return 'pandoc' }
    return $null
}

function Convert-HtmlToMarkdown {
    [CmdletBinding()]
    param(
        [string]$HtmlPath,
        [string]$MdPath
    )
    $toolsDir = Join-Path $script:ProjectRoot 'tools'
    $converterScript = Join-Path $toolsDir 'html_to_md.py'

    if (-not (Test-Path $converterScript)) {
        throw "Python converter not found at $converterScript"
    }

    if ($script:PythonExe) {
        $errFile = Join-Path ([System.IO.Path]::GetTempPath()) "spbackup_py_err_$PID.txt"
        $proc = Start-Process -FilePath $script:PythonExe `
            -ArgumentList @("`"$converterScript`"", '--in', "`"$HtmlPath`"", '--out', "`"$MdPath`"") `
            -Wait -PassThru -NoNewWindow -RedirectStandardError $errFile

        if ($proc.ExitCode -eq 0) {
            if (Test-Path $errFile) { Remove-Item $errFile -Force -ErrorAction SilentlyContinue }
            return $true
        }

        $errMsg = if (Test-Path $errFile) { Get-Content $errFile -Raw } else { 'Unknown error' }
        if (Test-Path $errFile) { Remove-Item $errFile -Force -ErrorAction SilentlyContinue }
        Write-LogJsonl -Level 'WARN' -Event 'python_convert_fail' -Message "Python converter failed (exit $($proc.ExitCode)): $errMsg"
    }

    if ($script:PandocExe) {
        $proc = Start-Process -FilePath $script:PandocExe `
            -ArgumentList @('-f', 'html', '-t', 'markdown', '-o', "`"$MdPath`"", "`"$HtmlPath`"") `
            -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 0) { return $true }
        Write-LogJsonl -Level 'WARN' -Event 'pandoc_convert_fail' -Message "Pandoc failed (exit $($proc.ExitCode))"
    }

    throw "No working HTML-to-Markdown converter available. Install Python deps (pip install -r tools/requirements.txt) or install Pandoc."
}

function Export-LoopItem {
    [CmdletBinding()]
    param(
        [object]$DriveItem,
        [string]$OutDir,
        [bool]$ExportRawLoop,
        [bool]$ExportHtml,
        [bool]$ExportMd,
        [hashtable]$State,
        [string]$StatePath,
        [bool]$DryRun
    )

    $parentRef = Get-SafeProp $DriveItem 'parentReference'
    $driveId   = Get-SafeProp $parentRef 'driveId'
    $itemId    = Get-SafeProp $DriveItem 'id'
    $diName    = Get-SafeProp $DriveItem 'name'
    $stateKey  = "${driveId}:${itemId}"
    $safeName  = New-SafeName -Name ($diName ?? 'unnamed') -Suffix ($itemId ?? 'unknown') -StripExtension
    $itemDir   = Join-Path $OutDir 'items' $safeName

    $result = [ordered]@{
        itemId               = $itemId
        driveId              = $driveId
        name                 = $diName
        webUrl               = Get-SafeProp $DriveItem 'webUrl'
        lastModifiedDateTime = Get-SafeProp $DriveItem 'lastModifiedDateTime'
        eTag                 = Get-SafeProp $DriveItem 'eTag'
        outputDir            = $safeName
        exported             = @()
        skipped              = $false
        error                = $null
    }

    # Incremental check
    $existing = $null
    if ($State.ContainsKey($stateKey)) { $existing = $State[$stateKey] }

    if ($existing) {
        $diETag     = Get-SafeProp $DriveItem 'eTag'
        $storedETag = $existing['eTag']

        $filesExist = $true
        if ($ExportHtml    -and -not (Test-Path (Join-Path $itemDir 'page.html')))  { $filesExist = $false }
        if ($ExportMd      -and -not (Test-Path (Join-Path $itemDir 'page.md')))    { $filesExist = $false }
        if ($ExportRawLoop -and -not (Test-Path (Join-Path $itemDir 'page.loop')))  { $filesExist = $false }

        if ($filesExist) {
            $etagMatch = $false
            if ($storedETag -and $diETag) {
                $etagMatch = ($storedETag -eq $diETag)
            } elseif (-not $storedETag -or -not $diETag) {
                $hashesMatch = $true
                if ($existing.ContainsKey('sha256_html') -and (Test-Path (Join-Path $itemDir 'page.html'))) {
                    $currentHash = Get-FileSHA256 -Path (Join-Path $itemDir 'page.html')
                    if ($currentHash -ne $existing['sha256_html']) { $hashesMatch = $false }
                }
                if ($existing.ContainsKey('sha256_md') -and (Test-Path (Join-Path $itemDir 'page.md'))) {
                    $currentHash = Get-FileSHA256 -Path (Join-Path $itemDir 'page.md')
                    if ($currentHash -ne $existing['sha256_md']) { $hashesMatch = $false }
                }
                $etagMatch = $hashesMatch
            }

            if ($etagMatch) {
                Write-LogJsonl -Level 'INFO' -Event 'skip_unchanged' -DriveId $driveId -ItemId $itemId -Message "Unchanged: $diName"
                $result.skipped = $true
                return $result
            }
        }
    }

    if ($DryRun) {
        Write-LogJsonl -Level 'INFO' -Event 'dry_run_would_export' -DriveId $driveId -ItemId $itemId -Message "Would export: $diName"
        return $result
    }

    if (-not (Test-Path $itemDir)) {
        New-Item -ItemType Directory -Path $itemDir -Force | Out-Null
    }

    Write-LogJsonl -Level 'INFO' -Event 'export_start' -DriveId $driveId -ItemId $itemId -Message "Exporting: $diName"

    $metaETag = $null

    # 1) Save metadata
    try {
        $metaPath = Join-Path $itemDir 'meta.json'
        $meta = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$driveId/items/${itemId}?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference" `
            -DriveId $driveId -ItemId $itemId -ExtraRetryStatusCodes @(404)
        $metaJson = $meta | ConvertTo-Json -Depth 10
        Write-AtomicFile -Path $metaPath -Content $metaJson
        $result.exported += 'meta.json'

        $metaETag = Get-SafeProp $meta 'eTag'
        if ($metaETag) { $result['eTag'] = $metaETag }

        $metaParent = Get-SafeProp $meta 'parentReference'
        if ($metaParent) {
            $metaDriveId = Get-SafeProp $metaParent 'driveId'
            if ($metaDriveId -and $metaDriveId -ne $driveId) {
                Write-LogJsonl -Level 'WARN' -Event 'export_driveId_corrected' -DriveId $driveId -ItemId $itemId `
                    -Message "Metadata returned driveId '$metaDriveId' (was '$driveId'); using corrected value"
                $driveId = $metaDriveId
                $result['driveId'] = $metaDriveId
            }
        }
    } catch {
        $conciseErr = Get-ConciseErrorMessage $_.Exception.Message
        Write-LogJsonl -Level 'ERROR' -Event 'export_meta_fail' -DriveId $driveId -ItemId $itemId -Message $conciseErr
        $result.error = "Metadata fetch failed: $conciseErr"
        return $result
    }

    # 2) Raw .loop download
    if ($ExportRawLoop) {
        try {
            $loopPath = Join-Path $itemDir 'page.loop'
            $tmpLoop  = "$loopPath.tmp.$PID"
            Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$driveId/items/$itemId/content" -Raw -OutFile $tmpLoop -DriveId $driveId -ItemId $itemId | Out-Null
            Move-FileReliable -Source $tmpLoop -Destination $loopPath
            $result.exported += 'page.loop'
            Write-LogJsonl -Level 'INFO' -Event 'export_loop' -DriveId $driveId -ItemId $itemId -Message 'Raw .loop downloaded'
        } catch {
            $conciseErr = Get-ConciseErrorMessage $_.Exception.Message
            Write-LogJsonl -Level 'ERROR' -Event 'export_loop_fail' -DriveId $driveId -ItemId $itemId -Message $conciseErr
            if (Test-Path -LiteralPath $tmpLoop) { Remove-Item -LiteralPath $tmpLoop -Force -ErrorAction SilentlyContinue }
            $result.error = "Raw loop download failed: $conciseErr"
        }
    }

    # 3) HTML export
    $htmlPath = Join-Path $itemDir 'page.html'
    if ($ExportHtml -or $ExportMd) {
        try {
            $tmpHtml = "$htmlPath.tmp.$PID"
            Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$driveId/items/$itemId/content?format=html" `
                -Raw -OutFile $tmpHtml -DriveId $driveId -ItemId $itemId `
                -ExtraRetryStatusCodes @(404, 403) | Out-Null
            Move-FileReliable -Source $tmpHtml -Destination $htmlPath
            $result.exported += 'page.html'
            Write-LogJsonl -Level 'INFO' -Event 'export_html' -DriveId $driveId -ItemId $itemId -Message 'HTML exported'
        } catch {
            $conciseErr = Get-ConciseErrorMessage $_.Exception.Message
            Write-LogJsonl -Level 'ERROR' -Event 'export_html_fail' -DriveId $driveId -ItemId $itemId -Message $conciseErr
            if (Test-Path -LiteralPath $tmpHtml) { Remove-Item -LiteralPath $tmpHtml -Force -ErrorAction SilentlyContinue }
            $result.error = "HTML export failed: $conciseErr"
            $ExportMd = $false
        }
    }

    # 4) Markdown conversion
    if ($ExportMd -and (Test-Path -LiteralPath $htmlPath)) {
        try {
            $mdPath = Join-Path $itemDir 'page.md'
            Convert-HtmlToMarkdown -HtmlPath $htmlPath -MdPath $mdPath | Out-Null
            $result.exported += 'page.md'
            Write-LogJsonl -Level 'INFO' -Event 'export_md' -DriveId $driveId -ItemId $itemId -Message 'Markdown converted'
        } catch {
            $conciseErr = Get-ConciseErrorMessage $_.Exception.Message
            Write-LogJsonl -Level 'ERROR' -Event 'export_md_fail' -DriveId $driveId -ItemId $itemId -Message $conciseErr
            $result.error = "Markdown conversion failed: $conciseErr"
        }
    }

    # Remove HTML if not explicitly requested (was only needed for MD)
    if ($ExportMd -and -not $ExportHtml -and (Test-Path -LiteralPath $htmlPath)) {
        Remove-Item -LiteralPath $htmlPath -Force -ErrorAction SilentlyContinue
        $result.exported = $result.exported | Where-Object { $_ -ne 'page.html' }
    }

    # Update state
    $stateETag = if ($metaETag) { $metaETag } else { Get-SafeProp $DriveItem 'eTag' }
    $stateEntry = @{
        eTag                 = $stateETag
        lastModifiedDateTime = Get-SafeProp $DriveItem 'lastModifiedDateTime'
        lastExportTime       = (Get-Date).ToUniversalTime().ToString('o')
        outputDir            = $safeName
    }
    if (Test-Path (Join-Path $itemDir 'page.html')) {
        $stateEntry['sha256_html'] = Get-FileSHA256 -Path (Join-Path $itemDir 'page.html')
    }
    if (Test-Path (Join-Path $itemDir 'page.md')) {
        $stateEntry['sha256_md'] = Get-FileSHA256 -Path (Join-Path $itemDir 'page.md')
    }
    $State[$stateKey] = $stateEntry
    Save-State -Path $StatePath -State $State

    Write-LogJsonl -Level 'INFO' -Event 'export_complete' -DriveId $driveId -ItemId $itemId -Message "Done: $diName"
    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# Commands
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-LoopResolveCommand {
    [CmdletBinding()]
    param([hashtable]$Opts)

    $url = $Opts['url']
    if ([string]::IsNullOrWhiteSpace($url)) { throw '--url is required for the resolve command.' }

    $resolution = Resolve-LoopUrl -Url $url

    $output = [ordered]@{ success = $resolution['success']; method = $resolution['method'] }

    if ($resolution['success']) {
        $di = $resolution['driveItem']
        $diParentRef = Get-SafeProp $di 'parentReference'
        $output['driveItem'] = [ordered]@{
            id                   = Get-SafeProp $di 'id'
            name                 = Get-SafeProp $di 'name'
            webUrl               = Get-SafeProp $di 'webUrl'
            lastModifiedDateTime = Get-SafeProp $di 'lastModifiedDateTime'
            isFolder             = [bool](Get-SafeProp $di 'folder')
        }
        if ($diParentRef) {
            $output['driveItem']['driveId'] = Get-SafeProp $diParentRef 'driveId'
            $output['driveItem']['siteId']  = Get-SafeProp $diParentRef 'siteId'
        }
        if ($resolution['candidates']) {
            $output['candidates'] = $resolution['candidates'] | ForEach-Object {
                $cdi = $_.driveItem
                [ordered]@{
                    id         = Get-SafeProp $cdi 'id'
                    name       = Get-SafeProp $cdi 'name'
                    webUrl     = Get-SafeProp $cdi 'webUrl'
                    confidence = $_.confidence
                }
            }
        }
        if ($resolution['ambiguous']) { $output['ambiguous'] = $true }
        if ($resolution['loopAppData']) {
            $lad = $resolution['loopAppData']
            $output['loopAppData'] = [ordered]@{
                pageUrl         = $lad['pageUrl']
                workspaceUrl    = $lad['workspaceUrl']
                itemGuid        = $lad['itemId']
                driveIdentifier = $lad['driveIdentifier']
                folderItemId    = $lad['folderItemId']
                candidateUrls   = @($lad['candidateUrls'])
            }
        }
    } else {
        $output['error'] = $resolution['error']
    }

    $json = $output | ConvertTo-Json -Depth 10
    Write-Output $json
}

function Invoke-LoopBackupCommand {
    [CmdletBinding()]
    param([hashtable]$Opts)

    $url         = $Opts['url']
    $outDir      = $Opts['out']
    $mode        = $Opts['mode']        ?? 'auto'
    $exportRaw   = [bool]$Opts['raw-loop']
    $exportHtml  = if ($null -eq $Opts['html'] -and $null -eq $Opts['md']) { $true } else { [bool]$Opts['html'] }
    $exportMd    = if ($null -eq $Opts['html'] -and $null -eq $Opts['md']) { $true } else { [bool]$Opts['md'] }
    $concurrency = [int]($Opts['concurrency'] ?? 4)
    $sinceDate   = $Opts['since']       ?? ''
    $statePath   = $Opts['state']       ?? (Join-Path $outDir '.state.json')
    $dryRun      = [bool]$Opts['dry-run']
    $force       = [bool]$Opts['force']

    if ([string]::IsNullOrWhiteSpace($url))    { throw '--url is required for the backup command.' }
    if ([string]::IsNullOrWhiteSpace($outDir)) { throw '--out is required for the backup command.' }

    if (-not $Opts.ContainsKey('html') -and -not $Opts.ContainsKey('md')) {
        $exportHtml = $true
        $exportMd   = $true
    }

    $script:Semaphore = [System.Threading.Semaphore]::new($concurrency, $concurrency)

    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $logsDir = Join-Path $outDir 'logs'
    if (-not (Test-Path $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }
    $script:LogFilePath = Join-Path $logsDir "run-$($script:RunTimestamp).log.jsonl"

    Write-LogJsonl -Level 'INFO' -Event 'backup_start' -Url $url -Message "Loop backup starting: mode=$mode, html=$exportHtml, md=$exportMd, raw=$exportRaw, concurrency=$concurrency"

    # Detect converters
    $script:PythonExe = Find-PythonExe
    $script:PandocExe = Find-PandocExe
    if ($exportMd) {
        if (-not $script:PythonExe -and -not $script:PandocExe) {
            throw "Markdown export requested but no converter found. Install Python 3.8+ (venv auto-setup failed or Python not on PATH) or install Pandoc."
        }
        Write-LogJsonl -Level 'INFO' -Event 'converter_detected' -Message "Python=$($script:PythonExe ?? 'N/A'), Pandoc=$($script:PandocExe ?? 'N/A')"
    }

    # 1) Resolve URL
    $resolution = Resolve-LoopUrl -Url $url
    if (-not $resolution['success']) {
        Write-LogJsonl -Level 'ERROR' -Event 'resolve_failed' -Url $url -Message $resolution['error']
        Write-Error $resolution['error']
        $script:ExitCode = 1
        return
    }

    $rootItem      = $resolution['driveItem']
    $rootParentRef = Get-SafeProp $rootItem 'parentReference'
    $driveId       = Get-SafeProp $rootParentRef 'driveId'
    $itemId        = Get-SafeProp $rootItem 'id'
    $rootName      = Get-SafeProp $rootItem 'name'

    # Prefer loopAppData.driveIdentifier when available
    $loopDriveId = $null
    if ($resolution['loopAppData'] -and $resolution['loopAppData']['driveIdentifier']) {
        $loopDriveId = $resolution['loopAppData']['driveIdentifier']
    }

    if ([string]::IsNullOrWhiteSpace($driveId) -and $loopDriveId) {
        $driveId = $loopDriveId
    } elseif ($loopDriveId -and $driveId -ne $loopDriveId) {
        Write-LogJsonl -Level 'WARN' -Event 'driveId_mismatch' -DriveId $driveId `
            -Message "parentReference.driveId ($driveId) differs from loopAppData.driveIdentifier ($loopDriveId); using loopAppData value"
        $driveId = $loopDriveId
    }

    if ([string]::IsNullOrWhiteSpace($driveId)) {
        Write-Error 'Could not determine the drive ID.'
        $script:ExitCode = 1
        return
    }

    Write-LogJsonl -Level 'INFO' -Event 'resolved' -DriveId $driveId -ItemId $itemId `
        -Message "Resolved: $rootName (folder=$([bool](Get-SafeProp $rootItem 'folder')))"

    # 2) Determine mode
    $isFolder = [bool](Get-SafeProp $rootItem 'folder')
    $effectiveMode = $mode
    if ($effectiveMode -eq 'auto') {
        if ($isFolder) {
            $effectiveMode = 'workspace'
        } elseif ($resolution['loopAppData']) {
            $effectiveMode = 'workspace'
            Write-LogJsonl -Level 'INFO' -Event 'mode_auto_workspace' -Message 'Loop app URL detected; auto-selecting workspace mode'
        } else {
            $effectiveMode = 'page'
        }
    }

    $enumerateFromId = $itemId
    if ($effectiveMode -eq 'workspace' -and -not $isFolder) {
        Write-LogJsonl -Level 'INFO' -Event 'workspace_navigate_root' -DriveId $driveId -ItemId $itemId `
            -Message "Resolved item is a file but workspace mode requested; navigating to drive root"
        try {
            $driveRoot = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/$driveId/root?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference" -DriveId $driveId
            $enumerateFromId = Get-SafeProp $driveRoot 'id'
            Write-LogJsonl -Level 'INFO' -Event 'workspace_root_found' -DriveId $driveId -ItemId $enumerateFromId `
                -Message "Drive root: $(Get-SafeProp $driveRoot 'name')"
        } catch {
            Write-LogJsonl -Level 'WARN' -Event 'workspace_root_fail' -DriveId $driveId -Message $_.Exception.Message
            $parentRefId = Get-SafeProp $rootParentRef 'id'
            if ($rootParentRef -and $parentRefId) {
                $enumerateFromId = $parentRefId
            } else {
                Write-Error 'Workspace mode requires a folder or a resolvable drive root.'
                $script:ExitCode = 1
                return
            }
        }
    }

    Write-LogJsonl -Level 'INFO' -Event 'mode_selected' -Message "Effective mode: $effectiveMode"

    # 3) Enumerate items
    $itemsToExport = @()
    if ($effectiveMode -eq 'page') {
        $itemsToExport = @($rootItem)
    } else {
        Write-LogJsonl -Level 'INFO' -Event 'enumerate_start' -DriveId $driveId -ItemId $enumerateFromId -Message 'Enumerating workspace items...'
        $itemsToExport = @(Get-LoopItemsRecursive -DriveId $driveId -ItemId $enumerateFromId -SinceDate $sinceDate)
        Write-LogJsonl -Level 'INFO' -Event 'enumerate_complete' -Message "Found $($itemsToExport.Count) .loop items"

        # Supplement with search API for completeness
        Write-LogJsonl -Level 'INFO' -Event 'enumerate_search_supplement' -DriveId $driveId `
            -Message 'Supplementing with search API...'
        $searchItems = [System.Collections.Generic.List[object]]::new()
        $searchUri = "$($script:GRAPH_BASE)/drives/$driveId/root/search(q='.loop')?`$select=id,name,webUrl,eTag,cTag,lastModifiedDateTime,size,file,folder,parentReference&`$top=200"
        while ($searchUri) {
            try {
                $searchResp = Invoke-GraphRequest -Uri $searchUri -DriveId $driveId
                $searchValues = @()
                if (Test-SafeProp $searchResp 'value') { $searchValues = @($searchResp.value) }
                foreach ($sv in $searchValues) {
                    $svName = Get-SafeProp $sv 'name'
                    if ($svName -and $svName.EndsWith('.loop', [System.StringComparison]::OrdinalIgnoreCase)) {
                        if ($svName -eq 'Untitled.loop') { continue }
                        if ($sinceDate -ne '') {
                            $svMod = Get-SafeProp $sv 'lastModifiedDateTime'
                            if ($svMod) {
                                $modified = [datetime]::Parse($svMod)
                                $since    = [datetime]::Parse($sinceDate)
                                if ($modified -lt $since) { continue }
                            }
                        }
                        # Normalize driveId
                        $svParent = Get-SafeProp $sv 'parentReference'
                        if ($svParent) {
                            $svDriveId = Get-SafeProp $svParent 'driveId'
                            if (-not $svDriveId -or $svDriveId -ne $driveId) { $svParent.driveId = $driveId }
                        } else {
                            $sv | Add-Member -NotePropertyName 'parentReference' -NotePropertyValue ([PSCustomObject]@{ driveId = $driveId }) -Force
                        }
                        $searchItems.Add($sv)
                    }
                }
                $searchUri = $null
                if (Test-SafeProp $searchResp '@odata.nextLink') {
                    $searchUri = Get-SafeProp $searchResp '@odata.nextLink'
                }
            } catch {
                Write-LogJsonl -Level 'WARN' -Event 'enumerate_search_fail' -DriveId $driveId `
                    -Message "Search API failed: $($_.Exception.Message)"
                $searchUri = $null
            }
        }

        # Merge: deduplicate by item ID
        if ($searchItems.Count -gt 0) {
            $existingIds = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($existing in $itemsToExport) {
                $eid = Get-SafeProp $existing 'id'
                if ($eid) { $existingIds.Add($eid) | Out-Null }
            }
            $newFromSearch = 0
            foreach ($si in $searchItems) {
                $sid = Get-SafeProp $si 'id'
                if ($sid -and -not $existingIds.Contains($sid)) { $newFromSearch++ }
            }
            if ($newFromSearch -gt 0 -or $itemsToExport.Count -eq 0) {
                $mergedIds = [System.Collections.Generic.HashSet[string]]::new()
                $merged = [System.Collections.Generic.List[object]]::new()
                foreach ($ci in $itemsToExport) {
                    $cid = Get-SafeProp $ci 'id'
                    if ($cid -and $mergedIds.Add($cid)) { $merged.Add($ci) }
                }
                foreach ($si in $searchItems) {
                    $sid = Get-SafeProp $si 'id'
                    if ($sid -and $mergedIds.Add($sid)) { $merged.Add($si) }
                }
                $itemsToExport = @($merged)
                Write-LogJsonl -Level 'INFO' -Event 'enumerate_search_merged' -DriveId $driveId `
                    -Message "Search found $($searchItems.Count) items, $newFromSearch new (total: $($itemsToExport.Count))"
                if ($newFromSearch -gt 0) {
                    Write-Warning "Search API found $newFromSearch additional .loop files not found by children enumeration."
                }
            }
        }
    }

    if ($itemsToExport.Count -eq 0) {
        Write-Warning 'No .loop items found to export.'
        return
    }

    if ($itemsToExport.Count -gt $script:LARGE_ITEM_THRESHOLD) {
        Write-Warning "Large export: $($itemsToExport.Count) items."
    }

    # Load state
    $state = @{}
    if (-not $force) { $state = Read-State -Path $statePath }

    # 4) Export each item
    $manifestItems = [System.Collections.Generic.List[object]]::new()
    $failures      = [System.Collections.Generic.List[object]]::new()
    $exportedCount = 0
    $skippedCount  = 0

    foreach ($item in $itemsToExport) {
        try {
            $exportResult = Export-LoopItem `
                -DriveItem $item `
                -OutDir $outDir `
                -ExportRawLoop $exportRaw `
                -ExportHtml $exportHtml `
                -ExportMd $exportMd `
                -State $state `
                -StatePath $statePath `
                -DryRun $dryRun

            if ($exportResult -is [System.Collections.IList]) {
                $exportResult = $exportResult[-1]
            }

            $rSkipped = $exportResult['skipped']
            $rError   = $exportResult['error']
            $rItemId  = $exportResult['itemId']
            $rName    = $exportResult['name']
            $rOutDir  = $exportResult['outputDir']

            if ($rSkipped) { $skippedCount++ }
            elseif ($rError) {
                $failures.Add([ordered]@{ itemId = $rItemId; name = $rName; reason = $rError })
            } else { $exportedCount++ }

            $hashes = @{}
            if ($rOutDir -and -not $dryRun) {
                $idir = Join-Path $outDir 'items' $rOutDir
                foreach ($fname in @('meta.json', 'page.loop', 'page.html', 'page.md')) {
                    $fpath = Join-Path $idir $fname
                    if (Test-Path -LiteralPath $fpath) { $hashes[$fname] = Get-FileSHA256 -Path $fpath }
                }
            }

            $manifestItems.Add([ordered]@{
                itemId               = $exportResult['itemId']
                driveId              = $exportResult['driveId']
                name                 = $exportResult['name']
                webUrl               = $exportResult['webUrl']
                lastModifiedDateTime = $exportResult['lastModifiedDateTime']
                eTag                 = $exportResult['eTag']
                outputDir            = $rOutDir
                exported             = $exportResult['exported']
                skipped              = $rSkipped
                hashes               = $hashes
            })
        } catch {
            $eItemId   = Get-SafeProp $item 'id'
            $eItemName = Get-SafeProp $item 'name'
            Write-LogJsonl -Level 'ERROR' -Event 'export_item_error' -ItemId $eItemId -Message $_.Exception.Message
            $failures.Add([ordered]@{ itemId = $eItemId; name = $eItemName; reason = $_.Exception.Message })
        }
    }

    # 5) Write manifest
    if (-not $dryRun) {
        $manifest = [ordered]@{
            toolVersion   = $script:SUITE_VERSION
            toolName      = 'spbackup-loop'
            runTimestamp   = $script:RunTimestamp
            inputUrl       = $url
            mode           = $effectiveMode
            root           = [ordered]@{
                itemId  = $itemId
                driveId = $driveId
                name    = $rootName
                webUrl  = Get-SafeProp $rootItem 'webUrl'
            }
            totalItems     = $itemsToExport.Count
            exportedCount  = $exportedCount
            skippedCount   = $skippedCount
            failureCount   = $failures.Count
            items          = $manifestItems
            failures       = $failures
        }
        $manifestPath = Join-Path $outDir 'manifest.json'
        $manifestJson = $manifest | ConvertTo-Json -Depth 20
        Write-AtomicFile -Path $manifestPath -Content $manifestJson
    }

    $summaryMsg = "Backup complete: $exportedCount exported, $skippedCount skipped, $($failures.Count) failed (of $($itemsToExport.Count) total)"
    Write-LogJsonl -Level 'INFO' -Event 'backup_complete' -Message $summaryMsg
    Write-Host $summaryMsg -ForegroundColor $(if ($failures.Count -gt 0) { 'Yellow' } else { 'Green' })

    if ($failures.Count -gt 0) {
        $script:ExitCode = 1
        Write-Host "Failures:" -ForegroundColor Yellow
        foreach ($f in $failures) {
            Write-Host "  - $($f.name): $($f.reason)" -ForegroundColor Yellow
        }
    }
}

function Invoke-LoopVerifyCommand {
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

    $manifestItems = if (Test-SafeProp $manifest 'items') { @($manifest.items) } else { @() }
    foreach ($item in $manifestItems) {
        $itemSkipped = Get-SafeProp $item 'skipped'
        if ($itemSkipped) { continue }
        $itemOutDir = Get-SafeProp $item 'outputDir'
        $itemDir = Join-Path $outDir 'items' $itemOutDir

        if (-not (Test-Path -LiteralPath $itemDir)) {
            Write-Host "MISSING DIR: $itemOutDir" -ForegroundColor Red
            $missingFiles++
            continue
        }

        $hashObj = Get-SafeProp $item 'hashes'
        if ($hashObj) {
            $props = if ($hashObj -is [hashtable]) { $hashObj.Keys } else { $hashObj.PSObject.Properties.Name }
            foreach ($fname in $props) {
                $expectedHash = if ($hashObj -is [hashtable]) { $hashObj[$fname] } else { $hashObj.$fname }
                $fpath = Join-Path $itemDir $fname
                $checkedFiles++

                if (-not (Test-Path -LiteralPath $fpath)) {
                    Write-Host "MISSING: $(Join-Path $itemOutDir $fname)" -ForegroundColor Red
                    $missingFiles++
                    continue
                }

                $actualHash = Get-FileSHA256 -Path $fpath
                if ($actualHash -ne $expectedHash) {
                    Write-Host "HASH MISMATCH: $(Join-Path $itemOutDir $fname) (expected=$expectedHash, actual=$actualHash)" -ForegroundColor Yellow
                    $hashMismatch++
                } else { $okFiles++ }
            }
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

function Invoke-LoopDiagnoseCommand {
    [CmdletBinding()]
    param([hashtable]$Opts)

    Write-Host ''
    Write-Host '=== Loop Backup Diagnostic ===' -ForegroundColor Cyan
    Write-Host ''

    # 1) Env vars & auth method
    Write-Host '1. Environment variables' -ForegroundColor Yellow
    $tenantId     = $env:TENANT_ID
    $clientId     = $env:CLIENT_ID
    $clientSecret = $env:CLIENT_SECRET
    Write-Host "   TENANT_ID:     $(if ($tenantId) { $tenantId } else { '(NOT SET)' })"
    Write-Host "   CLIENT_ID:     $(if ($clientId) { $clientId } else { '(NOT SET)' })"
    Write-Host "   CLIENT_SECRET: $(if ($clientSecret) { $clientSecret.Substring(0, [math]::Min(4, $clientSecret.Length)) + '***' + ' (' + $clientSecret.Length + ' chars)' } else { '(not set — will use certificate auth)' })"
    Write-Host "   SEARCH_REGION: $(if ($env:SEARCH_REGION) { $env:SEARCH_REGION } else { '(not set — needed for search with app-only auth)' })"

    $cert = Find-Certificate
    if ($cert) {
        Write-Host "   CERT:          $($cert.Subject) (thumbprint=$($cert.Thumbprint))" -ForegroundColor Green
        $cert.Dispose()
    } elseif (-not $clientSecret) {
        Write-Host '   CERT:          (not found)' -ForegroundColor Red
    }

    $authMethod = if ($clientSecret) { 'client_secret' } else { 'certificate' }
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
    Show-TokenDiagnostic -Token $token
    Write-Host ''

    # 4) Graph API
    Write-Host '4. Graph API connectivity tests' -ForegroundColor Yellow
    Show-GraphConnectivityTests
    Write-Host ''

    # 5) URL-specific tests
    $url = $Opts['url']
    if ($url) {
        Write-Host '5. URL-specific tests' -ForegroundColor Yellow
        $loopAppData = Expand-LoopAppUrl -Url $url

        if ($loopAppData -and $loopAppData['candidateUrls']) {
            Write-Host '   Candidate URLs decoded from Loop link:' -ForegroundColor DarkGray
            foreach ($cu in @($loopAppData['candidateUrls'])) {
                Write-Host "     $cu" -ForegroundColor DarkGray
            }
            Write-Host ''
        }

        if ($loopAppData -and $loopAppData['driveIdentifier']) {
            $diagDriveId = $loopAppData['driveIdentifier']
            Write-Host "   Drive ID: $diagDriveId"

            try {
                $drive = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/${diagDriveId}?`$select=id,name,webUrl,driveType"
                Write-Host "   /drives/{id}: OK — $(Get-SafeProp $drive 'name')" -ForegroundColor Green
            } catch {
                Write-Host "   /drives/{id}: FAIL — $($_.Exception.Message)" -ForegroundColor Red
            }

            $diagItemId = $loopAppData['pageItemId']
            if ($diagItemId) {
                try {
                    $item = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/${diagDriveId}/items/${diagItemId}?`$select=id,name,webUrl"
                    Write-Host "   /drives/{id}/items/{id}: OK — $(Get-SafeProp $item 'name')" -ForegroundColor Green
                } catch {
                    Write-Host "   /drives/{id}/items/{id}: FAIL — $($_.Exception.Message)" -ForegroundColor Red
                }
            }

            try {
                $rootChildren = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/${diagDriveId}/root/children?`$top=10&`$select=id,name,file,folder"
                [array]$rcItems = @()
                if ($rootChildren.PSObject.Properties['value']) { [array]$rcItems = @($rootChildren.value) }
                $rcCount = [int]$rcItems.Length
                if ($rcCount -gt 0) {
                    Write-Host "   /drives/{id}/root/children: OK — $rcCount item(s)" -ForegroundColor Green
                    foreach ($rci in $rcItems) {
                        $rciName = Get-SafeProp $rci 'name'
                        $rciType = if (Test-SafeProp $rci 'folder') { 'folder' } else { 'file' }
                        Write-Host "     - $rciName ($rciType)" -ForegroundColor DarkGray
                    }
                } else {
                    Write-Host "   /drives/{id}/root/children: RETURNED 0 ITEMS" -ForegroundColor Yellow
                    Write-Host "     → Drive root is accessible but no children visible." -ForegroundColor Yellow
                    Write-Host "     → This usually means your guest app has 'Read' but NOT 'ReadContent'." -ForegroundColor Yellow
                    Write-Host "     → Fix: re-run Set-SPOApplicationPermission with -PermissionAppOnly Read, ReadContent" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "   /drives/{id}/root/children: FAIL — $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "     → If this is 403: guest app permission may be missing entirely." -ForegroundColor Yellow
            }

            # Delta query — alternative enumeration
            try {
                $delta = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/${diagDriveId}/root/delta?`$top=10&`$select=id,name,file,folder,deleted"
                [array]$deltaItems = @()
                if ($delta.PSObject.Properties['value']) { [array]$deltaItems = @($delta.value) }
                $deltaCount = [int]$deltaItems.Length
                Write-Host "   /drives/{id}/root/delta: OK — $deltaCount item(s)" -ForegroundColor $(if ($deltaCount -gt 0) { 'Green' } else { 'Yellow' })
            } catch {
                Write-Host "   /drives/{id}/root/delta: FAIL — $($_.Exception.Message)" -ForegroundColor Red
            }

            try {
                $driveSearch = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/${diagDriveId}/root/search(q='.loop')?`$top=10&`$select=id,name"
                [array]$searchItems = @()
                if ($driveSearch.PSObject.Properties['value']) { [array]$searchItems = @($driveSearch.value) }
                Write-Host "   /drives/{id}/search(.loop): OK — $([int]$searchItems.Length) result(s)" -ForegroundColor $(if ($searchItems.Length -gt 0) { 'Green' } else { 'Yellow' })
                foreach ($si in $searchItems) {
                    Write-Host "     - $(Get-SafeProp $si 'name')" -ForegroundColor DarkGray
                }
            } catch {
                Write-Host "   /drives/{id}/search(.loop): FAIL — $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Extract container GUID from candidate URLs
        $containerGuid = $null
        if ($loopAppData) {
            foreach ($cu in @($loopAppData['candidateUrls'])) {
                if ($cu -match 'CSP_([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
                    $containerGuid = $Matches[1]
                    break
                }
            }
        }

        # Site by URL (the site backing this SPE container)
        if ($containerGuid) {
            $spHost = $null
            foreach ($cu in @($loopAppData['candidateUrls'])) {
                if ($cu -match 'https://([^/]+\.sharepoint\.com)') {
                    $spHost = $Matches[1]
                    break
                }
            }
            if (-not $spHost) { $spHost = 'unknown.sharepoint.com' }
            $siteUrl2 = $spHost + ':/contentstorage/CSP_' + $containerGuid + ':'
            Write-Host ''
            Write-Host "   Site by URL: /sites/$siteUrl2" -ForegroundColor DarkGray
            try {
                $siteUri = "$($script:GRAPH_BASE)/sites/$($siteUrl2)?`$select=id,displayName,webUrl"
                $site = Invoke-GraphRequest -Uri $siteUri
                $siteId = Get-SafeProp $site 'id'
                $siteName = Get-SafeProp $site 'displayName'
                Write-Host "   /sites/{url}: OK — $siteName (id=$siteId)" -ForegroundColor Green

                try {
                    $siteDrives = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/sites/${siteId}/drives?`$select=id,name,webUrl,driveType"
                    $drivesList = if ($siteDrives.PSObject.Properties['value']) { @($siteDrives.value) } else { @() }
                    Write-Host "   /sites/{id}/drives: OK — found $($drivesList.Count) drive(s)" -ForegroundColor Green
                    foreach ($d in $drivesList) {
                        $dId = Get-SafeProp $d 'id'
                        $dName = Get-SafeProp $d 'name'
                        $dType = Get-SafeProp $d 'driveType'
                        Write-Host "     - $dName (type=$dType, id=$dId)" -ForegroundColor DarkGray
                    }
                } catch {
                    Write-Host "   /sites/{id}/drives: FAIL — $($_.Exception.Message)" -ForegroundColor Red
                }
            } catch {
                Write-Host "   /sites/{url}: FAIL — $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Container API (fileStorage/containers)
        if ($containerGuid) {
            Write-Host ''
            Write-Host "   Container GUID: $containerGuid"
            try {
                $container = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/storage/fileStorage/containers/${containerGuid}?`$select=id,displayName,containerTypeId,status"
                Write-Host "   /storage/.../containers/{id}: OK — $(Get-SafeProp $container 'displayName') (status=$(Get-SafeProp $container 'status'))" -ForegroundColor Green
            } catch {
                Write-Host "   /storage/.../containers/{id}: FAIL — $($_.Exception.Message)" -ForegroundColor Red
            }

            # List all containers by discovering the containerTypeId
            Write-Host ''
            Write-Host '   Listing all accessible containers:' -ForegroundColor DarkGray
            try {
                $containerTypeId = $null
                $discoverySites = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/sites?search=CSP_&`$top=5&`$select=id,displayName,webUrl"
                $discSiteList = if ($discoverySites.PSObject.Properties['value']) { @($discoverySites.value) } else { @() }
                foreach ($ds in $discSiteList) {
                    $dsUrl = Get-SafeProp $ds 'webUrl'
                    if ($dsUrl -match '/contentstorage/CSP_([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
                        $dsCspGuid = $Matches[1]
                        try {
                            $dsContainer = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/storage/fileStorage/containers/${dsCspGuid}?`$select=id,containerTypeId"
                            if (Test-SafeProp $dsContainer 'containerTypeId') {
                                $containerTypeId = Get-SafeProp $dsContainer 'containerTypeId'
                                Write-Host "   Discovered Loop containerTypeId: $containerTypeId" -ForegroundColor DarkGray
                                break
                            }
                        } catch { <# skip #> }
                    }
                }

                if (-not $containerTypeId) {
                    Write-Host '     Cannot list containers: containerTypeId not discovered.' -ForegroundColor Yellow
                    throw 'containerTypeId not available'
                }

                $allContainers = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/storage/fileStorage/containers?`$filter=containerTypeId eq $containerTypeId"
                $cList = if ($allContainers.PSObject.Properties['value']) { @($allContainers.value) } else { @() }
                if ($cList.Count -eq 0) {
                    Write-Host '     (no containers visible)' -ForegroundColor Yellow
                } else {
                    $foundTarget = $false
                    foreach ($c in $cList) {
                        $cId = Get-SafeProp $c 'id'
                        $cName = Get-SafeProp $c 'displayName'
                        $cStatus = Get-SafeProp $c 'status'
                        $isTarget = ($cId -eq $containerGuid)
                        if ($isTarget) { $foundTarget = $true }
                        $color = if ($isTarget) { 'Green' } else { 'DarkGray' }
                        $marker = if ($isTarget) { ' ← TARGET' } else { '' }
                        Write-Host "     - $cName (id=$cId, status=$cStatus)$marker" -ForegroundColor $color
                    }
                    if (-not $foundTarget) {
                        Write-Host "     Target container $containerGuid NOT in list" -ForegroundColor Yellow
                    }
                }
            } catch {
                Write-Host "   /storage/.../containers (list): FAIL — $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Shares API tests
        if ($loopAppData -and $loopAppData['candidateUrls']) {
            Write-Host ''
            Write-Host '   Shares API tests:' -ForegroundColor DarkGray
            foreach ($cu in @($loopAppData['candidateUrls'])) {
                $shareId = ConvertTo-ShareId -Url $cu
                $shortShareId = if ($shareId.Length -gt 40) { $shareId.Substring(0, 40) + '...' } else { $shareId }
                try {
                    $shareItem = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/shares/${shareId}/driveItem?`$select=id,name,webUrl"
                    Write-Host "   shares/${shortShareId}: OK — $(Get-SafeProp $shareItem 'name')" -ForegroundColor Green
                } catch {
                    Write-Host "   shares/${shortShareId}: FAIL — $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }

        # Sites search for this container
        if ($containerGuid) {
            Write-Host ''
            Write-Host '   Searching for container site:' -ForegroundColor DarkGray
            try {
                $siteSearch = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/sites?search=CSP_${containerGuid}&`$select=id,displayName,webUrl"
                [array]$searchList = @()
                if ($siteSearch.PSObject.Properties['value']) { [array]$searchList = @($siteSearch.value) }
                $searchCount = [int]$searchList.Length
                if ($searchCount -eq 0) {
                    Write-Host '     No sites found matching this container GUID' -ForegroundColor Yellow
                    Write-Host '     This container may not be indexed, deleted, or a personal Loop workspace.' -ForegroundColor Yellow
                } else {
                    foreach ($sr in $searchList) {
                        Write-Host "     - $(Get-SafeProp $sr 'displayName') ($(Get-SafeProp $sr 'webUrl'))" -ForegroundColor DarkGray
                    }
                }
            } catch {
                Write-Host "   site search: FAIL — $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Cross-check: test a known-accessible workspace to confirm the tool works
        Write-Host ''
        Write-Host '   Cross-check with known-accessible sites:' -ForegroundColor Yellow
        try {
            $allSites = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/sites?search=*&`$top=50&`$select=id,displayName,webUrl"
            $allSitesList = if ($allSites.PSObject.Properties['value']) { @($allSites.value) } else { @() }
            $cspSites = @($allSitesList | Where-Object { (Get-SafeProp $_ 'webUrl') -match '/contentstorage/CSP_' })

            if ($cspSites.Count -eq 0) {
                Write-Host '     No SPE container sites found in tenant.' -ForegroundColor Yellow
            } else {
                Write-Host "     Found $($cspSites.Count) SPE container site(s) accessible to this app:" -ForegroundColor DarkGray
                foreach ($cs in $cspSites) {
                    $csId = Get-SafeProp $cs 'id'
                    $csName = Get-SafeProp $cs 'displayName'
                    $csUrl = Get-SafeProp $cs 'webUrl'
                    Write-Host "     - $csName ($csUrl)" -ForegroundColor DarkGray

                    try {
                        $csDrive = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/sites/${csId}/drive?`$select=id,name,driveType"
                        $csDriveId = Get-SafeProp $csDrive 'id'
                        $csDriveName = Get-SafeProp $csDrive 'name'
                        Write-Host "       Drive: $csDriveName (id=$csDriveId)" -ForegroundColor Green

                        try {
                            $csRoot = Invoke-GraphRequest -Uri "$($script:GRAPH_BASE)/drives/${csDriveId}/root/children?`$top=5&`$select=id,name,file,folder"
                            [array]$csItems = @()
                            if ($csRoot.PSObject.Properties['value']) { [array]$csItems = @($csRoot.value) }
                            $csItemCount = [int]$csItems.Length
                            Write-Host "       Drive items: $csItemCount item(s) — ACCESSIBLE" -ForegroundColor Green
                        } catch {
                            Write-Host "       Drive items: FAIL — $($_.Exception.Message)" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "       Drive (default): FAIL — $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host '       → Guest app permission not granted.' -ForegroundColor Yellow
                    }
                }

                $targetAccessible = $false
                if ($containerGuid) {
                    foreach ($cs in $cspSites) {
                        $csUrl2 = Get-SafeProp $cs 'webUrl'
                        if ($csUrl2 -match $containerGuid) { $targetAccessible = $true; break }
                    }
                }

                if (-not $targetAccessible -and $containerGuid) {
                    Write-Host ''
                    Write-Host "     FINDING: Target container CSP_$containerGuid is NOT among accessible containers." -ForegroundColor Yellow
                    Write-Host '     But other SPE containers ARE accessible. This means:' -ForegroundColor Yellow
                    Write-Host '     • This specific workspace has restricted access (personal workspace?)' -ForegroundColor Yellow
                    Write-Host '     • The tool DOES work — try backing up an accessible workspace instead.' -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "     Cross-check: FAIL — $($_.Exception.Message)" -ForegroundColor Red
        }

        Write-Host ''
        Write-Host '6. Diagnosis' -ForegroundColor Yellow
        Write-Host '   Loop workspaces use SharePoint Embedded (SPE) containers.' -ForegroundColor DarkGray
        Write-Host '   Even with Sites.ReadWrite.All, SPE containers require a "guest app" permission grant.' -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '   SPE container type permissions:' -ForegroundColor DarkGray
        Write-Host '     • Read        → container/drive metadata only' -ForegroundColor DarkGray
        Write-Host '     • ReadContent → list and download files (required for /root/children and file content)' -ForegroundColor DarkGray
        Write-Host '     → Both are required for backup.' -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '   FIX: Grant your app guest access to Loop containers (Windows PowerShell 5.1 only):' -ForegroundColor Green
        Write-Host '     Connect-SPOService -Url https://<tenant>-admin.sharepoint.com' -ForegroundColor Cyan
        Write-Host '     Set-SPOApplicationPermission `' -ForegroundColor Cyan
        Write-Host '       -OwningApplicationId "a187e399-0c36-4b98-8f04-1edc167a0996" `' -ForegroundColor Cyan
        Write-Host '       -GuestApplicationId "<YOUR_CLIENT_ID>" `' -ForegroundColor Cyan
        Write-Host '       -PermissionAppOnly Read, ReadContent' -ForegroundColor Cyan
        Write-Host ''
    }

    Write-Host ''
    Write-Host '=== Diagnostic complete ===' -ForegroundColor Cyan
    Write-Host ''
}

# ─────────────────────────────────────────────────────────────────────────────
# Usage & Entry Point
# ─────────────────────────────────────────────────────────────────────────────
function Show-LoopUsage {
    $usage = @"
$($script:TOOL_NAME) — Microsoft Loop Backup

USAGE:
  pwsh ./backup-loop.ps1 backup --url "<URL>" --out "<dir>" [OPTIONS]
  pwsh ./backup-loop.ps1 resolve --url "<URL>"
  pwsh ./backup-loop.ps1 verify --out "<dir>"
  pwsh ./backup-loop.ps1 diagnose [--url "<URL>"] [--verbose]

  Or via the unified entry point:
  pwsh ./spbackup.ps1 loop backup --url "<URL>" --out "<dir>"

COMMANDS:
  backup      Export Loop pages (HTML, Markdown, raw .loop)
  resolve     Resolve a URL to Graph resource(s) and print JSON
  verify      Verify backup integrity against manifest
  diagnose    Check auth, decode JWT, test Graph API access
  setup-venv  Force-(re)create Python venv in tools/.venv (auto-created on first backup)
  help        Show this usage information

BACKUP OPTIONS:
  --url <URL>           Loop URL, SharePoint sharing link, or loop.cloud.microsoft URL (required)
  --out <dir>           Output directory (required)
  --mode <mode>         page | workspace | auto (default: auto)
  --raw-loop            Download raw .loop file bytes
  --html                Export HTML (default: on)
  --no-html             Disable HTML export
  --md                  Export Markdown (default: on)
  --no-md               Disable Markdown export
  --concurrency <N>     Max parallel downloads (default: 4)
  --since <ISO date>    Only items modified after this date
  --state <path>        State file path (default: <out>/.state.json)
  --force               Re-export all items, ignoring state
  --dry-run             Resolve and enumerate only; no downloads
  --verbose             Human-readable console output

ENVIRONMENT VARIABLES (required):
  TENANT_ID             Azure AD / Entra tenant ID
  CLIENT_ID             App registration client ID

AUTHENTICATION (one of the following):
  CLIENT_SECRET         App registration client secret
  CERT_PATH             Path to .pfx certificate (auto-discovered from ./certs/)
  CERT_PASSWORD         Certificate password (if any)

OPTIONAL ENV VARS:
  PYTHON                Path to python executable
  PANDOC                Path to pandoc executable
  SEARCH_REGION         SharePoint region for search (e.g. NAM, EUR, APC)
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
        'resolve'    { Invoke-LoopResolveCommand -Opts $opts }
        'backup'     { Invoke-LoopBackupCommand -Opts $opts }
        'verify'     { Invoke-LoopVerifyCommand -Opts $opts }
        'setup-venv' { Initialize-PythonVenv -Force }
        'diagnose'   { Invoke-LoopDiagnoseCommand -Opts $opts }
        'help'       { Show-LoopUsage }
        default {
            Write-Error "Unknown command: $command"
            Show-LoopUsage
            exit 1
        }
    }
}

# Run when invoked directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.') {
    Main -RawArgs $args
    exit $script:ExitCode
}
