# ─────────────────────────────────────────────────────────────────────────────
# SiteResolver.ps1 — Resolve SharePoint site URLs to Graph site IDs
# ─────────────────────────────────────────────────────────────────────────────
# Provides:
#   Resolve-SiteFromUrl — resolves any SharePoint site URL to a Graph site object
# ─────────────────────────────────────────────────────────────────────────────

function Resolve-SiteFromUrl {
    <#
    .SYNOPSIS
        Resolve a SharePoint site URL to a Graph site ID.
        Handles subsites by trying progressively deeper paths.
        Accepts URLs like:
          https://contoso.sharepoint.com/sites/team
          https://contoso.sharepoint.com/sites/intranet/subsite/Lists/Tasks
    #>
    [CmdletBinding()]
    param([string]$Url)

    Write-LogJsonl -Level 'INFO' -Event 'resolve_site_start' -Url $Url -Message 'Resolving site from URL'

    $parsed = [uri]$Url
    $hostname = $parsed.Host
    $script:SharePointHostname = $hostname   # store for SP REST API token scoping
    $pathSegments = $parsed.AbsolutePath.TrimEnd('/') -split '/' | Where-Object { $_ -ne '' }

    # Find the /sites/<name> or /teams/<name> base index
    $siteBaseIdx = -1
    for ($i = 0; $i -lt $pathSegments.Count; $i++) {
        if ($pathSegments[$i] -in @('sites', 'teams') -and ($i + 1) -lt $pathSegments.Count) {
            $siteBaseIdx = $i
            break
        }
    }

    if ($siteBaseIdx -eq -1) {
        # No /sites/ or /teams/ — try root site
        $graphUri = "$($script:GRAPH_BASE)/sites/${hostname}?`$select=id,displayName,webUrl,name"
        Write-LogJsonl -Level 'DEBUG' -Event 'resolve_site_graph' -Url $graphUri -Message "Resolving root site via Graph"
        try {
            $site = Invoke-GraphRequest -Uri $graphUri
            Write-LogJsonl -Level 'INFO' -Event 'resolve_site_success' -SiteId (Get-SafeProp $site 'id') -Message "Resolved site: $(Get-SafeProp $site 'displayName')"
            return $site
        } catch {
            throw "Could not resolve site from URL '$Url': $($_.Exception.Message)"
        }
    }

    # Build candidate site paths from longest (deepest subsite) to shortest.
    $stopSegments = @('Lists', '_layouts', '_api', 'SitePages', 'SiteAssets',
                      'Shared%20Documents', 'Forms', '_catalogs', 'AllItems.aspx')

    $candidates = [System.Collections.Generic.List[string]]::new()

    $maxDepth = $pathSegments.Count - 1
    for ($depth = $maxDepth; $depth -ge ($siteBaseIdx + 1); $depth--) {
        $seg = $pathSegments[$depth]
        if ($seg -in $stopSegments -or $seg.EndsWith('.aspx')) { continue }
        $candidatePath = '/' + (($pathSegments[0..$depth]) -join '/')
        $candidates.Add($candidatePath)
    }

    # Try each candidate from deepest to shallowest
    $lastError = $null
    foreach ($candidatePath in $candidates) {
        $graphUri = "$($script:GRAPH_BASE)/sites/${hostname}:${candidatePath}?`$select=id,displayName,webUrl,name"
        Write-LogJsonl -Level 'DEBUG' -Event 'resolve_site_graph' -Url $graphUri -Message "Trying site path: $candidatePath"
        try {
            $site = Invoke-GraphRequest -Uri $graphUri
            $siteId = Get-SafeProp $site 'id'
            $siteName = Get-SafeProp $site 'displayName'
            Write-LogJsonl -Level 'INFO' -Event 'resolve_site_success' -SiteId $siteId -Message "Resolved site: $siteName (path=$candidatePath)"
            return $site
        } catch {
            $lastError = $_
            $code = 0
            if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
            Write-LogJsonl -Level 'DEBUG' -Event 'resolve_site_try_fail' -Url $graphUri -StatusCode $code `
                -Message "Path '$candidatePath' failed: $($_.Exception.Message)"
        }
    }

    Write-LogJsonl -Level 'ERROR' -Event 'resolve_site_fail' -Url $Url -Message "All candidate paths failed"
    throw "Could not resolve site from URL '$Url'. Tried paths: $($candidates -join ', '). Last error: $($lastError.Exception.Message)"
}
