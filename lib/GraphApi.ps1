# ─────────────────────────────────────────────────────────────────────────────
# GraphApi.ps1 — HTTP wrappers for Microsoft Graph & SharePoint REST APIs
# ─────────────────────────────────────────────────────────────────────────────
# Provides:
#   Invoke-GraphRequest       — Graph API call with retry / throttle handling
#   Invoke-SharePointRequest  — SharePoint REST API call with retry / throttle
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-GraphRequest {
    [CmdletBinding()]
    param(
        [string]$Uri,
        [string]$Method     = 'GET',
        [object]$Body       = $null,
        [string]$ContentType = 'application/json',
        [switch]$Raw,
        [string]$OutFile    = '',
        [string]$SiteId     = '',
        [string]$ListId     = '',
        [string]$DriveId    = '',
        [string]$ItemId     = '',
        [int[]]$ExtraRetryStatusCodes = @()
    )
    $attempt = 0
    while ($true) {
        $attempt++
        $token = Get-GraphToken

        $headers = @{ Authorization = "Bearer $token" }

        $params = @{
            Uri             = $Uri
            Method          = $Method
            Headers         = $headers
            ErrorAction     = 'Stop'
        }
        if ($Body) {
            if ($Body -is [string]) {
                $params['Body'] = $Body
            } else {
                $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
            }
            $params['ContentType'] = $ContentType
        }
        if ($OutFile -ne '') {
            $params['OutFile'] = $OutFile
        }

        $semaphoreAcquired = $false
        try {
            if ($script:Semaphore) {
                $script:Semaphore.WaitOne() | Out-Null
                $semaphoreAcquired = $true
            }

            if ($Raw) {
                $response = Invoke-WebRequest @params -UseBasicParsing
                if ($OutFile -ne '') {
                    Write-LogJsonl -Level 'DEBUG' -Event 'graph_download' -Url $Uri `
                        -SiteId $SiteId -ListId $ListId -DriveId $DriveId -ItemId $ItemId `
                        -Attempt $attempt -StatusCode 200 -Message "Downloaded to $OutFile"
                    return $true
                }
                return $response
            } else {
                $response = Invoke-RestMethod @params
                Write-LogJsonl -Level 'DEBUG' -Event 'graph_request' -Url $Uri `
                    -SiteId $SiteId -ListId $ListId -DriveId $DriveId -ItemId $ItemId `
                    -Attempt $attempt -StatusCode 200 -Message 'OK'
                return $response
            }
        } catch {
            $ex = $_.Exception
            $statusCode = 0
            $retryAfter = 0
            $errorBody  = ''

            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $errorBody = $_.ErrorDetails.Message
            }

            if ($ex.PSObject.Properties['Response'] -and $ex.Response) {
                $statusCode = [int]$ex.Response.StatusCode
                $raValues = $null
                if ($ex.Response.Headers -and $ex.Response.Headers.TryGetValues('Retry-After', [ref]$raValues)) {
                    $raFirst = $raValues | Select-Object -First 1
                    if ($raFirst -match '^\d+$') {
                        $retryAfter = [int]$raFirst
                    }
                }
            }

            # Retryable: 429 (throttled), 503/504 (overloaded), and 0 (TCP timeout / DNS / socket error)
            $retryable = $statusCode -in @(0, 429, 503, 504)
            $maxForThis = $script:MAX_RETRIES
            if (-not $retryable -and $ExtraRetryStatusCodes.Count -gt 0 -and $statusCode -in $ExtraRetryStatusCodes) {
                $retryable = $true
                $maxForThis = [math]::Min(3, $script:MAX_RETRIES)
            }
            # Network-level errors get a longer minimum backoff (SharePoint connection throttling)
            if ($statusCode -eq 0 -and $retryAfter -lt 10) {
                $retryAfter = 10
            }

            # Parse Graph error code from response body
            $graphErrorCode = ''
            $graphErrorMsg  = ''
            if ($errorBody) {
                try {
                    $errObj = $errorBody | ConvertFrom-Json
                    $errInner = Get-SafeProp $errObj 'error'
                    if ($errInner) {
                        $graphErrorCode = Get-SafeProp $errInner 'code'
                        $graphErrorMsg  = Get-SafeProp $errInner 'message'
                    }
                } catch {}
            }

            $displayMsg = $ex.Message -replace [regex]::Escape($env:CLIENT_SECRET), '***' -replace [regex]::Escape($script:CachedToken ?? '___'), '***'
            $displayMsg = Get-ConciseErrorMessage $displayMsg
            if ($graphErrorCode) {
                $graphErrorMsg = Get-ConciseErrorMessage $graphErrorMsg
                $displayMsg += " [Graph error: $graphErrorCode — $graphErrorMsg]"
            }

            Write-LogJsonl -Level $(if ($retryable) { 'WARN' } else { 'ERROR' }) `
                           -Event 'graph_error' -Url $Uri `
                           -SiteId $SiteId -ListId $ListId -DriveId $DriveId -ItemId $ItemId `
                           -Attempt $attempt -StatusCode $statusCode `
                           -Message $displayMsg

            if ($retryable -and $attempt -lt $maxForThis) {
                $baseDelay = [math]::Pow(2, $attempt)
                $jitter    = Get-Random -Minimum 0.0 -Maximum 1.0
                $delay     = [math]::Max($retryAfter, $baseDelay + $jitter)
                Write-LogJsonl -Level 'INFO' -Event 'retry_wait' -Url $Uri -Attempt $attempt -Message "Waiting ${delay}s before retry"
                Start-Sleep -Seconds $delay
                continue
            }

            throw
        } finally {
            if ($semaphoreAcquired -and $script:Semaphore) {
                $script:Semaphore.Release() | Out-Null
            }
        }
    }
}

function Invoke-SharePointRequest {
    <#
    .SYNOPSIS
        Call a SharePoint REST API endpoint with retry logic.
        Requires a certificate-based SharePoint token (Get-SharePointToken).
    #>
    [CmdletBinding()]
    param(
        [string]$Uri,
        [string]$Method  = 'GET',
        [string]$OutFile = '',
        [string]$SiteId  = '',
        [string]$ListId  = '',
        [string]$ItemId  = ''
    )
    $attempt = 0
    while ($true) {
        $attempt++
        $spToken = Get-SharePointToken
        if (-not $spToken) { throw 'No SharePoint token available.' }

        $headers = @{
            Authorization = "Bearer $spToken"
            Accept        = 'application/json;odata=nometadata'
        }

        try {
            $params = @{
                Uri             = $Uri
                Method          = $Method
                Headers         = $headers
                UseBasicParsing = $true
                ErrorAction     = 'Stop'
            }
            if ($OutFile) { $params['OutFile'] = $OutFile }

            if ($OutFile) {
                # Use Invoke-WebRequest for binary file downloads to avoid
                # Invoke-RestMethod attempting XML/JSON deserialization.
                Invoke-WebRequest @params | Out-Null
                Write-LogJsonl -Level 'DEBUG' -Event 'sp_download' -Url $Uri -Message "Downloaded to $OutFile"
                return $true
            }

            $response = Invoke-RestMethod @params
            Write-LogJsonl -Level 'DEBUG' -Event 'sp_request' -Url $Uri -Message 'OK'
            return $response
        } catch {
            $statusCode = 0
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }

            if ($statusCode -eq 429 -or $statusCode -ge 500 -or $statusCode -eq 0) {
                if ($attempt -ge $script:MAX_RETRIES) { throw }
                $retryAfter = 2 * [math]::Pow(2, $attempt - 1)
                # Network-level errors get a longer minimum backoff
                if ($statusCode -eq 0 -and $retryAfter -lt 10) { $retryAfter = 10 }
                try {
                    $ra = $_.Exception.Response.Headers['Retry-After']
                    if ($ra) { $retryAfter = [int]$ra }
                } catch { }
                $label = if ($statusCode -eq 0) { 'connection error' } else { "HTTP $statusCode" }
                Write-LogJsonl -Level 'WARN' -Event 'sp_throttle' -Url $Uri -Message "$label — retrying in ${retryAfter}s (attempt $attempt)"
                Start-Sleep -Seconds $retryAfter
                continue
            }

            $conciseErr = Get-ConciseErrorMessage $_.Exception.Message
            Write-LogJsonl -Level 'ERROR' -Event 'sp_error' -Url $Uri -SiteId $SiteId -ListId $ListId -ItemId $ItemId -Message $conciseErr
            throw
        }
    }
}
