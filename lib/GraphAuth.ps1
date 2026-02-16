# ─────────────────────────────────────────────────────────────────────────────
# GraphAuth.ps1 — Graph & SharePoint token acquisition
# ─────────────────────────────────────────────────────────────────────────────
# Provides:
#   Get-GraphToken        — client-secret-based OAuth2 token for Microsoft Graph
#   Get-SharePointToken   — certificate-based (JWT bearer) token for SharePoint REST API
#   Test-SharePointAccess — quick connectivity check for the SP REST API
# ─────────────────────────────────────────────────────────────────────────────

function Get-GraphToken {
    [CmdletBinding()]
    param()

    if ($script:CachedToken -and $script:CachedTokenExpiry -gt (Get-Date).AddMinutes(2)) {
        return $script:CachedToken
    }

    $tenantId     = $env:TENANT_ID
    $clientId     = $env:CLIENT_ID
    $clientSecret = $env:CLIENT_SECRET

    if ([string]::IsNullOrWhiteSpace($tenantId))     { throw 'Environment variable TENANT_ID is not set.' }
    if ([string]::IsNullOrWhiteSpace($clientId))      { throw 'Environment variable CLIENT_ID is not set.' }
    if ([string]::IsNullOrWhiteSpace($clientSecret))  { throw 'Environment variable CLIENT_SECRET is not set.' }

    $tokenUrl = $script:TOKEN_URL_TEMPLATE -f $tenantId
    $body = @{
        client_id     = $clientId
        client_secret = $clientSecret
        grant_type    = 'client_credentials'
        scope         = 'https://graph.microsoft.com/.default'
    }

    Write-LogJsonl -Level 'DEBUG' -Event 'auth_request' -Url $tokenUrl -Message 'Requesting access token'

    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded'
    } catch {
        Write-LogJsonl -Level 'ERROR' -Event 'auth_failure' -Url $tokenUrl -Message $_.Exception.Message
        throw "Failed to acquire Graph access token: $($_.Exception.Message)"
    }

    $accessToken = Get-SafeProp $response 'access_token'
    $expiresIn   = Get-SafeProp $response 'expires_in'
    if (-not $accessToken) {
        throw "Token response did not contain access_token. Response keys: $(($response.PSObject.Properties.Name) -join ', ')"
    }

    $script:CachedToken       = $accessToken
    $script:CachedTokenExpiry  = (Get-Date).AddSeconds(($expiresIn ?? 3600))

    Write-LogJsonl -Level 'INFO' -Event 'auth_success' -Message "Token acquired, expires in ${expiresIn}s"
    return $script:CachedToken
}

function Get-SharePointToken {
    <#
    .SYNOPSIS
        Acquire an OAuth2 token scoped to SharePoint Online using certificate-based
        client assertion (JWT Bearer).

        SharePoint Online requires certificate auth for app-only REST API access;
        client secrets are rejected with 401.  The certificate (.pfx) path is read
        from the CERT_PATH environment variable or auto-discovered from ./certs/.

        Required Azure AD setup:
          - SharePoint > Sites.Read.All (Application) with admin consent
          - Certificate public key (.cer) uploaded to the app registration
    #>
    [CmdletBinding()]
    param()

    if ($script:CachedSpToken -and $script:CachedSpTokenExpiry -gt (Get-Date).AddMinutes(2)) {
        return $script:CachedSpToken
    }

    if (-not $script:SharePointHostname) {
        throw 'SharePoint hostname not set. Resolve a site URL first.'
    }

    $tenantId = $env:TENANT_ID
    $clientId = $env:CLIENT_ID

    # --- Locate the certificate (.pfx) ---
    $certPath = $env:CERT_PATH
    if (-not $certPath) {
        $autoCert = Join-Path $script:ProjectRoot 'certs' 'spbackup.pfx'
        if (Test-Path $autoCert) { $certPath = $autoCert }
    }
    if (-not $certPath -or -not (Test-Path $certPath)) {
        Write-LogJsonl -Level 'WARN' -Event 'sp_auth_no_cert' `
            -Message 'No certificate found. Set CERT_PATH env var or place spbackup.pfx in ./certs/. SP REST API requires cert auth.'
        return $null
    }

    # --- Load certificate ---
    try {
        $certPassword = $env:CERT_PASSWORD
        if ($certPassword) {
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                $certPath, $certPassword,
                [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
        } else {
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                $certPath, [string]::Empty,
                [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
        }
    } catch {
        Write-LogJsonl -Level 'WARN' -Event 'sp_auth_cert_load_fail' -Message "Failed to load certificate: $($_.Exception.Message)"
        return $null
    }

    # --- Build JWT client assertion (RFC 7523) ---
    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Header
    $thumbprintBytes = $cert.GetCertHash()
    $x5t = [Convert]::ToBase64String($thumbprintBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $jwtHeader = @{ alg = 'RS256'; typ = 'JWT'; x5t = $x5t } | ConvertTo-Json -Compress
    $headerB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtHeader)).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    # Payload
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $jwtPayload = @{
        aud = $tokenUrl
        exp = $now + 600
        iss = $clientId
        jti = [guid]::NewGuid().ToString()
        nbf = $now
        sub = $clientId
    } | ConvertTo-Json -Compress
    $payloadB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtPayload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    # Signature
    $dataToSign = [System.Text.Encoding]::UTF8.GetBytes("$headerB64.$payloadB64")
    $rsaKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    $sigBytes = $rsaKey.SignData($dataToSign,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $sigB64 = [Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $clientAssertion = "$headerB64.$payloadB64.$sigB64"

    # --- Request token ---
    $spScope = "https://$($script:SharePointHostname)/.default"
    $body = @{
        grant_type            = 'client_credentials'
        client_id             = $clientId
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = $clientAssertion
        scope                 = $spScope
    }

    Write-LogJsonl -Level 'DEBUG' -Event 'sp_auth_request' -Url $tokenUrl `
        -Message "Requesting SharePoint token (cert auth, thumbprint=$($cert.Thumbprint)) for $spScope"

    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded'
    } catch {
        $errMsg = $_.Exception.Message
        try {
            if ($_.ErrorDetails.Message) {
                $detail = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($detail.error_description) { $errMsg = $detail.error_description }
            }
        } catch { }
        Write-LogJsonl -Level 'WARN' -Event 'sp_auth_failure' -Message "Certificate auth failed: $errMsg"
        return $null
    }

    $accessToken = Get-SafeProp $response 'access_token'
    $expiresIn   = Get-SafeProp $response 'expires_in'
    if (-not $accessToken) { return $null }

    $script:CachedSpToken       = $accessToken
    $script:CachedSpTokenExpiry  = (Get-Date).AddSeconds(($expiresIn ?? 3600))

    Write-LogJsonl -Level 'INFO' -Event 'sp_auth_success' `
        -Message "SharePoint token acquired (cert auth), expires in ${expiresIn}s"

    $cert.Dispose()
    return $script:CachedSpToken
}

function Test-SharePointAccess {
    <#
    .SYNOPSIS
        Test whether the app has SharePoint REST API access via certificate auth.
    #>
    [CmdletBinding()]
    param([string]$SiteWebUrl)

    $spToken = Get-SharePointToken
    if (-not $spToken) { return $false }

    try {
        $headers = @{
            Authorization = "Bearer $spToken"
            Accept        = 'application/json;odata=nometadata'
        }
        $testUri = "$SiteWebUrl/_api/web/title"
        $null = Invoke-RestMethod -Uri $testUri -Headers $headers -Method Get -UseBasicParsing -ErrorAction Stop
        return $true
    } catch {
        $errDetail = ''
        try {
            if ($_.Exception.Response) {
                $stream = $_.Exception.Response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
                $reader = [System.IO.StreamReader]::new($stream)
                $errDetail = $reader.ReadToEnd()
                $reader.Dispose()
            }
        } catch { }
        Write-LogJsonl -Level 'WARN' -Event 'sp_access_test_fail' -Message "SharePoint REST API not accessible: $errDetail"
        return $false
    }
}
