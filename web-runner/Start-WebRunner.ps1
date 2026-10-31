#Requires -Version 7.0
<#
.SYNOPSIS
    Web UI for Entra ID secret rotation and Barracuda XDR ATR permissions.

.DESCRIPTION
    Serves a browser UI (default http://127.0.0.1:8785/) that uses the same Graph app
    registrations as Exchange Online Analyzer (EOA / XOA) via GraphAppCredential.psm1.
    App create/update/delete stays in XOA. This runner only consumes EOA-prefixed creds.

.EXAMPLE
    .\Start-WebRunner.ps1
    .\Start-WebRunner.ps1 -NoBrowser -ListenLab
#>
param(
    [int]$Port = 8785,
    [switch]$NoBrowser,
    [switch]$ListenLab
)

$ErrorActionPreference = 'Stop'
$script:RunnerRoot = $PSScriptRoot
$script:ProjectRoot = Split-Path $PSScriptRoot -Parent
$script:ListenPort = $Port
$script:ListenLab = [bool]$ListenLab
$script:PublicUrl = "http://127.0.0.1:$Port/"
if (-not [string]::IsNullOrWhiteSpace([string]$env:ESR_PUBLIC_URL)) {
    $script:PublicUrl = [string]$env:ESR_PUBLIC_URL
    if (-not $script:PublicUrl.EndsWith('/')) { $script:PublicUrl += '/' }
}
$script:Sessions = @{}
$script:ShutdownRequested = $false
$script:AddRevokeSessions = $true
if ($env:ESR_ADD_REVOKE_SESSIONS -eq '0') { $script:AddRevokeSessions = $false }
$script:SecretDisplayNameTemplate = if ($env:ESR_SECRET_NAME_TEMPLATE) { [string]$env:ESR_SECRET_NAME_TEMPLATE } else { 'secret{YEAR}' }

# Prefer shared XOA GraphAppCredential module (Linux store + WCM)
$xoaRoot = [string]$env:ESR_XOA_ROOT
if ([string]::IsNullOrWhiteSpace($xoaRoot)) {
    $candidate = Join-Path (Split-Path $script:ProjectRoot -Parent) 'exchangeonlineanalyzer'
    if (Test-Path (Join-Path $candidate 'Modules\GraphAppCredential.psm1')) {
        $xoaRoot = $candidate
    } elseif (Test-Path '/opt/exchangeonlineanalyzer/Modules/GraphAppCredential.psm1') {
        $xoaRoot = '/opt/exchangeonlineanalyzer'
    }
}
$graphCredModule = $null
if ($xoaRoot) {
    $p = Join-Path $xoaRoot 'Modules\GraphAppCredential.psm1'
    if (Test-Path -LiteralPath $p) { $graphCredModule = $p }
}
if (-not $graphCredModule) {
    $local = Join-Path $script:ProjectRoot 'Modules\GraphAppCredential.psm1'
    if (Test-Path -LiteralPath $local) { $graphCredModule = $local }
}
if (-not $graphCredModule) {
    throw 'GraphAppCredential.psm1 not found. Set ESR_XOA_ROOT to the ExchangeOnlineAnalyzer install path.'
}
Import-Module $graphCredModule -Force

# Default Linux lab store (same file XOA uses when HOME-based path is empty)
if ([string]::IsNullOrWhiteSpace([string]$env:EOA_CREDENTIAL_STORE) -and -not [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) {
    $labStore = '/var/lib/exchangeonlineanalyzer/.exchangeonlineanalyzer/graph-apps.json'
    if (Test-Path -LiteralPath $labStore) {
        $env:EOA_CREDENTIAL_STORE = $labStore
    }
}

function Write-JsonResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        $Body,
        [int]$StatusCode = 200
    )
    try {
        if (-not $Response.OutputStream.CanWrite) { return }
        $json = $Body | ConvertTo-Json -Depth 10 -Compress
        $bytes = [Text.Encoding]::UTF8.GetBytes($json)
        $Response.StatusCode = $StatusCode
        $Response.ContentType = 'application/json; charset=utf-8'
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } finally {
        try { $Response.OutputStream.Close() } catch { }
        try { $Response.Close() } catch { }
    }
}

function Write-BytesResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [byte[]]$Bytes,
        [string]$ContentType,
        [int]$StatusCode = 200
    )
    try {
        if (-not $Response.OutputStream.CanWrite) { return }
        $Response.StatusCode = $StatusCode
        $Response.ContentType = $ContentType
        $Response.ContentLength64 = $Bytes.Length
        $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    } finally {
        try { $Response.OutputStream.Close() } catch { }
        try { $Response.Close() } catch { }
    }
}

function Get-RequestBodyObject {
    param([System.Net.HttpListenerRequest]$Request)
    if (-not $Request.HasEntityBody) { return $null }
    $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    try {
        $raw = $reader.ReadToEnd()
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return $raw | ConvertFrom-Json
    } finally {
        $reader.Dispose()
    }
}

function Get-SessionId {
    param([System.Net.HttpListenerRequest]$Request)
    $cookie = $Request.Cookies['esr_sid']
    if ($cookie -and $cookie.Value) { return [string]$cookie.Value }
    return $null
}

function Get-OrCreateSession {
    param(
        [System.Net.HttpListenerRequest]$Request,
        [System.Net.HttpListenerResponse]$Response
    )
    $sid = Get-SessionId -Request $Request
    if ($sid -and $script:Sessions.ContainsKey($sid)) {
        return $script:Sessions[$sid]
    }
    $sid = [guid]::NewGuid().ToString('N')
    $session = [ordered]@{
        Id        = $sid
        TenantId  = $null
        Token     = $null
        TokenExp  = [datetime]::MinValue
        OrgName   = $null
    }
    $script:Sessions[$sid] = $session
    $c = New-Object System.Net.Cookie('esr_sid', $sid, '/')
    $c.HttpOnly = $true
    $Response.SetCookie($c)
    return $session
}

function Invoke-GraphGet {
    param(
        [string]$Token,
        [string]$Uri
    )
    $headers = @{ Authorization = "Bearer $Token" }
    return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers -ErrorAction Stop
}

function Invoke-GraphJson {
    param(
        [string]$Token,
        [string]$Method,
        [string]$Uri,
        $Body = $null
    )
    $headers = @{
        Authorization = "Bearer $Token"
        'Content-Type' = 'application/json'
    }
    $params = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = $headers
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 12 -Compress)
    }
    return Invoke-RestMethod @params
}

function Get-GraphPaged {
    param(
        [string]$Token,
        [string]$Uri
    )
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-GraphGet -Token $Token -Uri $next
        if ($page.value) {
            foreach ($v in @($page.value)) { $items.Add($v) }
        }
        $next = $page.'@odata.nextLink'
    }
    return $items
}

function Ensure-SessionToken {
    param($Session)
    if (-not $Session.TenantId) { throw 'Not connected. Select a tenant first.' }
    if ($Session.Token -and $Session.TokenExp -gt (Get-Date).AddMinutes(2)) {
        return $Session.Token
    }
    $token = Get-GraphAppTokenFromWCM -TenantId $Session.TenantId -Prefix EOA
    if (-not $token) { throw "Failed to get Graph token for tenant $($Session.TenantId). Check XOA app registration / admin consent." }
    $Session.Token = $token
    $Session.TokenExp = (Get-Date).AddMinutes(50)
    return $token
}

function Get-SecretDisplayName {
    $year = (Get-Date).Year
    return ($script:SecretDisplayNameTemplate -replace '\{YEAR\}', "$year")
}

function Get-BarracudaAtrPermissions {
    $list = @(
        @{ Id = '741f803b-c850-494e-b5df-cde7c675a1ca'; Name = 'User.ReadWrite.All' }
        @{ Id = '3011c876-62b7-4ada-afa2-506cbbecc68c'; Name = 'User.EnableDisableAccount.All' }
    )
    if ($script:AddRevokeSessions) {
        $list += @{ Id = '77f3a031-c388-4f99-b373-dc68676a979e'; Name = 'User.RevokeSessions.All' }
    }
    return $list
}

function Handle-Api {
    param(
        [System.Net.HttpListenerContext]$Context
    )
    $req = $Context.Request
    $res = $Context.Response
    $path = $req.Url.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($path)) { $path = '/' }
    $method = $req.HttpMethod.ToUpperInvariant()

    try {
        if ($method -eq 'GET' -and ($path -eq '' -or $path -eq '/')) {
            $htmlPath = Join-Path $script:RunnerRoot 'www\index.html'
            $bytes = [System.IO.File]::ReadAllBytes($htmlPath)
            Write-BytesResponse -Response $res -Bytes $bytes -ContentType 'text/html; charset=utf-8'
            return
        }

        if ($method -eq 'GET' -and $path -match '^/(app\.(js|css)|favicon\.svg)$') {
            $file = Join-Path $script:RunnerRoot ("www\" + $Matches[0].TrimStart('/'))
            if (-not (Test-Path -LiteralPath $file)) {
                Write-JsonResponse -Response $res -Body @{ error = 'Not found' } -StatusCode 404
                return
            }
            $bytes = [System.IO.File]::ReadAllBytes($file)
            $ct = switch -Regex ($path) {
                '\.js$' { 'application/javascript; charset=utf-8' }
                '\.css$' { 'text/css; charset=utf-8' }
                '\.svg$' { 'image/svg+xml' }
                default { 'application/octet-stream' }
            }
            Write-BytesResponse -Response $res -Bytes $bytes -ContentType $ct
            return
        }

        if ($method -eq 'GET' -and $path -eq '/api/health') {
            Write-JsonResponse -Response $res -Body @{
                ok      = $true
                version = '1.0.0-web'
                public  = $script:PublicUrl
            }
            return
        }

        if ($method -eq 'GET' -and $path -eq '/api/tenants') {
            $list = @(Get-WCMTenantListWithNames -Prefix EOA -SkipGraphLookup -ErrorAction Stop)
            $rows = foreach ($t in $list) {
                [pscustomobject]@{
                    tenantId    = $t.TenantId
                    displayName = $t.DisplayName
                    displayText = $t.DisplayText
                }
            }
            Write-JsonResponse -Response $res -Body @{ tenants = @($rows) }
            return
        }

        $session = Get-OrCreateSession -Request $req -Response $res

        if ($method -eq 'GET' -and $path -eq '/api/status') {
            Write-JsonResponse -Response $res -Body @{
                connected = [bool]$session.TenantId
                tenantId  = $session.TenantId
                orgName  = $session.OrgName
            }
            return
        }

        if ($method -eq 'POST' -and $path -eq '/api/connect') {
            $body = Get-RequestBodyObject -Request $req
            $tid = [string]$body.tenantId
            if ([string]::IsNullOrWhiteSpace($tid)) { throw 'tenantId is required' }
            $token = Get-GraphAppTokenFromWCM -TenantId $tid -Prefix EOA
            if (-not $token) { throw "No token for tenant $tid. Create/update the Graph app in XOA first." }
            $session.TenantId = $tid
            $session.Token = $token
            $session.TokenExp = (Get-Date).AddMinutes(50)
            $orgName = $null
            try {
                $org = Invoke-GraphGet -Token $token -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName'
                if ($org.value -and $org.value[0].displayName) { $orgName = [string]$org.value[0].displayName }
            } catch { }
            $session.OrgName = $orgName
            Write-JsonResponse -Response $res -Body @{
                ok       = $true
                tenantId = $tid
                orgName = $orgName
            }
            return
        }

        if ($method -eq 'POST' -and $path -eq '/api/disconnect') {
            $session.TenantId = $null
            $session.Token = $null
            $session.TokenExp = [datetime]::MinValue
            $session.OrgName = $null
            Write-JsonResponse -Response $res -Body @{ ok = $true }
            return
        }

        if ($method -eq 'GET' -and $path -eq '/api/apps/expired') {
            $token = Ensure-SessionToken -Session $session
            $apps = Get-GraphPaged -Token $token -Uri 'https://graph.microsoft.com/v1.0/applications?$select=id,displayName,appId,passwordCredentials&$top=999'
            $now = [datetime]::UtcNow
            $expired = [System.Collections.Generic.List[object]]::new()
            foreach ($app in $apps) {
                $creds = @($app.passwordCredentials)
                if ($creds.Count -eq 0) { continue }
                $old = $creds | Where-Object { $_.endDateTime -and ([datetime]$_.endDateTime) -lt $now } | Sort-Object { [datetime]$_.endDateTime }
                if (-not $old -or @($old).Count -eq 0) { continue }
                $oldest = @($old)[0]
                $expired.Add([pscustomobject]@{
                    applicationId = $app.id
                    appId         = $app.appId
                    displayName   = $app.displayName
                    endDate       = ([datetime]$oldest.endDateTime).ToString('o')
                    expiredCount  = @($old).Count
                })
            }
            $sorted = $expired | Sort-Object displayName
            Write-JsonResponse -Response $res -Body @{ apps = @($sorted) }
            return
        }

        if ($method -eq 'GET' -and $path -eq '/api/apps') {
            $token = Ensure-SessionToken -Session $session
            $q = [string]$req.QueryString['q']
            $apps = Get-GraphPaged -Token $token -Uri 'https://graph.microsoft.com/v1.0/applications?$select=id,displayName,appId&$top=999'
            $rows = foreach ($a in ($apps | Sort-Object displayName)) {
                if ($q -and $a.displayName -and ($a.displayName.ToLowerInvariant().IndexOf($q.ToLowerInvariant()) -lt 0)) { continue }
                [pscustomobject]@{
                    applicationId = $a.id
                    appId         = $a.appId
                    displayName   = $a.displayName
                }
            }
            Write-JsonResponse -Response $res -Body @{ apps = @($rows) }
            return
        }

        if ($method -eq 'POST' -and $path -eq '/api/secrets/generate') {
            $token = Ensure-SessionToken -Session $session
            $body = Get-RequestBodyObject -Request $req
            $appId = [string]$body.applicationId
            if ([string]::IsNullOrWhiteSpace($appId)) { throw 'applicationId is required' }
            $displayName = Get-SecretDisplayName
            $payload = @{
                passwordCredential = @{
                    displayName = $displayName
                    endDateTime = (Get-Date).ToUniversalTime().AddYears(1).ToString('o')
                }
            }
            $result = Invoke-GraphJson -Token $token -Method Post -Uri "https://graph.microsoft.com/v1.0/applications/$appId/addPassword" -Body $payload
            Write-JsonResponse -Response $res -Body @{
                ok          = $true
                displayName = $displayName
                secretText  = $result.secretText
                keyId       = $result.keyId
                endDateTime = $result.endDateTime
            }
            return
        }

        if ($method -eq 'POST' -and $path -eq '/api/secrets/delete-expired') {
            $token = Ensure-SessionToken -Session $session
            $body = Get-RequestBodyObject -Request $req
            $appId = [string]$body.applicationId
            if ([string]::IsNullOrWhiteSpace($appId)) { throw 'applicationId is required' }
            $app = Invoke-GraphGet -Token $token -Uri "https://graph.microsoft.com/v1.0/applications/$appId`?`$select=id,displayName,passwordCredentials"
            $now = [datetime]::UtcNow
            $expired = @($app.passwordCredentials) | Where-Object { $_.endDateTime -and ([datetime]$_.endDateTime) -lt $now } | Sort-Object { [datetime]$_.endDateTime }
            if (-not $expired -or $expired.Count -eq 0) { throw 'No expired secrets found for this application.' }
            $oldest = $expired[0]
            $payload = @{ keyId = $oldest.keyId }
            Invoke-GraphJson -Token $token -Method Post -Uri "https://graph.microsoft.com/v1.0/applications/$appId/removePassword" -Body $payload | Out-Null
            Write-JsonResponse -Response $res -Body @{
                ok          = $true
                keyId       = $oldest.keyId
                endDateTime = $oldest.endDateTime
            }
            return
        }

        if ($method -eq 'POST' -and $path -eq '/api/atr/add') {
            $token = Ensure-SessionToken -Session $session
            $body = Get-RequestBodyObject -Request $req
            $appObjectId = [string]$body.applicationId
            if ([string]::IsNullOrWhiteSpace($appObjectId)) { throw 'applicationId is required' }
            $perms = @(Get-BarracudaAtrPermissions)
            $msGraphAppId = '00000003-0000-0000-c000-000000000000'
            $app = Invoke-GraphGet -Token $token -Uri "https://graph.microsoft.com/v1.0/applications/$appObjectId`?`$select=id,appId,displayName,requiredResourceAccess"
            $existingRra = @($app.requiredResourceAccess)
            $msGraphEntry = $existingRra | Where-Object { $_.resourceAppId -eq $msGraphAppId } | Select-Object -First 1
            $existingAccess = @()
            if ($msGraphEntry -and $msGraphEntry.resourceAccess) { $existingAccess = @($msGraphEntry.resourceAccess) }
            $existingIds = @($existingAccess | ForEach-Object { $_.id })
            $barracudaIds = @($perms | ForEach-Object { $_.Id })
            $toAdd = @($perms | Where-Object { $existingIds -notcontains $_.Id })
            $toFix = @($existingAccess | Where-Object { $barracudaIds -contains $_.id -and $_.type -eq 'Scope' })

            if ($toAdd.Count -gt 0 -or $toFix.Count -gt 0) {
                $newAccess = [System.Collections.Generic.List[object]]::new()
                foreach ($ra in $existingAccess) {
                    $ptype = $ra.type
                    if ($barracudaIds -contains $ra.id) { $ptype = 'Role' }
                    $newAccess.Add(@{ id = $ra.id; type = $ptype })
                }
                foreach ($p in $toAdd) {
                    $newAccess.Add(@{ id = $p.Id; type = 'Role' })
                }
                $merged = [System.Collections.Generic.List[object]]::new()
                foreach ($rra in $existingRra) {
                    if ($rra.resourceAppId -eq $msGraphAppId) { continue }
                    $merged.Add($rra)
                }
                $merged.Add(@{
                    resourceAppId  = $msGraphAppId
                    resourceAccess = @($newAccess)
                })
                Invoke-GraphJson -Token $token -Method Patch -Uri "https://graph.microsoft.com/v1.0/applications/$appObjectId" -Body @{
                    requiredResourceAccess = @($merged)
                } | Out-Null
            }

            $consentGranted = [System.Collections.Generic.List[string]]::new()
            $consentFailed = [System.Collections.Generic.List[string]]::new()
            $appClientId = [string]$app.appId
            $clientSp = $null
            try {
                $spPage = Invoke-GraphGet -Token $token -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$appClientId'"
                if ($spPage.value -and $spPage.value.Count -gt 0) { $clientSp = $spPage.value[0] }
            } catch { }
            $msGraphSp = $null
            try {
                $gPage = Invoke-GraphGet -Token $token -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$msGraphAppId'"
                if ($gPage.value -and $gPage.value.Count -gt 0) { $msGraphSp = $gPage.value[0] }
            } catch { }

            if ($clientSp -and $msGraphSp) {
                $existingAssign = @()
                try {
                    $asg = Get-GraphPaged -Token $token -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($clientSp.id)/appRoleAssignments"
                    $existingAssign = @($asg | ForEach-Object { $_.appRoleId })
                } catch { }
                foreach ($p in $perms) {
                    if ($existingAssign -contains $p.Id) {
                        $consentGranted.Add("$($p.Name) (already consented)")
                        continue
                    }
                    try {
                        Invoke-GraphJson -Token $token -Method Post -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($clientSp.id)/appRoleAssignments" -Body @{
                            principalId = $clientSp.id
                            resourceId  = $msGraphSp.id
                            appRoleId   = $p.Id
                        } | Out-Null
                        $consentGranted.Add($p.Name)
                    } catch {
                        $consentFailed.Add("$($p.Name): $($_.Exception.Message)")
                    }
                }
            } else {
                $consentFailed.Add('Service principal not found; grant admin consent in Entra portal.')
            }

            Write-JsonResponse -Response $res -Body @{
                ok             = $true
                appName        = $app.displayName
                permissionsAdded = @($toAdd | ForEach-Object { $_.Name })
                consentGranted = @($consentGranted)
                consentFailed  = @($consentFailed)
            }
            return
        }

        Write-JsonResponse -Response $res -Body @{ error = "Not found: $method $path" } -StatusCode 404
    } catch {
        $msg = $_.Exception.Message
        Write-JsonResponse -Response $res -Body @{ error = $msg } -StatusCode 400
    }
}

# --- Listen ---
$prefix = if ($script:ListenLab) { "http://+:$Port/" } else { "http://127.0.0.1:$Port/" }
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
try {
    $listener.Start()
} catch {
    throw "Failed to bind $prefix. On Linux lab, use -ListenLab and ensure nothing else owns port $Port. $_"
}

Write-Host "Entra Secret Rotate web runner listening on $prefix"
Write-Host "Public URL: $script:PublicUrl"
Write-Host "GraphAppCredential: $graphCredModule"
Write-Host "EOA_CREDENTIAL_STORE=$([string]$env:EOA_CREDENTIAL_STORE)"

if (-not $NoBrowser -and -not $script:ListenLab) {
    try { Start-Process $script:PublicUrl } catch { }
}

while (-not $script:ShutdownRequested) {
    $context = $null
    try {
        $async = $listener.BeginGetContext($null, $null)
        while (-not $async.IsCompleted -and -not $script:ShutdownRequested) {
            Start-Sleep -Milliseconds 200
        }
        if ($script:ShutdownRequested) { break }
        $context = $listener.EndGetContext($async)
        Handle-Api -Context $context
    } catch {
        if ($script:ShutdownRequested) { break }
        Write-Warning "Request error: $($_.Exception.Message)"
        if ($context) {
            try { Write-JsonResponse -Response $context.Response -Body @{ error = $_.Exception.Message } -StatusCode 500 } catch { }
        }
    }
}

try { $listener.Stop() } catch { }
try { $listener.Close() } catch { }
Write-Host 'Stopped.'
