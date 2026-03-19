<#
.SYNOPSIS
    Store and retrieve Graph app credentials (app-only) in Windows Credential Manager.
.DESCRIPTION
    Target formats: EOA-GraphApp-{tenantId} (ExchangeOnlineAnalyzer), ESR-GraphApp-{tenantId} (Entra Secret Rotate).
    UserName stores "TenantId|ClientId", Password stores ClientSecret.
.NOTES
    Requires: Install-Module CredentialManager
#>

$script:credTargetPrefixEOA = 'EOA-GraphApp-'
$script:credTargetPrefixESR = 'ESR-GraphApp-'

function Get-GraphAppCredentialFromWCM {
    <#
    .SYNOPSIS
        Retrieves Graph app credentials from Windows Credential Manager for a tenant.
    .OUTPUTS
        @{ TenantId; ClientId; ClientSecret } or $null if not found
    .NOTES
        Tries CredentialManager first, falls back to CredRead P/Invoke (for pwsh compatibility).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $target = "$credPrefix$TenantId"

    # Try CredentialManager first (works in Windows PowerShell 5.1)
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            $cred = Get-StoredCredential -Target $target -ErrorAction SilentlyContinue
            if ($cred) {
                $parts = $cred.UserName -split '\|', 2
                if ($parts.Count -ge 2) {
                    return [pscustomobject]@{
                        TenantId     = $parts[0]
                        ClientId     = $parts[1]
                        ClientSecret = $cred.GetNetworkCredential().Password
                    }
                }
            }
        } catch {
            # CredentialManager may fail in pwsh
        }
    }

    # Fallback: CredRead P/Invoke (works in pwsh)
    try {
        $credObj = _ReadCredentialViaCredRead -Target $target
        if (-not $credObj) { return $null }
        $parts = $credObj.UserName -split '\|', 2
        if ($parts.Count -lt 2) { return $null }
        return [pscustomobject]@{
            TenantId     = $parts[0]
            ClientId     = $parts[1]
            ClientSecret = $credObj.CredentialBlob
        }
    } catch {
        return $null
    }
}

function Get-WCMTenantIds {
    <#
    .SYNOPSIS
        Returns tenant IDs that have Graph app credentials stored in Windows Credential Manager.
    .PARAMETER Prefix
        'EOA' (ExchangeOnlineAnalyzer) or 'ESR' (Entra Secret Rotate). Omit for EOA (backward compat).
    .OUTPUTS
        [string[]] Tenant IDs, or @() if none found
    #>
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $tenantIds = @()
    try {
        $output = cmdkey /list 2>$null
        if ($output) {
            $text = $output | Out-String
            $pattern = [regex]::Escape($credPrefix) + '([a-fA-F0-9\-]{36})'
            $m = [regex]::Matches($text, $pattern)
            foreach ($match in $m) {
                if ($match.Success -and $match.Groups[1].Value) {
                    $tid = $match.Groups[1].Value
                    if ($tid -notin $tenantIds) { $tenantIds += $tid }
                }
            }
        }
    } catch {}
    return $tenantIds
}

function _Get-StoredDisplayName {
    param([string]$TenantId, [string]$Prefix = 'EOA')
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $target = "${credPrefix}${TenantId}-DisplayName"
    try {
        if (Get-Module -ListAvailable -Name CredentialManager) {
            Import-Module CredentialManager -ErrorAction Stop
            $c = Get-StoredCredential -Target $target -ErrorAction SilentlyContinue
            if ($c) { return $c.GetNetworkCredential().Password }
        }
        $obj = _ReadCredentialViaCredRead -Target $target
        if ($obj -and $obj.CredentialBlob) { return $obj.CredentialBlob }
    } catch {}
    return $null
}

function Get-TenantDisplayNameFromWCM {
    param([Parameter(Mandatory = $true)][string]$TenantId, [string]$Prefix = 'EOA')
    $token = Get-GraphAppTokenFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $token) { return $null }
    try {
        $headers = @{ Authorization = "Bearer $token" }
        $resp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers -Method Get -ErrorAction Stop
        if ($resp.value -and $resp.value.Count -gt 0 -and $resp.value[0].displayName) {
            return $resp.value[0].displayName
        }
    } catch {}
    return $null
}

function Get-WCMTenantListWithNames {
    <#
    .SYNOPSIS
        Returns WCM tenants with display names for dropdown display, sorted alphabetically by DisplayText.
    .PARAMETER Prefix
        'EOA' or 'ESR'. Omit for EOA.
    .OUTPUTS
        @(@{ TenantId; DisplayName; DisplayText; Source }, ...)
    #>
    param([Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA')
    $result = @()
    $ids = Get-WCMTenantIds -Prefix $Prefix
    $sourceLabel = if ($Prefix -eq 'ESR') { ' (ESR)' } else { '' }
    foreach ($tid in $ids) {
        $name = _Get-StoredDisplayName -TenantId $tid -Prefix $Prefix
        if (-not $name) { $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $Prefix }
        $displayText = if ($name) { "$name$sourceLabel" } else { "$tid$sourceLabel" }
        $result += [pscustomobject]@{ TenantId = $tid; DisplayName = $name; DisplayText = $displayText; Source = $Prefix }
    }
    return $result | Sort-Object -Property DisplayText
}

function Get-GraphAppTokenFromWCM {
    <#
    .SYNOPSIS
        Gets an app-only access token using credentials from WCM. Returns $null if not found
    #>
    param([Parameter(Mandatory = $true)][string]$TenantId, [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA')
    $cred = Get-GraphAppCredentialFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $cred) { return $null }
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $cred.ClientId
        client_secret = $cred.ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    try {
        $resp = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        return $resp.access_token
    } catch {
        return $null
    }
}

function _ReadCredentialViaCredRead {
    param([string]$Target)
    if (-not $Target) { return $null }
    $sig = @'
[DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool CredRead(string target, uint type, int reservedFlag, out IntPtr credentialPtr);

[DllImport("Advapi32.dll", EntryPoint = "CredFree", SetLastError = true)]
public static extern bool CredFree(IntPtr cred);

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct NativeCredential {
    public uint Flags;
    public uint Type;
    public IntPtr TargetName;
    public IntPtr Comment;
    public long LastWritten;
    public uint CredentialBlobSize;
    public IntPtr CredentialBlob;
    public uint Persist;
    public uint AttributeCount;
    public IntPtr Attributes;
    public IntPtr TargetAlias;
    public IntPtr UserName;
}
'@
    try {
        Add-Type -MemberDefinition $sig -Namespace 'EOACredRead' -Name 'Util' -ErrorAction Stop
    } catch {
        if ($_.Exception.Message -notmatch 'already exists') { return $null }
    }
    $ptr = [IntPtr]::Zero
    $ok = [EOACredRead.Util]::CredRead($Target, 1, 0, [ref]$ptr)
    if (-not $ok -or $ptr -eq [IntPtr]::Zero) { return $null }
    try {
        $ncred = [System.Runtime.InteropServices.Marshal]::PtrToStructure($ptr, [EOACredRead.Util+NativeCredential])
        $userName = if ($ncred.UserName -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ncred.UserName) } else { $null }
        $blob = $null
        if ($ncred.CredentialBlob -ne [IntPtr]::Zero -and $ncred.CredentialBlobSize -gt 0) {
            $blob = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ncred.CredentialBlob, [int]$ncred.CredentialBlobSize / 2)
        }
        [EOACredRead.Util]::CredFree($ptr) | Out-Null
        return [pscustomobject]@{ UserName = $userName; CredentialBlob = $blob }
    } catch {
        try { [EOACredRead.Util]::CredFree($ptr) | Out-Null } catch {}
        return $null
    }
}

function Save-GraphAppCredentialToWCM {
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$ClientSecret,
        [Parameter(Mandatory = $false)][string]$TenantDisplayName,
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'ESR'
    )
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $target = "$credPrefix$TenantId"
    $userName = "${TenantId}|${ClientId}"
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            $cred = New-Object PSCredential $userName, (ConvertTo-SecureString $ClientSecret -AsPlainText -Force)
            New-StoredCredential -Target $target -Credentials $cred -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning "CredentialManager failed; falling back to cmdkey. Install CredentialManager for secure storage: Install-Module CredentialManager -Scope CurrentUser"
            Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$target", "/user:$userName", "/pass:$ClientSecret" -Wait -PassThru -WindowStyle Hidden | Out-Null
        }
    } else {
        Write-Warning "CredentialManager not installed. Client secret may be visible in process argv. Install for secure storage: Install-Module CredentialManager -Scope CurrentUser"
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$target", "/user:$userName", "/pass:$ClientSecret" -Wait -PassThru -WindowStyle Hidden | Out-Null
    }
    if ($TenantDisplayName -and -not [string]::IsNullOrWhiteSpace($TenantDisplayName)) {
        $nameTarget = "${credPrefix}${TenantId}-DisplayName"
        try {
            if (Get-Module -ListAvailable -Name CredentialManager) {
                Import-Module CredentialManager -ErrorAction Stop
                $nameCred = New-Object PSCredential 'DisplayName', (ConvertTo-SecureString $TenantDisplayName -AsPlainText -Force)
                New-StoredCredential -Target $nameTarget -Credentials $nameCred -ErrorAction Stop | Out-Null
            } else {
                Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$nameTarget", "/user:DisplayName", "/pass:$TenantDisplayName" -Wait -PassThru -WindowStyle Hidden | Out-Null
            }
        } catch { }
    }
}

function Remove-GraphAppCredentialFromWCM {
    param([Parameter(Mandatory = $true)][string]$TenantId, [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'ESR')
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $target = "$credPrefix$TenantId"
    $nameTarget = "${credPrefix}${TenantId}-DisplayName"
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            Remove-StoredCredential -Target $target -ErrorAction SilentlyContinue
            Remove-StoredCredential -Target $nameTarget -ErrorAction SilentlyContinue
            return
        } catch { }
    }
    try {
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$target" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$nameTarget" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
    } catch { }
}

Export-ModuleMember -Function Get-GraphAppCredentialFromWCM, Get-GraphAppTokenFromWCM, Get-WCMTenantIds, Get-TenantDisplayNameFromWCM, Get-WCMTenantListWithNames, Save-GraphAppCredentialToWCM, Remove-GraphAppCredentialFromWCM
