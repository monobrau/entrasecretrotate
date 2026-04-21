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
# First Microsoft Graph /organization error when resolving tenant labels (for UI diagnostics)
$script:LastGraphTenantNameLookupFailure = $null

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

function Set-WCMTenantDisplayName {
    <#
    .SYNOPSIS
        Stores tenant organization display name in WCM (EOA/ESR-GraphApp-{tenantId}-DisplayName) for UI labels.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA'
    )
    if ([string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($DisplayName)) { return }
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $nameTarget = "${credPrefix}${TenantId}-DisplayName"
    $dn = $DisplayName.Trim()
    try {
        if (Get-Module -ListAvailable -Name CredentialManager) {
            Import-Module CredentialManager -ErrorAction Stop
            $nameCred = New-Object PSCredential 'DisplayName', (ConvertTo-SecureString $dn -AsPlainText -Force)
            New-StoredCredential -Target $nameTarget -Credentials $nameCred -ErrorAction Stop | Out-Null
        } else {
            Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$nameTarget", "/user:DisplayName", "/pass:$dn" -Wait -PassThru -WindowStyle Hidden | Out-Null
        }
    } catch { }
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

function Get-GraphTenantNameLookupLastError {
    <#
    .SYNOPSIS
        Returns the first Microsoft Graph error text captured while resolving organization names, or $null.
    #>
    return [string]$script:LastGraphTenantNameLookupFailure
}

function Get-TenantDisplayNameFromWCM {
    param([Parameter(Mandatory = $true)][string]$TenantId, [string]$Prefix = 'EOA')
    $token = Get-GraphAppTokenFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $token) {
        if (-not $script:LastGraphTenantNameLookupFailure) {
            $script:LastGraphTenantNameLookupFailure = 'Could not obtain an OAuth token (check tenant ID and client secret in Windows Credential Manager for this saved app).'
        }
        return $null
    }
    try {
        $headers = @{ Authorization = "Bearer $token" }
        $uri = 'https://graph.microsoft.com/v1.0/organization?$select=displayName,verifiedDomains'
        $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        # value may be a single object or an array after JSON deserialization — do not use .Count on value alone
        $orgs = @($resp.value)
        if ($orgs.Count -lt 1) { return $null }
        $org = $orgs[0]
        $dn = $null
        if ($null -ne $org.displayName -and -not [string]::IsNullOrWhiteSpace([string]$org.displayName)) {
            $dn = [string]$org.displayName
        }
        if ([string]::IsNullOrWhiteSpace($dn) -and $null -ne $org.verifiedDomains) {
            $domains = @($org.verifiedDomains)
            $defObj = $domains | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1
            if ($defObj -and $defObj.name) { $dn = [string]$defObj.name }
            elseif ($domains.Count -gt 0 -and $domains[0].name) { $dn = [string]$domains[0].name }
        }
        if ([string]::IsNullOrWhiteSpace($dn)) { return $null }
        $dn = $dn.Trim()
        Set-WCMTenantDisplayName -TenantId $TenantId -DisplayName $dn -Prefix $Prefix
        return $dn
    } catch {
        $msg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $raw = [string]$_.ErrorDetails.Message
            try {
                $j = $raw | ConvertFrom-Json -ErrorAction Stop
                if ($j.error.message) { $msg = [string]$j.error.message }
                elseif ($j.error_description) { $msg = [string]$j.error_description }
                else { $msg = $raw }
            } catch {
                $msg = $raw
            }
        }
        if (-not $script:LastGraphTenantNameLookupFailure) {
            $script:LastGraphTenantNameLookupFailure = $msg
        }
        return $null
    }
}

function Get-WCMTenantListWithNames {
    <#
    .SYNOPSIS
        Returns WCM tenants with display names for dropdown display, sorted alphabetically by DisplayText.
    .PARAMETER Prefix
        'EOA' or 'ESR'. Omit for EOA.
    .PARAMETER SkipGraphLookup
        If set, does not call Microsoft Graph to resolve organization display name for tenants missing
        a locally stored display name. Use for UI startup to avoid blocking on network/token per tenant.
    .OUTPUTS
        @(@{ TenantId; DisplayName; DisplayText; Source }, ...)
        DisplayText is the human-readable list label only (organization display name). TenantId is always set for selection and removal logic—GUIDs are not shown in DisplayText.
    #>
    param(
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA',
        [switch]$SkipGraphLookup
    )
    $noNameLabel = 'Organization name not loaded'
    $result = @()
    if (-not $SkipGraphLookup) { $script:LastGraphTenantNameLookupFailure = $null }
    $ids = @(Get-WCMTenantIds -Prefix $Prefix) | Sort-Object
    foreach ($tid in $ids) {
        $name = _Get-StoredDisplayName -TenantId $tid -Prefix $Prefix
        if (-not $name -and -not $SkipGraphLookup) {
            $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $Prefix
        }
        if ($name) {
            $displayText = $name.Trim()
        } else {
            $displayText = $noNameLabel
        }
        $result += [pscustomobject]@{ TenantId = $tid; DisplayName = $name; DisplayText = $displayText; Source = $Prefix }
    }
    return $result | Sort-Object -Property @{ Expression = { [string]$_.DisplayText }; Ascending = $true }, @{ Expression = { [string]$_.TenantId }; Ascending = $true }
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
        Set-WCMTenantDisplayName -TenantId $TenantId -DisplayName $TenantDisplayName -Prefix $Prefix
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

function _Get-GraphAppShortTargetsFromCmdKeyList {
    param([Parameter(Mandatory = $true)][string]$NamePrefix)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        $output = cmdkey /list 2>$null
        $text = if ($output -is [string]) { $output } else { [string]::Join([Environment]::NewLine, @($output)) }
        foreach ($line in $text -split '\r?\n') {
            if ($line -notmatch 'Target:\s*(.+)$') { continue }
            $rest = $Matches[1].Trim()
            $short = $null
            if ($rest -match 'target=(.+)$') { $short = $Matches[1].Trim() }
            elseif ($rest.StartsWith($NamePrefix, [StringComparison]::OrdinalIgnoreCase)) { $short = $rest }
            if ($short -and $short.StartsWith($NamePrefix, [StringComparison]::OrdinalIgnoreCase)) {
                [void]$set.Add($short)
            }
        }
    } catch {}
    return @($set)
}

function Get-WCMUnrecognizedGraphAppTargets {
    <#
    .SYNOPSIS
        EOA- or ESR- GraphApp credential targets that do not match GUID or GUID-DisplayName pattern.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $esc = [regex]::Escape($credPrefix)
    $validMain = "^$esc[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$"
    $validDisp = "^$esc[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}-DisplayName$"
    $all = _Get-GraphAppShortTargetsFromCmdKeyList -NamePrefix $credPrefix
    $orphans = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $all) {
        if ($t -notmatch $validMain -and $t -notmatch $validDisp) {
            $orphans.Add($t)
        }
    }
    return @($orphans | Sort-Object)
}

function Remove-WCMGraphCredentialTarget {
    param([Parameter(Mandatory = $true)][string]$TargetName)
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            Remove-StoredCredential -Target $TargetName -ErrorAction SilentlyContinue
        } catch { }
    }
    try {
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$TargetName" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
    } catch { }
}

function Remove-GraphAppCredentialsLocalOnly {
    <#
    .SYNOPSIS
        Removes WCM entries for tenant(s) only; does not change Entra.
    #>
    param(
        [Parameter(Mandatory = $true)][string[]]$TenantId,
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    foreach ($tid in $TenantId) {
        if ([string]::IsNullOrWhiteSpace($tid)) { continue }
        Remove-GraphAppCredentialFromWCM -TenantId $tid.Trim() -Prefix $Prefix
    }
}

function Show-ClearLocalGraphWcmPicker {
    <#
    .SYNOPSIS
        Lists EOA (ExchangeOnlineAnalyzer) and ESR stored Graph app credentials for local-only removal.
    #>
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        Write-Warning "Show-ClearLocalGraphWcmPicker: System.Windows.Forms not available: $($_.Exception.Message)"
        return 0
    }
    $prevWait = [System.Windows.Forms.Application]::UseWaitCursor
    [System.Windows.Forms.Application]::UseWaitCursor = $true
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    $rowList = [System.Collections.ArrayList]::new()
    try {
        foreach ($pfx in @('EOA', 'ESR')) {
            foreach ($t in Get-WCMTenantListWithNames -Prefix $pfx) {
                $label = if ($pfx -eq 'EOA') { "[EOA] $($t.DisplayText)" } else { "[ESR] $($t.DisplayText)" }
                [void]$rowList.Add([pscustomobject]@{ DisplayText = $label; Kind = 'Tenant'; TenantId = $t.TenantId; WcmPrefix = $pfx; OrphanTarget = [string]$null })
            }
            foreach ($o in Get-WCMUnrecognizedGraphAppTargets -Prefix $pfx) {
                [void]$rowList.Add([pscustomobject]@{ DisplayText = "Unrecognized WCM target [$pfx]: $o"; Kind = 'Orphan'; TenantId = [string]$null; WcmPrefix = $pfx; OrphanTarget = $o })
            }
        }
    } finally {
        [System.Windows.Forms.Application]::UseWaitCursor = $prevWait
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
    }
    if ($rowList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No Graph app credentials found in Windows Credential Manager (EOA or ESR).",
            "Clear local credentials",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return 0
    }
    $sorted = @($rowList | Sort-Object -Property DisplayText)
    $selForm = New-Object System.Windows.Forms.Form
    $selForm.Text = "Clear local credentials (this PC only)"
    $selForm.Size = New-Object System.Drawing.Size(540, 420)
    $selForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $selForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Removes entries from Windows Credential Manager only.`r`nDoes NOT delete app registrations in Entra ID.`r`n`r`n[EOA] = ExchangeOnlineAnalyzer app; [ESR] = Entra Secret Rotate. Rows list organization display names only.`r`nIf a row shows Organization name not loaded, run Update App Perms (Organization.Read.All), use Refresh names on the main window, then try again."
    $lbl.Location = New-Object System.Drawing.Point(10, 10)
    $lbl.Size = New-Object System.Drawing.Size(510, 92)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(10, 108)
    $clb.Size = New-Object System.Drawing.Size(510, 212)
    $clb.CheckOnClick = $true
    foreach ($r in $sorted) { [void]$clb.Items.Add($r.DisplayText, $false) }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Remove selected"
    $btnOk.Location = New-Object System.Drawing.Point(210, 330)
    $btnOk.Size = New-Object System.Drawing.Size(140, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(360, 330)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $selForm.AcceptButton = $btnOk
    $selForm.CancelButton = $btnCancel
    $selForm.Controls.AddRange(@($lbl, $clb, $btnOk, $btnCancel))
    if ($selForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return 0 }
    $picked = @()
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($clb.GetItemChecked($i)) { $picked += $sorted[$i] }
    }
    if ($picked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No rows selected.", "Clear local credentials", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return 0
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Remove $($picked.Count) stored credential entry/entries from this PC only?`n`nEntra app registrations will NOT be changed.",
        "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return 0 }
    $removed = 0
    foreach ($p in $picked) {
        if ($p.Kind -eq 'Tenant' -and $p.TenantId) {
            Remove-GraphAppCredentialsLocalOnly -TenantId @($p.TenantId) -Prefix $p.WcmPrefix
            $removed++
        }
        elseif ($p.Kind -eq 'Orphan' -and $p.OrphanTarget) {
            Remove-WCMGraphCredentialTarget -TargetName $p.OrphanTarget
            $removed++
        }
    }
    [System.Windows.Forms.MessageBox]::Show(
        "Removed $removed local credential entry/entries from Windows Credential Manager.",
        "Clear local credentials",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    return $removed
}

Export-ModuleMember -Function Get-GraphAppCredentialFromWCM, Get-GraphAppTokenFromWCM, Get-WCMTenantIds, Get-TenantDisplayNameFromWCM, Get-GraphTenantNameLookupLastError, Get-WCMTenantListWithNames, Set-WCMTenantDisplayName, Save-GraphAppCredentialToWCM, Remove-GraphAppCredentialFromWCM, Get-WCMUnrecognizedGraphAppTargets, Remove-WCMGraphCredentialTarget, Remove-GraphAppCredentialsLocalOnly, Show-ClearLocalGraphWcmPicker
