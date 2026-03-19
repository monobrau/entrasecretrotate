<#
.SYNOPSIS
    Adds Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All to existing XOA app.
.DESCRIPTION
    Updates the XOA app with permissions needed for Entra Secret Rotate (secret rotation + Add ATR).
    Does not recreate the app - only adds missing permissions and grants admin consent.
    Requires: Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All (interactive admin).
    Set $appDisplayName to match your ExchangeOnlineAnalyzer app registration name.
.EXAMPLE
    .\Update-XOAAppPermissions.ps1
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'
# Must match the display name of your XOA app in Entra ID (configure for your environment)
$appDisplayName = 'Exchange Online Analyzer App'
$graphAppId = '00000003-0000-0000-c000-000000000000'

$permsToAdd = @(
    @{ id = '1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9'; name = 'Application.ReadWrite.All' }
    @{ id = '06b708a9-e830-4db3-a914-8e69da51d44f'; name = 'AppRoleAssignment.ReadWrite.All' }
)

$scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
Write-Host "`n=== Update XOA App Permissions ===" -ForegroundColor Cyan
Write-Host "Connecting..." -ForegroundColor Yellow

try {
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
} catch {
    Write-Error "Graph connect failed: $_"
}

$tenantId = (Get-MgContext).TenantId
Write-Host "Connected. Tenant: $tenantId" -ForegroundColor Green

$app = Get-MgApplication -Filter "displayName eq '$appDisplayName'" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $app) {
    Write-Host "`nApp '$appDisplayName' not found. Use Add App (XOA) to create it first, or set `$appDisplayName to match your app." -ForegroundColor Yellow
    exit 1
}

Write-Host "`nFound app: $($app.AppId)" -ForegroundColor Gray

$existingRra = @($app.RequiredResourceAccess)
$msGraphEntry = $existingRra | Where-Object { $_.ResourceAppId -eq $graphAppId } | Select-Object -First 1
$existingResourceAccess = @()
if ($msGraphEntry -and $msGraphEntry.ResourceAccess) {
    $existingResourceAccess = [System.Collections.Generic.List[object]]::new()
    foreach ($ra in @($msGraphEntry.ResourceAccess)) {
        $existingResourceAccess.Add(@{ Id = $ra.Id; Type = $ra.Type })
    }
}

$existingIds = $existingResourceAccess | ForEach-Object { $_.Id }
$toAdd = $permsToAdd | Where-Object { $existingIds -notcontains $_.Id }

if ($toAdd.Count -eq 0) {
    Write-Host "`nApp already has Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All." -ForegroundColor Green
    Write-Host "Granting admin consent if needed..." -ForegroundColor Gray
} else {
    Write-Host "`nAdding permissions: $(($toAdd | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor Yellow
    foreach ($p in $toAdd) {
        $existingResourceAccess.Add(@{ Id = $p.Id; Type = 'Role' })
    }
    $newMsGraphRra = @{ ResourceAppId = $graphAppId; ResourceAccess = $existingResourceAccess }
    $otherRra = $existingRra | Where-Object { $_.ResourceAppId -ne $graphAppId }
    $mergedRra = [System.Collections.Generic.List[object]]::new()
    foreach ($rra in $otherRra) { $mergedRra.Add($rra) }
    $mergedRra.Add($newMsGraphRra)
    Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $mergedRra -ErrorAction Stop
    Write-Host "  Permissions added." -ForegroundColor Green
}

# Grant admin consent for the new permissions
$clientSp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $clientSp) {
    Write-Host "`nCreating service principal..." -ForegroundColor Yellow
    $clientSp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
}
$msGraphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -ErrorAction SilentlyContinue | Select-Object -First 1

$consentGranted = @()
foreach ($p in $permsToAdd) {
    try {
        $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $clientSp.Id -ErrorAction SilentlyContinue |
            Where-Object { $_.AppRoleId -eq $p.Id -and $_.ResourceId -eq $msGraphSp.Id }
        if (-not $existing) {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $clientSp.Id -PrincipalId $clientSp.Id -ResourceId $msGraphSp.Id -AppRoleId $p.Id -ErrorAction Stop | Out-Null
            $consentGranted += $p.Name
            Write-Host "  $($p.Name) - admin consent granted" -ForegroundColor Green
        }
    } catch { Write-Warning "  $($p.Name) - $($_.Exception.Message)" }
}

if ($consentGranted.Count -eq 0 -and $toAdd.Count -eq 0) {
    Write-Host "`nAll permissions already configured." -ForegroundColor Green
} else {
    Write-Host "`n=== Update Complete ===" -ForegroundColor Cyan
}
Write-Host ""
