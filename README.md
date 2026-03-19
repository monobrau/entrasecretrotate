# Entra ID Secret Management

PowerShell GUI for managing expired application secrets in Microsoft Entra ID.

## Features

- Connect to Entra ID via Microsoft Graph (interactive or app credentials from WCM)
- Find applications with expired secrets
- Select any application (without finding expired first)
- Generate new secrets (1-year expiration, named `secretYYYY` by default)
- Delete expired secrets
- Add Barracuda XDR ATR permissions for automatic remediation
- Copy ticket note template

## App Credentials (WCM)

Uses the same app as ExchangeOnlineAnalyzer, stored as `EOA-GraphApp-{tenantId}`. Configure `$appDisplayName` in Update-XOAAppPermissions.ps1 to match your app.

- **Add App** creates the XOA app (via ExchangeOnlineAnalyzer script) and saves to WCM. New installs include Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All for Add ATR.
- **Update App Perms** adds Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All to an *existing* XOA app (for tenants created before this update). Set `$appDisplayName` in Update-XOAAppPermissions.ps1 to match your app's display name in Entra ID.
- **Delete App** removes the XOA app per tenant.

## Prerequisites

- Windows PowerShell 5.1+
- Microsoft.Graph.Authentication, Microsoft.Graph.Applications
- CredentialManager (recommended for Add App—stores secrets without exposing them in process argv)
- Permissions: Application.Read.All, Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All

## Install

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser
Install-Module CredentialManager -Scope CurrentUser
```

## Run

```powershell
.\entrasecretrotate.ps1
```

1. Choose tenant: **Interactive (browser)** or a saved app (ESR/XOA)
2. Click **Connect**
3. Click **Find Expired Secrets** or **Select Application**
4. Select an app, then **Generate New Secret** or **Add ATR Permissions**
5. Copy the secret immediately—it is shown only once

## Configuration

Edit variables at the top of the script:

| Variable | Description |
|----------|-------------|
| `$secretDisplayNameTemplate` | Secret name pattern, `{YEAR}` = current year (default: `secret{YEAR}`) |
| `$showTicketNotePopup` | Show ticket note after generating secret |
| `$addRevokeSessionsPermission` | Include User.RevokeSessions.All for Barracuda ATR |

## Security

- Secrets are displayed once only. Copy and store securely.
- The script adds new secrets; it does not remove old ones automatically (use **Delete Expired Secret** when ready).
