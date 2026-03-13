# Entra ID Secret Management

PowerShell GUI for managing expired application secrets in Microsoft Entra ID.

## Features

- Connect to Entra ID via Microsoft Graph
- Find applications with expired secrets
- Generate new secrets (1-year expiration, named `skoutYYYY`)
- Delete expired secrets
- Add Barracuda XDR ATR permissions for automatic remediation
- Copy ticket note template

## Prerequisites

- Windows PowerShell 5.1+
- Microsoft.Graph.Authentication, Microsoft.Graph.Applications
- Permissions: Application.Read.All, Application.ReadWrite.All, Organization.Read.All, AppRoleAssignment.ReadWrite.All

## Install

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser
```

## Run

```powershell
.\entrasecretrotate.ps1
```

1. Click **Connect** and sign in
2. Click **Find Expired Secrets**
3. Select an app, then **Generate New Secret** or **Add ATR Permissions**
4. Copy the secret immediately—it is shown only once

## Configuration

Edit variables at the top of the script:

| Variable | Description |
|----------|-------------|
| `$secretDisplayNameTemplate` | Secret name pattern, `{YEAR}` = current year (default: `skout{YEAR}`) |
| `$showTicketNotePopup` | Show ticket note after generating secret |
| `$addRevokeSessionsPermission` | Include User.RevokeSessions.All for Barracuda ATR |

## Security

- Secrets are displayed once only. Copy and store securely.
- The script adds new secrets; it does not remove old ones automatically (use **Delete Expired Secret** when ready).
