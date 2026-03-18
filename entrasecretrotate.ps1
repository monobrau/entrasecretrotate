# Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Applications modules
# Install with: Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser

Write-Host "Script started."

# --- Configuration ---
# Import specific sub-modules instead of meta-module to avoid dependency issues
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications")

# Graph app sessions from Windows Credential Manager (EOA-GraphApp-{tenantId} format, shared with ExchangeOnlineAnalyzer)
$graphAppCredentialModulePath = Join-Path $PSScriptRoot "Modules\GraphAppCredential.psm1"

# ExchangeOnlineAnalyzer path for Add/Delete app registration scripts (per XOA)
# Default: sibling directory (e.g. c:\git\exchangeonlineanalyzer when entrasecretrotate is in c:\git\entrasecretrotate)
$xoaPath = Join-Path (Split-Path $PSScriptRoot -Parent) "exchangeonlineanalyzer"

# Secret naming configuration
# Customize the display name for new secrets. {YEAR} will be replaced with current year.
$secretDisplayNameTemplate = "skout{YEAR}"

# Ticket note configuration
# Set to $true to show the ticket note popup, $false to disable
$showTicketNotePopup = $true

# Customize the ticket note popup title
$ticketNotePopupTitle = "ConnectWise Ticket Note"

# Barracuda XDR ATR (Automatic Threat Response) permissions for automatic remediation
# Set to $true to also add User.RevokeSessions.All (revoke sessions when blocking users)
$addRevokeSessionsPermission = $true

# Customize the ticket note template. Available placeholders:
# {DISPLAYNAME} - The secret display name
# {DATETIME} - Current date/time
$ticketNoteTemplate = @"
Barracuda XDR O365 Monitoring Integration - Secret Update

Action Summary:
- Logged into Barracuda XDR portal
- Reviewed Microsoft Office 365 integration
- Integration reported that the Entra application secret had expired
- Logged into Microsoft Entra
- Reviewed the secret and confirmed it had expired
- Generated new secret with description: {DISPLAYNAME}
- Implemented new secret in Barracuda XDR portal
- Waited until change propagated
- Tested new secret
- Test was successful
- Saved secret in the portal
- All tasks complete

Date/Time of update: {DATETIME}

What is this?
The secret used for the secure connection between your Microsoft 365 environment and the Barracuda XDR monitoring system has expired. This process updates the expired secret to restore the connection and ensure that the Barracuda XDR system can continue to access and monitor your Microsoft 365 activity, keeping your integration functional and operational.

Why is this needed?
When secrets expire, the integration between Microsoft 365 and Barracuda XDR stops working, which means monitoring and alerting capabilities are disrupted. Updating the expired secret is essential to restore monitoring functionality and ensure that your organization's security monitoring and alerting systems remain operational. This is a critical maintenance task to keep your M365 monitoring functional and protect against cyber threats.
"@

# GUI Layout Constants
$GUI_MARGIN = 10
$GUI_SPACING = 5
$GUI_BUTTON_HEIGHT = 30
$GUI_BUTTON_WIDTH = 110
$GUI_BUTTON_WIDTH_WIDE = 180
$GUI_LABEL_HEIGHT = 20
$GUI_FORM_WIDTH = 850
$GUI_FORM_HEIGHT = 600

# --- Function Definitions for Module Management ---

Function Write-StatusMessage {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $prefix = "[$timestamp]"

    switch ($Type) {
        'Success' { Write-Host "$prefix $Message" -ForegroundColor Green }
        'Warning' { Write-Warning "$prefix $Message" }
        'Error'   { Write-Error "$prefix $Message" }
        default   { Write-Host "$prefix $Message" }
    }
}

Function Test-RequiredModules {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Modules
    )
    Write-Host "Checking for required modules: $($Modules -join ', ')..."
    $missingModules = @()
    foreach ($moduleName in $Modules) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            $missingModules += $moduleName
            Write-Host " - Module '$moduleName' not found." -ForegroundColor Red
        } else {
             Write-Host " - Module '$moduleName' found." -ForegroundColor Green
        }
    }
    return $missingModules
}

Function Install-MissingModules {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Modules
    )
    Write-Host "Attempting to install missing modules: $($Modules -join ', ')..." -ForegroundColor Yellow
    try {
        # Use -Scope CurrentUser as it generally doesn't require admin rights
        Install-Module -Name $Modules -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        Write-Host "Modules installed successfully." -ForegroundColor Green
        # Using Write-Host instead of MessageBox here to keep it simple
        Write-Host "Required modules installed successfully. Please close this PowerShell session and open a new one to run the script again." -ForegroundColor Information
        return $true # Indicate success
    } catch {
        Write-Error "Failed to install modules. Please install them manually: Install-Module -Name $($Modules -join ', ') -Scope CurrentUser"
        # Using Write-Host instead of MessageBox here
        Write-Host "Failed to install modules. Please install them manually from an elevated PowerShell session using:`n`nInstall-Module -Name $($Modules -join ', ') -Scope CurrentUser -Repository PSGallery -Force`n`nThen restart the script." -ForegroundColor Red
        return $false # Indicate failure
    }
}

Function Import-RequiredModules {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Modules
    )
    # Import only Microsoft.Graph.Applications to avoid "Assembly with same name is already loaded".
    # Applications depends on Microsoft.Graph.Authentication and loads it automatically.
    # Importing both explicitly causes a conflict when Applications tries to load Authentication again.
    $moduleToImport = "Microsoft.Graph.Applications"
    Write-Host "Attempting to import required modules (via $moduleToImport, which includes Authentication)..."
    try {
        Write-Host " - Importing '$moduleToImport'..."
        Import-Module -Name $moduleToImport -ErrorAction Stop
        Write-Host " - '$moduleToImport' imported successfully." -ForegroundColor Green
        Write-Host "All required modules imported." -ForegroundColor Green
        return $true # Indicate success
    } catch {
        Write-Error "Failed to import modules. Error: $($_.Exception.Message)"
        # Using Write-Host instead of MessageBox here
        Write-Host "Failed to import required PowerShell modules. Please ensure they are installed correctly and restart the script.`n`nError: $($_.Exception.Message)" -ForegroundColor Red
        return $false # Indicate failure
    }
}

# --- GUI Setup Function (Defined here, called later) ---
function Setup-GUI {
    Write-Host "Setting up GUI..."
    # Form
    $global:Form.Text = "Entra ID Secret Management"
    $global:Form.Size = [System.Drawing.Size]::new([int]$GUI_FORM_WIDTH, [int]$GUI_FORM_HEIGHT)
    $global:Form.StartPosition = "CenterScreen"
    $global:Form.FormBorderStyle = "FixedSingle" # Prevent resizing
    $global:Form.MaximizeBox = $false

    # Calculate positions using constants
    $row1Y = [int]$GUI_MARGIN
    $row2Y = [int]($row1Y + $GUI_BUTTON_HEIGHT + $GUI_SPACING)
    $row3Y = [int]($row2Y + $GUI_LABEL_HEIGHT + $GUI_SPACING)
    $row4Y = [int]($row3Y + $GUI_BUTTON_HEIGHT + $GUI_MARGIN)
    $row5Y = [int]($row4Y + $GUI_LABEL_HEIGHT)
    $listBoxHeight = 150
    $row6Y = [int]($row5Y + $listBoxHeight + $GUI_MARGIN)
    $row7Y = [int]($row6Y + $GUI_LABEL_HEIGHT)
    $row8Y = [int]($row7Y + $GUI_LABEL_HEIGHT + $GUI_MARGIN)
    $row9Y = [int]($row8Y + $GUI_BUTTON_HEIGHT + $GUI_MARGIN)

    # Tenant selector (Graph app sessions from WCM, shared with ExchangeOnlineAnalyzer)
    $tenantLabel = New-Object System.Windows.Forms.Label
    $tenantLabel.Text = "Tenant:"
    $tenantLabel.Location = [System.Drawing.Point]::new([int]$GUI_MARGIN, [int]($row1Y + 6))
    $tenantLabel.Size = [System.Drawing.Size]::new(45, [int]$GUI_LABEL_HEIGHT)
    $global:Form.Controls.Add($tenantLabel)

    $tenantComboX = [int]($GUI_MARGIN + 50)
    $global:TenantComboBox = New-Object System.Windows.Forms.ComboBox
    $global:TenantComboBox.Location = [System.Drawing.Point]::new($tenantComboX, $row1Y)
    $global:TenantComboBox.Size = [System.Drawing.Size]::new(220, 25)
    $global:TenantComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $global:TenantComboBox.TabIndex = 0
    Update-TenantComboBox
    $global:Form.Controls.Add($global:TenantComboBox)

    # Refresh tenants button (reload WCM app sessions after adding via ExchangeOnlineAnalyzer)
    $refreshTenantsBtn = New-Object System.Windows.Forms.Button
    $refreshTenantsBtn.Text = "↻"
    $refreshTenantsBtn.Location = [System.Drawing.Point]::new($tenantComboX + 222, $row1Y)
    $refreshTenantsBtn.Size = [System.Drawing.Size]::new(28, [int]$GUI_BUTTON_HEIGHT)
    $refreshTenantsBtn.Add_Click({ Update-TenantComboBox })
    $global:Form.Controls.Add($refreshTenantsBtn)

    # Connect Button
    $connectX = [int]($tenantComboX + 220 + 28 + $GUI_MARGIN)
    $global:ConnectButton.Location = [System.Drawing.Point]::new($connectX, $row1Y)
    $global:ConnectButton.Size = [System.Drawing.Size]::new([int]$GUI_BUTTON_WIDTH, [int]$GUI_BUTTON_HEIGHT)
    $global:ConnectButton.Text = "Connect"
    $global:Form.Controls.Add($global:ConnectButton)

    # Disconnect Button
    $disconnectX = [int]($connectX + $GUI_BUTTON_WIDTH + $GUI_MARGIN)
    $global:DisconnectButton.Location = [System.Drawing.Point]::new($disconnectX, $row1Y)
    $global:DisconnectButton.Size = [System.Drawing.Size]::new([int]$GUI_BUTTON_WIDTH, [int]$GUI_BUTTON_HEIGHT)
    $global:DisconnectButton.Text = "Disconnect"
    $global:DisconnectButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:DisconnectButton)

    # Add App button (create XOA app registration, save to WCM)
    $addAppX = [int]($disconnectX + $GUI_BUTTON_WIDTH + $GUI_MARGIN)
    $global:AddAppButton = New-Object System.Windows.Forms.Button
    $global:AddAppButton.Location = [System.Drawing.Point]::new($addAppX, $row1Y)
    $global:AddAppButton.Size = [System.Drawing.Size]::new(90, [int]$GUI_BUTTON_HEIGHT)
    $global:AddAppButton.Text = "Add App"
    $global:Form.Controls.Add($global:AddAppButton)

    # Delete App button (remove XOA app registration per tenant)
    $deleteAppX = [int]($addAppX + 90 + $GUI_SPACING)
    $global:DeleteAppButton = New-Object System.Windows.Forms.Button
    $global:DeleteAppButton.Location = [System.Drawing.Point]::new($deleteAppX, $row1Y)
    $global:DeleteAppButton.Size = [System.Drawing.Size]::new(90, [int]$GUI_BUTTON_HEIGHT)
    $global:DeleteAppButton.Text = "Delete App"
    $global:DeleteAppButton.BackColor = [System.Drawing.Color]::FromArgb(198, 40, 40)
    $global:DeleteAppButton.ForeColor = [System.Drawing.Color]::White
    $global:Form.Controls.Add($global:DeleteAppButton)

    # Copy Ticket Note Button
    $copyTicketX = [int]($deleteAppX + 90 + $GUI_MARGIN)
    $global:CopyTicketNoteButton.Location = [System.Drawing.Point]::new($copyTicketX, $row1Y)
    $global:CopyTicketNoteButton.Size = [System.Drawing.Size]::new(150, [int]$GUI_BUTTON_HEIGHT)
    $global:CopyTicketNoteButton.Text = "Copy Ticket Note"
    $global:CopyTicketNoteButton.Enabled = $true # Always enabled
    $global:Form.Controls.Add($global:CopyTicketNoteButton)

    # Status Label
    $statusWidth = [int]($GUI_FORM_WIDTH - (2 * $GUI_MARGIN))
    $global:StatusLabel.Location = [System.Drawing.Point]::new([int]$GUI_MARGIN, $row2Y)
    $global:StatusLabel.Size = [System.Drawing.Size]::new($statusWidth, [int]$GUI_LABEL_HEIGHT)
    $global:StatusLabel.Text = "Status: Disconnected"
    $global:Form.Controls.Add($global:StatusLabel)

    # Find Secrets Button
    $global:FindSecretsButton.Location = [System.Drawing.Point]::new([int]$GUI_MARGIN, $row3Y)
    $global:FindSecretsButton.Size = [System.Drawing.Size]::new([int]$GUI_BUTTON_WIDTH_WIDE, [int]$GUI_BUTTON_HEIGHT)
    $global:FindSecretsButton.Text = "Find Expired Secrets"
    $global:FindSecretsButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:FindSecretsButton)

    # Select Application button (add secret without finding expired first)
    $selectAppX = [int]($GUI_MARGIN + $GUI_BUTTON_WIDTH_WIDE + $GUI_MARGIN)
    $global:SelectAppButton = New-Object System.Windows.Forms.Button
    $global:SelectAppButton.Location = [System.Drawing.Point]::new($selectAppX, $row3Y)
    $global:SelectAppButton.Size = [System.Drawing.Size]::new(140, [int]$GUI_BUTTON_HEIGHT)
    $global:SelectAppButton.Text = "Select Application"
    $global:SelectAppButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:SelectAppButton)

    # Expired Secrets Label
    $global:ExpiredSecretsLabel.Location = [System.Drawing.Point]::new([int]$GUI_MARGIN, $row4Y)
    $global:ExpiredSecretsLabel.Size = [System.Drawing.Size]::new(300, [int]$GUI_LABEL_HEIGHT)
    $global:ExpiredSecretsLabel.Text = "Applications with Expired Secrets:"
    $global:Form.Controls.Add($global:ExpiredSecretsLabel)

    # Expired Secrets ListBox
    $listBoxWidth = [int]($GUI_FORM_WIDTH - (2 * $GUI_MARGIN))
    $global:ExpiredSecretsListBox.Location = [System.Drawing.Point]::new([int]$GUI_MARGIN, $row5Y)
    $global:ExpiredSecretsListBox.Size = [System.Drawing.Size]::new($listBoxWidth, $listBoxHeight)
    $global:ExpiredSecretsListBox.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:ExpiredSecretsListBox)

    # Selected Secret Label
    $global:SelectedSecretLabel.Location = [System.Drawing.Point]::new([int]$GUI_MARGIN, $row6Y)
    $global:SelectedSecretLabel.Size = [System.Drawing.Size]::new(150, [int]$GUI_LABEL_HEIGHT)
    $global:SelectedSecretLabel.Text = "Selected Application:"
    $global:Form.Controls.Add($global:SelectedSecretLabel)

    # Selected App Name Label
    $selectedAppX = [int]($GUI_MARGIN + 160)
    $global:SelectedAppNameLabel.Location = [System.Drawing.Point]::new($selectedAppX, $row6Y)
    $global:SelectedAppNameLabel.Size = [System.Drawing.Size]::new(400, [int]$GUI_LABEL_HEIGHT)
    $global:SelectedAppNameLabel.Text = ""
    $global:Form.Controls.Add($global:SelectedAppNameLabel)

    # Selected End Date Label
    $global:SelectedEndDateLabel.Location = [System.Drawing.Point]::new($selectedAppX, $row7Y)
    $global:SelectedEndDateLabel.Size = [System.Drawing.Size]::new(400, [int]$GUI_LABEL_HEIGHT)
    $global:SelectedEndDateLabel.Text = ""
    $global:Form.Controls.Add($global:SelectedEndDateLabel)

    # Generate Secret Button
    $global:GenerateSecretButton.Location = [System.Drawing.Point]::new([int]$GUI_MARGIN, $row8Y)
    $global:GenerateSecretButton.Size = [System.Drawing.Size]::new(150, [int]$GUI_BUTTON_HEIGHT)
    $global:GenerateSecretButton.Text = "Generate New Secret"
    $global:GenerateSecretButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:GenerateSecretButton)

    # Delete Expired Secret Button
    $deleteButtonX = [int]($GUI_MARGIN + 150 + $GUI_MARGIN)
    $global:DeleteSecretButton = New-Object System.Windows.Forms.Button
    $global:DeleteSecretButton.Location = [System.Drawing.Point]::new($deleteButtonX, $row8Y)
    $global:DeleteSecretButton.Size = [System.Drawing.Size]::new([int]$GUI_BUTTON_WIDTH_WIDE, [int]$GUI_BUTTON_HEIGHT)
    $global:DeleteSecretButton.Text = "Delete Expired Secret"
    $global:DeleteSecretButton.Enabled = $false
    $global:Form.Controls.Add($global:DeleteSecretButton)

    # Add ATR Permissions Button (Barracuda XDR automatic remediation)
    $addAtrButtonX = [int]($deleteButtonX + $GUI_BUTTON_WIDTH_WIDE + $GUI_MARGIN)
    $global:AddAtrPermissionsButton = New-Object System.Windows.Forms.Button
    $global:AddAtrPermissionsButton.Location = [System.Drawing.Point]::new($addAtrButtonX, $row8Y)
    $global:AddAtrPermissionsButton.Size = [System.Drawing.Size]::new(200, [int]$GUI_BUTTON_HEIGHT)
    $global:AddAtrPermissionsButton.Text = "Add ATR Permissions"
    $global:AddAtrPermissionsButton.Enabled = $false
    $global:AddAtrPermissionsButton.ForeColor = [System.Drawing.Color]::DarkBlue
    $global:Form.Controls.Add($global:AddAtrPermissionsButton)

    # New Secret Label
    $global:NewSecretLabel.Location = [System.Drawing.Point]::new([int]$GUI_MARGIN, $row9Y)
    $global:NewSecretLabel.Size = [System.Drawing.Size]::new(100, [int]$GUI_LABEL_HEIGHT)
    $global:NewSecretLabel.Text = "New Secret:"
    $global:Form.Controls.Add($global:NewSecretLabel)

    # New Secret TextBox
    $secretTextX = [int]($GUI_MARGIN + 110)
    $global:NewSecretTextBox.Location = [System.Drawing.Point]::new($secretTextX, [int]($row9Y - 3))
    $global:NewSecretTextBox.Size = [System.Drawing.Size]::new(450, 25)
    $global:NewSecretTextBox.ReadOnly = $true # Make it read-only
    $global:Form.Controls.Add($global:NewSecretTextBox)

    # Copy Secret Button
    $copySecretButtonX = [int]($secretTextX + 450 + $GUI_MARGIN)
    $global:CopySecretButton.Location = [System.Drawing.Point]::new($copySecretButtonX, [int]($row9Y - 3))
    $global:CopySecretButton.Size = [System.Drawing.Size]::new(120, 30)
    $global:CopySecretButton.Text = "Copy Secret"
    $global:CopySecretButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:CopySecretButton)

    # Tenant Label (at bottom)
    $global:TenantLabel.Location = [System.Drawing.Point]::new(10, 540)
    $global:TenantLabel.Size = [System.Drawing.Size]::new(820, 20)
    $global:TenantLabel.Text = "Tenant: Not Connected"
    $global:TenantLabel.ForeColor = [System.Drawing.Color]::Gray
    $global:Form.Controls.Add($global:TenantLabel)

    # --- Event Handlers ---

    # Connect Button Click
    $global:ConnectButton.Add_Click({
        Connect-Tenant
    })

    # Disconnect Button Click
    $global:DisconnectButton.Add_Click({
        Disconnect-Tenant
    })

    # Copy Ticket Note Button Click
    $global:CopyTicketNoteButton.Add_Click({
        Copy-TicketNoteTemplate
    })

    # Find Secrets Button Click
    $global:FindSecretsButton.Add_Click({
        Find-ExpiredSecrets
    })

    # Select Application Button Click
    $global:SelectAppButton.Add_Click({
        Show-SelectApplicationDialog
    })

    # ListBox Selection Change
    $global:ExpiredSecretsListBox.Add_SelectedValueChanged({
        Update-SelectedSecretInfo
    })

    # Generate Secret Button Click
    $global:GenerateSecretButton.Add_Click({
        Generate-NewSecret
    })
    # Delete Secret Button Click
    $global:DeleteSecretButton.Add_Click({ Delete-ExpiredSecret })

    # Add ATR Permissions Button Click
    $global:AddAtrPermissionsButton.Add_Click({ Add-BarracudaXdrPermissions })

    # Add App / Delete App button clicks
    $global:AddAppButton.Add_Click({ Add-XOAAppRegistration })
    $global:DeleteAppButton.Add_Click({ Delete-XOAAppRegistration })

    # Copy Secret Button Click
    $global:CopySecretButton.Add_Click({
        if ($global:NewSecretTextBox.Text -ne "") {
            [System.Windows.Forms.Clipboard]::SetText($global:NewSecretTextBox.Text)
            [System.Windows.Forms.MessageBox]::Show("Secret copied to clipboard!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("No secret to copy. Please generate a secret first.", "No Secret", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    Write-Host "GUI setup complete."
}


# --- Logic Functions ---

function Update-TenantComboBox {
    <#
    .SYNOPSIS
        Populates the tenant dropdown with Graph app sessions from Windows Credential Manager (EOA-GraphApp-* format).
    #>
    if (-not $global:TenantComboBox) { return }
    $global:TenantComboBox.Items.Clear()
    $global:TenantComboBox.Items.Add([pscustomobject]@{ DisplayText = "Interactive (browser)"; TenantId = $null }) | Out-Null
    try {
        if (Test-Path $graphAppCredentialModulePath) {
            Import-Module $graphAppCredentialModulePath -Force -ErrorAction SilentlyContinue
            if (Get-Command Get-WCMTenantListWithNames -ErrorAction SilentlyContinue) {
                $tenantList = Get-WCMTenantListWithNames | Sort-Object -Property DisplayText
                foreach ($t in $tenantList) {
                    $global:TenantComboBox.Items.Add([pscustomobject]@{ DisplayText = $t.DisplayText; TenantId = $t.TenantId }) | Out-Null
                }
            }
        }
    } catch { /* non-fatal */ }
    $global:TenantComboBox.DisplayMember = "DisplayText"
    $global:TenantComboBox.ValueMember = "TenantId"
    if ($global:TenantComboBox.Items.Count -gt 0) {
        $global:TenantComboBox.SelectedIndex = 0
    }
}

function Add-XOAAppRegistration {
    <#
    .SYNOPSIS
        Launches ExchangeOnlineAnalyzer's New-GraphInboxRulesApp.ps1 -SaveToWCM to create app registration and save to WCM.
    #>
    $scriptPath = Join-Path $xoaPath "New-GraphInboxRulesApp.ps1"
    if (-not (Test-Path $scriptPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "ExchangeOnlineAnalyzer script not found.`n`nExpected: $scriptPath`n`nEnsure ExchangeOnlineAnalyzer is installed at: $xoaPath",
            "Add App",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    try {
        $psExe = (Get-Process -Id $PID).Path
        Start-Process $psExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -SaveToWCM" -Wait
        Update-TenantComboBox
        [System.Windows.Forms.MessageBox]::Show("Add App completed. Tenant list refreshed.", "Add App", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to run Add App: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Delete-XOAAppRegistration {
    <#
    .SYNOPSIS
        Shows tenant selection dialog and runs Remove-GraphInboxRulesApp.ps1 for each selected tenant.
    #>
    $scriptPath = Join-Path $xoaPath "Remove-GraphInboxRulesApp.ps1"
    if (-not (Test-Path $scriptPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "ExchangeOnlineAnalyzer script not found.`n`nExpected: $scriptPath`n`nEnsure ExchangeOnlineAnalyzer is installed at: $xoaPath",
            "Delete App",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    $tenantList = @()
    try {
        if (Test-Path $graphAppCredentialModulePath) {
            Import-Module $graphAppCredentialModulePath -Force -ErrorAction SilentlyContinue
        }
        if (Get-Command Get-WCMTenantListWithNames -ErrorAction SilentlyContinue) {
            $tenantList = Get-WCMTenantListWithNames
        }
    } catch {}
    if ($tenantList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No app credentials found in Windows Credential Manager. Nothing to remove.", "Delete App", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selForm = New-Object System.Windows.Forms.Form
    $selForm.Text = "Select Tenant(s) to Remove App From"
    $selForm.Size = [System.Drawing.Size]::new(450, 380)
    $selForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $selForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Select which tenant(s) to remove the Graph app registration from (River Run Security Investigator):"
    $lbl.Location = [System.Drawing.Point]::new(10, 10)
    $lbl.Size = [System.Drawing.Size]::new(410, 35)
    $lbl.AutoSize = $true
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = [System.Drawing.Point]::new(10, 50)
    $clb.Size = [System.Drawing.Size]::new(410, 240)
    $clb.CheckOnClick = $true
    foreach ($t in $tenantList) { [void]$clb.Items.Add($t.DisplayText, $false) }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Remove Selected"
    $btnOk.Location = [System.Drawing.Point]::new(180, 300)
    $btnOk.Size = [System.Drawing.Size]::new(120, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = [System.Drawing.Point]::new(310, 300)
    $btnCancel.Size = [System.Drawing.Size]::new(90, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $selForm.AcceptButton = $btnOk
    $selForm.CancelButton = $btnCancel
    $selForm.Controls.AddRange(@($lbl, $clb, $btnOk, $btnCancel))
    if ($selForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $selected = @()
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($clb.GetItemChecked($i)) { $selected += $tenantList[$i].TenantId }
    }
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No tenants selected.", "Delete App", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Remove app registration for $($selected.Count) tenant(s)? This will delete the app, service principal, and stored credentials.",
        "Confirm Delete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        $psExe = (Get-Process -Id $PID).Path
        foreach ($tid in $selected) {
            Start-Process $psExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -TenantId `"$tid`" -Force" -Wait
        }
        Update-TenantComboBox
        [System.Windows.Forms.MessageBox]::Show("App removal completed for $($selected.Count) tenant(s). Tenant list refreshed.", "Delete App", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to run Delete App: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Connect-Tenant {
    Write-Host "Attempting to connect..."
    $global:StatusLabel.Text = "Status: Connecting..."
    $global:ConnectButton.Enabled = $false
    $global:DisconnectButton.Enabled = $false
    $global:FindSecretsButton.Enabled = $false
    $global:ExpiredSecretsListBox.Enabled = $false
    $global:GenerateSecretButton.Enabled = $false
    $global:AddAtrPermissionsButton.Enabled = $false
    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()
    $global:SelectedApplicationForSecret = $null

    # Required scopes for reading applications, adding secrets, and granting admin consent
    # Organization.Read.All is needed to get organization display name
    # AppRoleAssignment.ReadWrite.All is needed to grant admin consent for application permissions
    $scopes = "Application.Read.All", "Application.ReadWrite.All", "Organization.Read.All", "AppRoleAssignment.ReadWrite.All"

    $selectedTenantId = $null
    if ($global:TenantComboBox -and $global:TenantComboBox.SelectedItem) {
        $sel = $global:TenantComboBox.SelectedItem
        if ($sel.TenantId) { $selectedTenantId = $sel.TenantId }
    }

    try {
        if ($selectedTenantId) {
            # Use Graph app credentials from Windows Credential Manager (EOA-GraphApp-* format)
            Write-StatusMessage "Connecting via app credentials for tenant $selectedTenantId..." -Type Info
            if (Test-Path $graphAppCredentialModulePath) {
                Import-Module $graphAppCredentialModulePath -Force -ErrorAction SilentlyContinue
            }
            $token = $null
            if (Get-Command Get-GraphAppTokenFromWCM -ErrorAction SilentlyContinue) {
                $token = Get-GraphAppTokenFromWCM -TenantId $selectedTenantId
            }
            if (-not $token) {
                throw "No app credentials found for tenant $selectedTenantId in Windows Credential Manager. Add credentials via ExchangeOnlineAnalyzer or use Interactive (browser)."
            }
            $secToken = ConvertTo-SecureString $token -AsPlainText -Force
            Connect-MgGraph -AccessToken $secToken -NoWelcome -ErrorAction Stop
        } else {
            Write-StatusMessage "Connecting to Microsoft Graph (interactive)..." -Type Info
            Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        }
        $context = Get-MgContext
        $tenantId = $context.TenantId
        
        # Get organization name
        try {
            $organization = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($organization -and $organization.DisplayName) {
                $global:TenantLabel.Text = "Tenant: $($organization.DisplayName) ($tenantId)"
                $global:TenantLabel.ForeColor = [System.Drawing.Color]::Black
            } else {
                $global:TenantLabel.Text = "Tenant: $tenantId"
                $global:TenantLabel.ForeColor = [System.Drawing.Color]::Black
            }
        } catch {
            $global:TenantLabel.Text = "Tenant: $tenantId"
            $global:TenantLabel.ForeColor = [System.Drawing.Color]::Black
        }
        
        $global:StatusLabel.Text = "Status: Connected to Tenant ID '$tenantId'"
        $global:DisconnectButton.Enabled = $true
        $global:FindSecretsButton.Enabled = $true
        $global:SelectAppButton.Enabled = $true
        $global:ExpiredSecretsListBox.Enabled = $true
        if ($global:TenantComboBox) { $global:TenantComboBox.Enabled = $false }
        # Note: GenerateSecretButton is enabled only when an application is selected
        Write-StatusMessage "Successfully connected to tenant: $tenantId" -Type Success

    } catch {
        $errorMsg = $_.Exception.Message
        $global:StatusLabel.Text = "Status: Connection failed - $errorMsg"
        $global:ConnectButton.Enabled = $true
        if ($global:TenantComboBox) { $global:TenantComboBox.Enabled = $true }
        Write-StatusMessage "Connection failed: $errorMsg" -Type Error
    }
}

function Disconnect-Tenant {
    Write-StatusMessage "Disconnecting from Microsoft Graph..." -Type Info
    $global:StatusLabel.Text = "Status: Disconnecting..."
    $global:ConnectButton.Enabled = $false
    $global:DisconnectButton.Enabled = $false
    $global:FindSecretsButton.Enabled = $false
    $global:SelectAppButton.Enabled = $false
    $global:ExpiredSecretsListBox.Enabled = $false
    $global:GenerateSecretButton.Enabled = $false
    $global:AddAtrPermissionsButton.Enabled = $false
    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()
    $global:SelectedApplicationForSecret = $null

    try {
        Disconnect-MgGraph -ErrorAction Stop
        $global:StatusLabel.Text = "Status: Disconnected"
        $global:TenantLabel.Text = "Tenant: Not Connected"
        $global:TenantLabel.ForeColor = [System.Drawing.Color]::Gray
        $global:ConnectButton.Enabled = $true
        if ($global:TenantComboBox) { $global:TenantComboBox.Enabled = $true }
        Write-StatusMessage "Successfully disconnected from Microsoft Graph" -Type Success
    } catch {
        $errorMsg = $_.Exception.Message
        $global:StatusLabel.Text = "Status: Disconnection failed - $errorMsg"
        $global:DisconnectButton.Enabled = $true # Allow retry if disconnect itself fails
        if ($global:TenantComboBox) { $global:TenantComboBox.Enabled = $true }
        Write-StatusMessage "Disconnection failed: $errorMsg" -Type Error
    }
}

function Find-ExpiredSecrets {
    Write-StatusMessage "Searching for expired secrets..." -Type Info
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
         $global:StatusLabel.Text = "Status: Not connected. Please connect first."
         Write-StatusMessage "Cannot find secrets: Not connected to Graph" -Type Warning
         return
    }

    $global:StatusLabel.Text = "Status: Searching for expired secrets..."
    $global:FindSecretsButton.Enabled = $false
    $global:ExpiredSecretsListBox.Enabled = $false
    $global:GenerateSecretButton.Enabled = $false
    $global:AddAtrPermissionsButton.Enabled = $false
    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()
    $global:SelectedApplicationForSecret = $null

    $now = Get-Date

    try {
        Write-Host "Getting applications from Graph..."
        # Get all applications. -All switch handles pagination automatically.
        # Select only DisplayName, Id, and PasswordCredentials to improve performance
        $applications = Get-MgApplication -All -Property DisplayName, Id, PasswordCredentials -ErrorAction Stop
        Write-Host "Finished getting applications. Processing..."

        # Use ArrayList for better performance with large datasets
        $expiredAppsList = [System.Collections.ArrayList]::new()
        $listBoxItems = [System.Collections.ArrayList]::new()

        foreach ($app in $applications) {
            if ($app.PasswordCredentials) {
                $expiredSecrets = $app.PasswordCredentials | Where-Object { $_.EndDateTime -lt $now }
                if ($expiredSecrets.Count -gt 0) {
                    # Store the application object and relevant secret info
                    # Just showing the *first* expired secret's end date in the list for simplicity
                    $oldestExpiredSecretEndDate = $expiredSecrets | Sort-Object EndDateTime | Select-Object -First 1 | Select-Object -ExpandProperty EndDateTime
                    [void]$expiredAppsList.Add([pscustomobject]@{
                        ApplicationId = $app.Id
                        DisplayName   = $app.DisplayName
                        EndDate       = $oldestExpiredSecretEndDate # Store the oldest expired date
                    })
                    [void]$listBoxItems.Add("$($app.DisplayName) (Expired before: $($oldestExpiredSecretEndDate.ToShortDateString()))")
                }
            }
        }

        # Convert to array and assign
        $global:ExpiredApplicationsData = $expiredAppsList.ToArray()
        $global:ExpiredSecretsListBox.Items.AddRange($listBoxItems.ToArray())

        if ($global:ExpiredApplicationsData.Count -eq 0) {
            $global:StatusLabel.Text = "Status: No applications found with expired secrets."
            Write-StatusMessage "No applications found with expired secrets" -Type Info
        } else {
            $count = $global:ExpiredApplicationsData.Count
            $global:StatusLabel.Text = "Status: Found $count application(s) with expired secrets."
            Write-StatusMessage "Found $count application(s) with expired secrets" -Type Success
        }

    } catch {
        $errorMsg = $_.Exception.Message
        $global:StatusLabel.Text = "Status: Error finding secrets - $errorMsg"
        Write-StatusMessage "Error finding secrets: $errorMsg" -Type Error
    } finally {
        $global:FindSecretsButton.Enabled = $true
        $global:ExpiredSecretsListBox.Enabled = $true # Enable ListBox even if empty
        Write-Host "Find Expired Secrets process finished."
    }
}

function Update-SelectedSecretInfo {
    Write-Host "Updating selected secret info..."
    $selectedIndex = $global:ExpiredSecretsListBox.SelectedIndex

    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:GenerateSecretButton.Enabled = $false
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false
    $global:DeleteSecretButton.Enabled = $false
    $global:AddAtrPermissionsButton.Enabled = $false

    if ($selectedIndex -ge 0 -and $selectedIndex -lt $global:ExpiredApplicationsData.Count) {
        # Valid selection from Expired Secrets list
        $selectedApp = $global:ExpiredApplicationsData[$selectedIndex]
        $global:SelectedApplicationForSecret = [pscustomobject]@{ ApplicationId = $selectedApp.ApplicationId; DisplayName = $selectedApp.DisplayName; EndDate = $selectedApp.EndDate }
        $global:SelectedAppNameLabel.Text = $selectedApp.DisplayName
        $global:SelectedEndDateLabel.Text = "Oldest Expired Date: $($selectedApp.EndDate.ToShortDateString())"
        $global:GenerateSecretButton.Enabled = $true
        $global:DeleteSecretButton.Enabled = $true
        $global:AddAtrPermissionsButton.Enabled = $true
        Write-Host "Selected application: $($selectedApp.DisplayName)"
    } else {
        # No valid selection from list - preserve selection from Select Application if present
        if ($global:SelectedApplicationForSecret) {
            $global:SelectedAppNameLabel.Text = $global:SelectedApplicationForSecret.DisplayName
            $global:SelectedEndDateLabel.Text = "Selected for secret rotation"
            $global:GenerateSecretButton.Enabled = $true
            $global:AddAtrPermissionsButton.Enabled = $true
        } else {
            $global:SelectedApplicationForSecret = $null
        }
        $global:DeleteSecretButton.Enabled = $false
        Write-Host "No valid application selected from list."
    }
}

function Show-SelectApplicationDialog {
    <#
    .SYNOPSIS
        Browse and select any application to add a new secret, without finding expired secrets first.
    #>
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        $global:StatusLabel.Text = "Status: Not connected. Please connect first."
        [System.Windows.Forms.MessageBox]::Show("Please connect to a tenant first.", "Select Application", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $global:StatusLabel.Text = "Status: Loading applications..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $applications = Get-MgApplication -All -Property DisplayName, Id, AppId -ErrorAction Stop
    } catch {
        $global:StatusLabel.Text = "Status: Error loading applications"
        [System.Windows.Forms.MessageBox]::Show("Failed to load applications: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $global:StatusLabel.Text = "Status: Connected"
    $appList = $applications | Sort-Object DisplayName | ForEach-Object { [pscustomobject]@{ DisplayName = $_.DisplayName; Id = $_.Id; AppId = $_.AppId } }
    if ($appList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No applications found in this tenant.", "Select Application", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selForm = New-Object System.Windows.Forms.Form
    $selForm.Text = "Select Application (Add Secret)"
    $selForm.Size = [System.Drawing.Size]::new(500, 450)
    $selForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $selForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Select an application to add a new secret:"
    $lbl.Location = [System.Drawing.Point]::new(10, 10)
    $lbl.Size = [System.Drawing.Size]::new(460, 20)
    $selForm.Controls.Add($lbl)
    $searchLbl = New-Object System.Windows.Forms.Label
    $searchLbl.Text = "Search:"
    $searchLbl.Location = [System.Drawing.Point]::new(10, 35)
    $searchLbl.Size = [System.Drawing.Size]::new(45, 20)
    $selForm.Controls.Add($searchLbl)
    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = [System.Drawing.Point]::new(60, 33)
    $searchBox.Size = [System.Drawing.Size]::new(410, 20)
    $selForm.Controls.Add($searchBox)
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = [System.Drawing.Point]::new(10, 60)
    $listBox.Size = [System.Drawing.Size]::new(460, 300)
    $listBox.DisplayMember = "DisplayName"
    foreach ($a in $appList) { [void]$listBox.Items.Add($a) }
    $selForm.Controls.Add($listBox)
    $filterScript = {
        $txt = $searchBox.Text.Trim().ToLower()
        $listBox.Items.Clear()
        if ([string]::IsNullOrEmpty($txt)) {
            foreach ($a in $script:allApps) { [void]$listBox.Items.Add($a) }
        } else {
            foreach ($a in $script:allApps) {
                if ($a.DisplayName -and $a.DisplayName.ToLower().Contains($txt)) { [void]$listBox.Items.Add($a) }
            }
        }
    }
    $script:allApps = $appList
    $searchBox.Add_TextChanged($filterScript)
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Select"
    $btnOk.Location = [System.Drawing.Point]::new(200, 370)
    $btnOk.Size = [System.Drawing.Size]::new(90, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = [System.Drawing.Point]::new(300, 370)
    $btnCancel.Size = [System.Drawing.Size]::new(90, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $selForm.AcceptButton = $btnOk
    $selForm.CancelButton = $btnCancel
    $selForm.Controls.AddRange(@($btnOk, $btnCancel))
    $listBox.DoubleClick += { $selForm.DialogResult = [System.Windows.Forms.DialogResult]::OK; $selForm.Close() }
    if ($selForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $selected = $listBox.SelectedItem
    if (-not $selected) {
        [System.Windows.Forms.MessageBox]::Show("Please select an application.", "Select Application", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $global:SelectedApplicationForSecret = [pscustomobject]@{ ApplicationId = $selected.Id; DisplayName = $selected.DisplayName; EndDate = $null }
    $global:SelectedAppNameLabel.Text = $selected.DisplayName
    $global:SelectedEndDateLabel.Text = "Selected for secret rotation"
    $global:GenerateSecretButton.Enabled = $true
    $global:NewSecretTextBox.Text = ""
    $global:DeleteSecretButton.Enabled = $false
    $global:AddAtrPermissionsButton.Enabled = $true
    Write-StatusMessage "Selected application: $($selected.DisplayName)" -Type Info
}

function Generate-NewSecret {
    Write-StatusMessage "Initiating secret generation..." -Type Info
     if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
         $global:StatusLabel.Text = "Status: Not connected. Please connect first."
         Write-StatusMessage "Cannot generate secret: Not connected to Graph" -Type Warning
         return
    }

    $selectedApp = $global:SelectedApplicationForSecret
    if (-not $selectedApp) {
        $selectedIndex = $global:ExpiredSecretsListBox.SelectedIndex
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $global:ExpiredApplicationsData.Count) {
            $selectedApp = $global:ExpiredApplicationsData[$selectedIndex]
        }
    }
    if (-not $selectedApp) {
        $global:StatusLabel.Text = "Status: Please select an application first (use Find Expired Secrets or Select Application)."
        Write-StatusMessage "Cannot generate secret: No application selected" -Type Warning
        return
    }

    $appId = $selectedApp.ApplicationId
    $appName = $selectedApp.DisplayName

    # Validate required data
    if ([string]::IsNullOrWhiteSpace($appId)) {
        $global:StatusLabel.Text = "Status: Error - Invalid application ID"
        Write-StatusMessage "Generate secret failed: Invalid application ID" -Type Error
        return
    }

    $global:StatusLabel.Text = "Status: Generating new secret for '$appName'..."
    $global:GenerateSecretButton.Enabled = $false
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false # Disable copy button when clearing
    Write-StatusMessage "Generating secret for App ID: $appId, Name: $appName" -Type Info

    # Generate display name from template (skoutYYYY - current year)
    $currentYear = (Get-Date).Year
    $displayName = $secretDisplayNameTemplate -replace '\{YEAR\}', $currentYear

    # Create secret with 1-year (365 days) expiration
    $passwordCred = @{
        displayName = $displayName
        endDateTime = (Get-Date).AddDays(365)
    }

    try {
        Write-Host "Calling Add-MgApplicationPassword with displayName '$displayName' and 365-day expiration..."
        $newSecret = Add-MgApplicationPassword -ApplicationId $appId -PasswordCredential $passwordCred -ErrorAction Stop
        Write-Host "Add-MgApplicationPassword returned."

        # The actual secret value is in the SecretText property and is only returned NOW
        $secretValue = $newSecret.SecretText

        $global:NewSecretTextBox.Text = $secretValue
        $global:StatusLabel.Text = "Status: New secret generated for '$appName'. COPY IMMEDIATELY!"
        $global:CopySecretButton.Enabled = $true # Enable copy button when secret is generated
        Write-StatusMessage "New secret generated successfully for '$appName'. Secret displayed in textbox." -Type Success

        # Show a popup with a professional summary for ticketing system (if enabled)
        if ($showTicketNotePopup) {
            $now = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            # Replace placeholders in the ticket note template
            $ticketNote = $ticketNoteTemplate -replace '\{DISPLAYNAME\}', $displayName -replace '\{DATETIME\}', $now

            # Show a custom popup with a read-only textbox, a copy button, a paste screenshot button, and a PictureBox
            $popupForm = New-Object System.Windows.Forms.Form
            $popupForm.Text = $ticketNotePopupTitle
        $popupForm.Size = [System.Drawing.Size]::new(600, 600)
        $popupForm.StartPosition = "CenterScreen"
        $popupForm.Topmost = $true

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ReadOnly = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Size = [System.Drawing.Size]::new(560, 200)
        $textBox.Location = [System.Drawing.Point]::new(10, 10)
        $textBox.Text = $ticketNote
        $textBox.Font = New-Object System.Drawing.Font("Consolas", 10)

        $copyButton = New-Object System.Windows.Forms.Button
        $copyButton.Text = "Copy to Clipboard"
        $copyButton.Size = [System.Drawing.Size]::new(150, 30)
        $copyButton.Location = [System.Drawing.Point]::new(10, 220)
        $copyButton.Add_Click({
            [System.Windows.Forms.Clipboard]::SetText($textBox.Text)
            if ($pictureBox.Image) {
                # Inform the user that the image is now on the clipboard for a second paste
                [System.Windows.Forms.MessageBox]::Show("Text copied! Now click in your document and paste again to insert the screenshot.", "Image Ready to Paste", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                [System.Windows.Forms.Clipboard]::SetImage($pictureBox.Image)
            }
        })

        $pasteScreenshotButton = New-Object System.Windows.Forms.Button
        $pasteScreenshotButton.Text = "Paste Screenshot"
        $pasteScreenshotButton.Size = [System.Drawing.Size]::new(150, 30)
        $pasteScreenshotButton.Location = [System.Drawing.Point]::new(170, 220)

        $pictureBox = New-Object System.Windows.Forms.PictureBox
        $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pictureBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $pictureBox.Location = [System.Drawing.Point]::new(10, 260)
        $pictureBox.Size = [System.Drawing.Size]::new(560, 280)

        $pasteScreenshotButton.Add_Click({
            if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
                $img = [System.Windows.Forms.Clipboard]::GetImage()
                $pictureBox.Image = $img
            } else {
                [System.Windows.Forms.MessageBox]::Show("Clipboard does not contain an image.", "No Image", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        })

        $closeButton = New-Object System.Windows.Forms.Button
        $closeButton.Text = "Close"
        $closeButton.Size = [System.Drawing.Size]::new(100, 30)
        $closeButton.Location = [System.Drawing.Point]::new(330, 220)
        $closeButton.Add_Click({ $popupForm.Close() })

        $popupForm.Controls.Add($textBox)
        $popupForm.Controls.Add($copyButton)
        $popupForm.Controls.Add($pasteScreenshotButton)
        $popupForm.Controls.Add($closeButton)
        $popupForm.Controls.Add($pictureBox)
        $popupForm.AcceptButton = $closeButton
        $popupForm.ShowDialog()
        } # End if ($showTicketNotePopup)

        # Re-enable Generate button in case user wants another one,
        # but warn them the previous one is lost if not copied.
        # Or, keep it disabled until a new item is selected. Let's keep it disabled
        # until a new item is selected for safety.
        # $global:GenerateSecretButton.Enabled = $true # Commented out

    } catch {
        $errorMsg = $_.Exception.Message
        $global:StatusLabel.Text = "Status: Error generating secret for '$appName' - $errorMsg"
        Write-StatusMessage "Error generating secret for '$appName': $errorMsg" -Type Error
    }
}

function Delete-ExpiredSecret {
    Write-StatusMessage "Attempting to delete expired secret..." -Type Info

    # Validate connection state
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show("Not connected to Microsoft Graph. Please connect first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Write-StatusMessage "Delete operation failed: Not connected to Graph" -Type Warning
        return
    }

    $selectedIndex = $global:ExpiredSecretsListBox.SelectedIndex
    if ($selectedIndex -lt 0 -or $selectedIndex -ge $global:ExpiredApplicationsData.Count) {
        [System.Windows.Forms.MessageBox]::Show("Please select an application with an expired secret.", "No Application Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $selectedApp = $global:ExpiredApplicationsData[$selectedIndex]
    $appId = $selectedApp.ApplicationId
    $appName = $selectedApp.DisplayName
    $now = Get-Date
    # Get the full app object to find all expired secrets
    try {
        $app = Get-MgApplication -ApplicationId $appId -ErrorAction Stop
        $expiredSecrets = $app.PasswordCredentials | Where-Object { $_.EndDateTime -lt $now }
        if (-not $expiredSecrets -or $expiredSecrets.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No expired secrets found for this application.", "No Expired Secrets", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        # Delete the oldest expired secret
        $oldestSecret = $expiredSecrets | Sort-Object EndDateTime | Select-Object -First 1
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the oldest expired secret for '" + $appName + "'?\nEnd Date: " + $oldestSecret.EndDateTime.ToString() + "\nKeyId: " + $oldestSecret.KeyId + "", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-MgApplicationPassword -ApplicationId $appId -KeyId $oldestSecret.KeyId -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show("Expired secret deleted successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            # Refresh the expired secrets list
            Find-ExpiredSecrets
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error deleting expired secret: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Add-BarracudaXdrPermissions {
    Write-StatusMessage "Adding Barracuda XDR ATR permissions..." -Type Info

    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show("Not connected to Microsoft Graph. Please connect first.", "Not Connected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Write-StatusMessage "Add permissions failed: Not connected to Graph" -Type Warning
        return
    }

    # Accept selection from either Expired Secrets list or Select Application dialog
    $selectedApp = $global:SelectedApplicationForSecret
    if (-not $selectedApp) {
        $selectedIndex = $global:ExpiredSecretsListBox.SelectedIndex
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $global:ExpiredApplicationsData.Count) {
            $selectedApp = $global:ExpiredApplicationsData[$selectedIndex]
        }
    }
    if (-not $selectedApp) {
        [System.Windows.Forms.MessageBox]::Show("Please select an application first (use Find Expired Secrets or Select Application).", "No Application Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $appId = $selectedApp.ApplicationId
    $appName = $selectedApp.DisplayName

    # Barracuda XDR ATR permissions (Application permissions for automatic remediation)
    # See: https://campus.barracuda.com/product/xdr/doc/319684663/setting-up-atr-for-microsoft-365-cloud/
    $barracudaPermissions = @(
        @{ Id = "741f803b-c850-494e-b5df-cde7c675a1ca"; Type = "Role"; Name = "User.ReadWrite.All" }
        @{ Id = "3011c876-62b7-4ada-afa2-506cbbecc68c"; Type = "Role"; Name = "User.EnableDisableAccount.All" }
    )
    if ($addRevokeSessionsPermission) {
        $barracudaPermissions += @{ Id = "77f3a031-c388-4f99-b373-dc68676a979e"; Type = "Role"; Name = "User.RevokeSessions.All" }
    }

    $confirmMsg = "Add the following API permissions to '$appName' for Barracuda XDR automatic remediation?`n`n" +
        (($barracudaPermissions | ForEach-Object { "  - $($_.Name)" }) -join "`n") +
        "`n`nAdmin consent will be granted automatically via PowerShell."
    $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Add ATR Permissions", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $global:StatusLabel.Text = "Status: Adding ATR permissions to '$appName'..."
        $global:AddAtrPermissionsButton.Enabled = $false

        $app = Get-MgApplication -ApplicationId $appId -Property Id, AppId, RequiredResourceAccess -ErrorAction Stop
        $appClientId = $app.AppId  # Client ID for service principal lookup (appId != Object Id)
        $msGraphAppId = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph

        # Build merged requiredResourceAccess: preserve all existing, add/merge Microsoft Graph permissions
        $existingRra = @()
        if ($app.RequiredResourceAccess) {
            $existingRra = @($app.RequiredResourceAccess)
        }

        $msGraphEntry = $existingRra | Where-Object { $_.ResourceAppId -eq $msGraphAppId } | Select-Object -First 1
        $existingResourceAccess = @()
        if ($msGraphEntry -and $msGraphEntry.ResourceAccess) {
            $existingResourceAccess = @($msGraphEntry.ResourceAccess)
        }

        $barracudaIds = $barracudaPermissions | ForEach-Object { $_.Id }
        $existingIds = $existingResourceAccess | ForEach-Object { $_.Id }
        $toAdd = $barracudaPermissions | Where-Object { $existingIds -notcontains $_.Id }
        # Permissions that exist but as delegated (Scope) - must be application (Role) for Barracuda ATR
        $toFix = $existingResourceAccess | Where-Object { $barracudaIds -contains $_.Id -and $_.Type -eq "Scope" }

        # Add missing permissions or fix delegated->application (only if needed)
        if ($toAdd.Count -gt 0 -or $toFix.Count -gt 0) {
            $newResourceAccess = [System.Collections.Generic.List[object]]::new()
            foreach ($ra in $existingResourceAccess) {
                $permType = $ra.Type
                if ($barracudaIds -contains $ra.Id) { $permType = "Role" }  # Ensure application permission
                $newResourceAccess.Add(@{ Id = $ra.Id; Type = $permType })
            }
            foreach ($p in $toAdd) {
                $newResourceAccess.Add(@{ Id = $p.Id; Type = "Role" })  # Application permission
            }

            $newMsGraphRra = @{
                ResourceAppId  = $msGraphAppId
                ResourceAccess = $newResourceAccess
            }

            $otherRra = $existingRra | Where-Object { $_.ResourceAppId -ne $msGraphAppId }
            $mergedRra = [System.Collections.Generic.List[object]]::new()
            foreach ($rra in $otherRra) {
                $mergedRra.Add($rra)
            }
            $mergedRra.Add($newMsGraphRra)

            Update-MgApplication -ApplicationId $appId -RequiredResourceAccess $mergedRra -ErrorAction Stop
        }

        # Grant admin consent via app role assignments - check ALL Barracuda permissions,
        # not just newly added ones (permissions may have been added before but never consented)
        # Use appClientId (Application Client ID) for SP lookup - NOT appId (Object ID)
        $clientSp = Get-MgServicePrincipal -Filter "appId eq '$appClientId'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $clientSp) {
            # Try direct API lookup (appId is alternate key) - filter can miss SPs in some tenants
            try {
                $clientSp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$appClientId')" -ErrorAction Stop
            } catch { }
        }
        if (-not $clientSp) {
            Write-StatusMessage "Service principal not found for app. Creating it..." -Type Info
            try {
                $clientSp = New-MgServicePrincipal -AppId $appClientId -ErrorAction Stop
                Write-StatusMessage "Service principal created for '$appName'" -Type Success
            } catch {
                if ($_.Exception.Message -like "*already exists*") {
                    $clientSp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$appClientId')" -ErrorAction SilentlyContinue
                }
                if (-not $clientSp) {
                    Write-StatusMessage "Could not create or find service principal: $($_.Exception.Message)" -Type Warning
                }
            }
        }
        $msGraphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $msGraphSp) {
            $msGraphSp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='00000003-0000-0000-c000-000000000000')" -ErrorAction SilentlyContinue
        }
        $consentGranted = @()
        $consentFailed = @()
        # Normalize Id (Invoke-MgGraphRequest returns hashtable with "id", cmdlets return .Id)
        $clientSpId = if ($clientSp) { if ($clientSp.Id) { $clientSp.Id } elseif ($clientSp["id"]) { $clientSp["id"] } else { $clientSp.id } } else { $null }
        $msGraphSpId = if ($msGraphSp) { if ($msGraphSp.Id) { $msGraphSp.Id } elseif ($msGraphSp["id"]) { $msGraphSp["id"] } else { $msGraphSp.id } } else { $null }
        if ($clientSpId -and $msGraphSpId) {
            foreach ($p in $barracudaPermissions) {
                try {
                    $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $clientSpId -ErrorAction SilentlyContinue |
                        Where-Object { $_.AppRoleId -eq $p.Id -and $_.ResourceId -eq $msGraphSpId }
                    if (-not $existing) {
                        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $clientSpId -PrincipalId $clientSpId -ResourceId $msGraphSpId -AppRoleId $p.Id -ErrorAction Stop
                        $consentGranted += $p.Name
                    }
                } catch {
                    $consentFailed += "$($p.Name): $($_.Exception.Message)"
                }
            }
        }

        $tenantName = "your tenant"
        try {
            $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($org -and $org.DisplayName) { $tenantName = $org.DisplayName }
        } catch { }

        $msg = ""
        if ($toAdd.Count -gt 0) {
            $msg = "Successfully added permissions:`n" + (($toAdd | ForEach-Object { $_.Name }) -join "`n")
        }
        if ($consentGranted.Count -gt 0) {
            if ($msg) { $msg += "`n`n" }
            $msg += "Admin consent granted for:`n" + ($consentGranted -join "`n")
        }
        if ($consentGranted.Count -eq 0 -and $toAdd.Count -eq 0 -and $consentFailed.Count -eq 0 -and $clientSpId -and $msGraphSpId) {
            $msg = "All Barracuda XDR ATR permissions are already configured and consented for this application."
        }
        if ($consentFailed.Count -gt 0) {
            if ($msg) { $msg += "`n`n" }
            $msg += "Admin consent failed. Grant manually in portal: App registrations > $appName > API permissions > Grant admin consent for $tenantName.`n" + ($consentFailed -join "`n")
        } elseif (-not $clientSpId -or -not $msGraphSpId) {
            if ($msg) { $msg += "`n`n" }
            $msg += "Could not grant admin consent (service principal not found). Grant manually in portal: App registrations > $appName > API permissions > Grant admin consent for $tenantName."
        }
        if (-not $msg) { $msg = "No changes needed." }
        [System.Windows.Forms.MessageBox]::Show($msg, "Permissions Added", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusSuffix = if ($consentGranted.Count -gt 0) { " Admin consent granted." } elseif ($toAdd.Count -gt 0) { " Grant admin consent in portal if needed." } else { "" }
        $global:StatusLabel.Text = "Status: ATR permissions for '$appName'." + $statusSuffix
        Write-StatusMessage "ATR permissions for '$appName'. Admin consent: $(if ($consentGranted.Count -gt 0) { 'granted' } else { 'grant manually in portal' })" -Type Success

    } catch {
        $errMsg = $_.Exception.Message
        $hint = ""
        if ($errMsg -match "Authorization_RequestDenied|Insufficient privileges") {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx -and $ctx.AuthType -eq "App-only") {
                $hint = @"

The saved app does not have the required permissions (Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All).

Options:
• Use 'Interactive (browser)' in the tenant dropdown—your Global Admin role applies when you sign in interactively.
• Or add those permissions to the saved app in Entra: App registrations > [your app] > API permissions > Add permission > Microsoft Graph > Application permissions, then Grant admin consent.
"@
            } else {
                $hint = "`n`nEnsure your account has Application Administrator (or higher) in Entra. If using a saved app, try 'Interactive (browser)' instead."
            }
        }
        [System.Windows.Forms.MessageBox]::Show("Error adding permissions: $errMsg$hint", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $global:StatusLabel.Text = "Status: Error adding permissions - $errMsg"
        Write-StatusMessage "Error adding ATR permissions: $errMsg" -Type Error
    } finally {
        $global:AddAtrPermissionsButton.Enabled = $true
    }
}

function Copy-TicketNoteTemplate {
    Write-Host "Generating ticket note template..."
    # Get local time (not UTC) for the timestamp - [DateTime]::Now explicitly returns local time
    $now = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    # Display name (same logic as in Generate-NewSecret): skoutYYYY
    $currentYear = (Get-Date).Year
    $displayName = "skout$currentYear"
    
    $ticketNote = @"
Barracuda XDR O365 Monitoring Integration - Secret Update

Action Summary:
- Logged into Barracuda XDR portal
- Reviewed Microsoft Office 365 integration
- Integration reported that the Entra application secret had expired
- Logged into Microsoft Entra
- Reviewed the secret and confirmed it had expired
- Generated new secret with description: $displayName
- Implemented new secret in Barracuda XDR portal
- Waited until change propagated
- Tested new secret
- Test was successful
- Saved secret in the portal
- All tasks complete

Date/Time of update: $now

What is this?
The secret used for the secure connection between your Microsoft 365 environment and the Barracuda XDR monitoring system has expired. This process updates the expired secret to restore the connection and ensure that the Barracuda XDR system can continue to access and monitor your Microsoft 365 activity, keeping your integration functional and operational.

Why is this needed?
When secrets expire, the integration between Microsoft 365 and Barracuda XDR stops working, which means monitoring and alerting capabilities are disrupted. Updating the expired secret is essential to restore monitoring functionality and ensure that your organization's security monitoring and alerting systems remain operational. This is a critical maintenance task to keep your M365 monitoring functional and protect against cyber threats.
"@
    
    try {
        [System.Windows.Forms.Clipboard]::SetText($ticketNote)
        [System.Windows.Forms.MessageBox]::Show("Ticket note template copied to clipboard!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Write-Host "Ticket note template copied to clipboard successfully."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Host "Error copying to clipboard: $($_.Exception.Message)"
    }
}


# --- Main Execution ---

# Check and Install Modules
$missing = Test-RequiredModules -Modules $requiredModules
if ($missing.Count -gt 0) {
    Write-Host "Required PowerShell modules are missing: $($missing -join ', '). Attempting installation..." -ForegroundColor Warning
    # Using Write-Host instead of MessageBox here to keep it simple for this phase
    Write-Host "Please install the missing modules manually from an elevated PowerShell session using:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name $($missing -join ', ') -Scope CurrentUser -Repository PSGallery -Force" -ForegroundColor Yellow
    Write-Host "Then restart the script." -ForegroundColor Yellow
    exit # Always exit if modules are missing
}

# Import Modules (if found)
# Check if all required modules are already imported
$allModulesImported = $true
foreach ($moduleName in $requiredModules) {
    if (Get-Module -Name $moduleName) {
        Write-Host "Module '$moduleName' already imported." -ForegroundColor Green
    } else {
        $allModulesImported = $false
    }
}

if (-not $allModulesImported) {
    # Check if modules are available
    $allAvailable = $true
    foreach ($moduleName in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            $allAvailable = $false
            Write-Host "Module '$moduleName' is not installed. Please install it and try again." -ForegroundColor Red
        }
    }
    
    if ($allAvailable) {
        Write-Host "Attempting to import required modules: $($requiredModules -join ', ')..."
        if (-not (Import-RequiredModules -Modules $requiredModules)) {
            Write-Host "Module import failed. Please resolve the errors above and try again." -ForegroundColor Red
            exit # Exit if import failed
        }
    } else {
        Write-Host "One or more required modules are not installed. Please install them and try again." -ForegroundColor Red
        Write-Host "Install with: Install-Module -Name $($requiredModules -join ', ') -Scope CurrentUser" -ForegroundColor Yellow
        exit
    }
}

# --- Load GUI Assemblies ---
Write-Host "Attempting to load System.Windows.Forms..."
try {
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "System.Windows.Forms loaded."

    Write-Host "Attempting to load System.Drawing..."
    Add-Type -AssemblyName System.Drawing
    Write-Host "System.Drawing loaded."
} catch {
    Write-Host "Error loading .NET types: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Failed to load required .NET components for the GUI. Ensure your PowerShell environment is healthy." -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit # Exit if essential types can't be loaded
}

# --- Global Variables (Declare AFTER Add-Type) ---
# These need to be declared *after* the Add-Type calls for System.Windows.Forms and System.Drawing
$global:Form = New-Object System.Windows.Forms.Form
$global:ConnectButton = New-Object System.Windows.Forms.Button
$global:DisconnectButton = New-Object System.Windows.Forms.Button
$global:StatusLabel = New-Object System.Windows.Forms.Label
$global:FindSecretsButton = New-Object System.Windows.Forms.Button
$global:ExpiredSecretsLabel = New-Object System.Windows.Forms.Label
$global:ExpiredSecretsListBox = New-Object System.Windows.Forms.ListBox
$global:SelectedSecretLabel = New-Object System.Windows.Forms.Label
$global:SelectedAppNameLabel = New-Object System.Windows.Forms.Label
$global:SelectedEndDateLabel = New-Object System.Windows.Forms.Label
$global:GenerateSecretButton = New-Object System.Windows.Forms.Button
$global:NewSecretLabel = New-Object System.Windows.Forms.Label
$global:NewSecretTextBox = New-Object System.Windows.Forms.TextBox
$global:CopyTicketNoteButton = New-Object System.Windows.Forms.Button
$global:CopySecretButton = New-Object System.Windows.Forms.Button
$global:TenantLabel = New-Object System.Windows.Forms.Label
$global:ExpiredApplicationsData = @() # Store application objects with expired secrets
$global:SelectedApplicationForSecret = $null # Set by ExpiredSecretsListBox selection OR Select Application dialog


# Setup the GUI form and controls (Call function AFTER variables are declared)
Setup-GUI

Write-Host "Showing the GUI form..."
# Show the form
[void]$global:Form.ShowDialog()
Write-Host "GUI form closed."

Write-Host "Cleaning up resources..."
# Clean up objects when the form is closed
$global:Form.Dispose()
if ($global:TenantComboBox) { $global:TenantComboBox.Dispose() }
$global:ConnectButton.Dispose()
$global:DisconnectButton.Dispose()
$global:StatusLabel.Dispose()
$global:FindSecretsButton.Dispose()
$global:ExpiredSecretsLabel.Dispose()
$global:ExpiredSecretsListBox.Dispose()
$global:SelectedSecretLabel.Dispose()
$global:SelectedAppNameLabel.Dispose()
$global:SelectedEndDateLabel.Dispose()
$global:GenerateSecretButton.Dispose()
$global:NewSecretLabel.Dispose()
$global:NewSecretTextBox.Dispose()
$global:CopySecretButton.Dispose()
$global:DeleteSecretButton.Dispose()
$global:AddAtrPermissionsButton.Dispose()
$global:CopyTicketNoteButton.Dispose()
$global:TenantLabel.Dispose()
if ($global:AddAppButton) { $global:AddAppButton.Dispose() }
if ($global:DeleteAppButton) { $global:DeleteAppButton.Dispose() }
if ($global:SelectAppButton) { $global:SelectAppButton.Dispose() }
Write-Host "Resources cleaned up."

# Optional: Disconnect on script exit if still connected
Write-Host "Checking if still connected to Graph for cleanup disconnect..."
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host "Disconnected from Graph during cleanup."
} catch {
    Write-Host "Error during cleanup disconnect (might not have been connected): $($_.Exception.Message)"
} # Ignore errors during cleanup disconnect

Write-Host "Script finished."
