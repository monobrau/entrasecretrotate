# Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Applications modules
# Install with: Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser

Write-Host "Script started."

# --- Configuration ---
# Import specific sub-modules instead of meta-module to avoid dependency issues
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications")

# --- Function Definitions for Module Management ---

Function Test-RequiredModules {
    param($Modules)
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
    param($Modules)
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
    param($Modules)
    Write-Host "Attempting to import required modules: $($Modules -join ', ')..."
    try {
        foreach ($moduleName in $Modules) {
            Write-Host " - Importing '$moduleName'..."
            Import-Module -Name $moduleName -ErrorAction Stop
            Write-Host " - '$moduleName' imported successfully." -ForegroundColor Green
        }
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
    $global:Form.Size = New-Object System.Drawing.Size(850, 600)
    $global:Form.StartPosition = "CenterScreen"
    $global:Form.FormBorderStyle = "FixedSingle" # Prevent resizing
    $global:Form.MaximizeBox = $false

    # Connect Button
    $global:ConnectButton.Location = New-Object System.Drawing.Point(10, 10)
    $global:ConnectButton.Size = New-Object System.Drawing.Size(110, 30)
    $global:ConnectButton.Text = "Connect"
    $global:Form.Controls.Add($global:ConnectButton)

    # Disconnect Button
    $global:DisconnectButton.Location = New-Object System.Drawing.Point(130, 10)
    $global:DisconnectButton.Size = New-Object System.Drawing.Size(110, 30)
    $global:DisconnectButton.Text = "Disconnect"
    $global:DisconnectButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:DisconnectButton)

    # Copy Ticket Note Button
    $global:CopyTicketNoteButton.Location = New-Object System.Drawing.Point(250, 10)
    $global:CopyTicketNoteButton.Size = New-Object System.Drawing.Size(150, 30)
    $global:CopyTicketNoteButton.Text = "Copy Ticket Note"
    $global:CopyTicketNoteButton.Enabled = $true # Always enabled
    $global:Form.Controls.Add($global:CopyTicketNoteButton)

    # Status Label
    $global:StatusLabel.Location = New-Object System.Drawing.Point(10, 45)
    $global:StatusLabel.Size = New-Object System.Drawing.Size(820, 20)
    $global:StatusLabel.Text = "Status: Disconnected"
    $global:Form.Controls.Add($global:StatusLabel)

    # Find Secrets Button
    $global:FindSecretsButton.Location = New-Object System.Drawing.Point(10, 70)
    $global:FindSecretsButton.Size = New-Object System.Drawing.Size(180, 30)
    $global:FindSecretsButton.Text = "Find Expired Secrets"
    $global:FindSecretsButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:FindSecretsButton)

    # Expired Secrets Label
    $global:ExpiredSecretsLabel.Location = New-Object System.Drawing.Point(10, 110)
    $global:ExpiredSecretsLabel.Size = New-Object System.Drawing.Size(300, 20)
    $global:ExpiredSecretsLabel.Text = "Applications with Expired Secrets:"
    $global:Form.Controls.Add($global:ExpiredSecretsLabel)

    # Expired Secrets ListBox
    $global:ExpiredSecretsListBox.Location = New-Object System.Drawing.Point(10, 130)
    $global:ExpiredSecretsListBox.Size = New-Object System.Drawing.Size(820, 150)
    $global:ExpiredSecretsListBox.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:ExpiredSecretsListBox)

    # Selected Secret Label
    $global:SelectedSecretLabel.Location = New-Object System.Drawing.Point(10, 270)
    $global:SelectedSecretLabel.Size = New-Object System.Drawing.Size(150, 20)
    $global:SelectedSecretLabel.Text = "Selected Application:"
    $global:Form.Controls.Add($global:SelectedSecretLabel)

    # Selected App Name Label
    $global:SelectedAppNameLabel.Location = New-Object System.Drawing.Point(170, 270)
    $global:SelectedAppNameLabel.Size = New-Object System.Drawing.Size(400, 20)
    $global:SelectedAppNameLabel.Text = ""
    $global:Form.Controls.Add($global:SelectedAppNameLabel)

    # Selected End Date Label
    $global:SelectedEndDateLabel.Location = New-Object System.Drawing.Point(170, 290)
    $global:SelectedEndDateLabel.Size = New-Object System.Drawing.Size(400, 20)
    $global:SelectedEndDateLabel.Text = ""
    $global:Form.Controls.Add($global:SelectedEndDateLabel)

    # Generate Secret Button
    $global:GenerateSecretButton.Location = New-Object System.Drawing.Point(10, 320)
    $global:GenerateSecretButton.Size = New-Object System.Drawing.Size(150, 30)
    $global:GenerateSecretButton.Text = "Generate New Secret"
    $global:GenerateSecretButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:GenerateSecretButton)

    # New Secret Label
    $global:NewSecretLabel.Location = New-Object System.Drawing.Point(10, 360)
    $global:NewSecretLabel.Size = New-Object System.Drawing.Size(100, 20)
    $global:NewSecretLabel.Text = "New Secret:"
    $global:Form.Controls.Add($global:NewSecretLabel)

    # New Secret TextBox
    $global:NewSecretTextBox.Location = New-Object System.Drawing.Point(120, 357)
    $global:NewSecretTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $global:NewSecretTextBox.ReadOnly = $true # Make it read-only
    $global:Form.Controls.Add($global:NewSecretTextBox)

    # Copy Secret Button
    $global:CopySecretButton.Location = New-Object System.Drawing.Point(580, 355)
    $global:CopySecretButton.Size = New-Object System.Drawing.Size(120, 30)
    $global:CopySecretButton.Text = "Copy Secret"
    $global:CopySecretButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:CopySecretButton)

    # Delete Expired Secret Button
    $global:DeleteSecretButton = New-Object System.Windows.Forms.Button
    $global:DeleteSecretButton.Location = New-Object System.Drawing.Point(170, 320)
    $global:DeleteSecretButton.Size = New-Object System.Drawing.Size(180, 30)
    $global:DeleteSecretButton.Text = "Delete Expired Secret"
    $global:DeleteSecretButton.Enabled = $false
    $global:Form.Controls.Add($global:DeleteSecretButton)

    # Tenant Label (at bottom)
    $global:TenantLabel.Location = New-Object System.Drawing.Point(10, 540)
    $global:TenantLabel.Size = New-Object System.Drawing.Size(820, 20)
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
# (Connect-Tenant, Disconnect-Tenant, Find-ExpiredSecrets, Update-SelectedSecretInfo functions remain the same)

function Connect-Tenant {
    Write-Host "Attempting to connect..."
    $global:StatusLabel.Text = "Status: Connecting..."
    $global:ConnectButton.Enabled = $false
    $global:DisconnectButton.Enabled = $false
    $global:FindSecretsButton.Enabled = $false
    $global:ExpiredSecretsListBox.Enabled = $false
    $global:GenerateSecretButton.Enabled = $false
    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()

    # Required scopes for reading applications and adding secrets
    # Organization.Read.All is needed to get organization display name
    $scopes = "Application.Read.All", "Application.ReadWrite.All", "Organization.Read.All"

    try {
        Write-Host "Calling Connect-MgGraph..."
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        Write-Host "Connect-MgGraph returned."
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
        $global:ExpiredSecretsListBox.Enabled = $true
        $global:GenerateSecretButton.Enabled = $true # Enable generate button after successful connection
        Write-Host "Connection successful."

        # Check for User.EnableDisableAccount.All role/permission

    } catch {
        $global:StatusLabel.Text = "Status: Connection failed - $($_.Exception.Message)"
        $global:ConnectButton.Enabled = $true
        Write-Host "Connection failed: $($_.Exception.Message)"
    }
}

function Disconnect-Tenant {
    Write-Host "Attempting to disconnect..."
    $global:StatusLabel.Text = "Status: Disconnecting..."
    $global:ConnectButton.Enabled = $false
    $global:DisconnectButton.Enabled = $false
    $global:FindSecretsButton.Enabled = $false
    $global:ExpiredSecretsListBox.Enabled = $false
    $global:GenerateSecretButton.Enabled = $false
    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()

    try {
        Write-Host "Calling Disconnect-MgGraph..."
        Disconnect-MgGraph -ErrorAction Stop
        Write-Host "Disconnect-MgGraph returned."
        $global:StatusLabel.Text = "Status: Disconnected"
        $global:TenantLabel.Text = "Tenant: Not Connected"
        $global:TenantLabel.ForeColor = [System.Drawing.Color]::Gray
        $global:ConnectButton.Enabled = $true
        Write-Host "Disconnection successful."
    } catch {
        $global:StatusLabel.Text = "Status: Disconnection failed - $($_.Exception.Message)"
        $global:DisconnectButton.Enabled = $true # Allow retry if disconnect itself fails
        Write-Host "Disconnection failed: $($_.Exception.Message)"
    }
}

function Find-ExpiredSecrets {
    Write-Host "Attempting to find expired secrets..."
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
         $global:StatusLabel.Text = "Status: Not connected. Please connect first."
         Write-Host "Not connected. Cannot find secrets."
         return
    }

    $global:StatusLabel.Text = "Status: Searching for expired secrets..."
    $global:FindSecretsButton.Enabled = $false
    $global:ExpiredSecretsListBox.Enabled = $false
    $global:GenerateSecretButton.Enabled = $false
    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()

    $now = Get-Date

    try {
        Write-Host "Getting applications from Graph..."
        # Get all applications. -All switch handles pagination automatically.
        # Select only DisplayName, Id, and PasswordCredentials to improve performance
        $applications = Get-MgApplication -All -Property DisplayName, Id, PasswordCredentials -ErrorAction Stop
        Write-Host "Finished getting applications. Processing..."

        $expiredCount = 0
        foreach ($app in $applications) {
            if ($app.PasswordCredentials) {
                $expiredSecrets = $app.PasswordCredentials | Where-Object { $_.EndDateTime -lt $now }
                if ($expiredSecrets.Count -gt 0) {
                    $expiredCount++
                    # Store the application object and relevant secret info
                    # Just showing the *first* expired secret's end date in the list for simplicity
                    $oldestExpiredSecretEndDate = $expiredSecrets | Sort-Object EndDateTime | Select-Object -First 1 | Select-Object -ExpandProperty EndDateTime
                    $global:ExpiredApplicationsData += [pscustomobject]@{
                        ApplicationId = $app.Id
                        DisplayName   = $app.DisplayName
                        EndDate       = $oldestExpiredSecretEndDate # Store the oldest expired date
                        # Could potentially store all expired secrets for this app if needed later
                        # ExpiredSecrets = $expiredSecrets
                    }
                    $global:ExpiredSecretsListBox.Items.Add("$($app.DisplayName) (Expired before: $($oldestExpiredSecretEndDate.ToShortDateString()))")
                }
            }
        }

        if ($global:ExpiredApplicationsData.Count -eq 0) {
            $global:StatusLabel.Text = "Status: No applications found with expired secrets."
            Write-Host "No applications found with expired secrets."
        } else {
            $global:StatusLabel.Text = "Status: Found $($global:ExpiredApplicationsData.Count) application(s) with expired secrets."
            Write-Host "Found $($global:ExpiredApplicationsData.Count) application(s) with expired secrets."
        }

    } catch {
        $global:StatusLabel.Text = "Status: Error finding secrets - $($_.Exception.Message)"
        Write-Host "Error finding secrets: $($_.Exception.Message)"
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

    if ($selectedIndex -ge 0 -and $selectedIndex -lt $global:ExpiredApplicationsData.Count) {
        $selectedApp = $global:ExpiredApplicationsData[$selectedIndex]
        $global:SelectedAppNameLabel.Text = $selectedApp.DisplayName
        $global:SelectedEndDateLabel.Text = "Oldest Expired Date: $($selectedApp.EndDate.ToShortDateString())"
        $global:GenerateSecretButton.Enabled = $true
        $global:DeleteSecretButton.Enabled = $true
        Write-Host "Selected application: $($selectedApp.DisplayName)"
    } else {
        $global:DeleteSecretButton.Enabled = $false
        Write-Host "No valid application selected."
    }
}

function Generate-NewSecret {
    Write-Host "Attempting to generate new secret..."
     if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
         $global:StatusLabel.Text = "Status: Not connected. Please connect first."
         Write-Host "Not connected. Cannot generate secret."
         return
    }

    $selectedIndex = $global:ExpiredSecretsListBox.SelectedIndex

    if ($selectedIndex -lt 0 -or $selectedIndex -ge $global:ExpiredApplicationsData.Count) {
        $global:StatusLabel.Text = "Status: Please select an application first."
        Write-Host "No application selected for secret generation."
        return
    }

    $selectedApp = $global:ExpiredApplicationsData[$selectedIndex]
    $appId = $selectedApp.ApplicationId
    $appName = $selectedApp.DisplayName

    $global:StatusLabel.Text = "Status: Generating new secret for '$appName'..."
    $global:GenerateSecretButton.Enabled = $false
    $global:NewSecretTextBox.Text = ""
    $global:CopySecretButton.Enabled = $false # Disable copy button when clearing
    Write-Host "Generating secret for App ID: $appId, Name: $appName"

    # Calculate next year for the DisplayName
    $nextYear = (Get-Date).Year + 1
    $displayName = "SKOUT$nextYear"

    # Detect if -DisplayName is supported
    $supportsDisplayName = ($null -ne (Get-Command Add-MgApplicationPassword | Select-Object -ExpandProperty Parameters | Where-Object { $_.Name -eq 'DisplayName' }))

    try {
        if ($supportsDisplayName) {
            Write-Host "Calling Add-MgApplicationPassword with DisplayName '$displayName'..."
            $newSecret = Add-MgApplicationPassword -ApplicationId $appId -DisplayName $displayName -ErrorAction Stop
        } else {
            Write-Host "Calling Add-MgApplicationPassword without DisplayName (parameter not supported in this module version)..."
            $newSecret = Add-MgApplicationPassword -ApplicationId $appId -ErrorAction Stop
        }
        Write-Host "Add-MgApplicationPassword returned."

        # The actual secret value is in the SecretText property and is only returned NOW
        $secretValue = $newSecret.SecretText

        $global:NewSecretTextBox.Text = $secretValue
        $global:StatusLabel.Text = "Status: New secret generated for '$appName'. COPY IMMEDIATELY!"
        $global:CopySecretButton.Enabled = $true # Enable copy button when secret is generated
        Write-Host "New secret value obtained. Displayed in textbox."

        # Show a popup with a professional summary for ConnectWise PSA ticket
        # Get local time (not UTC) for the timestamp - [DateTime]::Now explicitly returns local time
        $now = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
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
        # Show a custom popup with a read-only textbox, a copy button, a paste screenshot button, and a PictureBox
        $popupForm = New-Object System.Windows.Forms.Form
        $popupForm.Text = "ConnectWise Ticket Note"
        $popupForm.Size = New-Object System.Drawing.Size(600, 600)
        $popupForm.StartPosition = "CenterScreen"
        $popupForm.Topmost = $true

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ReadOnly = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Size = New-Object System.Drawing.Size(560, 200)
        $textBox.Location = New-Object System.Drawing.Point(10, 10)
        $textBox.Text = $ticketNote
        $textBox.Font = New-Object System.Drawing.Font("Consolas", 10)

        $copyButton = New-Object System.Windows.Forms.Button
        $copyButton.Text = "Copy to Clipboard"
        $copyButton.Size = New-Object System.Drawing.Size(150, 30)
        $copyButton.Location = New-Object System.Drawing.Point(10, 220)
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
        $pasteScreenshotButton.Size = New-Object System.Drawing.Size(150, 30)
        $pasteScreenshotButton.Location = New-Object System.Drawing.Point(170, 220)

        $pictureBox = New-Object System.Windows.Forms.PictureBox
        $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pictureBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $pictureBox.Location = New-Object System.Drawing.Point(10, 260)
        $pictureBox.Size = New-Object System.Drawing.Size(560, 280)

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
        $closeButton.Size = New-Object System.Drawing.Size(100, 30)
        $closeButton.Location = New-Object System.Drawing.Point(330, 220)
        $closeButton.Add_Click({ $popupForm.Close() })

        $popupForm.Controls.Add($textBox)
        $popupForm.Controls.Add($copyButton)
        $popupForm.Controls.Add($pasteScreenshotButton)
        $popupForm.Controls.Add($closeButton)
        $popupForm.Controls.Add($pictureBox)
        $popupForm.AcceptButton = $closeButton
        $popupForm.ShowDialog()

        # Re-enable Generate button in case user wants another one,
        # but warn them the previous one is lost if not copied.
        # Or, keep it disabled until a new item is selected. Let's keep it disabled
        # until a new item is selected for safety.
        # $global:GenerateSecretButton.Enabled = $true # Commented out

    } catch {
        $global:StatusLabel.Text = "Status: Error generating secret for '$appName' - $($_.Exception.Message)"
        Write-Host "Error generating secret: $($_.Exception.Message)"
    }
}

function Delete-ExpiredSecret {
    Write-Host "Attempting to delete expired secret..."
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

function Copy-TicketNoteTemplate {
    Write-Host "Generating ticket note template..."
    # Get local time (not UTC) for the timestamp - [DateTime]::Now explicitly returns local time
    $now = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    # Calculate next year for the DisplayName (same logic as in Generate-NewSecret)
    $nextYear = (Get-Date).Year + 1
    $displayName = "SKOUT$nextYear"
    
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


# Setup the GUI form and controls (Call function AFTER variables are declared)
Setup-GUI

Write-Host "Showing the GUI form..."
# Show the form
[void]$global:Form.ShowDialog()
Write-Host "GUI form closed."

Write-Host "Cleaning up resources..."
# Clean up objects when the form is closed
$global:Form.Dispose()
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
$global:CopyTicketNoteButton.Dispose()
$global:TenantLabel.Dispose()
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
