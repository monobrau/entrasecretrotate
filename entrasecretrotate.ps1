# Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Applications modules
# Install with: Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser

Write-Host "Script started."

# --- Configuration ---
# Import specific sub-modules instead of meta-module to avoid dependency issues
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications")

# Secret naming configuration
# Customize the display name for new secrets. {YEAR} will be replaced with next year.
$secretDisplayNameTemplate = "SKOUT{YEAR}"

# Ticket note configuration
# Set to $true to show the ticket note popup, $false to disable
$showTicketNotePopup = $true

# Customize the ticket note popup title
$ticketNotePopupTitle = "ConnectWise Ticket Note"

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
This process updates the secure connection between your Microsoft 365 environment and the Barracuda XDR monitoring system. By rotating (changing) the secret, we ensure that only authorized systems can access and monitor your Microsoft 365 activity, keeping your integration healthy and up to date.

Why is this needed?
Regularly updating these secrets is a best practice for security. It helps prevent unauthorized access by making sure old credentials cannot be used if they are ever exposed. This approach strengthens your organization's protection against cyber threats and ensures that your monitoring and alerting systems remain reliable.
"@

# GUI Layout Constants
$GUI_MARGIN = 10
$GUI_SPACING = 5
$GUI_BUTTON_HEIGHT = 30
$GUI_BUTTON_WIDTH = 110
$GUI_BUTTON_WIDTH_WIDE = 180
$GUI_LABEL_HEIGHT = 20
$GUI_FORM_WIDTH = 850
$GUI_FORM_HEIGHT = 550

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
    $global:Form.Size = New-Object System.Drawing.Size($GUI_FORM_WIDTH, $GUI_FORM_HEIGHT)
    $global:Form.StartPosition = "CenterScreen"
    $global:Form.FormBorderStyle = "FixedSingle" # Prevent resizing
    $global:Form.MaximizeBox = $false

    # Calculate positions using constants
    $row1Y = $GUI_MARGIN
    $row2Y = $row1Y + $GUI_BUTTON_HEIGHT + $GUI_SPACING
    $row3Y = $row2Y + $GUI_LABEL_HEIGHT + $GUI_SPACING
    $row4Y = $row3Y + $GUI_BUTTON_HEIGHT + $GUI_MARGIN
    $row5Y = $row4Y + $GUI_LABEL_HEIGHT
    $listBoxHeight = 150
    $row6Y = $row5Y + $listBoxHeight + $GUI_MARGIN
    $row7Y = $row6Y + $GUI_LABEL_HEIGHT
    $row8Y = $row7Y + $GUI_LABEL_HEIGHT + $GUI_MARGIN
    $row9Y = $row8Y + $GUI_BUTTON_HEIGHT + $GUI_MARGIN

    # Connect Button
    $global:ConnectButton.Location = New-Object System.Drawing.Point($GUI_MARGIN, $row1Y)
    $global:ConnectButton.Size = New-Object System.Drawing.Size($GUI_BUTTON_WIDTH, $GUI_BUTTON_HEIGHT)
    $global:ConnectButton.Text = "Connect"
    $global:Form.Controls.Add($global:ConnectButton)

    # Disconnect Button
    $disconnectX = $GUI_MARGIN + $GUI_BUTTON_WIDTH + $GUI_MARGIN
    $global:DisconnectButton.Location = New-Object System.Drawing.Point($disconnectX, $row1Y)
    $global:DisconnectButton.Size = New-Object System.Drawing.Size($GUI_BUTTON_WIDTH, $GUI_BUTTON_HEIGHT)
    $global:DisconnectButton.Text = "Disconnect"
    $global:DisconnectButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:DisconnectButton)

    # Status Label
    $statusWidth = $GUI_FORM_WIDTH - (2 * $GUI_MARGIN)
    $global:StatusLabel.Location = New-Object System.Drawing.Point($GUI_MARGIN, $row2Y)
    $global:StatusLabel.Size = New-Object System.Drawing.Size($statusWidth, $GUI_LABEL_HEIGHT)
    $global:StatusLabel.Text = "Status: Disconnected"
    $global:Form.Controls.Add($global:StatusLabel)

    # Find Secrets Button
    $global:FindSecretsButton.Location = New-Object System.Drawing.Point($GUI_MARGIN, $row3Y)
    $global:FindSecretsButton.Size = New-Object System.Drawing.Size($GUI_BUTTON_WIDTH_WIDE, $GUI_BUTTON_HEIGHT)
    $global:FindSecretsButton.Text = "Find Expired Secrets"
    $global:FindSecretsButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:FindSecretsButton)

    # Expired Secrets Label
    $global:ExpiredSecretsLabel.Location = New-Object System.Drawing.Point($GUI_MARGIN, $row4Y)
    $global:ExpiredSecretsLabel.Size = New-Object System.Drawing.Size(300, $GUI_LABEL_HEIGHT)
    $global:ExpiredSecretsLabel.Text = "Applications with Expired Secrets:"
    $global:Form.Controls.Add($global:ExpiredSecretsLabel)

    # Expired Secrets ListBox
    $listBoxWidth = $GUI_FORM_WIDTH - (2 * $GUI_MARGIN)
    $global:ExpiredSecretsListBox.Location = New-Object System.Drawing.Point($GUI_MARGIN, $row5Y)
    $global:ExpiredSecretsListBox.Size = New-Object System.Drawing.Size($listBoxWidth, $listBoxHeight)
    $global:ExpiredSecretsListBox.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:ExpiredSecretsListBox)

    # Selected Secret Label
    $global:SelectedSecretLabel.Location = New-Object System.Drawing.Point($GUI_MARGIN, $row6Y)
    $global:SelectedSecretLabel.Size = New-Object System.Drawing.Size(150, $GUI_LABEL_HEIGHT)
    $global:SelectedSecretLabel.Text = "Selected Application:"
    $global:Form.Controls.Add($global:SelectedSecretLabel)

    # Selected App Name Label
    $selectedAppX = $GUI_MARGIN + 160
    $global:SelectedAppNameLabel.Location = New-Object System.Drawing.Point($selectedAppX, $row6Y)
    $global:SelectedAppNameLabel.Size = New-Object System.Drawing.Size(400, $GUI_LABEL_HEIGHT)
    $global:SelectedAppNameLabel.Text = ""
    $global:Form.Controls.Add($global:SelectedAppNameLabel)

    # Selected End Date Label
    $global:SelectedEndDateLabel.Location = New-Object System.Drawing.Point($selectedAppX, $row7Y)
    $global:SelectedEndDateLabel.Size = New-Object System.Drawing.Size(400, $GUI_LABEL_HEIGHT)
    $global:SelectedEndDateLabel.Text = ""
    $global:Form.Controls.Add($global:SelectedEndDateLabel)

    # Generate Secret Button
    $global:GenerateSecretButton.Location = New-Object System.Drawing.Point($GUI_MARGIN, $row8Y)
    $global:GenerateSecretButton.Size = New-Object System.Drawing.Size(150, $GUI_BUTTON_HEIGHT)
    $global:GenerateSecretButton.Text = "Generate New Secret"
    $global:GenerateSecretButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:GenerateSecretButton)

    # Delete Expired Secret Button
    $deleteButtonX = $GUI_MARGIN + 150 + $GUI_MARGIN
    $global:DeleteSecretButton = New-Object System.Windows.Forms.Button
    $global:DeleteSecretButton.Location = New-Object System.Drawing.Point($deleteButtonX, $row8Y)
    $global:DeleteSecretButton.Size = New-Object System.Drawing.Size($GUI_BUTTON_WIDTH_WIDE, $GUI_BUTTON_HEIGHT)
    $global:DeleteSecretButton.Text = "Delete Expired Secret"
    $global:DeleteSecretButton.Enabled = $false
    $global:Form.Controls.Add($global:DeleteSecretButton)

    # New Secret Label
    $global:NewSecretLabel.Location = New-Object System.Drawing.Point($GUI_MARGIN, $row9Y)
    $global:NewSecretLabel.Size = New-Object System.Drawing.Size(100, $GUI_LABEL_HEIGHT)
    $global:NewSecretLabel.Text = "New Secret:"
    $global:Form.Controls.Add($global:NewSecretLabel)

    # New Secret TextBox
    $secretTextX = $GUI_MARGIN + 110
    $global:NewSecretTextBox.Location = New-Object System.Drawing.Point($secretTextX, $row9Y - 3)
    $global:NewSecretTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $global:NewSecretTextBox.ReadOnly = $true # Make it read-only
    $global:Form.Controls.Add($global:NewSecretTextBox)

    # --- Event Handlers ---

    # Connect Button Click
    $global:ConnectButton.Add_Click({
        Connect-Tenant
    })

    # Disconnect Button Click
    $global:DisconnectButton.Add_Click({
        Disconnect-Tenant
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
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()

    # Required scopes for reading applications and adding secrets
    $scopes = "Application.Read.All", "Application.ReadWrite.All"

    try {
        Write-StatusMessage "Connecting to Microsoft Graph..." -Type Info
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        $context = Get-MgContext
        $tenantId = $context.TenantId
        $global:StatusLabel.Text = "Status: Connected to Tenant ID '$tenantId'"
        $global:DisconnectButton.Enabled = $true
        $global:FindSecretsButton.Enabled = $true
        $global:ExpiredSecretsListBox.Enabled = $true
        # Note: GenerateSecretButton is enabled only when an application is selected
        Write-StatusMessage "Successfully connected to tenant: $tenantId" -Type Success

    } catch {
        $errorMsg = $_.Exception.Message
        $global:StatusLabel.Text = "Status: Connection failed - $errorMsg"
        $global:ConnectButton.Enabled = $true
        Write-StatusMessage "Connection failed: $errorMsg" -Type Error
    }
}

function Disconnect-Tenant {
    Write-StatusMessage "Disconnecting from Microsoft Graph..." -Type Info
    $global:StatusLabel.Text = "Status: Disconnecting..."
    $global:ConnectButton.Enabled = $false
    $global:DisconnectButton.Enabled = $false
    $global:FindSecretsButton.Enabled = $false
    $global:ExpiredSecretsListBox.Enabled = $false
    $global:GenerateSecretButton.Enabled = $false
    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:NewSecretTextBox.Text = ""
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()

    try {
        Disconnect-MgGraph -ErrorAction Stop
        $global:StatusLabel.Text = "Status: Disconnected"
        $global:ConnectButton.Enabled = $true
        Write-StatusMessage "Successfully disconnected from Microsoft Graph" -Type Success
    } catch {
        $errorMsg = $_.Exception.Message
        $global:StatusLabel.Text = "Status: Disconnection failed - $errorMsg"
        $global:DisconnectButton.Enabled = $true # Allow retry if disconnect itself fails
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
    $global:SelectedAppNameLabel.Text = ""
    $global:SelectedEndDateLabel.Text = ""
    $global:NewSecretTextBox.Text = ""
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()

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
    Write-StatusMessage "Initiating secret generation..." -Type Info
     if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
         $global:StatusLabel.Text = "Status: Not connected. Please connect first."
         Write-StatusMessage "Cannot generate secret: Not connected to Graph" -Type Warning
         return
    }

    $selectedIndex = $global:ExpiredSecretsListBox.SelectedIndex

    if ($selectedIndex -lt 0 -or $selectedIndex -ge $global:ExpiredApplicationsData.Count) {
        $global:StatusLabel.Text = "Status: Please select an application first."
        Write-StatusMessage "Cannot generate secret: No application selected" -Type Warning
        return
    }

    $selectedApp = $global:ExpiredApplicationsData[$selectedIndex]
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
    Write-StatusMessage "Generating secret for App ID: $appId, Name: $appName" -Type Info

    # Generate display name from template
    $nextYear = (Get-Date).Year + 1
    $displayName = $secretDisplayNameTemplate -replace '\{YEAR\}', $nextYear

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
        Write-StatusMessage "New secret generated successfully for '$appName'. Secret displayed in textbox." -Type Success

        # Show a popup with a professional summary for ticketing system (if enabled)
        if ($showTicketNotePopup) {
            $now = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            # Replace placeholders in the ticket note template
            $ticketNote = $ticketNoteTemplate -replace '\{DISPLAYNAME\}', $displayName -replace '\{DATETIME\}', $now

            # Show a custom popup with a read-only textbox, a copy button, a paste screenshot button, and a PictureBox
        $popupForm = New-Object System.Windows.Forms.Form
        $popupForm.Text = $ticketNotePopupTitle
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
$global:DeleteSecretButton.Dispose()
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
