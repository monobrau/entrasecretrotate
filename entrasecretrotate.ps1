# Requires the Microsoft.Graph module: Install-Module Microsoft.Graph

Write-Host "Script started."

# --- Configuration ---
$requiredModules = @("Microsoft.Graph")

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
    $global:Form.Size = New-Object System.Drawing.Size(600, 500)
    $global:Form.StartPosition = "CenterScreen"
    $global:Form.FormBorderStyle = "FixedSingle" # Prevent resizing
    $global:Form.MaximizeBox = $false

    # Connect Button
    $global:ConnectButton.Location = New-Object System.Drawing.Point(10, 10)
    $global:ConnectButton.Size = New-Object System.Drawing.Size(100, 30)
    $global:ConnectButton.Text = "Connect"
    $global:Form.Controls.Add($global:ConnectButton)

    # Disconnect Button
    $global:DisconnectButton.Location = New-Object System.Drawing.Point(120, 10)
    $global:DisconnectButton.Size = New-Object System.Drawing.Size(100, 30)
    $global:DisconnectButton.Text = "Disconnect"
    $global:DisconnectButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:DisconnectButton)

    # Status Label
    $global:StatusLabel.Location = New-Object System.Drawing.Point(230, 15)
    $global:StatusLabel.Size = New-Object System.Drawing.Size(350, 20)
    $global:StatusLabel.Text = "Status: Disconnected"
    $global:Form.Controls.Add($global:StatusLabel)

    # Find Secrets Button
    $global:FindSecretsButton.Location = New-Object System.Drawing.Point(10, 50)
    $global:FindSecretsButton.Size = New-Object System.Drawing.Size(150, 30)
    $global:FindSecretsButton.Text = "Find Expired Secrets"
    $global:FindSecretsButton.Enabled = $false # Disabled initially
    $global:Form.Controls.Add($global:FindSecretsButton)

    # Expired Secrets Label
    $global:ExpiredSecretsLabel.Location = New-Object System.Drawing.Point(10, 90)
    $global:ExpiredSecretsLabel.Size = New-Object System.Drawing.Size(200, 20)
    $global:ExpiredSecretsLabel.Text = "Applications with Expired Secrets:"
    $global:Form.Controls.Add($global:ExpiredSecretsLabel)

    # Expired Secrets ListBox
    $global:ExpiredSecretsListBox.Location = New-Object System.Drawing.Point(10, 110)
    $global:ExpiredSecretsListBox.Size = New-Object System.Drawing.Size(560, 150)
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
        Write-Host "Calling Connect-MgGraph..."
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        Write-Host "Connect-MgGraph returned."
        $context = Get-MgContext
        $tenantId = $context.TenantId
        $global:StatusLabel.Text = "Status: Connected to Tenant ID '$tenantId'"
        $global:DisconnectButton.Enabled = $true
        $global:FindSecretsButton.Enabled = $true
        $global:ExpiredSecretsListBox.Enabled = $true
        Write-Host "Connection successful."
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
    $global:ExpiredSecretsListBox.Items.Clear()
    $global:ExpiredApplicationsData = @()

    try {
        Write-Host "Calling Disconnect-MgGraph..."
        Disconnect-MgGraph -ErrorAction Stop
        Write-Host "Disconnect-MgGraph returned."
        $global:StatusLabel.Text = "Status: Disconnected"
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

    if ($selectedIndex -ge 0 -and $selectedIndex -lt $global:ExpiredApplicationsData.Count) {
        $selectedApp = $global:ExpiredApplicationsData[$selectedIndex]
        $global:SelectedAppNameLabel.Text = $selectedApp.DisplayName
        $global:SelectedEndDateLabel.Text = "Oldest Expired Date: $($selectedApp.EndDate.ToShortDateString())"
        $global:GenerateSecretButton.Enabled = $true
        Write-Host "Selected application: $($selectedApp.DisplayName)"
    } else {
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
    Write-Host "Generating secret for App ID: $appId, Name: $appName"

    try {
        Write-Host "Calling Add-MgApplicationPassword (without DisplayName)..."
        # Add a new password credential (secret)
        # Removed -DisplayName for compatibility with older module versions
        $newSecret = Add-MgApplicationPassword -ApplicationId $appId -ErrorAction Stop
        Write-Host "Add-MgApplicationPassword returned."

        # The actual secret value is in the SecretText property and is only returned NOW
        $secretValue = $newSecret.SecretText

        $global:NewSecretTextBox.Text = $secretValue
        $global:StatusLabel.Text = "Status: New secret generated for '$appName'. COPY IMMEDIATELY!"
        Write-Host "New secret value obtained. Displayed in textbox."

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
if (-not (Import-RequiredModules -Modules $requiredModules)) {
    # Import failed, message printed by Import-RequiredModules
    Write-Host "Module import failed. Please resolve the errors above and try again." -ForegroundColor Red
    exit # Exit if import failed
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
