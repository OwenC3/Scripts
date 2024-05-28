<#
.SYNOPSIS
    Sets a NinjaOne Custom Field called 'AssetTag' with a 4-digit code.
.DESCRIPTION
    Sets a NinjaOne Custom Field called 'AssetTag' with a 4-digit code. 
    If the custom field is already set, it will prompt the user to update the value and display the current value.
    If the custom field is not set, it will prompt the user to set the value.
    The ability to close the form is disabled when the custom field is not set.

    This script MUST be run as SYSTEM
.EXAMPLE
    .\Set-NinjaAssetTag.ps1
    This will prompt the user to enter a 4-digit code and set the NinjaOne Custom Field 'AssetTag' with the value.
.OUTPUTS
    Pop-up form to enter 4-digit code
.NOTES
    Release Notes: Initial Release
    Written By: RunningFreak
.COMPONENT
    RunAsUser, NuGet, NinjaRMM
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-ExecutionPolicy Bypass -Scope Process -Force

try {
    # Check if NuGet is installed and its version
    $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if ($null -eq $nugetProvider -or $nugetProvider.Version -lt [version]"2.8.5.201") {
        Write-Host "Installing or updating NuGet..."
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
    } else {
        Write-Host "NuGet is already up to date."
    }

    # Check if RunAsUser module is installed and its version
    $runAsUserModule = Get-Module -Name RunAsUser -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $runAsUserModule) {
        Write-Host "RunAsUser module is not installed. Installing..."
        Install-Module RunAsUser -Confirm:$False -Force -ErrorAction Stop
    } else {
        # Assuming you want to ensure it's the latest version available in the repository
        $installedVersion = $runAsUserModule.Version
        $availableVersion = Find-Module -Name RunAsUser | Select-Object -ExpandProperty Version
        if ($availableVersion -gt $installedVersion) {
            Write-Host "Updating RunAsUser module to the latest version..."
            Install-Module RunAsUser -Confirm:$False -Force -ErrorAction Stop
        } else {
            Write-Host "RunAsUser module is already up to date."
        }
    }

    Write-Host "Installation and update process completed successfully."
} catch {
    Write-Host "An error occurred during installation: $_"
}


# Grab Asset Tag Custom Field
$AssetTag = Ninja-Property-Get AssetTag

if ($null -ne $AssetTag) {
    Write-Host "Current Asset Tag: $AssetTag"
} else {
    Write-Host "Asset Tag not set"
}

$scriptblock = {

    # Grab Asset Tag Custom Field
    $AssetTag = Ninja-Property-Get AssetTag

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Create form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Enter 4-Digit Asset Tag"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true  # Form stays on top of other applications

    # Create label
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(380, 40)
    $label.Text = "Enter 4-Digit Asset Tag:"
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($label)

    # Create textbox
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(75, 60)
    $textBox.Size = New-Object System.Drawing.Size(250, 25) 
    $textBox.Font = New-Object System.Drawing.Font("Arial", 10) 
    $form.Controls.Add($textBox)

    # Create OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(150, 100)
    $okButton.Size = New-Object System.Drawing.Size(100, 30) 
    $okButton.Text = "OK"
    $okButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold) 

    # Display the current custom field value underneath the OK button
    if ($AssetTag -ne $null) {
        $currentKeyLabel = New-Object System.Windows.Forms.Label
        $currentKeyLabel.Location = New-Object System.Drawing.Point(10, 140)
        $currentKeyLabel.Size = New-Object System.Drawing.Size(380, 20)
        $currentKeyLabel.Text = "Current Asset Tag: $AssetTag"
        $currentKeyLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $currentKeyLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $form.Controls.Add($currentKeyLabel)
    }

    # Event handler for OK button and Enter key
    $okButton.Add_Click({
            $enteredCode = $textBox.Text
            if ($enteredCode -eq "" -or $enteredCode -match '\D' -or $enteredCode.Length -ne 4) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a valid 4-Digit Asset Tag.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } else {
                # Check if the custom field is set
                if ($AssetTag -eq $null) {
                    # Set the custom field if not set
                    Ninja-Property-Set AssetTag $enteredCode
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $form.Close()
                } else {
                    $confirmationMessage = "The existing 4-Digit Asset Tag is $AssetTag.`r`nYou entered $enteredCode.`r`nDo you want to update the value?"
                    $confirmationResult = [System.Windows.Forms.MessageBox]::Show($confirmationMessage, "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            
                    if ($confirmationResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                        # Set the custom field to the new value
                        Ninja-Property-Set AssetTag $enteredCode
                        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $form.Close()
                    }
                }
            }
        })

    $form.AcceptButton = $okButton  # Set the form's accept button to the OK button

    # Hide the close button (red button) when the custom field is not set
    if ($AssetTag -eq $null) {
        $form.ControlBox = $false
    }

    $form.Controls.Add($okButton)

    # Event handler for Enter key
    $form.Add_KeyDown({
            if ($_.KeyCode -eq 'Enter') {
                $enteredCode = $textBox.Text
                if ($enteredCode -eq "" -or $enteredCode -match '\D' -or $enteredCode.Length -ne 4) {
                    [System.Windows.Forms.MessageBox]::Show("Please enter a valid 4-Digit Asset Tag.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                } else {
                    # Check if the custom field is set
                    if ($AssetTag -eq $null) {
                        # Set the custom field
                        Ninja-Property-Set AssetTag $enteredCode
                        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $form.Close()
                    } else {
                        $confirmationMessage = "The existing 4-Digit Asset Tag is $AssetTag.`r`nYou entered $enteredCode.`r`nDo you want to update the value?"
                        $confirmationResult = [System.Windows.Forms.MessageBox]::Show($confirmationMessage, "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                
                        if ($confirmationResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                            # Set the custom field
                            Ninja-Property-Set AssetTag $enteredCode
                            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                            $form.Close()
                        }
                    }
                }
            }
        })

    # Show form and wait for user input
    $result = $form.ShowDialog()

    # Check if OK button was clicked
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $enteredCode = $textBox.Text
        Write-Host "Entered Code: $enteredCode"
    } else {
        Write-Host "Operation Cancelled"
    }

} # End Script Block

# Run Function to prompt user for asset tag
invoke-ascurrentuser -scriptblock $scriptblock

# Grab new asset tag
$AssetTag = Ninja-Property-Get AssetTag
if ($null -ne $AssetTag) {
    Write-Host "Ninja AssetTag Custom Field set to: $AssetTag"
} else {
    Write-Host "Asset Tag not set"
}
