# Account Expiration GUI Tool created by Gavin Liddell.
# Requires the Microsoft.Graph module to function.

$scriptModules = @( # Define needed modules for script.
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users"
)

$currentDate = Get-Date -Format yyyy-MM-dd
$checkLogPath = Test-Path "$PSScriptRoot\gui-logs\" -ErrorAction Stop
If ($checkLogPath -eq $false) {New-Item -Path "$PSScriptRoot" -Name "gui-logs" -ItemType "directory"}

# Script Variables
$logFile = "$PSScriptRoot\gui-logs\GUI-$currentDate.log" # Set location of log file.

# Import core functions to the GUI tool.
. $PSScriptRoot\BuildGUI-Functions.ps1

# Create log timestamp and print script info.
Add-Content -Path $logFile -Value "`nAccessed by $env:USERNAME from $env:USERDOMAIN on $(Get-Date)"
Write-Output "Account Expiration GUI Tool created by Gavin Liddell.`nOpening GUI..."

# Enables accounts and updates their expiration date.
Function Enable-Account {
    Param (
        [string]$userID,
        [string]$acctExp,
        [string]$logFile
    )

    Try {
        # Create the new extension property.
        $extensionProperty = @{
            <insert extension ID here> = $acctExp
        }
        # Update user with new info.
        Update-MgUser -UserId $userID -BodyParameter $extensionProperty
        Update-MgUser -UserId $userID -AccountEnabled:$true
    }
    Catch {
        Show-Error -errorMsg "Error: Failed to update the user. $_"
        Add-Content -Path $logFile -Value "ERROR: Failed to update the user. $_"
    }
}

# Deletes the extension if user contains extension.
Function Remove-Extension {
    Param (
        [string]$userID,
        [string]$logFile
    )

    Try {
        # Create the new extension property.
        $extensionProperty = @{
            <insert extension ID here> = $null
        }
        # Update user with new info.
        Update-MgUser -UserId $userID -BodyParameter $extensionProperty
    }
    Catch {
        Show-Error -errorMsg "Error: Failed to remove extension. $_"
        Add-Content -Path $logFile -Value "ERROR: Failed to remove extension. $_"
    }
}

<#
# Uncomment code to restrict access to only allow users signed-in to the Entra tenant.
$domain = "<domain name here>"
If ($env:USERDOMAIN -ne $domain) {
    Show-Error -errorMsg "Please connect to the Entra tenant to run this tool."
    Add-Content -Path $logFile -Value "`n$env:USERNAME attempted to acces the script on $env:USERDOMAIN domain on $(Get-Date), but was denied due to being on wrong domain."
    Exit
}
#>

# Main function that executes the primary functions of script.
Function Main {
    Do {
        # Pull information from GUI objects.
        $getResult = Get-GUI -previousText $previousText

        If ($getResult.button -eq [System.Windows.Forms.DialogResult]::OK -and $getResult.Id) {
            # Save inputted text to re-insert into text box if user needs to update something.
            $previousText = $getResult.Id
            
            # Import modules function
            . $PSScriptRoot\Modules-Functions.ps1

            # Execute modules function.
            Install-RequiredModules -requiredModules $scriptModules

            # Import and execute tenant login function
            . $PSScriptRoot\<script name here>
            $loginStatus = Get-TenantLogin

            # Determin if the login process succeeded or not.
            If (!($loginStatus)) {
                Show-Error -errorMsg "Failed to connect to Entra tenant. Please confirm that credentials and tenant ID are all valid."
                Add-Content -Path $logFile -Value "Failed to connect to Entra tenant. Please confirm that credentials and tenant ID are all valid."
                Exit
            } Else {
                Write-Output "Connected to Entra tenant successfully."
                Add-Content -Path $logFile -Value "Connected to Entra tenant successfully."
            }

            Try {
                # Get account information.
                $pullAcctInfo = Get-MgUser -UserId $getResult.Id

                # Check if account information exists.
                If ($pullAcctInfo) {
                    # Set information about account.
                    $name = "$($pullAcctInfo.GivenName) $($pullAcctInfo.Surname)"
                    $email = $pullAcctInfo.UserPrincipalName
                    $job = $pullAcctInfo.JobTitle
                    $location = $pullAcctInfo.OfficeLocation

                    # Check if delete property check box was checked.
                    If ($getResult.deleteProperty) {
                        Write-Output "$($getResult.Id) was found in the system."
                        Add-Content -Path $logFile -Value "$($getResult.Id) was found in the system."

                        # Show confirmation prompt.
                        $confirmationResult = Show-Confirmation -accountName $name -accountMail $email -accountJob $job -accountLocation $location -expirationDate "deleted"

                        # Check if user confirms the results of the information pulled.
                        If ($confirmationResult -eq [System.Windows.Forms.DialogResult]::OK) {
                            # Create new extension
                            Remove-Extension -userID $getResult.Id -logFile $logFile
                            $userFound = $true
                        }
                        Else {
                            $userFound = $false
                        }
                    }
                    Else {
                        # Set date to whatever is stored in the extension.
                        $getExtension = Get-MgUser -UserId $getResult.Id -Property "<insert extension ID here>"
                        $getDate = $getExtension.AdditionalProperties["<insert extension ID here>"]
                        
                        # Check if date was found.
                        If ($getDate) {
                            Write-Output "$($getResult.Id) was found in the system."
                            Add-Content -Path $logFile -Value "$($getResult.Id) was found in the system."

                            # Parse date stored in account and display confirmation.
                            $pulledDate = [datetime]::Parse($getDate) # Convert into datetime format.
                            $newDate = ($pulledDate).AddDays($getResult.daysExtended) | Get-Date -Format yyyy-MM-dd # Add days and convert format back to yyyy-MM-dd.
                            $confirmationResult = Show-Confirmation -accountName $name -accountMail $email -accountJob $job -accountLocation $location -expirationDate $newDate # Show conformation with information.
                            
                            # Check if user confirms the results of the information pulled.
                            If ($confirmationResult -eq [System.Windows.Forms.DialogResult]::OK) {
                                # Enable account and update expiration date.
                                Enable-Account -userID $getResult.Id -acctExp $newDate -logFile $logFile
                                $userFound = $true
                            }
                            Else {
                                $userFound = $false
                            }
                        } Else {
                            Write-Output "$($getResult.Id) was found in the system."
                            Add-Content -Path $logFile -Value "$($getResult.Id) was found in the system."

                            # Use current date as base for extension and display confirmation.
                            $newDate = (Get-Date).AddDays($getResult.daysExtended) | Get-Date -Format yyyy-MM-dd # Add days and convert format back to yyyy-MM-dd.
                            $confirmationResult = Show-Confirmation -accountName $name -accountMail $email -accountJob $job -accountLocation $location -expirationDate $newDate # Show conformation with information.

                            # Check if user confirms the results of the information pulled.
                            If ($confirmationResult -eq [System.Windows.Forms.DialogResult]::OK) {
                                # Enable account and update expiration date.
                                Enable-Account -userID $getResult.Id -acctExp $newDate -logFile $logFile
                                $userFound = $true
                            }
                            Else {
                                $userFound = $false
                            }
                        }
                    }
                }
                Else {
                    # Display error if account could not be found.
                    Show-Error -errorMsg "Error: could not find the account '$($getResult.Id)'."
                    Add-Content -Path $logFile -Value "ERROR: could not find the account '$($getResult.Id)'."
                    $userFound = $false
                }
            }
            Catch {
                # Display error if Try function fails.
                Show-Error -errorMsg "Error: $_"
                Add-Content -Path $logFile -Value "ERROR: $_"
                $userFound = $false
            }
            # Disconnect from Graph.
            $disconnect = Disconnect-MgGraph
            If ($disconnect -ne $error) {
                Write-Output "Disconnected from Entra tenant successfully."
                Add-Content -Path $logFile -Value "Disconnected from Entra tenant successfully."
            }
            Else {
                Write-Output "Failed to disconnect from Entra tenant."
                Add-Content -Path $logFile -Value "ERROR: Failed to disconnect from Entra tenant."
            }
        }
        Else {
            Exit
        }
    } While (-Not $userFound)
}

# Execute Main function.
Main