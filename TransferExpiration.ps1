# Account Expiration Migration Tool by Gavin Liddell.
# Requires the Microsoft.Graph module to function.
# Must have access to on-prem AD tenant.

# Import script modules.
$scriptModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users"
)

$currentDate = Get-Date -Format yyyy-MM-dd
$checkLogPath = Test-Path "$PSScriptRoot\transfer-logs\" -ErrorAction Stop
If ($checkLogPath -eq $false) {New-Item -Path "$PSScriptRoot" -Name "transfer-logs" -ItemType "directory"}

# Static variables
$extensionId = "<insert extension ID here>" # Name for the directory extension.
$logFile = "$PSScriptRoot\transfer-logs\transfer-$currentDate.log" # Set location of log file.

# Filter criterea variables.
$filterUPN = "*@example1.com"
$filterName = "<Name filter 1>"
$filterName2 = "<Name filter 2>"
$filterName3 = "<Name filter 3>"

# Add AccountExpirationDate attribute to Entra accounts.
Function Update-Expiration {
    # Set parameters for variables to reduce risk of rogue value types.
    Param (
        [string]$userExp,
        [string]$userID,
        [string]$extID
    )

    # Set extension property values
    $extensionProperty = @{
        $extID = $userExp
    }
    
    # Update user with new extension
    Update-MgUser -UserId $userID -BodyParameter $extensionProperty
    Write-Output "Successfully moved property for $userID."
}

# Write out script information and current date.
Write-Output "Account Expiration Migration Tool by Gavin Liddell`nRunning script..."
Add-Content -Path $logFile -Value "`n$(Get-Date)"

# Check and install required modules.
. $PSScriptRoot\Modules-Functions.ps1
Install-RequiredModules -requiredModules $scriptModules

# Get Contractors from on-prem AD.
$getADUsernames = Get-ADUser -Filter {Enabled -eq $true} -Properties AccountExpirationDate | `
    Where-Object {$_.UserPrincipalName -like $filterUPN -and ($_.Name -like $filterName -or $_.Name -like $filterName2 -or $_.Name -like $filterName3)}

# Check if user accounts were pulled.
If ($getADUsernames) {
    Add-Content -Path $logFile -Value "Successfully aquired AD accounts."
    Write-Output "Successfully aquired AD accounts."
}
# Exit if not pulled.
Else {
    Add-Content -Path $logFile -Value "Failed to retrieve AD accounts. Stopped."
    Write-Warning "ERROR: Failed to retrieve AD accounts. Stopped."
    Exit
}

# Import credentials function.
. $PSScriptRoot\<insert login script name here>

# Connect to Entra tenant.
$getLogin = Get-TenantLogin
If (!($getlogin)) {
    Write-Warning "Failed to connect to Entra tenant. Please confirm that credentials and tenant ID are all correct."
    Add-Content -Path $logFile -Value "Failed to connect to Entra tenant. Please confirm that credentials and tenant ID are all correct."
    Exit
}
Else {
    Write-Output "Connected to Entra tenant successfully."
    Add-Content -Path $logFile -Value "Connected to Entra tenant successfully."
}

# Sort users based on if they match or not and add a new property if they match.
ForEach ($user in $getADUsernames) {
    Try {
        # Check for existing user.
        $existingUser = Get-MgUser -Filter "userPrincipalName eq '$($user.UserPrincipalName)'" -Property "UserPrincipalName"

        # Update account expiration date if user exists.
        If ($existingUser) {
            $userExp = $user.AccountExpirationDate.ToString("yyyy-MM-dd")
            Update-Expiration -userExp $userExp -userID $user.UserPrincipalName -extID $extensionId
        }
        # Run a second check to see if user exists.
        Else {
            # Extract the username from the existing email address.
            $username = ($user.UserPrincipalName -split "@")[0]
            
            # Set alternative domain name.
            $newDomain = "example2.com"

            # Create the new email handle.
            $newEmail = "$username@$newDomain"

            # Second check for existing user.
            $existingUser = Get-MgUser -Filter "userPrincipalName eq '$newEmail'" -Property "UserPrincipalName"

            # Update account expiration date if user exists.
            If ($existingUser) {
                $userExp = $user.AccountExpirationDate.ToString("yyyy-MM-dd")
                Update-Expiration -userExp $userExp -userID $newEmail -extID $extensionId
            }
            # Log error if user is not found.
            Else {
                Write-Warning "Could not find $($user.UserPrincipalName) on new tenant. Skipping."
                Add-Content -Path $logFile -Value "ERROR: Could not find $($user.UserPrincipalName) on new tenant. Skipping."
            }
        }
        }
    # Log error if the main function fails.
    Catch {
        Write-Warning "Could not update $($user.UserPrincipalName). Logged Error: $_"
        Add-Content -Path $logFile -Value "ERROR: Could not update $($user.UserPrincipalName). Logged Error: $_"
    }
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

# Display where log file can be found.
Write-Output "The script has completed. Logs saved to $logFile."
Add-Content -Path $logFile -Value "The script has completed."

Exit