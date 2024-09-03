# Account Expiration Automation by Gavin Liddell.
# Requires the Microsoft.Graph module to function.

# Import Microsoft Graph modules.
$scriptModules = @(
    "Microsoft.Graph.Authentication"
    "Microsoft.Graph.Users"
    "Microsoft.Graph.Users.Actions"
)

# Date variables.
$formattedDate = Get-Date -Format yyyy-MM-dd # Get formatted date for expiration function.
$oneWeek = (Get-Date).AddDays(7) | Get-Date -Format yyyy-MM-dd # Get one week from current day.
$twoWeeks = (Get-Date).AddDays(14) | Get-Date -Format yyyy-MM-dd # Get two weeks from current day.

# ID variables.
$extensionId = "<insert extension ID here>" # Name for the directory extension.
$senderId = "<insert sender ID here>" # Service Account's ID.

$checkLogPath = Test-Path "$PSScriptRoot\main-logs\" -ErrorAction Stop

If ($checkLogPath -eq $false) {
    New-Item -Path "$PSScriptRoot" -Name "main-logs" -ItemType "directory"
}

# Path variables.
$logFile = "$PSScriptRoot\main-logs\$formattedDate.log" # Set log path.
$emailFile = "$PSScriptRoot\<email file name>" # HTML file used for email template.

# Array variables.
$expiringAccounts = @() # Initialize an array to store accounts with the accountExpirationDate property.
$pairedAccounts = @() # Initialize an array to store accounts with managers assigned to them.
$expiredAccounts = @() # Initialize an array to store accounts that need to be disabled.
$ccEmailList = @( # Initialize an array to store emails that need to be cc'd when accounts expire.
    "HR@example2.com"
)

# Filter criterea variables.
$filterName = "<displayName filter 1>"
$filterName2 = "<displayName filter 2>"
$filterName3 = "<displayName filter 3>"

# Import functions.
. $PSScriptRoot\Modules-Functions.ps1
. $PSScriptRoot\SendEmail-Functions.ps1
. $PSScriptRoot\<login script name>

# Script information and beginning of log file.
Write-Output "Account Expiration Automation by Gavin Liddell`nRunning script..."
Add-Content -Path $logFile -Value "$(Get-Date)"

# Check and install required modules.
Install-RequiredModules -requiredModules $scriptModules

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

# Pull all expiring accounts.
$tenantUsers = Get-MgUser -All -Filter "accountEnabled eq true" -Property "displayName,UserPrincipalName,$extensionId,mail" | `
    Where-Object {$_.displayName -like $filterName -or $_.displayName -like $filterName2 -or $_.displayName -like $filterName3}

# Check if user accounts were pulled.
If ($tenantUsers) {
    Write-Output "Successfully pulled expiring accounts."
}
# Exit script if not pulled.
Else {
    Write-Warning "Failed to retrieve expiring accounts. Stopped."
    Add-Content -Path $logFile -Value "ERROR: Failed to retrieve expiring accounts. Stopped."
    Exit
}

# Loop through each account and check for the specific extension and property.
ForEach ($user in $tenantUsers) {
    Try {
        # Fetch the specific extension for the account.
        $userID = $user.UserPrincipalName
        $userName = $user.displayName
        $expirationDate = $user.AdditionalProperties["$extensionId"]
        
        # Check if the extension and the property exist.
        If ($expirationDate) {
            # Create a temporary object for script to manipulate.
            $expiring = [PSCustomObject]@{
                Name = $userName
                UserPrincipalName = $userID
                Email = $user.mail
                AccountExpirationDate = $expirationDate
            }
            # Add users to array once property has been added to new object.
            $expiringAccounts += $expiring
        }
        # Exit if extension property not found.
        Else {
            Write-Warning "accountExpirationDate property not found for $userID. Please add property if user needs it."
            Add-Content -Path $logFile -Value "ERROR: accountExpirationDate property not found for $userName ($userID). Please add property if user needs it."
        }
    } 
    # Log if main function fails.
    Catch {
        Write-Warning "Failed to retrieve extension for $userName ($userID). Error: $_"
        Add-Content -Path $logFile -Value "ERROR: Failed to retrieve extension for $userName ($userID). Error: $_"
    }
}

# Pull manager information and add to each user object.
ForEach ($user in $expiringAccounts) {
    Try {
        $userID = $user.UserPrincipalName
        $userName = $user.Name
        
        # Get Manager IDs.
        $manager = Get-MgUserManager -UserId $userID -ErrorAction SilentlyContinue

        # Run code if manager is found.
        If ($manager) {
            # Pull manager information.
            $managerInfo = Get-MgUser -UserId $manager.Id | Select-Object DisplayName,Mail
            
            # Assign manager to expiring account.
            $assignedManagerName = $managerInfo.DisplayName
            $assignedManagerEmail = $managerInfo.Mail
            $user | Add-Member -NotePropertyName ManagerName -NotePropertyValue $assignedManagerName
            $user | Add-Member -NotePropertyName ManagerEmail -NotePropertyValue $assignedManagerEmail
            
            # Add expiring account to new array.
            $pairedAccounts += $user
            
            # Report success of assigning the account to manager.
            Write-Output "Successfully assigned $userName to $assignedManagerName."
        }
        # Log error if manager is not found.
        Else {
            Write-Warning "Manager not found for $userName. Please assign a manager to this user."
            Add-Content -Path $logFile -Value "ERROR: Manager not found for $userName ($userID). Please assign a manager to this user."
        }
    }
    # Log error if main function fails.
    Catch {
        Write-Warning "Failed to fetch manager info for $userName. Error: $_"
        Add-Content -Path $logFile -Value "ERROR: Failed to fetch manager info for $userName ($userID). Error: $_"
    }
}

# Group expiring accounts by manager's email.
$managerGroups = $pairedAccounts | Group-Object -Property ManagerName -AsHashTable -AsString

# Divide accounts by expiration and send emails to managers.
ForEach ($manager in $managerGroups.Keys) {
    Try {
        $groupedUsers = $managerGroups[$manager] # Pull hash table.
        $managerName = $groupedUsers[0].ManagerName # Pull manager's name.
        $managerEmail = $groupedUsers[0].ManagerEmail # Pull manager's email.

        # Filter accounts into one week, two weeks, and expired categories.
        $usersTwoWeeks = $groupedUsers | `
            Where-Object {$_.AccountExpirationDate -eq "$twoWeeks"} | `
            Select-Object Email -Unique
        $usersOneWeek = $groupedUsers | `
            Where-Object {$_.AccountExpirationDate -eq "$oneWeek"} | `
            Select-Object Email -Unique
        $usersExpired = $groupedUsers | `
            Where-Object {$_.AccountExpirationDate -eq "$formattedDate"} | `
            Select-Object Email -Unique

        # Add commas in between each accounts's email.
        $usersTwoWeeksString = $usersTwoWeeks.Email -join ", "
        $usersOneWeekString = $usersOneWeek.Email -join ", "
        $usersExpiredString = $usersExpired.Email -join ", "

        # Call SendEmail function if accounts are found.
        If ($usersTwoWeeks -or $usersOneWeek -or $usersExpired) {
            Invoke-SendEmail `
                -usersTwoWeeks $usersTwoWeeksString `
                -usersOneWeek $usersOneWeekString `
                -usersExpired $usersExpiredString `
                -managerName $managerName `
                -managerEmail $managerEmail `
                -senderID $senderId `
                -formsLink $formsLink `
                -logoLink $logoLink `
                -htmlFile $emailFile `
                -ccEmails $ccEmailList
            Write-Output "`nFound expiring accounts tied to $managerName. Email sent successfully."
        }
        # Log error if manager does not have any expiring accounts that are about to expire.
        Else {
            Write-Output "`nNo expiring users found for $managerName. Skipping."
        }
    }
    # Log error if main function fails.
    Catch {
        Write-Warning "Failed to group users and/or send email to $managerName. Error: $_"
        Add-Content -Path $logFile -Value "ERROR Failed to group users and/or send email to $managerName. Error: $_"
    }
}

# Find expired users and put them in their own array.
$expiredAccounts += $expiringAccounts | Where-Object {$_.AccountExpirationDate -eq "$formattedDate"}

# Check to see if there are any expired accounts.
If ($expiredAccounts) {
    # Disable each account that matches the current date.
    ForEach ($user in $expiredAccounts) {
        Try {
            # Fetch the specific extension for the account.
            $userID = $user.UserPrincipalName
            $userName = $user.displayName
            Update-MgUser -UserId $userID -AccountEnabled:$false
            $disabledUsers += $userName
        } Catch {
            # Log the error.
            Write-Warning "Failed to disable account for $userName. Error: $_"
            Add-Content -Path $logFile -Value "ERROR: Failed to disable account for $userName ($userID). Error: $_"
        }
    }
    # Log accounts that expired and were successfully disabled.
    Write-Output "`nSuccessfully disabled accounts for these users:`n$disabledUsers"
    Add-Content -Path $logFile -Value "Successfully disabled accounts for these users:`n$disabledUsers"
}
# Report that there are no accounts expiring today.
Else {
    Write-Output "`nNo accounts expired today."
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
Write-Output "The script has completed. Logs saved to '$logFile'."
Add-Content -Path $logFile -Value "The script has completed. Exiting."

Exit