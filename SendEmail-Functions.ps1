# Send email function when values are determined.
Function Invoke-SendEmail {
    Param(
        [string]$htmlFile,
        [string]$managerName,
        [string]$usersTwoWeeks,
        [string]$usersOneWeek,
        [string]$usersExpired,
        [string]$managerEmail,
        [string]$senderId,
        [string[]]$ccEmails
    )
    # Pull HTML file contents.
    $getHTML = Get-Content -Path $htmlFile -Raw

    # Set Email parameters.
    $params = @{
        message = @{
            subject = "Account Expiration Notice"
            importance = "High"
            body = @{
                contentType = "HTML"
                content = $getHTML `
                    -replace '{{managerName}}', $managerName `
                    -replace '{{usersTwoWeeks}}', $usersTwoWeeks `
                    -replace '{{usersOneWeek}}', $usersOneWeek `
                    -replace '{{usersExpired}}', $usersExpired
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = "$managerEmail"
                    }
                }
            )
        }
    }
    
    # Check if $usersExpired is not empty
    If ($usersExpired) {
        $ccList = @()

        # Add additional emails
        ForEach ($email in $ccEmails) {
            $ccList += @{
                emailAddress = @{
                    address = $email
                }
            }
        }

        # Assign ccRecipients to the message
        $params.message.ccRecipients = $ccList
    }

    Send-MgUserMail -UserId $senderId -BodyParameter $params # Send email to manager with parameters set above.
    Start-Sleep -Seconds 0.2
}