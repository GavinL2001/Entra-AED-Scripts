# Builds and displays the GUI.
Function Get-GUI {
    Param (
        [string]$previousText
    )

    # Set window size
    $windowWidth = 375
    $windowHeight = 330

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Creates window.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Expiring Account Management Tool"
    $form.Size = New-Object System.Drawing.Size($windowWidth, $windowHeight)
    $form.StartPosition = 'CenterScreen'
    $form.Icon = "$PSScriptRoot\AET.ico"

    # Creates OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(($windowWidth - 175), ($windowHeight - 75))
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # Creates Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(($windowWidth - 100), ($windowHeight - 75))
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    # Creates description of the tool.
    $description = New-Object System.Windows.Forms.Label
    $description.Location = New-Object System.Drawing.Point(10, 20)
    $description.Size = New-Object System.Drawing.Size(($windowWidth - 20), 40)
    $description.Text = "Interacts with Microsoft Entra to extend expiring accounts.`nCreated by Gavin Liddell."
    $form.Controls.Add($description)

    # Creates label for text box.
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 70)
    $label.Size = New-Object System.Drawing.Size(($windowWidth - 20), 20)
    $label.Text = "Enter account's UserPrincipalName (ex: first-last@example2.com):"
    $form.Controls.Add($label)

    # Creates text box.
    $textBox = New-Object System.Windows.Forms.TextBox
    If ($previousText) { $textBox.Text = $previousText }
    $textBox.Location = New-Object System.Drawing.Point(10, 90)
    $textBox.Size = New-Object System.Drawing.Size(305, 20)
    $form.Controls.Add($textBox)

    # Creates check box for deleting extension function.
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Location = New-Object System.Drawing.Point(10, 125)
    $checkBox.Size = New-Object System.Drawing.Size(($windowWidth - 50), 20)
    $checkBox.Text = "Remove AccountExpirationDate property."
    $form.Controls.Add($checkBox)

    # Creates label for radio button options.
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10, 160)
    $label2.Size = New-Object System.Drawing.Size(($windowWidth - 20), 20)
    $label2.Text = "How long do you want to extend the account for? (Default: 90 Days)"
    $form.Controls.Add($label2)

    # Creates radio button for "30 Days" option.
    $radioButton1 = New-Object System.Windows.Forms.RadioButton
    $radioButton1.Location = New-Object System.Drawing.Point(10, 180)
    $radioButton1.Size = New-Object System.Drawing.Size(($windowWidth - 50), 20)
    $radioButton1.Text = "30 Days"
    $form.Controls.Add($radioButton1)

    # Creates radio button for "60 Days" option.
    $radioButton2 = New-Object System.Windows.Forms.RadioButton
    $radioButton2.Location = New-Object System.Drawing.Point(10, 200)
    $radioButton2.Size = New-Object System.Drawing.Size(($windowWidth - 50), 20)
    $radioButton2.Text = "60 Days"
    $form.Controls.Add($radioButton2)

    # Creates radio button for "90 Days" option.
    $radioButton3 = New-Object System.Windows.Forms.RadioButton
    $radioButton3.Location = New-Object System.Drawing.Point(10, 220)
    $radioButton3.Size = New-Object System.Drawing.Size(($windowWidth - 50), 20)
    $radioButton3.Text = "90 Days"
    $form.Controls.Add($radioButton3)

    # Displays textbox and captures results from OK and Cancel buttons.
    $form.Add_Shown({$textBox.Select()})
    $buttonResult = $form.ShowDialog()

    # Determines which option was selected with radio buttons.
    $days = switch ($true) {
        $radioButton1.Checked {30}
        $radioButton2.Checked {60}
        $radioButton3.Checked {90}
        default {90}
    }

    # Creates custom object with various properties for script to pull from.
    [PSCustomObject]@{
        button = $buttonResult
        Id = $textBox.Text
        daysExtended = $days
        deleteProperty = $checkBox.Checked
    }
}

# Displays confirmation of selections and inputted information.
Function Show-Confirmation {
    Param (
        [string]$accountName,
        [string]$accountMail,
        [string]$accountJob,
        [string]$accountLocation,
        [string]$expirationDate
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Set window size
    $confirmWindowWidth = 250
    $confirmWindowHeight = 225

    # Creates window.
    $confirm = New-Object System.Windows.Forms.Form
    $confirm.Text = 'Confirmation'
    $confirm.Size = New-Object System.Drawing.Size($confirmWindowWidth, $confirmWindowHeight)
    $confirm.StartPosition = 'CenterScreen'

    # Creates Confirm button.
    $okButton1 = New-Object System.Windows.Forms.Button
    $okButton1.Location = New-Object System.Drawing.Point(($confirmWindowWidth - 175), ($confirmWindowHeight - 75))
    $okButton1.Size = New-Object System.Drawing.Size(75, 23)
    $okButton1.Text = 'Confirm'
    $okButton1.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $confirm.AcceptButton = $okButton1
    $confirm.Controls.Add($okButton1)

    # Creates Cancel button.
    $cancelButton1 = New-Object System.Windows.Forms.Button
    $cancelButton1.Location = New-Object System.Drawing.Point(($confirmWindowWidth - 100), ($confirmWindowHeight - 75))
    $cancelButton1.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton1.Text = 'Cancel'
    $cancelButton1.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $confirm.CancelButton = $cancelButton1
    $confirm.Controls.Add($cancelButton1)

    # Creates label to display information pulled.
    $label4 = New-Object System.Windows.Forms.Label
    $label4.Location = New-Object System.Drawing.Point(10, 20)
    $label4.Size = New-Object System.Drawing.Size(180, 200)
    $label4.Text = @"
Confirm that the information displayed here is correct:

Name: $accountName
Email: $accountMail
Job Title: $accountJob
Office Location: $accountLocation
New Expiration Date: $expirationDate
"@
    $confirm.Controls.Add($label4)

    # Collect button result.
    $confirm.ShowDialog()
}

# Displays errors if any occur during the script.
Function Show-Error {
    Param (
        [string]$errorMsg
    )

        # Set window size
        $errorWindowWidth = 350
        $errorWindowHeight = 200

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Creates window.
    $errorbox = New-Object System.Windows.Forms.Form
    $errorbox.Text = 'Error'
    $errorbox.Size = New-Object System.Drawing.Size($errorWindowWidth, $errorWindowHeight)
    $errorbox.StartPosition = 'CenterScreen'

    # Creates OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(($errorWindowWidth / 2 - 37), ($errorWindowHeight - 75))
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $errorbox.AcceptButton = $okButton
    $errorbox.Controls.Add($okButton)

    # Creates label to display error information.
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(($errorWindowWidth - 30), ($errorWindowHeight - 100))
    $label.Text = $errorMsg
    $errorbox.Controls.Add($label)

    # Collects button result.
    $errorbox.ShowDialog()
}