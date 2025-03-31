<#
.SYNOPSIS
    Dark theme GUI tool to assign FullAccess and Send As permissions from an AD group to a Shared Mailbox.

.DESCRIPTION
    - Accepts an AD Group and Shared Mailbox name.
    - Validates that both exist before proceeding.
    - Assigns FullAccess and Send As rights to each user in the group.
    - Logs results in GUI and to a .log file.

.NOTES
    Author  : Farid Amghar
    Created : 2025-03-31
    Version : 2.0
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --------- COLORS (Dark Theme) ---------
$backColor   = [System.Drawing.Color]::FromArgb(30,30,30)
$textColor   = [System.Drawing.Color]::WhiteSmoke
$buttonColor = [System.Drawing.Color]::FromArgb(45,45,45)
$inputColor  = [System.Drawing.Color]::FromArgb(50,50,50)

# --------- MAIN FORM ---------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Shared Mailbox Permission Tool"
$form.Size = New-Object System.Drawing.Size(600, 520)
$form.StartPosition = "CenterScreen"
$form.BackColor = $backColor
$form.ForeColor = $textColor

# --------- CONTROLS ---------
function Add-Label($text, $x, $y) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Location = New-Object System.Drawing.Point($x, $y)
    $label.Size = New-Object System.Drawing.Size(160, 20)
    $label.ForeColor = $textColor
    $form.Controls.Add($label)
    return $label
}

function Add-TextBox($x, $y, $default="") {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point($x, $y)
    $box.Size = New-Object System.Drawing.Size(370, 20)
    $box.BackColor = $inputColor
    $box.ForeColor = $textColor
    $box.Text = $default
    $form.Controls.Add($box)
    return $box
}

Add-Label "AD Group Name:" 20 20
$textGroup = Add-TextBox 180 18

Add-Label "Shared Mailbox Name:" 20 60
$textMailbox = Add-TextBox 180 58

Add-Label "Log File Path:" 20 100
$textLogPath = Add-TextBox 180 98 "C:\Logs\SharedMailbox_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(20, 180)
$logBox.Size = New-Object System.Drawing.Size(530, 260)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.BackColor = $inputColor
$logBox.ForeColor = $textColor
$form.Controls.Add($logBox)

# --------- LOG FUNCTION ---------
function Write-Log {
    param($msg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] $msg"
    $logBox.AppendText("$line`r`n")

    if ($textLogPath.Text -ne "") {
        $logFolder = Split-Path $textLogPath.Text
        if (!(Test-Path $logFolder)) {
            New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $textLogPath.Text -Value $line
    }
}

# --------- VALIDATION FUNCTION ---------
function Validate-Inputs {
    param($groupName, $sharedMailbox, $dc)

    try {
        $exists = Get-ADGroup -Identity $groupName -Server $dc -ErrorAction Stop
        Write-Log "✅ Group '$groupName' found in AD."
    } catch {
        Write-Log "❌ Group '$groupName' not found in Active Directory."
        return $false
    }

    try {
        $mbx = Get-Mailbox -Identity $sharedMailbox -ErrorAction Stop
        if ($mbx.RecipientTypeDetails -ne "SharedMailbox") {
            Write-Log "❌ '$sharedMailbox' is not a Shared Mailbox."
            return $false
        }
        Write-Log "✅ Shared Mailbox '$sharedMailbox' found."
    } catch {
        Write-Log "❌ Shared Mailbox '$sharedMailbox' not found in Exchange."
        return $false
    }

    return $true
}

# --------- BUTTON "Run" ---------
$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Text = "Run"
$buttonRun.Location = New-Object System.Drawing.Point(250, 140)
$buttonRun.Size = New-Object System.Drawing.Size(100, 30)
$buttonRun.BackColor = $buttonColor
$buttonRun.ForeColor = $textColor
$form.Controls.Add($buttonRun)

$buttonRun.Add_Click({
    $groupName = $textGroup.Text.Trim()
    $sharedMailbox = $textMailbox.Text.Trim()
    $logFile = $textLogPath.Text.Trim()
    $dc = "FRDRPDC8269.GB.intra.corp"

    if ($groupName -eq "" -or $sharedMailbox -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Please fill in both fields.","Missing Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    if (-not (Validate-Inputs $groupName $sharedMailbox $dc)) {
        return
    }

    try {
        $members = Get-ADGroupMember -Identity $groupName -Server $dc
        Write-Log "Found $($members.Count) member(s) in group '$groupName'."
    } catch {
        Write-Log "❌ Error retrieving members from group: $_"
        return
    }

    foreach ($member in $members) {
        if ($member.objectClass -ne "user") {
            Write-Log "⏩ Skipped (not a user): $($member.Name)"
            continue
        }

        try {
            Add-MailboxPermission -Identity $sharedMailbox `
                                  -User $member.SamAccountName `
                                  -AccessRights FullAccess `
                                  -InheritanceType All `
                                  -DomainController $dc -ErrorAction Stop
            Write-Log "✅ FullAccess → $($member.SamAccountName)"
        } catch {
            Write-Log "❌ FullAccess error: $($member.SamAccountName) → $_"
        }

        try {
            Add-ADPermission -Identity $sharedMailbox `
                             -User $member.SamAccountName `
                             -ExtendedRights "Send As" `
                             -DomainController $dc -ErrorAction Stop
            Write-Log "✅ Send As → $($member.SamAccountName)"
        } catch {
            Write-Log "❌ Send As error: $($member.SamAccountName) → $_"
        }
    }

    Write-Log "✔️ Permission assignment completed."
})

# --------- SHOW FORM ---------
[void]$form.ShowDialog()
