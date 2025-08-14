# ================================
# Disable ActiveSync for users not in "ActiveSync-Access-Allowed" group
# Scheduled Task Compatible Script
# ================================

# --- Configurable Parameters ---
$SecurityGroup = "ActiveSync-Access-Allowed"
$LogPath = "C:\scripts\ActiveSyncScheduled.log"
$SendEmail = $true  # Set to $true to enable email alerts
$SmtpServer = "YOUR SMTP SERVER"
$From = "services@yourdomain.com"
$To = "someone@yourdomain.com"
$Subject = "ActiveSync Restriction Report"

# --- Load Required Modules ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
} catch {
    Add-Content -Path $LogPath -Value "[$(Get-Date)] ERROR: Failed to load modules - $_"
    exit 1
}

# --- Get Group Members ---
try {
    $AllowedUsers = Get-ADGroupMember -Identity $SecurityGroup -Recursive | Where-Object { $_.objectClass -eq 'user' } | Select-Object -ExpandProperty SamAccountName
} catch {
    Add-Content -Path $LogPath -Value "[$(Get-Date)] ERROR: Failed to retrieve group members - $_"
    exit 1
}

# --- Get All Mailboxes ---
try {
    $AllMailboxes = Get-CASMailbox -ResultSize Unlimited
} catch {
    Add-Content -Path $LogPath -Value "[$(Get-Date)] ERROR: Failed to retrieve mailboxes - $_"
    exit 1
}

# --- Process Each Mailbox ---
$Changes = @()
foreach ($Mailbox in $AllMailboxes) {
    $Sam = $Mailbox.SamAccountName
    if ($AllowedUsers -notcontains $Sam -and $Mailbox.ActiveSyncEnabled) {
        try {
            Set-CASMailbox -Identity $Mailbox.Identity -ActiveSyncEnabled $false
            $Changes += "$Sam disabled"
        } catch {
            Add-Content -Path $LogPath -Value "[$(Get-Date)] ERROR disabling $Sam - $_"
        }
    }
}

# --- Logging ---
$Timestamp = Get-Date
Add-Content -Path $LogPath -Value "`n[$Timestamp] Task executed. Changes made: $($Changes.Count)"
if ($Changes.Count -gt 0) {
    Add-Content -Path $LogPath -Value ($Changes -join "`n")
}

# --- Optional Email Notification ---
if ($SendEmail -and $Changes.Count -gt 0) {
    $Body = "The following users had ActiveSync disabled:`n`n" + ($Changes -join "`n")
    try {
        Send-MailMessage -SmtpServer $SmtpServer -From $From -To $To -Subject $Subject -Body $Body
    } catch {
        Add-Content -Path $LogPath -Value "[$Timestamp] ERROR sending email - $_"
    }
}