# Check if the module is installed
if (-not (Get-Module -Name Mailozaurr -ListAvailable)) {
    Write-Host "Mailozaurr module is not installed. Attempting to install..."
    Install-Module -Name Mailozaurr -AllowClobber -Force
    if ($?) {
        Write-Host "Mailozaurr module installed successfully."
        Import-Module -Name Mailozaurr -Force
        Write-Host "Mailozaurr module imported."
    } else {
        Write-Host "Failed to install Mailozaurr module. Please check for errors."
        exit
    }
} else {
    Write-Host "Mailozaurr module is already installed."
    Import-Module -Name Mailozaurr -Force
    Write-Host "Mailozaurr module imported."
}

# Update the module
Write-Host "Checking for updates to Mailozaurr module..."
Update-Module -Name Mailozaurr
if ($?) {
    Write-Host "Mailozaurr module is up to date."
} else {
    Write-Host "Failed to update Mailozaurr module. Please check for errors."
}
# Send email with CSV attachment
$smtpServer = "mail.smtp2go.com"
$smtpPort = 2525
$from = "NOC-Alerts@loffler.com"
$to = "lofflervision@loffler.com"
$SMTPUsername = "Loffler-NOCAlerts"
$SMTPPassword = $env:SMTPEmailPassword
[securestring]$secStringPassword = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
[pscredential]$EmailCredential = New-Object System.Management.Automation.PSCredential ($SMTPUsername, $secStringPassword)
$subject = "Abandoned devices in deployment bucket"
$body = "$env:computername was left in the deployment bucket. Please get it out! Go to this link to see all devices in the deployment bucket: https://zinfandel.rmm.datto.com/site/284972/deployment-bucket"


Send-EmailMessage `
    -SmtpServer $smtpServer `
    -Port $smtpPort `
    -From $from `
    -To $to `
    -Credential $EmailCredential `
    -Subject $subject `
    -Body $body
