# Export Microsoft Defender for Office 365 Configuration
# Requires Exchange Online Management module

# Install module if needed
# Install-Module -Name ExchangeOnlineManagement -Force

# Connect to Exchange Online
Connect-ExchangeOnline

# Create export directory
$exportPath = "C:\temp\DefenderForOffice365Export"
if (!(Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Export Anti-Spam Policies
Write-Host "Exporting Anti-Spam policies..." -ForegroundColor Green
Get-HostedContentFilterPolicy | Export-Clixml "$exportPath\AntiSpam_Policies_$timestamp.xml"
Get-HostedContentFilterPolicy | ConvertTo-Json -Depth 10 | Out-File "$exportPath\AntiSpam_Policies_$timestamp.json"

# Export Anti-Spam Rules
Get-HostedContentFilterRule | Export-Clixml "$exportPath\AntiSpam_Rules_$timestamp.xml"
Get-HostedContentFilterRule | ConvertTo-Json -Depth 10 | Out-File "$exportPath\AntiSpam_Rules_$timestamp.json"

# Export Anti-Malware Policies
Write-Host "Exporting Anti-Malware policies..." -ForegroundColor Green
Get-MalwareFilterPolicy | Export-Clixml "$exportPath\AntiMalware_Policies_$timestamp.xml"
Get-MalwareFilterPolicy | ConvertTo-Json -Depth 10 | Out-File "$exportPath\AntiMalware_Policies_$timestamp.json"

# Export Anti-Malware Rules
Get-MalwareFilterRule | Export-Clixml "$exportPath\AntiMalware_Rules_$timestamp.xml"
Get-MalwareFilterRule | ConvertTo-Json -Depth 10 | Out-File "$exportPath\AntiMalware_Rules_$timestamp.json"

# Export Anti-Phishing Policies
Write-Host "Exporting Anti-Phishing policies..." -ForegroundColor Green
Get-AntiPhishPolicy | Export-Clixml "$exportPath\AntiPhish_Policies_$timestamp.xml"
Get-AntiPhishPolicy | ConvertTo-Json -Depth 10 | Out-File "$exportPath\AntiPhish_Policies_$timestamp.json"

# Export Anti-Phishing Rules
Get-AntiPhishRule | Export-Clixml "$exportPath\AntiPhish_Rules_$timestamp.xml"
Get-AntiPhishRule | ConvertTo-Json -Depth 10 | Out-File "$exportPath\AntiPhish_Rules_$timestamp.json"

# Export Safe Attachments Policies (if Defender for Office 365 Plan 1/2)
Write-Host "Exporting Safe Attachments policies..." -ForegroundColor Green
Get-SafeAttachmentPolicy | Export-Clixml "$exportPath\SafeAttachments_Policies_$timestamp.xml"
Get-SafeAttachmentPolicy | ConvertTo-Json -Depth 10 | Out-File "$exportPath\SafeAttachments_Policies_$timestamp.json"

# Export Safe Attachments Rules
Get-SafeAttachmentRule | Export-Clixml "$exportPath\SafeAttachments_Rules_$timestamp.xml"
Get-SafeAttachmentRule | ConvertTo-Json -Depth 10 | Out-File "$exportPath\SafeAttachments_Rules_$timestamp.json"

# Export Safe Links Policies (if Defender for Office 365 Plan 1/2)
Write-Host "Exporting Safe Links policies..." -ForegroundColor Green
Get-SafeLinksPolicy | Export-Clixml "$exportPath\SafeLinks_Policies_$timestamp.xml"
Get-SafeLinksPolicy | ConvertTo-Json -Depth 10 | Out-File "$exportPath\SafeLinks_Policies_$timestamp.json"

# Export Safe Links Rules
Get-SafeLinksRule | Export-Clixml "$exportPath\SafeLinks_Rules_$timestamp.xml"
Get-SafeLinksRule | ConvertTo-Json -Depth 10 | Out-File "$exportPath\SafeLinks_Rules_$timestamp.json"

# Export Outbound Spam Filter Policy
Write-Host "Exporting Outbound Spam policies..." -ForegroundColor Green
Get-HostedOutboundSpamFilterPolicy | Export-Clixml "$exportPath\OutboundSpam_Policies_$timestamp.xml"
Get-HostedOutboundSpamFilterPolicy | ConvertTo-Json -Depth 10 | Out-File "$exportPath\OutboundSpam_Policies_$timestamp.json"

# Export Connection Filter Policy
Write-Host "Exporting Connection Filter policies..." -ForegroundColor Green
Get-HostedConnectionFilterPolicy | Export-Clixml "$exportPath\ConnectionFilter_Policies_$timestamp.xml"
Get-HostedConnectionFilterPolicy | ConvertTo-Json -Depth 10 | Out-File "$exportPath\ConnectionFilter_Policies_$timestamp.json"

# Export DKIM Configuration
Write-Host "Exporting DKIM configuration..." -ForegroundColor Green
Get-DkimSigningConfig | Export-Clixml "$exportPath\DKIM_Config_$timestamp.xml"
Get-DkimSigningConfig | ConvertTo-Json -Depth 10 | Out-File "$exportPath\DKIM_Config_$timestamp.json"

# Export ATP (Advanced Threat Protection) Policies if available
Write-Host "Exporting ATP policies..." -ForegroundColor Green
Get-AtpPolicyForO365 | Export-Clixml "$exportPath\ATP_Policy_$timestamp.xml"
Get-AtpPolicyForO365 | ConvertTo-Json -Depth 10 | Out-File "$exportPath\ATP_Policy_$timestamp.json"

Write-Host "`nExport completed successfully!" -ForegroundColor Cyan
Write-Host "Files saved to: $exportPath" -ForegroundColor Cyan

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false