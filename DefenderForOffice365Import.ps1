# Import Microsoft Defender for Office 365 Configuration to New Tenant
# Requires Exchange Online Management module

# Connect to the NEW tenant
Connect-ExchangeOnline

# Set the import path (where you exported the files)
$importPath = "C:\temp\DefenderForOffice365Export"

# Get the most recent export files
$timestamp = Get-ChildItem "$importPath\AntiSpam_Policies_*.xml" | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object -First 1 | 
    ForEach-Object { $_.Name -replace 'AntiSpam_Policies_|\.xml' }

Write-Host "Importing configuration from timestamp: $timestamp" -ForegroundColor Cyan

# Function to create policy with error handling
function Import-PolicyWithErrorHandling {
    param($PolicyObject, $PolicyType)
    try {
        Write-Host "Creating $PolicyType : $($PolicyObject.Name)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error creating $PolicyType : $($PolicyObject.Name) - $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# ========================================
# IMPORT ANTI-SPAM POLICIES AND RULES
# ========================================
Write-Host "`n=== Importing Anti-Spam Configuration ===" -ForegroundColor Yellow
$antiSpamPolicies = Import-Clixml "$importPath\AntiSpam_Policies_$timestamp.xml"

foreach ($policy in $antiSpamPolicies) {
    # Skip default policy if it exists
    if ($policy.Name -eq "Default") {
        Write-Host "Updating existing Default anti-spam policy..." -ForegroundColor Cyan
        Set-HostedContentFilterPolicy -Identity "Default" `
            -BulkSpamAction $policy.BulkSpamAction `
            -BulkThreshold $policy.BulkThreshold `
            -HighConfidenceSpamAction $policy.HighConfidenceSpamAction `
            -SpamAction $policy.SpamAction `
            -PhishSpamAction $policy.PhishSpamAction `
            -HighConfidencePhishAction $policy.HighConfidencePhishAction `
            -EnableLanguageBlockList $policy.EnableLanguageBlockList `
            -EnableRegionBlockList $policy.EnableRegionBlockList `
            -IncreaseScoreWithImageLinks $policy.IncreaseScoreWithImageLinks `
            -IncreaseScoreWithNumericIps $policy.IncreaseScoreWithNumericIps `
            -IncreaseScoreWithRedirectToOtherPort $policy.IncreaseScoreWithRedirectToOtherPort `
            -IncreaseScoreWithBizOrInfoUrls $policy.IncreaseScoreWithBizOrInfoUrls `
            -MarkAsSpamBulkMail $policy.MarkAsSpamBulkMail `
            -QuarantineRetentionPeriod $policy.QuarantineRetentionPeriod
    }
    else {
        # Create new custom policy
        try {
            New-HostedContentFilterPolicy -Name $policy.Name `
                -BulkSpamAction $policy.BulkSpamAction `
                -BulkThreshold $policy.BulkThreshold `
                -HighConfidenceSpamAction $policy.HighConfidenceSpamAction `
                -SpamAction $policy.SpamAction `
                -PhishSpamAction $policy.PhishSpamAction `
                -HighConfidencePhishAction $policy.HighConfidencePhishAction `
                -EnableLanguageBlockList $policy.EnableLanguageBlockList `
                -EnableRegionBlockList $policy.EnableRegionBlockList `
                -QuarantineRetentionPeriod $policy.QuarantineRetentionPeriod
            Write-Host "Created anti-spam policy: $($policy.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating anti-spam policy $($policy.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Import Anti-Spam Rules
$antiSpamRules = Import-Clixml "$importPath\AntiSpam_Rules_$timestamp.xml"
foreach ($rule in $antiSpamRules) {
    try {
        New-HostedContentFilterRule -Name $rule.Name `
            -HostedContentFilterPolicy $rule.HostedContentFilterPolicy `
            -Priority $rule.Priority `
            -RecipientDomainIs $rule.RecipientDomainIs `
            -SentTo $rule.SentTo `
            -SentToMemberOf $rule.SentToMemberOf `
            -Enabled $rule.State
        Write-Host "Created anti-spam rule: $($rule.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error creating anti-spam rule $($rule.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ========================================
# IMPORT ANTI-MALWARE POLICIES AND RULES
# ========================================
Write-Host "`n=== Importing Anti-Malware Configuration ===" -ForegroundColor Yellow
$antiMalwarePolicies = Import-Clixml "$importPath\AntiMalware_Policies_$timestamp.xml"

foreach ($policy in $antiMalwarePolicies) {
    if ($policy.Name -eq "Default") {
        Write-Host "Updating existing Default anti-malware policy..." -ForegroundColor Cyan
        Set-MalwareFilterPolicy -Identity "Default" `
            -Action $policy.Action `
            -EnableFileFilter $policy.EnableFileFilter `
            -EnableInternalSenderAdminNotifications $policy.EnableInternalSenderAdminNotifications `
            -EnableInternalSenderNotifications $policy.EnableInternalSenderNotifications `
            -ZapEnabled $policy.ZapEnabled
    }
    else {
        try {
            New-MalwareFilterPolicy -Name $policy.Name `
                -Action $policy.Action `
                -EnableFileFilter $policy.EnableFileFilter `
                -EnableInternalSenderAdminNotifications $policy.EnableInternalSenderAdminNotifications `
                -EnableInternalSenderNotifications $policy.EnableInternalSenderNotifications `
                -ZapEnabled $policy.ZapEnabled
            Write-Host "Created anti-malware policy: $($policy.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating anti-malware policy $($policy.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Import Anti-Malware Rules
$antiMalwareRules = Import-Clixml "$importPath\AntiMalware_Rules_$timestamp.xml"
foreach ($rule in $antiMalwareRules) {
    try {
        New-MalwareFilterRule -Name $rule.Name `
            -MalwareFilterPolicy $rule.MalwareFilterPolicy `
            -Priority $rule.Priority `
            -RecipientDomainIs $rule.RecipientDomainIs `
            -SentTo $rule.SentTo `
            -SentToMemberOf $rule.SentToMemberOf `
            -Enabled $rule.State
        Write-Host "Created anti-malware rule: $($rule.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error creating anti-malware rule $($rule.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ========================================
# IMPORT ANTI-PHISHING POLICIES AND RULES
# ========================================
Write-Host "`n=== Importing Anti-Phishing Configuration ===" -ForegroundColor Yellow
$antiPhishPolicies = Import-Clixml "$importPath\AntiPhish_Policies_$timestamp.xml"

foreach ($policy in $antiPhishPolicies) {
    if ($policy.Name -like "*Office365 AntiPhish Default*") {
        Write-Host "Updating existing Default anti-phishing policy..." -ForegroundColor Cyan
        Set-AntiPhishPolicy -Identity $policy.Identity `
            -EnableMailboxIntelligence $policy.EnableMailboxIntelligence `
            -EnableMailboxIntelligenceProtection $policy.EnableMailboxIntelligenceProtection `
            -EnableSpoofIntelligence $policy.EnableSpoofIntelligence `
            -EnableUnauthenticatedSender $policy.EnableUnauthenticatedSender `
            -PhishThresholdLevel $policy.PhishThresholdLevel
    }
    else {
        try {
            New-AntiPhishPolicy -Name $policy.Name `
                -EnableMailboxIntelligence $policy.EnableMailboxIntelligence `
                -EnableMailboxIntelligenceProtection $policy.EnableMailboxIntelligenceProtection `
                -EnableOrganizationDomainsProtection $policy.EnableOrganizationDomainsProtection `
                -EnableSpoofIntelligence $policy.EnableSpoofIntelligence `
                -EnableTargetedDomainsProtection $policy.EnableTargetedDomainsProtection `
                -EnableTargetedUserProtection $policy.EnableTargetedUserProtection `
                -EnableUnauthenticatedSender $policy.EnableUnauthenticatedSender `
                -PhishThresholdLevel $policy.PhishThresholdLevel `
                -TargetedDomainProtectionAction $policy.TargetedDomainProtectionAction `
                -TargetedUserProtectionAction $policy.TargetedUserProtectionAction
            Write-Host "Created anti-phishing policy: $($policy.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating anti-phishing policy $($policy.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Import Anti-Phishing Rules
$antiPhishRules = Import-Clixml "$importPath\AntiPhish_Rules_$timestamp.xml"
foreach ($rule in $antiPhishRules) {
    try {
        New-AntiPhishRule -Name $rule.Name `
            -AntiPhishPolicy $rule.AntiPhishPolicy `
            -Priority $rule.Priority `
            -RecipientDomainIs $rule.RecipientDomainIs `
            -SentTo $rule.SentTo `
            -SentToMemberOf $rule.SentToMemberOf `
            -Enabled $rule.State
        Write-Host "Created anti-phishing rule: $($rule.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error creating anti-phishing rule $($rule.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ========================================
# IMPORT SAFE ATTACHMENTS (Defender for Office 365)
# ========================================
Write-Host "`n=== Importing Safe Attachments Configuration ===" -ForegroundColor Yellow
try {
    $safeAttachPolicies = Import-Clixml "$importPath\SafeAttachments_Policies_$timestamp.xml"
    
    foreach ($policy in $safeAttachPolicies) {
        if ($policy.Name -like "*Built-In Protection Policy*") {
            Write-Host "Skipping built-in Safe Attachments policy: $($policy.Name)" -ForegroundColor Gray
            continue
        }
        
        try {
            New-SafeAttachmentPolicy -Name $policy.Name `
                -Enable $policy.Enable `
                -Action $policy.Action `
                -Redirect $policy.Redirect `
                -RedirectAddress $policy.RedirectAddress
            Write-Host "Created Safe Attachments policy: $($policy.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating Safe Attachments policy $($policy.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Import Safe Attachments Rules
    $safeAttachRules = Import-Clixml "$importPath\SafeAttachments_Rules_$timestamp.xml"
    foreach ($rule in $safeAttachRules) {
        try {
            New-SafeAttachmentRule -Name $rule.Name `
                -SafeAttachmentPolicy $rule.SafeAttachmentPolicy `
                -Priority $rule.Priority `
                -RecipientDomainIs $rule.RecipientDomainIs `
                -SentTo $rule.SentTo `
                -SentToMemberOf $rule.SentToMemberOf `
                -Enabled $rule.State
            Write-Host "Created Safe Attachments rule: $($rule.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating Safe Attachments rule $($rule.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "Safe Attachments not available or error importing: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ========================================
# IMPORT SAFE LINKS (Defender for Office 365)
# ========================================
Write-Host "`n=== Importing Safe Links Configuration ===" -ForegroundColor Yellow
try {
    $safeLinksPolicies = Import-Clixml "$importPath\SafeLinks_Policies_$timestamp.xml"
    
    foreach ($policy in $safeLinksPolicies) {
        if ($policy.Name -like "*Built-In Protection Policy*") {
            Write-Host "Skipping built-in Safe Links policy: $($policy.Name)" -ForegroundColor Gray
            continue
        }
        
        try {
            New-SafeLinksPolicy -Name $policy.Name `
                -EnableSafeLinksForEmail $policy.EnableSafeLinksForEmail `
                -EnableSafeLinksForTeams $policy.EnableSafeLinksForTeams `
                -EnableSafeLinksForOffice $policy.EnableSafeLinksForOffice `
                -TrackClicks $policy.TrackClicks `
                -AllowClickThrough $policy.AllowClickThrough `
                -ScanUrls $policy.ScanUrls `
                -EnableForInternalSenders $policy.EnableForInternalSenders `
                -DeliverMessageAfterScan $policy.DeliverMessageAfterScan `
                -DisableUrlRewrite $policy.DisableUrlRewrite
            Write-Host "Created Safe Links policy: $($policy.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating Safe Links policy $($policy.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Import Safe Links Rules
    $safeLinksRules = Import-Clixml "$importPath\SafeLinks_Rules_$timestamp.xml"
    foreach ($rule in $safeLinksRules) {
        try {
            New-SafeLinksRule -Name $rule.Name `
                -SafeLinksPolicy $rule.SafeLinksPolicy `
                -Priority $rule.Priority `
                -RecipientDomainIs $rule.RecipientDomainIs `
                -SentTo $rule.SentTo `
                -SentToMemberOf $rule.SentToMemberOf `
                -Enabled $rule.State
            Write-Host "Created Safe Links rule: $($rule.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating Safe Links rule $($rule.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "Safe Links not available or error importing: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ========================================
# IMPORT CONNECTION FILTER POLICY
# ========================================
Write-Host "`n=== Importing Connection Filter Configuration ===" -ForegroundColor Yellow
try {
    $connFilterPolicies = Import-Clixml "$importPath\ConnectionFilter_Policies_$timestamp.xml"
    foreach ($policy in $connFilterPolicies) {
        Set-HostedConnectionFilterPolicy -Identity "Default" `
            -EnableSafeList $policy.EnableSafeList `
            -IPAllowList $policy.IPAllowList `
            -IPBlockList $policy.IPBlockList
        Write-Host "Updated Connection Filter policy" -ForegroundColor Green
    }
}
catch {
    Write-Host "Error updating Connection Filter: $($_.Exception.Message)" -ForegroundColor Red
}

# ========================================
# IMPORT OUTBOUND SPAM FILTER
# ========================================
Write-Host "`n=== Importing Outbound Spam Filter Configuration ===" -ForegroundColor Yellow
try {
    $outboundSpamPolicies = Import-Clixml "$importPath\OutboundSpam_Policies_$timestamp.xml"
    foreach ($policy in $outboundSpamPolicies) {
        Set-HostedOutboundSpamFilterPolicy -Identity "Default" `
            -RecipientLimitExternalPerHour $policy.RecipientLimitExternalPerHour `
            -RecipientLimitInternalPerHour $policy.RecipientLimitInternalPerHour `
            -RecipientLimitPerDay $policy.RecipientLimitPerDay `
            -ActionWhenThresholdReached $policy.ActionWhenThresholdReached `
            -BccSuspiciousOutboundMail $policy.BccSuspiciousOutboundMail `
            -NotifyOutboundSpam $policy.NotifyOutboundSpam
        Write-Host "Updated Outbound Spam Filter policy" -ForegroundColor Green
    }
}
catch {
    Write-Host "Error updating Outbound Spam Filter: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Import Process Completed ===" -ForegroundColor Cyan
Write-Host "Please review any errors above and verify the configuration in the new tenant." -ForegroundColor Cyan
Write-Host "`nNote: Some settings may require manual configuration:" -ForegroundColor Yellow
Write-Host "- User impersonation lists in anti-phishing policies" -ForegroundColor Yellow
Write-Host "- Domain impersonation lists" -ForegroundColor Yellow
Write-Host "- Notification email addresses" -ForegroundColor Yellow
Write-Host "- DKIM signing (requires DNS configuration)" -ForegroundColor Yellow

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false