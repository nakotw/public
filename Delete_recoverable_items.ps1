# Connect with the required flag for purge operations
Connect-ExchangeOnline
Connect-IPPSSession -EnableSearchOnlySession

# Improved script to create a Compliance Search for recoverable items with hold policy management
# MODIFIED: Only targets DiscoveryHolds folder

# Prompt for the mailbox email address
Write-Host "Enter the email address of the mailbox to search:" -ForegroundColor Cyan
$Mailbox = Read-Host "Email address"

# Validate email format
if ($Mailbox -notmatch '^[\w\.-]+@[\w\.-]+\.\w+$') {
    Write-Host "Invalid email address format. Exiting script." -ForegroundColor Red
    exit
}

Write-Host "Processing mailbox: $Mailbox" -ForegroundColor Green
Write-Host "`n"

# Check for hold policies
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "CHECKING FOR HOLD POLICIES" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan

try {
    $mailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
    $hasHolds = $false
    $holdPolicies = @()

    # Check Litigation Hold
    if ($mailboxInfo.LitigationHoldEnabled -eq $true) {
        $hasHolds = $true
        Write-Host "✓ Litigation Hold: " -NoNewline -ForegroundColor Yellow
        Write-Host "ENABLED" -ForegroundColor Red
        $holdPolicies += [PSCustomObject]@{
            Type = "Litigation Hold"
            Name = "Litigation Hold"
            Details = "Enabled on: $($mailboxInfo.LitigationHoldDate)"
        }
    }

    # Check In-Place Hold
    if ($mailboxInfo.InPlaceHolds -and $mailboxInfo.InPlaceHolds.Count -gt 0) {
        Write-Host "✓ In-Place Holds Found: " -NoNewline -ForegroundColor Yellow
        Write-Host $mailboxInfo.InPlaceHolds.Count -ForegroundColor Red
        
        $activeHoldCount = 0
        
        foreach ($hold in $mailboxInfo.InPlaceHolds) {
            # Check for exclusion holds (those starting with -)
            if ($hold -match '^-UniH') {
                # Retention Policy Exclusion - mailbox is EXCLUDED, not under hold
                Write-Host "  - Retention Policy Exclusion: $hold " -NoNewline -ForegroundColor Green
                Write-Host "(mailbox is excluded - WILL NOT block deletion)" -ForegroundColor Gray
                continue  # Skip to next hold, don't add to holdPolicies
            }
            elseif ($hold -match '^-mbx') {
                # eDiscovery Hold Exclusion - mailbox is EXCLUDED, not under hold
                Write-Host "  - eDiscovery Hold Exclusion: $hold " -NoNewline -ForegroundColor Green
                Write-Host "(mailbox is excluded - WILL NOT block deletion)" -ForegroundColor Gray
                continue  # Skip to next hold, don't add to holdPolicies
            }
            elseif ($hold -match '^UniH') {
                # Unified Hold (Retention Policy) - ACTIVE HOLD
                $activeHoldCount++
                $policyName = "Retention Policy"
                try {
                    $policy = Get-RetentionCompliancePolicy | Where-Object { $_.Guid -eq $hold.Replace('UniH', '') }
                    if ($policy) { $policyName = $policy.Name }
                } catch { }
                
                $holdPolicies += [PSCustomObject]@{
                    Type = "Retention Policy"
                    Name = $policyName
                    Details = "Hold ID: $hold"
                    HoldId = $hold
                }
                Write-Host "  - $($holdPolicies[-1].Type): $($holdPolicies[-1].Name)" -ForegroundColor Yellow
            }
            elseif ($hold -match '^mbx') {
                # eDiscovery Hold - ACTIVE HOLD
                $activeHoldCount++
                $policyName = "eDiscovery Hold"
                try {
                    $case = Get-CaseHoldPolicy | Where-Object { $_.Guid -eq $hold.Replace('mbx', '') }
                    if ($case) { $policyName = $case.Name }
                } catch { }
                
                $holdPolicies += [PSCustomObject]@{
                    Type = "eDiscovery Hold"
                    Name = $policyName
                    Details = "Hold ID: $hold"
                    HoldId = $hold
                }
                Write-Host "  - $($holdPolicies[-1].Type): $($holdPolicies[-1].Name)" -ForegroundColor Yellow
            }
            else {
                # Unknown Hold Type
                $activeHoldCount++
                $holdPolicies += [PSCustomObject]@{
                    Type = "Unknown Hold"
                    Name = "Unknown"
                    Details = "Hold ID: $hold"
                    HoldId = $hold
                }
                Write-Host "  - $($holdPolicies[-1].Type): $($holdPolicies[-1].Name)" -ForegroundColor Yellow
            }
        }
        
        # Only set hasHolds if there are active holds (not exclusions)
        if ($activeHoldCount -gt 0) {
            $hasHolds = $true
        }
    }

    # Check Delay Hold (might be present after removing a hold)
    if ($mailboxInfo.DelayHoldApplied -eq $true) {
        $hasHolds = $true
        Write-Host "✓ Delay Hold: " -NoNewline -ForegroundColor Yellow
        Write-Host "APPLIED" -ForegroundColor Red
        Write-Host "  (Temporary hold after policy removal - can take up to 30 days to clear)" -ForegroundColor Gray
        $holdPolicies += [PSCustomObject]@{
            Type = "Delay Hold"
            Name = "Delay Hold"
            Details = "Temporary hold after policy removal"
        }
    }

    if (-not $hasHolds) {
        Write-Host "✓ No holds found on this mailbox" -ForegroundColor Green
    }

    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "`n"

    # If holds exist, ask user what to do
    if ($hasHolds) {
        Write-Host "WARNING: This mailbox has active holds that may prevent deletion of recoverable items." -ForegroundColor Red
        Write-Host "`nOptions:" -ForegroundColor Cyan
        Write-Host "1. Remove holds and proceed with search/purge" -ForegroundColor Yellow
        Write-Host "2. Proceed with search only (without removing holds - deletion will likely fail)" -ForegroundColor Yellow
        Write-Host "3. Exit script" -ForegroundColor Yellow
        Write-Host "`n"
        
        $choice = Read-Host "Enter your choice (1, 2, or 3)"
        
        switch ($choice) {
            "1" {
                Write-Host "`nRemoving hold policies..." -ForegroundColor Yellow
                
                # Remove Litigation Hold
                if ($mailboxInfo.LitigationHoldEnabled -eq $true) {
                    Write-Host "Removing Litigation Hold..." -ForegroundColor Yellow
                    try {
                        Set-Mailbox -Identity $Mailbox -LitigationHoldEnabled $false
                        Write-Host "✓ Litigation Hold removed" -ForegroundColor Green
                    } catch {
                        Write-Host "✗ Error removing Litigation Hold: $_" -ForegroundColor Red
                    }
                }
                
                # Remove In-Place Holds
                if ($mailboxInfo.InPlaceHolds -and $mailboxInfo.InPlaceHolds.Count -gt 0) {
                    Write-Host "Removing In-Place Holds..." -ForegroundColor Yellow
                    
                    foreach ($holdPolicy in $holdPolicies | Where-Object { $_.HoldId }) {
                        $hold = $holdPolicy.HoldId
                        
                        if ($hold -match '^UniH') {
                            # Retention Policy - need to exclude user
                            $policyId = $hold.Replace('UniH', '')
                            try {
                                $policy = Get-RetentionCompliancePolicy | Where-Object { $_.Guid -eq $policyId }
                                if ($policy) {
                                    Write-Host "  Excluding user from Retention Policy: $($policy.Name)..." -ForegroundColor Yellow
                                    
                                    # Get current exclusions and add this mailbox
                                    $currentExclusions = @()
                                    if ($policy.ExchangeLocationException) {
                                        $currentExclusions = @($policy.ExchangeLocationException)
                                    }
                                    
                                    # Only add if not already excluded
                                    if ($currentExclusions -notcontains $Mailbox) {
                                        $newExclusions = $currentExclusions + $Mailbox
                                        Set-RetentionCompliancePolicy -Identity $policy.Identity -ExchangeLocationException $newExclusions
                                        Write-Host "  ✓ User excluded from $($policy.Name)" -ForegroundColor Green
                                    } else {
                                        Write-Host "  ℹ User already excluded from $($policy.Name)" -ForegroundColor Cyan
                                    }
                                }
                            } catch {
                                Write-Host "  ✗ Error excluding from retention policy: $_" -ForegroundColor Red
                            }
                        }
                        elseif ($hold -match '^mbx') {
                            # eDiscovery Hold
                            $policyId = $hold.Replace('mbx', '')
                            try {
                                $casePolicy = Get-CaseHoldPolicy | Where-Object { $_.Guid -eq $policyId }
                                if ($casePolicy) {
                                    Write-Host "  Removing mailbox from eDiscovery Hold: $($casePolicy.Name)..." -ForegroundColor Yellow
                                    
                                    # Check if mailbox is in the policy
                                    $locations = @($casePolicy.ExchangeLocation)
                                    if ($locations -contains $Mailbox) {
                                        Set-CaseHoldPolicy -Identity $casePolicy.Identity -RemoveExchangeLocation $Mailbox
                                        Write-Host "  ✓ User removed from $($casePolicy.Name)" -ForegroundColor Green
                                    } else {
                                        Write-Host "  ℹ User not found in $($casePolicy.Name)" -ForegroundColor Cyan
                                    }
                                }
                            } catch {
                                Write-Host "  ✗ Error removing from eDiscovery hold: $_" -ForegroundColor Red
                            }
                        }
                    }
                }
                
                Write-Host "`n✓ Hold policies processed. Waiting 30 seconds for changes to propagate..." -ForegroundColor Green
                Start-Sleep -Seconds 30
                
                # Verify holds were removed
                Write-Host "Verifying hold removal..." -ForegroundColor Cyan
                $updatedMailbox = Get-Mailbox -Identity $Mailbox
                
                $remainingHolds = @()
                if ($updatedMailbox.LitigationHoldEnabled -eq $true) {
                    $remainingHolds += "Litigation Hold"
                }
                
                # Check for active holds only (ignore exclusions)
                if ($updatedMailbox.InPlaceHolds -and $updatedMailbox.InPlaceHolds.Count -gt 0) {
                    $activeInPlaceHolds = @($updatedMailbox.InPlaceHolds | Where-Object { $_ -notmatch '^-' })
                    if ($activeInPlaceHolds.Count -gt 0) {
                        $remainingHolds += "In-Place Holds ($($activeInPlaceHolds.Count))"
                        Write-Host "  Active holds (not exclusions):" -ForegroundColor Yellow
                        foreach ($hold in $activeInPlaceHolds) {
                            Write-Host "    $hold" -ForegroundColor Yellow
                        }
                    }
                }
                
                if ($updatedMailbox.DelayHoldApplied -eq $true) {
                    $remainingHolds += "Delay Hold"
                }
                
                if ($remainingHolds.Count -eq 0) {
                    Write-Host "✓ All holds successfully removed" -ForegroundColor Green
                } else {
                    Write-Host "⚠ The following holds are still present:" -ForegroundColor Yellow
                    foreach ($hold in $remainingHolds) {
                        Write-Host "  - $hold" -ForegroundColor Yellow
                    }
                    
                    if ($updatedMailbox.DelayHoldApplied -eq $true) {
                        Write-Host "`n⚠ IMPORTANT: Delay Hold detected!" -ForegroundColor Red
                        Write-Host "Delay Hold can take up to 30 days to automatically clear." -ForegroundColor Yellow
                        Write-Host "However, you can try to remove it immediately with the following command:" -ForegroundColor Cyan
                        Write-Host "Set-Mailbox -Identity $Mailbox -RemoveDelayHoldApplied" -ForegroundColor White
                        Write-Host "`nDo you want to attempt removing Delay Hold now? (Y/N): " -NoNewline -ForegroundColor Cyan
                        $removeDelayHold = Read-Host
                        
                        if ($removeDelayHold -eq "Y" -or $removeDelayHold -eq "y") {
                            try {
                                Write-Host "Attempting to remove Delay Hold..." -ForegroundColor Yellow
                                Set-Mailbox -Identity $Mailbox -RemoveDelayHoldApplied
                                Write-Host "✓ Delay Hold removal command executed" -ForegroundColor Green
                                Write-Host "Waiting 30 seconds for changes to take effect..." -ForegroundColor Gray
                                Start-Sleep -Seconds 30
                                
                                # Verify again
                                $finalCheck = Get-Mailbox -Identity $Mailbox
                                if ($finalCheck.DelayHoldApplied -eq $false) {
                                    Write-Host "✓ Delay Hold successfully removed!" -ForegroundColor Green
                                } else {
                                    Write-Host "⚠ Delay Hold still present. May require more time to clear." -ForegroundColor Yellow
                                }
                            } catch {
                                Write-Host "✗ Error removing Delay Hold: $_" -ForegroundColor Red
                                Write-Host "The hold may require more time to clear naturally." -ForegroundColor Yellow
                            }
                        }
                    }
                    
                    Write-Host "`n⚠ WARNING: Purge operations may fail while holds are active." -ForegroundColor Yellow
                    Write-Host "Do you want to continue anyway? (Y/N): " -NoNewline -ForegroundColor Cyan
                    $continueWithHolds = Read-Host
                    
                    if ($continueWithHolds -ne "Y" -and $continueWithHolds -ne "y") {
                        Write-Host "Exiting script. Please resolve holds before retrying." -ForegroundColor Yellow
                        exit
                    }
                }
                Write-Host "`n"
            }
            "2" {
                Write-Host "`nProceeding with search only (holds remain active)..." -ForegroundColor Yellow
            }
            "3" {
                Write-Host "Exiting script..." -ForegroundColor Yellow
                exit
            }
            default {
                Write-Host "Invalid choice. Exiting script." -ForegroundColor Red
                exit
            }
        }
    }
}
catch {
    Write-Host "Error checking mailbox: $_" -ForegroundColor Red
    exit
}

$SearchName = "RecoverableItemsSearch_DiscoveryHolds_$Mailbox"

# Check user RecoverableItems statistics
Write-Host "Checking RecoverableItems statistics..." -ForegroundColor Cyan
Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope RecoverableItems | Select-Object Name, FolderAndSubfolderSize, ItemsInFolderAndSubfolders

# Get ONLY the DiscoveryHolds folder - THIS IS THE KEY CHANGE
Write-Host "Retrieving DiscoveryHolds folder..." -ForegroundColor Cyan
$FolderIds = Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope RecoverableItems | Where-Object { $_.Name -eq "DiscoveryHolds" -and $_.FolderId -ne $null } | Select-Object FolderId, Name

# Verify we found the DiscoveryHolds folder
if ($FolderIds -eq $null -or $FolderIds.Count -eq 0) {
    Write-Host "DiscoveryHolds folder not found or is empty for $Mailbox." -ForegroundColor Yellow
    Write-Host "This may be normal if there are no items in DiscoveryHolds." -ForegroundColor Gray
    
    # Show what folders were found
    Write-Host "`nAvailable RecoverableItems folders:" -ForegroundColor Cyan
    Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope RecoverableItems | Select-Object Name, ItemsInFolderAndSubfolders | Format-Table -AutoSize
    
    exit
}

Write-Host "✓ Found DiscoveryHolds folder" -ForegroundColor Green

# Define the nibbler array (hexadecimal characters)
$nibbler = [byte[]]@(0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66)

# Initialize query string
$Query = ""

# Process each folder and build the search query
Write-Host "Building search query for DiscoveryHolds..." -ForegroundColor Cyan
foreach ($Folder in $FolderIds) {
    Write-Host "Processing folder: $($Folder.Name)" -ForegroundColor Yellow
    
    $folderIdBase64 = $Folder.FolderId
    $folderIdBytes = [Convert]::FromBase64String($folderIdBase64)
    $indexIdBytes = New-Object byte[] 48
    $indexIdIdx = 0
    
    $folderIdBytes | Select-Object -Skip 23 -First 24 | ForEach-Object {
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF]
    }
    
    $folderId = [System.Text.Encoding]::ASCII.GetString($indexIdBytes)
    
    # Add to query with OR operator if not the first folder
    if ($Query -ne "") {
        $Query += " OR "
    }
    $Query += "folderid:$folderId"
}

# Check if we found any folders to search
if ($Query -eq "") {
    Write-Host "No DiscoveryHolds folder ID could be generated. Exiting script." -ForegroundColor Red
    exit
}

Write-Host "Query generated: $Query" -ForegroundColor Gray

# Create and start the compliance search
Write-Host "Creating compliance search '$SearchName'..." -ForegroundColor Green
$existingSearch = Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue
if ($existingSearch) {
    Write-Host "A search with name '$SearchName' already exists. Removing it before creating a new one." -ForegroundColor Yellow
    Remove-ComplianceSearch -Identity $SearchName -Confirm:$false
}

New-ComplianceSearch -Name $SearchName -ExchangeLocation $Mailbox -ContentMatchQuery $Query

Write-Host "Starting compliance search..." -ForegroundColor Green
Start-ComplianceSearch -Identity $SearchName

# Monitor search status
Write-Host "Monitoring search status (press Ctrl+C to exit monitoring)..." -ForegroundColor Cyan
do {
    Start-Sleep -Seconds 5
    $searchStatus = Get-ComplianceSearch -Identity $SearchName
    Write-Host "Search status: $($searchStatus.Status)" -ForegroundColor Yellow
} while ($searchStatus.Status -ne "Completed")

# Show results
Write-Host "Search completed. Details:" -ForegroundColor Green
Start-Sleep -Seconds 2
$searchResults = Get-ComplianceSearch -Identity $SearchName
$searchResults | Format-List Name, Status, Items, Size

# Display summary
Write-Host "`n"
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "SEARCH RESULTS SUMMARY (DiscoveryHolds Only)" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "Mailbox: " -NoNewline -ForegroundColor White
Write-Host $Mailbox -ForegroundColor Yellow
Write-Host "Number of items found: " -NoNewline -ForegroundColor White
Write-Host $searchResults.Items -ForegroundColor Yellow
Write-Host "Total size: " -NoNewline -ForegroundColor White
Write-Host $searchResults.Size -ForegroundColor Yellow
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "`n"

# Simplified version WITHOUT loop
# Use this if you're confident holds are removed and just want to submit the purge

# ... [All your existing hold checking and removal code] ...
# ... [All your existing search creation code] ...

# After search completes and shows results:

if ($searchResults.Items -gt 0) {
    Write-Host "Do you want to PURGE these items from DiscoveryHolds? This action is IRREVERSIBLE!" -ForegroundColor Red
    Write-Host "Type 'YES' to confirm purge, or anything else to exit: " -NoNewline -ForegroundColor Yellow
    $confirmation = Read-Host
    
    if ($confirmation -eq "YES") {
        Write-Host "`nSubmitting purge request for $($searchResults.Items) items..." -ForegroundColor Red
        Write-Host "Note: Microsoft will process these at ~10 items per minute." -ForegroundColor Yellow
        Write-Host "Estimated completion time: $([math]::Round($searchResults.Items / 10 / 60, 1)) hours" -ForegroundColor Yellow
        Write-Host "`nThis is a ONE-TIME submission. The purge will continue in the background." -ForegroundColor Cyan
        Write-Host "You can close this window after the purge action is created.`n" -ForegroundColor Cyan
        
        try {
            # Submit the purge (one time only)
            $purgeActionName = "${SearchName}_Purge"
            
            # Remove any existing purge action
            $existingPurgeAction = Get-ComplianceSearchAction -Identity $purgeActionName -ErrorAction SilentlyContinue
            if ($existingPurgeAction) {
                Write-Host "Removing previous purge action..." -ForegroundColor Yellow
                Remove-ComplianceSearchAction -Identity $purgeActionName -Confirm:$false
                Start-Sleep -Seconds 5
            }
            
            # Create the purge action
            Write-Host "Creating purge action..." -ForegroundColor Yellow
            New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType HardDelete -Confirm:$false | Out-Null
            
            # Wait for it to start
            Write-Host "Waiting for purge to initialize..." -ForegroundColor Cyan
            Start-Sleep -Seconds 30
            
            # Check initial status
            $purgeStatus = Get-ComplianceSearchAction -Identity $purgeActionName
            
            Write-Host "`n" 
            Write-Host "=" * 60 -ForegroundColor Green
            Write-Host "PURGE SUBMITTED SUCCESSFULLY" -ForegroundColor Green
            Write-Host "=" * 60 -ForegroundColor Green
            Write-Host "Purge Action Name: " -NoNewline -ForegroundColor White
            Write-Host $purgeActionName -ForegroundColor Yellow
            Write-Host "Status: " -NoNewline -ForegroundColor White
            Write-Host $purgeStatus.Status -ForegroundColor Yellow
            Write-Host "Items to delete: " -NoNewline -ForegroundColor White
            Write-Host $searchResults.Items -ForegroundColor Yellow
            Write-Host "`nThe purge will continue in the background." -ForegroundColor Cyan
            Write-Host "You can monitor progress with these commands:" -ForegroundColor Cyan
            Write-Host "`n  Get-ComplianceSearchAction -Identity '$purgeActionName' | FL Status,Results" -ForegroundColor White
            Write-Host "  Start-ComplianceSearch -Identity '$SearchName' -Force" -ForegroundColor White
            Write-Host "  Get-ComplianceSearch -Identity '$SearchName' | FL Items,Size" -ForegroundColor White
            Write-Host "`nCheck back in $([math]::Round($searchResults.Items / 10 / 60, 1)) hours to verify completion." -ForegroundColor Yellow
            Write-Host "=" * 60 -ForegroundColor Green
            
        } catch {
            Write-Host "✗ Error creating purge action: $_" -ForegroundColor Red
            if ($_.Exception.Message -like "*hold*") {
                Write-Host "`nThis error suggests holds may still be active." -ForegroundColor Yellow
                Write-Host "Run this command to check: Get-Mailbox -Identity $Mailbox | FL *Hold*" -ForegroundColor Cyan
            }
            elseif ($_.Exception.Message -like "*EnableSearchOnlySession*") {
                Write-Host "`nERROR: The PowerShell session was not created with -EnableSearchOnlySession flag." -ForegroundColor Red
                Write-Host "Please disconnect and reconnect using: Connect-IPPSSession -EnableSearchOnlySession" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "Purge cancelled. Items remain in DiscoveryHolds." -ForegroundColor Yellow
    }
} else {
    Write-Host "No items found to purge in DiscoveryHolds." -ForegroundColor Green
}

Write-Host "`nScript complete. You can now close this window." -ForegroundColor Green
