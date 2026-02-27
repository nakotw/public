<#  ------------------------------------------------------------
 Friendly On-Screen UX Version (same logic, better output)
-------------------------------------------------------------#>

# -----------------------------
# UI helpers (friendly output)
# -----------------------------
function Write-Banner {
    param([string]$Title)
    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor Cyan
    Write-Host ("  " + $Title) -ForegroundColor Cyan
    Write-Host ("=" * 72) -ForegroundColor Cyan
    Write-Host ""
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("-" * 72) -ForegroundColor DarkCyan
    Write-Host ("  " + $Title) -ForegroundColor Cyan
    Write-Host ("-" * 72) -ForegroundColor DarkCyan
}

function Write-Step {
    param(
        [int]$Number,
        [string]$Text
    )
    Write-Host ("[$Number] " + $Text) -ForegroundColor Cyan
}

function Write-Info {
    param([string]$Text)
    Write-Host ("ℹ  " + $Text) -ForegroundColor Gray
}

function Write-Ok {
    param([string]$Text)
    Write-Host ("✅ " + $Text) -ForegroundColor Green
}

function Write-Warn {
    param([string]$Text)
    Write-Host ("⚠️  " + $Text) -ForegroundColor Yellow
}

function Write-Err {
    param([string]$Text)
    Write-Host ("❌ " + $Text) -ForegroundColor Red
}

function Ask-Choice {
    param(
        [string]$Title,
        [hashtable]$Options
    )

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
    foreach ($k in ($Options.Keys | Sort-Object)) {
        Write-Host ("  {0}. {1}" -f $k, $Options[$k]) -ForegroundColor Yellow
    }
    Write-Host ""

    do {
        $choice = Read-Host "Your choice"
        if (-not $Options.ContainsKey($choice)) {
            Write-Warn "Please enter one of: $($Options.Keys -join ', ')"
        }
    } while (-not $Options.ContainsKey($choice))

    return $choice
}

function Ask-ConfirmExact {
    param(
        [string]$Prompt,
        [string]$Exact = "YES"
    )
    Write-Host ""
    Write-Warn $Prompt
    $v = Read-Host "Type '$Exact' to confirm"
    return ($v -eq $Exact)
}

function Show-SpinnerWhile {
    param(
        [scriptblock]$ConditionScript,
        [int]$IntervalSeconds = 3,
        [string]$Label = "Working"
    )

    $spinner = @('|','/','-','\')
    $i = 0
    while (& $ConditionScript) {
        $c = $spinner[$i % $spinner.Count]
        Write-Host -NoNewline "`r$c $Label... " -ForegroundColor Yellow
        Start-Sleep -Seconds $IntervalSeconds
        $i++
    }
    Write-Host "`r✅ $Label... done.     " -ForegroundColor Green
}

# -----------------------------
# Connect sessions
# -----------------------------
Write-Banner "Exchange / Purview DiscoveryHolds Cleanup"

Write-Step 1 "Connecting to Exchange Online + Purview session..."
try {
    Connect-ExchangeOnline -ErrorAction Stop | Out-Null
    Connect-IPPSSession -EnableSearchOnlySession -ErrorAction Stop | Out-Null
    Write-Ok "Connected successfully."
} catch {
    Write-Err "Connection failed: $($_.Exception.Message)"
    return
}

# -----------------------------
# Prompt mailbox
# -----------------------------
Write-Section "Mailbox selection"
Write-Info "Tip: use the full UPN (example: user@company.com)."

$Mailbox = Read-Host "Enter the mailbox email address"
if ($Mailbox -notmatch '^[\w\.-]+@[\w\.-]+\.\w+$') {
    Write-Err "That doesn't look like a valid email format. Exiting."
    return
}

Write-Ok "Target mailbox: $Mailbox"

# -----------------------------
# Add mailbox to main hold policy exception
# -----------------------------
Write-Section "Retention policy: adding mailbox to MAIN hold policy exception"
Write-Step 2 "Updating retention policy exception list..."
try {
    Set-RetentionCompliancePolicy `
      -Identity "7 years retention policy" `
      -AddExchangeLocationException $Mailbox

    Write-Ok "$Mailbox added to the main hold policy exception list."
} catch {
    Write-Err "Failed to update retention policy: $($_.Exception.Message)"
    return
}

# -----------------------------
# Check holds
# -----------------------------
Write-Section "Checking holds (Litigation / In-Place / Delay Hold)"
Write-Step 3 "Collecting mailbox hold information..."

try {
    $mailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
    $hasHolds = $false
    $holdPolicies = @()

    # Litigation Hold
    if ($mailboxInfo.LitigationHoldEnabled -eq $true) {
        $hasHolds = $true
        Write-Warn "Litigation Hold is ENABLED (may block purge)."
        $holdPolicies += [PSCustomObject]@{
            Type    = "Litigation Hold"
            Name    = "Litigation Hold"
            Details = "Enabled on: $($mailboxInfo.LitigationHoldDate)"
        }
    } else {
        Write-Ok "Litigation Hold: not enabled."
    }

    # In-Place Holds
    if ($mailboxInfo.InPlaceHolds -and $mailboxInfo.InPlaceHolds.Count -gt 0) {

        Write-Info "In-Place Holds found: $($mailboxInfo.InPlaceHolds.Count)"
        $activeHoldCount = 0

        foreach ($hold in $mailboxInfo.InPlaceHolds) {

            if ($hold -match '^-UniH') {
                Write-Ok "Retention Policy exclusion detected: $hold (this will NOT block deletion)"
                continue
            }
            elseif ($hold -match '^-mbx') {
                Write-Ok "eDiscovery Hold exclusion detected: $hold (this will NOT block deletion)"
                continue
            }
            elseif ($hold -match '^UniH') {
                $activeHoldCount++
                $policyName = "Retention Policy"
                try {
                    $policy = Get-RetentionCompliancePolicy | Where-Object { $_.Guid -eq $hold.Replace('UniH', '') }
                    if ($policy) { $policyName = $policy.Name }
                } catch { }

                $holdPolicies += [PSCustomObject]@{
                    Type    = "Retention Policy"
                    Name    = $policyName
                    Details = "Hold ID: $hold"
                    HoldId  = $hold
                }
                Write-Warn "Active Retention Hold: $policyName"
            }
            elseif ($hold -match '^mbx') {
                $activeHoldCount++
                $policyName = "eDiscovery Hold"
                try {
                    $case = Get-CaseHoldPolicy | Where-Object { $_.Guid -eq $hold.Replace('mbx', '') }
                    if ($case) { $policyName = $case.Name }
                } catch { }

                $holdPolicies += [PSCustomObject]@{
                    Type    = "eDiscovery Hold"
                    Name    = $policyName
                    Details = "Hold ID: $hold"
                    HoldId  = $hold
                }
                Write-Warn "Active eDiscovery Hold: $policyName"
            }
            else {
                $activeHoldCount++
                $holdPolicies += [PSCustomObject]@{
                    Type    = "Unknown Hold"
                    Name    = "Unknown"
                    Details = "Hold ID: $hold"
                    HoldId  = $hold
                }
                Write-Warn "Active Unknown Hold: $hold"
            }
        }

        if ($activeHoldCount -gt 0) { $hasHolds = $true }
        if ($activeHoldCount -eq 0) { Write-Ok "Only exclusions found (no active In-Place holds blocking purge)." }
    } else {
        Write-Ok "In-Place Holds: none found."
    }

    # Delay Hold
    if ($mailboxInfo.DelayHoldApplied -eq $true) {
        $hasHolds = $true
        Write-Warn "Delay Hold is APPLIED (can take time to clear; may block purge)."
        Write-Info "Delay Hold is usually temporary after hold removal (can take up to ~30 days)."
        $holdPolicies += [PSCustomObject]@{
            Type    = "Delay Hold"
            Name    = "Delay Hold"
            Details = "Temporary hold after policy removal"
        }
    } else {
        Write-Ok "Delay Hold: not applied."
    }

    if (-not $hasHolds) {
        Write-Ok "No active holds detected that should block purge."
    } else {
        Write-Warn "This mailbox has holds that may prevent deletion."
    }

    # Hold action choice
    if ($hasHolds) {
        $choice = Ask-Choice -Title "What would you like to do next?" -Options @{
            "1" = "Try to remove/neutralize holds, then continue (recommended if you need to purge)"
            "2" = "Continue with search only (purge will likely fail)"
            "3" = "Exit safely"
        }

        switch ($choice) {
            "1" {
                Write-Section "Hold removal / exclusions"
                Write-Step 4 "Processing hold changes..."

                # Remove Litigation Hold
                if ($mailboxInfo.LitigationHoldEnabled -eq $true) {
                    Write-Info "Removing Litigation Hold..."
                    try {
                        Set-Mailbox -Identity $Mailbox -LitigationHoldEnabled $false
                        Write-Ok "Litigation Hold removed."
                    } catch {
                        Write-Err "Error removing Litigation Hold: $($_.Exception.Message)"
                    }
                }

                # Remove In-Place Holds
                if ($mailboxInfo.InPlaceHolds -and $mailboxInfo.InPlaceHolds.Count -gt 0) {

                    foreach ($holdPolicy in $holdPolicies | Where-Object { $_.HoldId }) {
                        $hold = $holdPolicy.HoldId

                        if ($hold -match '^UniH') {
                            $policyId = $hold.Replace('UniH', '')
                            try {
                                $policy = Get-RetentionCompliancePolicy | Where-Object { $_.Guid -eq $policyId }
                                if ($policy) {
                                    Write-Info "Excluding mailbox from Retention Policy: $($policy.Name)"
                                    $currentExclusions = @()
                                    if ($policy.ExchangeLocationException) { $currentExclusions = @($policy.ExchangeLocationException) }

                                    if ($currentExclusions -notcontains $Mailbox) {
                                        $newExclusions = $currentExclusions + $Mailbox
                                        Set-RetentionCompliancePolicy -Identity $policy.Identity -ExchangeLocationException $newExclusions
                                        Write-Ok "Excluded from $($policy.Name)."
                                    } else {
                                        Write-Info "Already excluded from $($policy.Name)."
                                    }
                                }
                            } catch {
                                Write-Err "Error excluding from retention policy: $($_.Exception.Message)"
                            }
                        }
                        elseif ($hold -match '^mbx') {
                            $policyId = $hold.Replace('mbx', '')
                            try {
                                $casePolicy = Get-CaseHoldPolicy | Where-Object { $_.Guid -eq $policyId }
                                if ($casePolicy) {
                                    Write-Info "Removing mailbox from eDiscovery Hold: $($casePolicy.Name)"
                                    $locations = @($casePolicy.ExchangeLocation)

                                    if ($locations -contains $Mailbox) {
                                        Set-CaseHoldPolicy -Identity $casePolicy.Identity -RemoveExchangeLocation $Mailbox
                                        Write-Ok "Removed from $($casePolicy.Name)."
                                    } else {
                                        Write-Info "Mailbox not listed in $($casePolicy.Name)."
                                    }
                                }
                            } catch {
                                Write-Err "Error removing from eDiscovery hold: $($_.Exception.Message)"
                            }
                        }
                    }
                }

                Write-Info "Waiting 30 seconds for changes to propagate..."
                Start-Sleep -Seconds 30

                Write-Step 5 "Verifying remaining holds..."
                $updatedMailbox = Get-Mailbox -Identity $Mailbox

                $remainingHolds = @()
                if ($updatedMailbox.LitigationHoldEnabled -eq $true) { $remainingHolds += "Litigation Hold" }

                if ($updatedMailbox.InPlaceHolds -and $updatedMailbox.InPlaceHolds.Count -gt 0) {
                    $activeInPlaceHolds = @($updatedMailbox.InPlaceHolds | Where-Object { $_ -notmatch '^-' })
                    if ($activeInPlaceHolds.Count -gt 0) {
                        $remainingHolds += "In-Place Holds ($($activeInPlaceHolds.Count))"
                        Write-Warn "Active In-Place holds still present:"
                        foreach ($h in $activeInPlaceHolds) { Write-Host ("   - " + $h) -ForegroundColor Yellow }
                    }
                }

                if ($updatedMailbox.DelayHoldApplied -eq $true) { $remainingHolds += "Delay Hold" }

                if ($remainingHolds.Count -eq 0) {
                    Write-Ok "All blocking holds appear cleared."
                } else {
                    Write-Warn "Some holds are still present: $($remainingHolds -join ', ')"

                    if ($updatedMailbox.DelayHoldApplied -eq $true) {
                        Write-Warn "Delay Hold detected."
                        Write-Info "You can attempt immediate removal with:"
                        Write-Host "  Set-Mailbox -Identity $Mailbox -RemoveDelayHoldApplied" -ForegroundColor White

                        $tryDelay = Read-Host "Attempt RemoveDelayHoldApplied now? (Y/N)"
                        if ($tryDelay -match '^(Y|y)$') {
                            try {
                                Set-Mailbox -Identity $Mailbox -RemoveDelayHoldApplied
                                Write-Ok "Delay Hold removal command executed."
                                Start-Sleep -Seconds 30
                            } catch {
                                Write-Err "Error removing Delay Hold: $($_.Exception.Message)"
                            }
                        }
                    }

                    $cont = Read-Host "Continue anyway (purge may fail)? (Y/N)"
                    if ($cont -notmatch '^(Y|y)$') {
                        Write-Ok "Exiting safely. Resolve holds then re-run."
                        return
                    }
                }
            }
            "2" {
                Write-Warn "Continuing with SEARCH ONLY. Purge will likely fail while holds remain."
            }
            "3" {
                Write-Ok "Exiting safely. No changes beyond what already ran."
                return
            }
        }
    }

} catch {
    Write-Err "Error checking mailbox: $($_.Exception.Message)"
    return
}

# -----------------------------
# Search DiscoveryHolds only
# -----------------------------
$SearchName = "RecoverableItemsSearch_DiscoveryHolds_$Mailbox"

Write-Section "DiscoveryHolds search scope"
Write-Step 6 "Showing RecoverableItems folder statistics (for visibility)"
Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope RecoverableItems |
    Select-Object Name, FolderAndSubfolderSize, ItemsInFolderAndSubfolders |
    Format-Table -AutoSize

Write-Step 7 "Retrieving DiscoveryHolds folder id..."
$FolderIds = Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope RecoverableItems |
    Where-Object { $_.Name -eq "DiscoveryHolds" -and $_.FolderId -ne $null } |
    Select-Object FolderId, Name

if ($FolderIds -eq $null -or $FolderIds.Count -eq 0) {
    Write-Warn "DiscoveryHolds folder not found (or empty). This can be normal."
    Write-Info "Available RecoverableItems folders:"
    Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope RecoverableItems |
        Select-Object Name, ItemsInFolderAndSubfolders |
        Format-Table -AutoSize
    return
}

Write-Ok "DiscoveryHolds folder found."

# Build folderid query
$nibbler = [byte[]]@(0x30,0x31,0x32,0x33,0x34,0x35,0x36,0x37,0x38,0x39,0x61,0x62,0x63,0x64,0x65,0x66)
$Query = ""

Write-Step 8 "Building ContentMatchQuery..."
foreach ($Folder in $FolderIds) {
    Write-Info "Processing: $($Folder.Name)"
    $folderIdBytes = [Convert]::FromBase64String($Folder.FolderId)
    $indexIdBytes = New-Object byte[] 48
    $indexIdIdx = 0

    $folderIdBytes | Select-Object -Skip 23 -First 24 | ForEach-Object {
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF]
    }

    $folderId = [System.Text.Encoding]::ASCII.GetString($indexIdBytes)
    if ($Query -ne "") { $Query += " OR " }
    $Query += "folderid:$folderId"
}

if ([string]::IsNullOrWhiteSpace($Query)) {
    Write-Err "Could not generate folder query. Exiting."
    return
}

Write-Info "Query: $Query"

# Create search
Write-Section "Compliance Search"
Write-Step 9 "Preparing compliance search: $SearchName"

$existingSearch = Get-ComplianceSearch -Identity $SearchName -ErrorAction SilentlyContinue
if ($existingSearch) {
    Write-Warn "A search with the same name already exists. Removing it first..."
    Remove-ComplianceSearch -Identity $SearchName -Confirm:$false
    Write-Ok "Old search removed."
}

try {
    New-ComplianceSearch -Name $SearchName -ExchangeLocation $Mailbox -ContentMatchQuery $Query | Out-Null
    Write-Ok "Compliance search created."
} catch {
    Write-Err "Failed to create compliance search: $($_.Exception.Message)"
    return
}

Write-Step 10 "Starting compliance search..."
Start-ComplianceSearch -Identity $SearchName | Out-Null

# Monitor with spinner
Show-SpinnerWhile -Label "Compliance search running" -IntervalSeconds 5 -ConditionScript {
    $s = Get-ComplianceSearch -Identity $SearchName
    return ($s.Status -ne "Completed")
}

# Results
Write-Section "Search results (DiscoveryHolds only)"
$searchResults = Get-ComplianceSearch -Identity $SearchName
$searchResults | Format-List Name, Status, Items, Size

Write-Banner "SEARCH SUMMARY"
Write-Host ("Mailbox: " + $Mailbox) -ForegroundColor Yellow
Write-Host ("Items found: " + $searchResults.Items) -ForegroundColor Yellow
Write-Host ("Total size: " + $searchResults.Size) -ForegroundColor Yellow

# -----------------------------
# Purge (one-time submit)
# -----------------------------
if ($searchResults.Items -gt 0) {

    $doPurge = Ask-ConfirmExact -Prompt "Purge $($searchResults.Items) item(s) from DiscoveryHolds? This is IRREVERSIBLE." -Exact "YES"
    if (-not $doPurge) {
        Write-Warn "Purge cancelled. Items remain in DiscoveryHolds."
        Write-Banner "Done"
        Write-Ok "Script complete. You can close this window."
        return
    }

    Write-Section "Purge submission"
    Write-Step 11 "Submitting purge request..."
    Write-Info "Microsoft processes purge roughly ~10 items/minute (varies)."
    $estHours = [math]::Round($searchResults.Items / 10 / 60, 1)
    Write-Info "Rough estimate: ~$estHours hour(s)."

    try {
        $purgeActionName = "${SearchName}_Purge"

        $existingPurgeAction = Get-ComplianceSearchAction -Identity $purgeActionName -ErrorAction SilentlyContinue
        if ($existingPurgeAction) {
            Write-Warn "Previous purge action found. Removing it..."
            Remove-ComplianceSearchAction -Identity $purgeActionName -Confirm:$false
            Start-Sleep -Seconds 5
            Write-Ok "Previous purge action removed."
        }

        Write-Info "Creating new purge action..."
        New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType HardDelete -Confirm:$false | Out-Null

        Write-Info "Waiting 30 seconds for purge to initialize..."
        Start-Sleep -Seconds 30

        $purgeStatus = Get-ComplianceSearchAction -Identity $purgeActionName

        Write-Banner "PURGE SUBMITTED"
        Write-Ok "Purge Action Name: $purgeActionName"
        Write-Info "Status: $($purgeStatus.Status)"
        Write-Info "Items targeted: $($searchResults.Items)"
        Write-Info "Monitor progress with:"
        Write-Host "  Get-ComplianceSearchAction -Identity '$purgeActionName' | FL Status,Results" -ForegroundColor White
        Write-Host "  Get-ComplianceSearch -Identity '$SearchName' | FL Items,Size,Status" -ForegroundColor White
        Write-Host ""
        Write-Info "Check back in ~${estHours} hour(s) to verify completion." -ForegroundColor Yellow

    } catch {
        Write-Err "Error creating purge action: $($_.Exception.Message)"
        if ($_.Exception.Message -like "*hold*") {
            Write-Warn "This suggests holds may still be active."
            Write-Info "Run: Get-Mailbox -Identity $Mailbox | FL *Hold*"
        }
        if ($_.Exception.Message -like "*EnableSearchOnlySession*") {
            Write-Warn "Your session may not be connected with -EnableSearchOnlySession."
            Write-Info "Reconnect using: Connect-IPPSSession -EnableSearchOnlySession"
        }
        return
    }

} else {
    Write-Ok "No items found to purge in DiscoveryHolds."
}

Write-Banner "Done"
Write-Ok "Script complete. You can close this window."
