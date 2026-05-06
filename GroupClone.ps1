<#
.SYNOPSIS
Clones group memberships, group ownerships, and mailbox delegations
from one user to another using Microsoft Graph and Exchange Online.

.DESCRIPTION
- Graph handles security + M365 (Unified) groups
- Mail-enabled DLs and mail-enabled security groups are deferred to EXO
- Copies mailbox permissions (FullAccess, SendAs, SendOnBehalf)
- -WhatIf supported everywhere
- -ShowDebug enables verbose enumeration output
- CSV export of all actions
#>

[CmdletBinding()]
param (
    [string]$SourceUser,
    [string]$TargetUser,
    [switch]$WhatIf,
    [switch]$ShowDebug
)

#region Helper Functions
function Prompt-IfMissing {
    param ($Value, $Prompt)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        Read-Host $Prompt
    } else {
        $Value
    }
}

function Ensure-Module {
    param ([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -ErrorAction Stop
}
#endregion

#region Input
$SourceUser = Prompt-IfMissing $SourceUser "Enter SOURCE user UPN"
$TargetUser = Prompt-IfMissing $TargetUser "Enter TARGET user UPN"
#endregion

#region Tracking
$Results   = [System.Collections.Generic.List[PSCustomObject]]::new()
$EXOGroups = [System.Collections.Generic.List[PSCustomObject]]::new()
#endregion

#region Modules
Ensure-Module Microsoft.Graph.Authentication
Ensure-Module Microsoft.Graph.Users
Ensure-Module Microsoft.Graph.Groups
Ensure-Module ExchangeOnlineManagement
#endregion

#region Connect
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes @(
    "User.Read.All",
    "Group.ReadWrite.All",
    "Directory.ReadWrite.All"
) -NoWelcome

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false -Device
#endregion

#region Resolve Users
$Source = Get-MgUser -UserId $SourceUser -ErrorAction Stop
$Target = Get-MgUser -UserId $TargetUser -ErrorAction Stop
#endregion

#region Graph Group Memberships
Write-Host "`nCopying group memberships (Graph)..." -ForegroundColor Cyan

$Groups = Get-MgUserMemberOf -UserId $Source.Id -All
if ($ShowDebug) { Write-Host "[DEBUG] MemberOf count: $($Groups.Count)" -ForegroundColor Magenta }

foreach ($Obj in $Groups) {

    if ($Obj.AdditionalProperties['@odata.type'] -ne "#microsoft.graph.group") { continue }

    $Group = Get-MgGroup -GroupId $Obj.Id -ErrorAction SilentlyContinue
    if (-not $Group) {
        Write-Host "[ERROR] Unable to resolve group $($Obj.Id)" -ForegroundColor Red
        continue
    }

    $IsDL = $Group.MailEnabled -and ($Group.GroupTypes -notcontains "Unified")

    if ($ShowDebug) {
        Write-Host "[DEBUG] $($Group.DisplayName) | MailEnabled=$($Group.MailEnabled) | GroupTypes=$($Group.GroupTypes -join ',')" -ForegroundColor DarkGray
    }

    if ($IsDL) {
        $EXOGroups.Add([PSCustomObject]@{
            Name = $Group.DisplayName
            Id   = $Group.Id
            Type = "Membership"
        })
        continue
    }

    if ($WhatIf) {
        Write-Host "[WhatIf] Would add member: $($Group.DisplayName)" -ForegroundColor Yellow
        $Results.Add([PSCustomObject]@{ Name=$Group.DisplayName; Action="Graph Member"; Status="WhatIf" })
        continue
    }

    try {
        New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $Target.Id -ErrorAction Stop
        Write-Host "[OK] Added member: $($Group.DisplayName)" -ForegroundColor Green
        $Results.Add([PSCustomObject]@{ Name=$Group.DisplayName; Action="Graph Member"; Status="Copied" })
    } catch {
        Write-Host "[FAIL] $($Group.DisplayName) — $($_.Exception.Message)" -ForegroundColor Red
        $Results.Add([PSCustomObject]@{ Name=$Group.DisplayName; Action="Graph Member"; Status="Failed"; Error=$_.Exception.Message })
    }
}
#endregion

#region Graph Group Ownerships
Write-Host "`nCopying group ownerships (Graph)..." -ForegroundColor Cyan

$OwnedGroups = Get-MgUserOwnedObject -UserId $Source.Id -All |
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq "#microsoft.graph.group" }

foreach ($Obj in $OwnedGroups) {

    $Group = Get-MgGroup -GroupId $Obj.Id -ErrorAction SilentlyContinue
    if (-not $Group) { continue }

    $IsDL = $Group.MailEnabled -and ($Group.GroupTypes -notcontains "Unified")

    if ($IsDL) {
        $EXOGroups.Add([PSCustomObject]@{
            Name = $Group.DisplayName
            Id   = $Group.Id
            Type = "Ownership"
        })
        continue
    }

    if ($WhatIf) {
        Write-Host "[WhatIf] Would add owner: $($Group.DisplayName)" -ForegroundColor Yellow
        $Results.Add([PSCustomObject]@{ Name=$Group.DisplayName; Action="Graph Owner"; Status="WhatIf" })
        continue
    }

    try {
        New-MgGroupOwner -GroupId $Group.Id -DirectoryObjectId $Target.Id -ErrorAction Stop
        Write-Host "[OK] Added owner: $($Group.DisplayName)" -ForegroundColor Green
        $Results.Add([PSCustomObject]@{ Name=$Group.DisplayName; Action="Graph Owner"; Status="Copied" })
    } catch {
        Write-Host "[FAIL] Owner $($Group.DisplayName) — $($_.Exception.Message)" -ForegroundColor Red
        $Results.Add([PSCustomObject]@{ Name=$Group.DisplayName; Action="Graph Owner"; Status="Failed"; Error=$_.Exception.Message })
    }
}
#endregion

#region EXO Distribution Groups
if ($EXOGroups.Count -gt 0) {
    Write-Host "`nProcessing distribution groups (EXO)..." -ForegroundColor Cyan

    foreach ($DG in $EXOGroups) {

        if ($WhatIf) {
            Write-Host "[WhatIf] Would update DL: $($DG.Name)" -ForegroundColor Yellow
            $Results.Add([PSCustomObject]@{ Name=$DG.Name; Action="EXO $($DG.Type)"; Status="WhatIf" })
            continue
        }

        try {
            if ($DG.Type -eq "Membership") {
                Add-DistributionGroupMember -Identity $DG.Name -Member $TargetUser -ErrorAction Stop
            }
            if ($DG.Type -eq "Ownership") {
                Set-DistributionGroup -Identity $DG.Name -ManagedBy @{Add=$TargetUser} -ErrorAction Stop
            }
            Write-Host "[OK] EXO $($DG.Type): $($DG.Name)" -ForegroundColor Green
            $Results.Add([PSCustomObject]@{ Name=$DG.Name; Action="EXO $($DG.Type)"; Status="Copied" })
        } catch {
            Write-Host "[FAIL] EXO $($DG.Name) — $($_.Exception.Message)" -ForegroundColor Red
            $Results.Add([PSCustomObject]@{ Name=$DG.Name; Action="EXO $($DG.Type)"; Status="Failed"; Error=$_.Exception.Message })
        }
    }
}
#endregion

#region Mailbox Delegations
Write-Host "`nCopying mailbox delegations..." -ForegroundColor Cyan

$Mailboxes = Get-Mailbox -ResultSize Unlimited |
    Where-Object { $_.RecipientTypeDetails -in @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox") }

foreach ($MB in $Mailboxes) {

    # FullAccess
    $FA = Get-MailboxPermission $MB.Identity |
        Where-Object { $_.User -like $SourceUser -and $_.AccessRights -contains "FullAccess" }

    if ($FA) {
        if ($WhatIf) {
            Write-Host "[WhatIf] FullAccess: $($MB.DisplayName)" -ForegroundColor Yellow
        } else {
            try {
                Add-MailboxPermission -Identity $MB.Identity -User $TargetUser -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop
                Write-Host "[OK] FullAccess: $($MB.DisplayName)" -ForegroundColor Green
            } catch {
                Write-Host "[FAIL] FullAccess $($MB.DisplayName) — $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # SendAs
    $SA = Get-RecipientPermission $MB.Identity |
        Where-Object { $_.Trustee -like $SourceUser -and $_.AccessRights -contains "SendAs" }

    if ($SA) {
        if ($WhatIf) {
            Write-Host "[WhatIf] SendAs: $($MB.DisplayName)" -ForegroundColor Yellow
        } else {
            try {
                Add-RecipientPermission -Identity $MB.Identity -Trustee $TargetUser -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                Write-Host "[OK] SendAs: $($MB.DisplayName)" -ForegroundColor Green
            } catch {
                Write-Host "[FAIL] SendAs $($MB.DisplayName) — $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # SendOnBehalf
    if ($MB.GrantSendOnBehalfTo -contains $Source.DisplayName) {
        if ($WhatIf) {
            Write-Host "[WhatIf] SendOnBehalf: $($MB.DisplayName)" -ForegroundColor Yellow
        } else {
            try {
                Set-Mailbox $MB.Identity -GrantSendOnBehalfTo @{Add=$TargetUser} -ErrorAction Stop
                Write-Host "[OK] SendOnBehalf: $($MB.DisplayName)" -ForegroundColor Green
            } catch {
                Write-Host "[FAIL] SendOnBehalf $($MB.DisplayName) — $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
#endregion

#region Export CSV
$CsvPath = ".\GroupClone_Results_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
$Results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvPath
Write-Host "`nResults exported to $CsvPath" -ForegroundColor Cyan
#endregion

#region Cleanup
Disconnect-MgGraph | Out-Null
Disconnect-ExchangeOnline -Confirm:$false
#endregion

Write-Host "`nCompleted successfully." -ForegroundColor Green