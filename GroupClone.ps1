<#
.SYNOPSIS
Copies group memberships, group ownerships, and Exchange mailbox delegations
from one user to another using Microsoft Graph (minimal modules) and EXO.
.DESCRIPTION
- Interactive by default
- Supports command-line parameters
- PS 7.6+ compatible
- Minimal Graph module imports
- Graceful install/import
- -WhatIf mode: simulates changes without applying them
- Exports copied and failed groups to CSV
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)][string]$SourceUser,
    [Parameter(Mandatory = $false)][string]$TargetUser,
    [switch]$WhatIf
)

#region Helper Functions
function Ensure-Module {
    param ([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing $Name..." -ForegroundColor Yellow
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -ErrorAction Stop
}

function Prompt-IfMissing {
    param ([string]$Value, [string]$Prompt)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return Read-Host $Prompt
    }
    return $Value
}
#endregion

#region Input Handling
$SourceUser = Prompt-IfMissing $SourceUser "Enter SOURCE user UPN"
$TargetUser = Prompt-IfMissing $TargetUser "Enter TARGET user UPN"
#endregion

#region Results Tracking
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
#endregion

#region Module Prep
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

#region Group Memberships
Write-Host "`nCopying group memberships..." -ForegroundColor Cyan

$Groups = Get-MgUserMemberOf -UserId $Source.Id -All
Write-Host "[DEBUG] Total MemberOf objects returned: $($Groups.Count)" -ForegroundColor Magenta

foreach ($Group in $Groups) {
    $ODataType = $Group.AdditionalProperties['@odata.type']
    Write-Host "[DEBUG] Object: $($Group.Id) | Type: $ODataType" -ForegroundColor DarkGray

    if ($ODataType -ne "#microsoft.graph.group") {
        Write-Host "[DEBUG]   Skipped (not a group)" -ForegroundColor DarkGray
        continue
    }

    $GroupDetail = Get-MgGroup -GroupId $Group.Id -ErrorAction SilentlyContinue
    if (-not $GroupDetail) {
        Write-Host "[DEBUG]   ERROR: Get-MgGroup returned null for $($Group.Id)" -ForegroundColor Red
    }
    $DisplayName = if ($GroupDetail) { $GroupDetail.DisplayName } else { $Group.Id }
    $GroupTypes  = if ($GroupDetail) { ($GroupDetail.GroupTypes -join ", ") } else { "unknown" }

    Write-Host "[DEBUG]   Name: $DisplayName | GroupTypes: $GroupTypes | MailEnabled: $($GroupDetail.MailEnabled) | SecurityEnabled: $($GroupDetail.SecurityEnabled)" -ForegroundColor DarkGray

    if ($WhatIf) {
        Write-Host "[WhatIf] Would add as member: $DisplayName ($($Group.Id))" -ForegroundColor Yellow
        $Results.Add([PSCustomObject]@{
            GroupId   = $Group.Id
            GroupName = $DisplayName
            Type      = "Membership"
            Status    = "WhatIf"
            Error     = ""
        })
    } else {
        try {
            New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $Target.Id -ErrorAction Stop
            Write-Host "[OK] Added as member: $DisplayName" -ForegroundColor Green
            $Results.Add([PSCustomObject]@{
                GroupId   = $Group.Id
                GroupName = $DisplayName
                Type      = "Membership"
                Status    = "Copied"
                Error     = ""
            })
        } catch {
            Write-Host "[FAIL] Member: $DisplayName — $($_.Exception.Message)" -ForegroundColor Red
            $Results.Add([PSCustomObject]@{
                GroupId   = $Group.Id
                GroupName = $DisplayName
                Type      = "Membership"
                Status    = "Failed"
                Error     = $_.Exception.Message
            })
        }
    }
}

Write-Host "[DEBUG] Membership pass complete. Groups processed: $(($Results | Where-Object Type -eq 'Membership').Count)" -ForegroundColor Magenta
#endregion

#region Group Ownerships
Write-Host "`nCopying group ownerships..." -ForegroundColor Cyan

$OwnedGroups = Get-MgUserOwnedObject -UserId $Source.Id -All |
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq "#microsoft.graph.group" }

Write-Host "[DEBUG] Owned groups (filtered to type group): $( ($OwnedGroups | Measure-Object).Count )" -ForegroundColor Magenta

foreach ($Group in $OwnedGroups) {
    $ODataType = $Group.AdditionalProperties['@odata.type']
    Write-Host "[DEBUG] Object: $($Group.Id) | Type: $ODataType" -ForegroundColor DarkGray

    $GroupDetail = Get-MgGroup -GroupId $Group.Id -ErrorAction SilentlyContinue
    if (-not $GroupDetail) {
        Write-Host "[DEBUG]   ERROR: Get-MgGroup returned null for $($Group.Id)" -ForegroundColor Red
    }
    $DisplayName = if ($GroupDetail) { $GroupDetail.DisplayName } else { $Group.Id }
    $GroupTypes  = if ($GroupDetail) { ($GroupDetail.GroupTypes -join ", ") } else { "unknown" }

    Write-Host "[DEBUG]   Name: $DisplayName | GroupTypes: $GroupTypes | MailEnabled: $($GroupDetail.MailEnabled) | SecurityEnabled: $($GroupDetail.SecurityEnabled)" -ForegroundColor DarkGray

    if ($WhatIf) {
        Write-Host "[WhatIf] Would add as owner: $DisplayName ($($Group.Id))" -ForegroundColor Yellow
        $Results.Add([PSCustomObject]@{
            GroupId   = $Group.Id
            GroupName = $DisplayName
            Type      = "Ownership"
            Status    = "WhatIf"
            Error     = ""
        })
    } else {
        try {
            New-MgGroupOwner -GroupId $Group.Id -DirectoryObjectId $Target.Id -ErrorAction Stop
            Write-Host "[OK] Added as owner: $DisplayName" -ForegroundColor Green
            $Results.Add([PSCustomObject]@{
                GroupId   = $Group.Id
                GroupName = $DisplayName
                Type      = "Ownership"
                Status    = "Copied"
                Error     = ""
            })
        } catch {
            Write-Host "[FAIL] Owner: $DisplayName — $($_.Exception.Message)" -ForegroundColor Red
            $Results.Add([PSCustomObject]@{
                GroupId   = $Group.Id
                GroupName = $DisplayName
                Type      = "Ownership"
                Status    = "Failed"
                Error     = $_.Exception.Message
            })
        }
    }
}

Write-Host "[DEBUG] Ownership pass complete. Groups processed: $(($Results | Where-Object Type -eq 'Ownership').Count)" -ForegroundColor Magenta
#endregion

#region Exchange Mailbox Delegations
Write-Host "`nCopying mailbox permissions..." -ForegroundColor Cyan

$Mailboxes = Get-Mailbox -ResultSize Unlimited |
    Where-Object {
        $_.RecipientTypeDetails -in @(
            "SharedMailbox",
            "RoomMailbox",
            "EquipmentMailbox",
            "UserMailbox"
        )
    }

foreach ($Mailbox in $Mailboxes) {
    # FullAccess
    $FA = Get-MailboxPermission -Identity $Mailbox.Identity |
        Where-Object { $_.User -like $SourceUser -and $_.AccessRights -contains "FullAccess" }

    if ($FA) {
        if ($WhatIf) {
            Write-Host "[WhatIf] Would grant FullAccess on $($Mailbox.DisplayName)" -ForegroundColor Yellow
            $Results.Add([PSCustomObject]@{
                GroupId   = $Mailbox.Identity
                GroupName = $Mailbox.DisplayName
                Type      = "FullAccess"
                Status    = "WhatIf"
                Error     = ""
            })
        } else {
            try {
                Add-MailboxPermission -Identity $Mailbox.Identity -User $TargetUser -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop | Out-Null
                Write-Host "Granted FullAccess: $($Mailbox.DisplayName)" -ForegroundColor Green
                $Results.Add([PSCustomObject]@{
                    GroupId   = $Mailbox.Identity
                    GroupName = $Mailbox.DisplayName
                    Type      = "FullAccess"
                    Status    = "Copied"
                    Error     = ""
                })
            } catch {
                Write-Host "Failed FullAccess: $($Mailbox.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
                $Results.Add([PSCustomObject]@{
                    GroupId   = $Mailbox.Identity
                    GroupName = $Mailbox.DisplayName
                    Type      = "FullAccess"
                    Status    = "Failed"
                    Error     = $_.Exception.Message
                })
            }
        }
    }

    # SendAs
    $SA = Get-RecipientPermission -Identity $Mailbox.Identity |
        Where-Object { $_.Trustee -like $SourceUser -and $_.AccessRights -contains "SendAs" }

    if ($SA) {
        if ($WhatIf) {
            Write-Host "[WhatIf] Would grant SendAs on $($Mailbox.DisplayName)" -ForegroundColor Yellow
            $Results.Add([PSCustomObject]@{
                GroupId   = $Mailbox.Identity
                GroupName = $Mailbox.DisplayName
                Type      = "SendAs"
                Status    = "WhatIf"
                Error     = ""
            })
        } else {
            try {
                Add-RecipientPermission -Identity $Mailbox.Identity -Trustee $TargetUser -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Host "Granted SendAs: $($Mailbox.DisplayName)" -ForegroundColor Green
                $Results.Add([PSCustomObject]@{
                    GroupId   = $Mailbox.Identity
                    GroupName = $Mailbox.DisplayName
                    Type      = "SendAs"
                    Status    = "Copied"
                    Error     = ""
                })
            } catch {
                Write-Host "Failed SendAs: $($Mailbox.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
                $Results.Add([PSCustomObject]@{
                    GroupId   = $Mailbox.Identity
                    GroupName = $Mailbox.DisplayName
                    Type      = "SendAs"
                    Status    = "Failed"
                    Error     = $_.Exception.Message
                })
            }
        }
    }

    # SendOnBehalf
    if ($Mailbox.GrantSendOnBehalfTo -contains $Source.DisplayName) {
        if ($WhatIf) {
            Write-Host "[WhatIf] Would grant SendOnBehalf on $($Mailbox.DisplayName)" -ForegroundColor Yellow
            $Results.Add([PSCustomObject]@{
                GroupId   = $Mailbox.Identity
                GroupName = $Mailbox.DisplayName
                Type      = "SendOnBehalf"
                Status    = "WhatIf"
                Error     = ""
            })
        } else {
            try {
                Set-Mailbox -Identity $Mailbox.Identity -GrantSendOnBehalfTo @{Add = $TargetUser} -ErrorAction Stop
                Write-Host "Granted SendOnBehalf: $($Mailbox.DisplayName)" -ForegroundColor Green
                $Results.Add([PSCustomObject]@{
                    GroupId   = $Mailbox.Identity
                    GroupName = $Mailbox.DisplayName
                    Type      = "SendOnBehalf"
                    Status    = "Copied"
                    Error     = ""
                })
            } catch {
                Write-Host "Failed SendOnBehalf: $($Mailbox.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
                $Results.Add([PSCustomObject]@{
                    GroupId   = $Mailbox.Identity
                    GroupName = $Mailbox.DisplayName
                    Type      = "SendOnBehalf"
                    Status    = "Failed"
                    Error     = $_.Exception.Message
                })
            }
        }
    }
}
#endregion

#region Export CSV
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvPath = ".\GroupClone_Results_${Timestamp}.csv"
$Results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
Write-Host "`nResults exported to: $CsvPath" -ForegroundColor Cyan
Write-Host "Total: $($Results.Count) | Copied: $(($Results | Where-Object Status -eq 'Copied').Count) | Failed: $(($Results | Where-Object Status -eq 'Failed').Count) | WhatIf: $(($Results | Where-Object Status -eq 'WhatIf').Count)" -ForegroundColor Cyan
#endregion

#region Cleanup
Disconnect-MgGraph | Out-Null
Disconnect-ExchangeOnline -Confirm:$false
#endregion

Write-Host "`nCompleted successfully." -ForegroundColor Green