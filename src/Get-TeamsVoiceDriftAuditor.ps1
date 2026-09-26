<#
================================================================================
SCRIPT NAME : Get-TeamsVoiceDriftAuditor.ps1
DESCRIPTION : Standalone diagnostic auditor for Microsoft Teams Auto Attendants 
              and Call Queues. Detects orphaned agents, empty target groups, 
              missing nested schedules, and broken fallback routes.
AUTHOR      : Soutik Bera
REPOSITORY  : TeamsVoice-DriftAuditor
MODULES REQ : MicrosoftTeams (v4.0.0+), Microsoft.Graph.Authentication, 
              Microsoft.Graph.Users, Microsoft.Graph.Groups
================================================================================
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$ReportOutputDir = ".\reports",

    [Parameter(Mandatory = $false)]
    [switch]$SkipAutoAttendants,

    [Parameter(Mandatory = $false)]
    [switch]$SkipCallQueues
)

$ErrorActionPreference = 'Stop'

# Ensure report output folder exists
if (-not (Test-Path -Path $ReportOutputDir)) {
    New-Item -ItemType Directory -Path $ReportOutputDir -Force | Out-Null
}

$Timestamp = (Get-Date -Format "yyyyMMdd-HHmmss")

# ==============================================================================
# SECTION 0: SESSION CONNECTION & PREREQUISITES
# ==============================================================================
Write-Host "[*] Checking session connections..." -ForegroundColor Cyan

# Teams Connection Check
try {
    $teamsSession = Get-CsTenant -ErrorAction SilentlyContinue
    if (-not $teamsSession) {
        Write-Host "[+] Connecting to Microsoft Teams..." -ForegroundColor Yellow
        Connect-MicrosoftTeams
    }
}
catch {
    Connect-MicrosoftTeams
}

# Graph Connection Check
$graphContext = Get-MgContext -ErrorAction SilentlyContinue
if (-not $graphContext) {
    Write-Host "[+] Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -NoWelcome
}

# Cache lookups to prevent Graph API rate limits (HTTP 429)
$UserStatusCache  = @{}
$GroupMemberCache = @{}

function Get-CachedUserStatus {
    param ([string]$UserId)
    if (-not $UserStatusCache.ContainsKey($UserId)) {
        $u = Get-MgUser -UserId $UserId -Property AccountEnabled, UserPrincipalName -ErrorAction SilentlyContinue
        $UserStatusCache[$UserId] = if ($u) { [bool]$u.AccountEnabled } else { $false }
    }
    return $UserStatusCache[$UserId]
}

function Get-CachedGroupMembers {
    param ([string]$GroupId)
    if (-not $GroupMemberCache.ContainsKey($GroupId)) {
        $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue
        $GroupMemberCache[$GroupId] = if ($members) { @($members.Id) } else { @() }
    }
    return $GroupMemberCache[$GroupId]
}

# ==============================================================================
# SECTION 1: CALL QUEUE AUDIT
# ==============================================================================
if (-not $SkipCallQueues) {
    Write-Host "`n[*] Auditing Call Queues for roster drift and routing orphans..." -ForegroundColor Cyan
    $CallQueues = Get-CsCallQueue;
    $CQAuditReport = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($cq in $CallQueues) {
        $totalAgents = 0
        $disabledAgents = 0
        $hasEmptyGroup =$false

        # Direct Users
        if ($cq.Users -and $cq.Users.Count -gt 0) {
            $totalAgents += $cq.Users.Count
            foreach ($uid in $cq.Users) {
                if (-not (Get-CachedUserStatus -UserId $uid)) {$disabledAgents++ 
                }
            }
        }

        # Groups / Teams
        if ($cq.DistributionLists -and $cq.DistributionLists.Count -gt 0) {
            foreach ($gid in $cq.DistributionLists) {
                $groupMembers = Get-CachedGroupMembers -GroupId $gid
                if ($groupMembers.Count -eq 0) {
                    $hasEmptyGroup =$true
                }
                else {
                    $totalAgents += $groupMembers.Count
                    foreach ($mid in $groupMembers) {
                        if (-not (Get-CachedUserStatus -UserId $mid)) {$disabledAgents++ 
                        }
                    }
                }
            }
        }

        $isOrphaned = ($totalAgents -eq 0) -or ($totalAgents -eq $disabledAgents)

        $CQAuditReport.Add([PSCustomObject]@{
            QueueName           = $cq.Name
            QueueId             = $cq.Identity
            DistributionMethod  = $cq.DistributionMethod
            TotalAgents         = $totalAgents
            ActiveAgents        = ($totalAgents -$disabledAgents)
            DisabledAgents      = $disabledAgents
            EmptyGroupAttached  = $hasEmptyGroup
            IsOrphaned          = $isOrphaned
            RoutingAction       = $cq.OverflowAction
            RoutingTarget       = $cq.OverflowActionTarget
        })
    }

    $cqExportPath = Join-Path $ReportOutputDir "CallQueue_DriftReport_$Timestamp.csv"
    $CQAuditReport | Export-Csv -Path $cqExportPath -NoTypeInformation 
    $CQAuditReport | Format-Table QueueName, ActiveAgents, DisabledAgents, EmptyGroupAttached, IsOrphaned -AutoSize
    Write-Host "[+] Call Queue report generated: $cqExportPath" -ForegroundColor Green
}

# ==============================================================================
# SECTION 2: AUTO ATTENDANT AUDIT
# ==============================================================================
if (-not $SkipAutoAttendants) {
    Write-Host "`n[*] Auditing Auto Attendants for broken menus and missing schedules..." -ForegroundColor Cyan
    $AutoAttendants = Get-CsAutoAttendant
    $AAAuditReport = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($aa in $AutoAttendants) {
        $warnings = [System.Collections.Generic.List[string]]::new()

        # Menu options check
        if ($aa.DefaultCallFlow -and $aa.DefaultCallFlow.Menu) {
            $menu = $aa.DefaultCallFlow.Menu
            if (-not $menu.MenuOptions -or $menu.MenuOptions.Count -eq 0) {
                $warnings.Add("Default call flow has no active menu options.")
            }
        }

        # Holiday & Business Hours schedules
        if ($aa.CallHandlingAssociations -and $aa.CallHandlingAssociations.Count -gt 0) {
            foreach ($assoc in $aa.CallHandlingAssociations) {
                if ($assoc.Type -eq "HolidaySchedule") {
                    $sched = Get-CsOnlineSchedule -Identity $assoc.ScheduleId -ErrorAction SilentlyContinue
                    if (-not $sched) {
                        $warnings.Add("Missing or deleted Holiday Schedule GUID: $($assoc.ScheduleId)")
                    }
                }
            }
        }

        # Operator target check
        # Operator target check (Direct Graph + Teams Verification)
    if ($aa.Operator) {
        $target = $aa.Operator.Target
        if ($aa.Operator.Type -eq "User") {
            # Check active Entra ID status via Microsoft Graph
            $graphUser = Get-MgUser -UserId $target -Property AccountEnabled, Id -ErrorAction SilentlyContinue

            if (-not $graphUser) {
            $warnings.Add("Assigned operator identity '$target' does not exist or has been soft-deleted.")
            }
            elseif (-not $graphUser.AccountEnabled) {
            $warnings.Add("Assigned operator identity '$target' is disabled in Entra ID.")
            }
        }
    }

        $AAAuditReport.Add([PSCustomObject]@{
            AutoAttendantName = $aa.Name
            Identity          = $aa.Identity
            Status            = if ($warnings.Count -gt 0) { "Attention Required" } else { "Clean" }
            WarningCount      = $warnings.Count
            DriftWarnings     = ($warnings -join " | ")
        })
    }

    $aaExportPath = Join-Path $ReportOutputDir "AutoAttendant_DriftReport_$Timestamp.csv"
    $AAAuditReport | Export-Csv -Path $aaExportPath -NoTypeInformation
    $AAAuditReport | Format-Table AutoAttendantName, Status, WarningCount, DriftWarnings -AutoSize
    Write-Host "[+] Auto Attendant report generated: $aaExportPath" -ForegroundColor Green
}

Write-Host "`n[+] Audit run finished." -ForegroundColor Green