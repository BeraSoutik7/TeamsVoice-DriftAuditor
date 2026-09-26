# TeamsVoice-DriftAuditor

Enterprise diagnostic engine to detect configuration drift, orphaned call routes, stale agent rosters, and broken fallback destinations across Microsoft Teams Auto Attendants (AA) and Call Queues (CQ).

## Key Audit Capabilities

* **Call Queue Agent Health:** Flags empty direct rosters, empty M365/Security distribution groups, and disabled Entra ID user accounts.
* **Orphan Detection:** Flags active queues where 100% of attached agents are disabled.
* **Auto Attendant Fallback Verification:** Validates that nested Holiday Schedules and Business Hours GUIDs still exist.
* **Operator Target Validation:** Flags operator routing configurations pointing to non-existent or deprovisioned user accounts.
* **Batch Performance Optimized:** Uses in-memory caching to eliminate repetitive Microsoft Graph calls and avoid HTTP 429 throttling.

## USAGE
# Run full audit (Exports CSVs to .\reports\)
.\src\Get-TeamsVoiceDriftAuditor.ps1

# Audit Call Queues only
.\src\Get-TeamsVoiceDriftAuditor.ps1 -SkipAutoAttendants

# Custom output directory
.\src\Get-TeamsVoiceDriftAuditor.ps1 -ReportOutputDir "C:\VoiceAudits"

## Prerequisites

PowerShell 7+ is recommended. Install required modules:

```powershell
Install-Module -Name MicrosoftTeams -MinimumVersion 4.0.0 -Scope CurrentUser
Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
Install-Module -Name Microsoft.Graph.Groups -Scope CurrentUser

