#Requires -Modules Microsoft.Graph.Teams, Microsoft.Graph.Users

<#
.SYNOPSIS
    Comprehensive Microsoft Teams usage analytics and governance reporting
.DESCRIPTION
    Analyzes Teams usage patterns, member activity, guest access, and compliance
    across the organization. Identifies optimization opportunities and governance issues.
.PARAMETER IncludeGuestAnalysis
    Include detailed guest user analysis
.PARAMETER AnalyzePolicies
    Analyze Teams policies and compliance settings
.PARAMETER ExportToExcel
    Export detailed results to Excel format
.PARAMETER DaysToAnalyze
    Number of days of activity to analyze (default: 30)
.EXAMPLE
    .\Get-TeamsAnalytics.ps1 -IncludeGuestAnalysis -AnalyzePolicies -ExportToExcel -DaysToAnalyze 60
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$IncludeGuestAnalysis,
    
    [Parameter(Mandatory = $false)]
    [switch]$AnalyzePolicies,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportToExcel,
    
    [Parameter(Mandatory = $false)]
    [int]$DaysToAnalyze = 30
)

# Connect to Microsoft Graph
try {
    $scopes = @(
        "Team.ReadBasic.All",
        "TeamMember.Read.All",
        "Channel.ReadBasic.All",
        "User.Read.All",
        "Reports.Read.All",
        "Directory.Read.All"
    )
    
    Connect-MgGraph -Scopes $scopes
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

Write-Host "=== MICROSOFT TEAMS ANALYTICS ===" -ForegroundColor Cyan

# Get all teams
Write-Host "Retrieving Teams information..." -ForegroundColor Yellow
$teams = Get-MgTeam -All

if ($teams.Count -eq 0) {
    Write-Host "No Teams found" -ForegroundColor Yellow
    Disconnect-MgGraph
    exit 0
}

Write-Host "Analyzing $($teams.Count) Teams..." -ForegroundColor Yellow

$teamsAnalytics = @()
$teamCounter = 0

foreach ($team in $teams) {
    $teamCounter++
    Write-Progress -Activity "Analyzing Teams" -Status "$($team.DisplayName) ($teamCounter/$($teams.Count))" -PercentComplete (($teamCounter / $teams.Count) * 100)
    
    try {
        # Get team details
        $teamDetails = Get-MgTeam -TeamId $team.Id
        
        # Get team members
        $members = Get-MgTeamMember -TeamId $team.Id -All
        $owners = $members | Where-Object Roles -contains "owner"
        $regularMembers = $members | Where-Object { -not $_.Roles -or $_.Roles.Count -eq 0 }
        
        # Get guest members
        $guestMembers = @()
        if ($IncludeGuestAnalysis) {
            foreach ($member in $members) {
                try {
                    $user = Get-MgUser -UserId $member.UserId -Property "UserPrincipalName,UserType,CreatedDateTime"
                    if ($user.UserType -eq "Guest") {
                        $guestMembers += $user
                    }
                }
                catch {
                    # User details not accessible
                }
            }
        }
        
        # Get channels
        $channels = Get-MgTeamChannel -TeamId $team.Id -All
        $standardChannels = $channels | Where-Object ChannelType -eq "standard"
        $privateChannels = $channels | Where-Object ChannelType -eq "private"
        $sharedChannels = $channels | Where-Object ChannelType -eq "shared"
        
        # Calculate team age
        $teamAge = if ($teamDetails.CreatedDateTime) {
            (Get-Date) - (Get-Date $teamDetails.CreatedDateTime) | Select-Object -ExpandProperty Days
        } else {
            $null
        }
        
        # Analyze team structure
        $teamIssues = @()
        
        if ($owners.Count -eq 0) {
            $teamIssues += "No owners assigned"
        } elseif ($owners.Count -eq 1) {
            $teamIssues += "Single owner risk"
        }
        
        if ($members.Count -eq 0) {
            $teamIssues += "No members"
        } elseif ($members.Count -gt 5000) {
            $teamIssues += "Large team (>5000 members)"
        }
        
        if ($guestMembers.Count -gt 0 -and $guestMembers.Count -gt ($members.Count * 0.3)) {
            $teamIssues += "High guest ratio (>30%)"
        }
        
        if ($channels.Count -gt 50) {
            $teamIssues += "Many channels (>50)"
        }
        
        # Calculate governance score
        $governanceScore = 100
        $governanceScore -= ($teamIssues.Count * 15)
        $governanceScore = [math]::Max($governanceScore, 0)
        
        $teamAnalysis = [PSCustomObject]@{
            TeamName = $teamDetails.DisplayName
            TeamId = $team.Id
            Description = $teamDetails.Description
            CreatedDateTime = $teamDetails.CreatedDateTime
            TeamAgeInDays = $teamAge
            Visibility = $teamDetails.Visibility
            
            # Membership
            TotalMembers = $members.Count
            OwnerCount = $owners.Count
            MemberCount = $regularMembers.Count
            GuestCount = $guestMembers.Count
            GuestPercentage = if ($members.Count -gt 0) { 
                [math]::Round(($guestMembers.Count / $members.Count) * 100, 1) 
            } else { 0 }
            
            # Channels
            TotalChannels = $channels.Count
            StandardChannels = $standardChannels.Count
            PrivateChannels = $privateChannels.Count
            SharedChannels = $sharedChannels.Count
            
            # Settings
            IsMembershipLimitedToOwners = $teamDetails.IsMembershipLimitedToOwners
            MemberSettings_AllowCreateUpdateChannels = $teamDetails.MemberSettings.AllowCreateUpdateChannels
            MemberSettings_AllowDeleteChannels = $teamDetails.MemberSettings.AllowDeleteChannels
            MemberSettings_AllowAddRemoveApps = $teamDetails.MemberSettings.AllowAddRemoveApps
            GuestSettings_AllowCreateUpdateChannels = $teamDetails.GuestSettings.AllowCreateUpdateChannels
            GuestSettings_AllowDeleteChannels = $teamDetails.GuestSettings.AllowDeleteChannels
            
            # Fun settings
            FunSettings_AllowGiphy = $teamDetails.FunSettings.AllowGiphy
            FunSettings_GiphyContentRating = $teamDetails.FunSettings.GiphyContentRating
            FunSettings_AllowStickersAndMemes = $teamDetails.FunSettings.AllowStickersAndMemes
            FunSettings_AllowCustomMemes = $teamDetails.FunSettings.AllowCustomMemes
            
            # Messaging settings
            MessagingSettings_AllowUserEditMessages = $teamDetails.MessagingSettings.AllowUserEditMessages
            MessagingSettings_AllowUserDeleteMessages = $teamDetails.MessagingSettings.AllowUserDeleteMessages
            MessagingSettings_AllowOwnerDeleteMessages = $teamDetails.MessagingSettings.AllowOwnerDeleteMessages
            MessagingSettings_AllowTeamMentions = $teamDetails.MessagingSettings.AllowTeamMentions
            MessagingSettings_AllowChannelMentions = $teamDetails.MessagingSettings.AllowChannelMentions
            
            # Governance
            GovernanceScore = $governanceScore
            GovernanceIssues = ($teamIssues -join "; ")
            NeedsAttention = ($teamIssues.Count -gt 0)
            
            # Classification
            IsSmallTeam = ($members.Count -le 10)
            IsMediumTeam = ($members.Count -gt 10 -and $members.Count -le 100)
            IsLargeTeam = ($members.Count -gt 100)
            IsExternalCollaboration = ($guestMembers.Count -gt 0)
        }
        
        $teamsAnalytics += $teamAnalysis
    }
    catch {
        Write-Host "Error analyzing team $($team.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Add basic info even if detailed analysis fails
        $basicAnalysis = [PSCustomObject]@{
            TeamName = $team.DisplayName
            TeamId = $team.Id
            GovernanceScore = 0
            GovernanceIssues = "Analysis failed: $($_.Exception.Message)"
            NeedsAttention = $true
        }
        
        $teamsAnalytics += $basicAnalysis
    }
}

Write-Progress -Activity "Analyzing Teams" -Completed

# Generate summary analytics
Write-Host "`n--- TEAMS ANALYTICS SUMMARY ---" -ForegroundColor Cyan

$totalTeams = $teamsAnalytics.Count
$teamsNeedingAttention = ($teamsAnalytics | Where-Object NeedsAttention -eq $true).Count
$totalMembers = ($teamsAnalytics | Measure-Object TotalMembers -Sum).Sum
$totalGuests = ($teamsAnalytics | Measure-Object GuestCount -Sum).Sum
$totalChannels = ($teamsAnalytics | Measure-Object TotalChannels -Sum).Sum

Write-Host "Total Teams: $totalTeams" -ForegroundColor White
Write-Host "Teams Needing Attention: $teamsNeedingAttention" -ForegroundColor $(if ($teamsNeedingAttention -gt 0) { "Yellow" } else { "Green" })
Write-Host "Total Members: $totalMembers" -ForegroundColor White
Write-Host "Total Guest Users: $totalGuests" -ForegroundColor White
Write-Host "Total Channels: $totalChannels" -ForegroundColor White

# Team size distribution
Write-Host "`nTeam Size Distribution:" -ForegroundColor Yellow
$smallTeams = ($teamsAnalytics | Where-Object IsSmallTeam -eq $true).Count
$mediumTeams = ($teamsAnalytics | Where-Object IsMediumTeam -eq $true).Count
$largeTeams = ($teamsAnalytics | Where-Object IsLargeTeam -eq $true).Count

Write-Host "  Small Teams (≤10 members): $smallTeams" -ForegroundColor White
Write-Host "  Medium Teams (11-100 members): $mediumTeams" -ForegroundColor White
Write-Host "  Large Teams (>100 members): $largeTeams" -ForegroundColor White

# External collaboration analysis
$externalCollabTeams = ($teamsAnalytics | Where-Object IsExternalCollaboration -eq $true).Count
Write-Host "`nExternal Collaboration:" -ForegroundColor Yellow
Write-Host "  Teams with guests: $externalCollabTeams" -ForegroundColor White
Write-Host "  Average guest percentage: $([math]::Round(($teamsAnalytics | Measure-Object GuestPercentage -Average).Average, 1))%" -ForegroundColor White

# Teams needing attention
if ($teamsNeedingAttention -gt 0) {
    Write-Host "`n--- TEAMS NEEDING ATTENTION ---" -ForegroundColor Red
    $teamsAnalytics | Where-Object NeedsAttention -eq $true | 
        Sort-Object GovernanceScore | 
        Select-Object -First 10 | 
        Select-Object TeamName, TotalMembers, OwnerCount, GovernanceScore, GovernanceIssues | 
        Format-Table -Wrap
}

# Guest analysis (if enabled)
if ($IncludeGuestAnalysis) {
    Write-Host "`n--- GUEST USER ANALYSIS ---" -ForegroundColor Yellow
    
    $highGuestTeams = $teamsAnalytics | Where-Object GuestPercentage -gt 30 | Sort-Object GuestPercentage -Descending
    if ($highGuestTeams) {
        Write-Host "Teams with high guest ratios (>30%):" -ForegroundColor Yellow
        $highGuestTeams | Select-Object TeamName, TotalMembers, GuestCount, GuestPercentage | Format-Table
    }
}

# Channel analysis
Write-Host "`n--- CHANNEL ANALYSIS ---" -ForegroundColor Yellow
$avgChannelsPerTeam = [math]::Round(($teamsAnalytics | Measure-Object TotalChannels -Average).Average, 1)
$totalPrivateChannels = ($teamsAnalytics | Measure-Object PrivateChannels -Sum).Sum
$totalSharedChannels = ($teamsAnalytics | Measure-Object SharedChannels -Sum).Sum

Write-Host "Average channels per team: $avgChannelsPerTeam" -ForegroundColor White
Write-Host "Total private channels: $totalPrivateChannels" -ForegroundColor White
Write-Host "Total shared channels: $totalSharedChannels" -ForegroundColor White

$teamsWithManyChannels = ($teamsAnalytics | Where-Object TotalChannels -gt 20).Count
if ($teamsWithManyChannels -gt 0) {
    Write-Host "Teams with many channels (>20): $teamsWithManyChannels" -ForegroundColor Yellow
}

# Policy analysis (if enabled)
if ($AnalyzePolicies) {
    Write-Host "`n--- POLICY COMPLIANCE ANALYSIS ---" -ForegroundColor Yellow
    
    # Analyze common security settings
    $allowGiphyTeams = ($teamsAnalytics | Where-Object FunSettings_AllowGiphy -eq $true).Count
    $allowCustomMemesTeams = ($teamsAnalytics | Where-Object FunSettings_AllowCustomMemes -eq $true).Count
    $guestChannelCreationTeams = ($teamsAnalytics | Where-Object GuestSettings_AllowCreateUpdateChannels -eq $true).Count
    
    Write-Host "Teams allowing Giphy: $allowGiphyTeams" -ForegroundColor White
    Write-Host "Teams allowing custom memes: $allowCustomMemesTeams" -ForegroundColor White
    Write-Host "Teams allowing guest channel creation: $guestChannelCreationTeams" -ForegroundColor White
}

# Recommendations
Write-Host "`n--- RECOMMENDATIONS ---" -ForegroundColor Cyan

$recommendations = @()

if ($teamsNeedingAttention -gt 0) {
    $recommendations += "Review $teamsNeedingAttention teams with governance issues"
}

$orphanedTeams = ($teamsAnalytics | Where-Object OwnerCount -eq 0).Count
if ($orphanedTeams -gt 0) {
    $recommendations += "Assign owners to $orphanedTeams orphaned teams"
}

$singleOwnerTeams = ($teamsAnalytics | Where-Object OwnerCount -eq 1).Count
if ($singleOwnerTeams -gt 0) {
    $recommendations += "Add additional owners to $singleOwnerTeams teams with single owners"
}

if ($totalGuests -gt 0) {
    $recommendations += "Review guest access policies and periodic access reviews for external users"
}

$oldTeams = ($teamsAnalytics | Where-Object { $_.TeamAgeInDays -gt 365 -and $_.TotalMembers -eq 0 }).Count
if ($oldTeams -gt 0) {
    $recommendations += "Consider archiving $oldTeams old teams with no members"
}

if ($recommendations.Count -gt 0) {
    $recommendations | ForEach-Object { Write-Host "• $_" -ForegroundColor Cyan }
} else {
    Write-Host "• Teams governance appears to be in good shape!" -ForegroundColor Green
}

# Export results
if ($ExportToExcel) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $outputPath = ".\Teams-Analytics-$timestamp.xlsx"
    
    if (Get-Module -ListAvailable -Name ImportExcel) {
        # Main analytics data
        $teamsAnalytics | Export-Excel -Path $outputPath -WorksheetName "Teams Analytics" -AutoSize -FreezeTopRow
        
        # Summary statistics
        $summaryData = @(
            [PSCustomObject]@{ Metric = "Total Teams"; Value = $totalTeams }
            [PSCustomObject]@{ Metric = "Teams Needing Attention"; Value = $teamsNeedingAttention }
            [PSCustomObject]@{ Metric = "Total Members"; Value = $totalMembers }
            [PSCustomObject]@{ Metric = "Total Guests"; Value = $totalGuests }
            [PSCustomObject]@{ Metric = "Total Channels"; Value = $totalChannels }
            [PSCustomObject]@{ Metric = "Small Teams"; Value = $smallTeams }
            [PSCustomObject]@{ Metric = "Medium Teams"; Value = $mediumTeams }
            [PSCustomObject]@{ Metric = "Large Teams"; Value = $largeTeams }
            [PSCustomObject]@{ Metric = "Teams with Guests"; Value = $externalCollabTeams }
        )
        
        $summaryData | Export-Excel -Path $outputPath -WorksheetName "Summary" -AutoSize
        
        # Teams needing attention
        if ($teamsNeedingAttention -gt 0) {
            $teamsAnalytics | Where-Object NeedsAttention -eq $true | 
                Export-Excel -Path $outputPath -WorksheetName "Needs Attention" -AutoSize
        }
        
        Write-Host "`nDetailed report exported to: $outputPath" -ForegroundColor Green
    } else {
        Write-Host "ImportExcel module not available. Exporting to CSV..." -ForegroundColor Yellow
        $teamsAnalytics | Export-Csv -Path ".\Teams-Analytics-$timestamp.csv" -NoTypeInformation
        Write-Host "Report exported to: .\Teams-Analytics-$timestamp.csv" -ForegroundColor Green
    }
}

Write-Host "`n--- NEXT STEPS ---" -ForegroundColor Cyan
Write-Host "• Implement Teams governance policies and lifecycle management" -ForegroundColor White
Write-Host "• Set up regular access reviews for guest users" -ForegroundColor White
Write-Host "• Establish naming conventions and templates for new teams" -ForegroundColor White
Write-Host "• Monitor Teams usage with Microsoft 365 usage reports" -ForegroundColor White

Disconnect-MgGraph
Write-Host "Teams analytics complete" -ForegroundColor Green