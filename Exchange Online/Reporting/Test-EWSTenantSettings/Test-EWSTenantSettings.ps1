#Requires -Version 5.1
<#
.SYNOPSIS
    Reports the current Exchange Web Services (EWS) configuration of an Exchange Online tenant.

.DESCRIPTION
    Collects and displays the tenant-level EWS settings that are relevant for the EWS retirement
    transition in Exchange Online:

    - EWSEnabled (the organization-wide EWS switch)
    - EwsAllowedAppIDs (the application-ID based allow list, retrieved with
      Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy)
    - EwsApplicationAccessPolicy with EwsAllowList and EwsBlockList (the older user-agent
      based controls, which can also affect REST traffic)
    - The EwsAllowOutlook / EwsAllowMacOutlook / EwsAllowEntourage / EwsAllowMowa client switches

    The script writes a formatted summary to the console, including the tenant display name and
    the MOERA domain (the initial <tenant>.onmicrosoft.com domain), an interpretation of the
    effective EWS state, and stores a snapshot of the configuration in a reports folder. The
    snapshot is intended as the documented baseline before EwsAllowedAppIDs changes are made.

    The script is read-only. It never modifies the organization configuration.

    Based on the guidance in "Notes from the field: testing EWSAllowedAppIDs safely":
    https://techcommunity.microsoft.com/blog/exchange/notes-from-the-field-testing-ewsallowedappids-safely/4548568

.PARAMETER OutputPath
    Directory the configuration snapshot is written to. The directory is created if it does not
    exist. Defaults to '.\Reports'.

.PARAMETER ReportFormat
    Format of the configuration snapshot. Valid values are 'Json', 'Csv', and 'Both'.
    Defaults to 'Json'.

.PARAMETER SkipReport
    Displays the configuration on screen only and does not write a snapshot file.

.EXAMPLE
    .\Test-EWSTenantSettings.ps1

    Verifies the Exchange Online connection, displays the current EWS configuration and writes a
    JSON snapshot to .\Reports.

.EXAMPLE
    .\Test-EWSTenantSettings.ps1 -OutputPath C:\Reports\EWS -ReportFormat Both

    Writes both a JSON and a CSV snapshot to C:\Reports\EWS.

.EXAMPLE
    .\Test-EWSTenantSettings.ps1 -SkipReport

    Displays the current EWS configuration without writing a snapshot file.

.NOTES
    Author: Thomas Stensitzki
    Target platform: Exchange Online
    Required modules: ExchangeOnlineManagement
    Required permissions: View-Only Organization Configuration (Get-OrganizationConfig,
                          Get-AcceptedDomain)
    Change log:
        1.0.0 - Initial version
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = '.\Reports',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Json', 'Csv', 'Both')]
    [string]$ReportFormat = 'Json',

    [Parameter(Mandatory = $false)]
    [switch]$SkipReport
)

begin {
    $ErrorActionPreference = 'Stop'

    $script:LabelWidth = 26

    function Write-Line {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Label,
            [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Value = '',
            [Parameter(Mandatory = $false)][string]$ValueColor = 'Gray'
        )

        Write-Host ('  {0} : ' -f $Label.PadRight($script:LabelWidth)) -NoNewline -ForegroundColor DarkGray
        Write-Host $Value -ForegroundColor $ValueColor
    }

    function Write-Section {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$Title
        )

        Write-Host ''
        Write-Host ('  {0}' -f $Title) -ForegroundColor White
        Write-Host ('  {0}' -f ('-' * 72)) -ForegroundColor DarkGray
    }

    function Format-ListValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)][AllowNull()]$Value
        )

        # EwsAllowedAppIDs is returned as a comma-separated string, the user-agent lists as collections
        $items = @(
            if ($null -eq $Value) { }
            elseif ($Value -is [string]) { $Value -split ',' }
            else { $Value }
        ) | ForEach-Object { "$_".Trim() } | Where-Object { $_ }

        , [string[]]$items
    }

    function Confirm-ExchangeOnlineConnection {
        [CmdletBinding()]
        param()

        $connection = $null

        if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -ne 'Expired' } |
                Select-Object -First 1
        }

        if ($connection) {
            return $connection
        }

        Write-Host ''
        Write-Host '  No active Exchange Online PowerShell connection was detected.' -ForegroundColor Yellow

        if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
            throw 'The ExchangeOnlineManagement module is not installed. Install it using "Install-Module ExchangeOnlineManagement -Scope CurrentUser", connect to Exchange Online manually, and run this script again.'
        }

        Import-Module -Name ExchangeOnlineManagement -ErrorAction Stop

        $answer = Read-Host -Prompt '  Run Connect-ExchangeOnline now? [Y] Yes  [N] No, I will connect manually'

        if ($answer -notmatch '^\s*(y|yes|j|ja)\s*$') {
            throw 'Connect to Exchange Online using Connect-ExchangeOnline and run this script again.'
        }

        $userPrincipalName = Read-Host -Prompt '  User principal name (leave empty for interactive sign-in)'

        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        else {
            Connect-ExchangeOnline -UserPrincipalName $userPrincipalName.Trim() -ShowBanner:$false -ErrorAction Stop
        }

        $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' } |
            Select-Object -First 1

        if (-not $connection) {
            throw 'Connecting to Exchange Online failed.'
        }

        return $connection
    }
}

process {
    Write-Host ''
    Write-Host '  Exchange Online EWS configuration report' -ForegroundColor Cyan
    Write-Host ('  {0}' -f ('=' * 72)) -ForegroundColor DarkGray

    $connection = Confirm-ExchangeOnlineConnection

    Write-Verbose -Message 'Fetching organization configuration'
    $organizationConfig = Get-OrganizationConfig
    $ewsPolicyConfig = Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy

    $moeraDomain = $null

    try {
        $moeraDomain = (Get-AcceptedDomain -ErrorAction Stop |
                Where-Object { $_.InitialDomain } |
                Select-Object -First 1).DomainName
    }
    catch {
        Write-Verbose -Message ('Get-AcceptedDomain failed, falling back to the organization name: {0}' -f $_.Exception.Message)
    }

    if (-not $moeraDomain) {
        # The organization name of an Exchange Online tenant is the MOERA domain
        $moeraDomain = $organizationConfig.Name
    }

    $moeraDomain = "$moeraDomain"

    $allowedAppIds = Format-ListValue -Value $ewsPolicyConfig.EwsAllowedAppIDs
    $ewsAllowList = Format-ListValue -Value $organizationConfig.EwsAllowList
    $ewsBlockList = Format-ListValue -Value $organizationConfig.EwsBlockList

    $ewsEnabled = $organizationConfig.EWSEnabled

    if ($null -eq $ewsEnabled) {
        $ewsEnabledText = '$null (not configured)'
        $ewsEnabledColor = 'Yellow'
        $effectiveState = 'EWS is unrestricted by the application-ID gate. The EwsAllowedAppIDs allow list is ignored while EWSEnabled is $null.'
        $effectiveColor = 'Yellow'
    }
    elseif ($ewsEnabled -eq $false) {
        $ewsEnabledText = 'False'
        $ewsEnabledColor = 'Red'
        $effectiveState = 'EWS is disabled for the entire organization. No application can access EWS, regardless of EwsAllowedAppIDs.'
        $effectiveColor = 'Red'
    }
    elseif ($allowedAppIds.Count -gt 0) {
        $ewsEnabledText = 'True'
        $ewsEnabledColor = 'Green'
        $effectiveState = ('EWS is enabled and restricted to {0} allowed application ID(s). All other applications are blocked.' -f $allowedAppIds.Count)
        $effectiveColor = 'Green'
    }
    else {
        $ewsEnabledText = 'True'
        $ewsEnabledColor = 'Green'
        $effectiveState = 'EWS is enabled and the application-ID allow list is empty, therefore all applications are permitted.'
        $effectiveColor = 'Yellow'
    }

    Write-Section -Title 'Tenant'
    Write-Line -Label 'Tenant display name' -Value "$($organizationConfig.DisplayName)" -ValueColor Cyan
    Write-Line -Label 'MOERA domain' -Value $moeraDomain -ValueColor Cyan
    Write-Line -Label 'Tenant ID' -Value "$($connection.TenantID)"
    Write-Line -Label 'Connected as' -Value "$($connection.UserPrincipalName)"
    Write-Line -Label 'Report time (UTC)' -Value ([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss'))

    Write-Section -Title 'EWS application-ID control (EWS retirement gate)'
    Write-Line -Label 'EWSEnabled' -Value $ewsEnabledText -ValueColor $ewsEnabledColor
    Write-Line -Label 'EwsAllowedAppIDs count' -Value $allowedAppIds.Count

    if ($allowedAppIds.Count -eq 0) {
        Write-Line -Label 'EwsAllowedAppIDs' -Value '<empty>' -ValueColor DarkGray
    }
    else {
        $index = 1

        foreach ($appId in $allowedAppIds) {
            Write-Line -Label ('EwsAllowedAppIDs [{0}]' -f $index) -Value $appId
            $index++
        }
    }

    Write-Section -Title 'User-agent controls (evaluated independently, can also affect REST)'
    Write-Line -Label 'EwsApplicationAccessPolicy' -Value $(if ($null -eq $organizationConfig.EwsApplicationAccessPolicy) { '<not set>' } else { "$($organizationConfig.EwsApplicationAccessPolicy)" })
    Write-Line -Label 'EwsAllowList count' -Value $ewsAllowList.Count
    Write-Line -Label 'EwsBlockList count' -Value $ewsBlockList.Count

    foreach ($entry in $ewsAllowList) { Write-Line -Label 'EwsAllowList entry' -Value $entry }
    foreach ($entry in $ewsBlockList) { Write-Line -Label 'EwsBlockList entry' -Value $entry }

    Write-Section -Title 'Client switches'
    Write-Line -Label 'EwsAllowOutlook' -Value "$($organizationConfig.EwsAllowOutlook)"
    Write-Line -Label 'EwsAllowMacOutlook' -Value "$($organizationConfig.EwsAllowMacOutlook)"
    Write-Line -Label 'EwsAllowEntourage' -Value "$($organizationConfig.EwsAllowEntourage)"
    Write-Line -Label 'EwsAllowMowa' -Value "$($organizationConfig.EwsAllowMowa)"

    Write-Section -Title 'Effective state'
    Write-Host ('  {0}' -f $effectiveState) -ForegroundColor $effectiveColor
    Write-Host '  Note: EwsAllowedAppIDs changes need up to 24 hours, EWSEnabled changes about 1 hour to take effect.' -ForegroundColor DarkGray

    $snapshot = [PSCustomObject]@{
        ReportTimeUtc              = [datetime]::UtcNow.ToString('o')
        TenantDisplayName          = "$($organizationConfig.DisplayName)"
        MoeraDomain                = $moeraDomain
        TenantId                   = "$($connection.TenantID)"
        ConnectedAs                = "$($connection.UserPrincipalName)"
        EWSEnabled                 = $ewsEnabled
        EffectiveState             = $effectiveState
        EwsAllowedAppIDsCount      = $allowedAppIds.Count
        EwsAllowedAppIDs           = $allowedAppIds
        EwsApplicationAccessPolicy = "$($organizationConfig.EwsApplicationAccessPolicy)"
        EwsAllowList               = $ewsAllowList
        EwsBlockList               = $ewsBlockList
        EwsAllowOutlook            = $organizationConfig.EwsAllowOutlook
        EwsAllowMacOutlook         = $organizationConfig.EwsAllowMacOutlook
        EwsAllowEntourage          = $organizationConfig.EwsAllowEntourage
        EwsAllowMowa               = $organizationConfig.EwsAllowMowa
    }

    if ($SkipReport) {
        Write-Host ''
        Write-Host '  SkipReport was specified, no snapshot file has been written.' -ForegroundColor DarkGray
        Write-Host ''
        return
    }

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        $null = New-Item -Path $OutputPath -ItemType Directory -Force
    }

    $resolvedOutputPath = (Resolve-Path -LiteralPath $OutputPath).ProviderPath
    $safeDomain = ($moeraDomain -replace '[^A-Za-z0-9\.\-]', '_')
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $baseFileName = 'EWSConfig_{0}_{1}' -f $safeDomain, $timestamp

    Write-Section -Title 'Snapshot'

    if ($ReportFormat -eq 'Json' -or $ReportFormat -eq 'Both') {
        $jsonFile = Join-Path -Path $resolvedOutputPath -ChildPath ('{0}.json' -f $baseFileName)
        $snapshot | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonFile -Encoding UTF8
        Write-Line -Label 'JSON snapshot' -Value $jsonFile -ValueColor Green
    }

    if ($ReportFormat -eq 'Csv' -or $ReportFormat -eq 'Both') {
        $csvFile = Join-Path -Path $resolvedOutputPath -ChildPath ('{0}.csv' -f $baseFileName)
        $snapshot |
            Select-Object -Property ReportTimeUtc, TenantDisplayName, MoeraDomain, TenantId, ConnectedAs,
                EWSEnabled, EffectiveState, EwsAllowedAppIDsCount,
                @{Name = 'EwsAllowedAppIDs'; Expression = { $_.EwsAllowedAppIDs -join ';' } },
                EwsApplicationAccessPolicy,
                @{Name = 'EwsAllowList'; Expression = { $_.EwsAllowList -join ';' } },
                @{Name = 'EwsBlockList'; Expression = { $_.EwsBlockList -join ';' } },
                EwsAllowOutlook, EwsAllowMacOutlook, EwsAllowEntourage, EwsAllowMowa |
            Export-Csv -LiteralPath $csvFile -NoTypeInformation -Encoding UTF8
        Write-Line -Label 'CSV snapshot' -Value $csvFile -ValueColor Green
    }

    Write-Host ''
}