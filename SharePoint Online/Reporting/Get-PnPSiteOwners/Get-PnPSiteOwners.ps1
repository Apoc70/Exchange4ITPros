#Requires -Modules PnP.PowerShell

<#

    .SYNOPSIS
    This script retrieves all sites in a SharePoint Online tenant and their owners.
    It generates three reports: Sites Report, Owners Report, and Sites Error Report.

    .DESCRIPTION
    The script connects to the SharePoint Online tenant using PnP PowerShell and retrieves all sites, including OneDrive sites if specified.
    It then fetches the site owners and generates reports in CSV and HTML formats.
    The reports include site URLs, titles, owner emails, and whether the site is a group site or not.
    The script also handles errors and generates a Sites Error Report for sites that could not be processed.
    The reports are saved in the same directory as the script.
    The script requires the PnP PowerShell module to be installed and configured with the necessary permissions.

    .PARAMETER IncludeOneDriveSites
    If specified, the script will include OneDrive sites in the report. By default, OneDrive sites are excluded.

    .PARAMETER ConfigFile
    The path to the configuration file in JSON format. If not specified, defaults to '<ScriptBaseName>.config' in the script directory.
    The configuration file can optionally include a 'TenantDisplayName' property, which is used to build readable report file names.
    If 'TenantDisplayName' is not present, a short tenant name derived from the admin URL is used instead.

    .PARAMETER ReportDirectory
    The directory where the reports will be saved. The default value is 'Reports'.

    .PARAMETER ExportCsv
    If specified, the script will export the reports in CSV format. By default, reports are generated in HTML format.

    .PARAMETER ClearReports
    If specified, the script will clear the existing reports in the report directory before generating new reports.
    By default, existing reports are not cleared.

    .EXAMPLE
    Get-PnPSiteOwners.ps1 -IncludeOneDriveSites

    This script will connect to the SharePoint Online tenant and retrieve all sites, including OneDrive sites, and their owners.

    .EXAMPLE
    Get-PnPSiteOwners.ps1

    This script will connect to the SharePoint Online tenant and retrieve all sites, excluding OneDrive sites, and their owners.

    .EXAMPLE
    Get-PnPSiteOwners.ps1 -ConfigFile 'contoso.config' -ExportCsv -ClearReports

    This script will use a custom configuration file, export CSV reports in addition to HTML, and clear existing reports first.

    .NOTES
    Target platform: SharePoint Online
    Required modules: PnP.PowerShell
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [switch]$IncludeOneDriveSites,
    [string]$ConfigFile,
    [string]$ReportDirectory = 'Reports',
    [switch]$ExportCsv,
    [switch]$ClearReports
)

function Resolve-EntraIdOwner {
    <#
        .SYNOPSIS
        Resolves an Entra ID user by UserPrincipalName and caches the result.

        .DESCRIPTION
        Looks up a user via Get-PnPEntraIDUser and stores the result in the supplied cache
        so that repeated owners across many sites are only resolved once per script run.

        .PARAMETER Connection
        The PnP connection used to query Entra ID.

        .PARAMETER UserPrincipalName
        The UserPrincipalName to resolve.

        .PARAMETER Cache
        A hashtable used to cache already resolved users across calls.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [hashtable]$Cache
    )

    if ($Cache.ContainsKey($UserPrincipalName)) {
        return $Cache[$UserPrincipalName]
    }

    $resolvedUser = Get-PnPEntraIDUser -Connection $Connection -Identity $UserPrincipalName -ErrorAction SilentlyContinue
    $Cache[$UserPrincipalName] = $resolvedUser

    return $resolvedUser
}

# Default the config file name to '<ScriptBaseName>.config' if not explicitly specified
$scriptDir = $PSScriptRoot
$scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)

if ([string]::IsNullOrWhiteSpace($ConfigFile)) {
    $ConfigFile = ('{0}.config' -f $scriptBaseName)
}

# Load configuration from the specified config file
$configFilePath = Join-Path -Path $scriptDir -ChildPath $ConfigFile

if (Test-Path -Path $configFilePath) {
    Write-Host ('Loading configuration from {0}' -f $configFilePath)
    try {
        $config = Get-Content -Path $configFilePath | ConvertFrom-Json
        # Extract configuration values
        $tenantId = $config.tenantid
        $adminUrl = $config.AdminUrl
        $clientId = $config.ClientId
        $certThumbprint = $config.CertThumbprint
    }
    catch {
        Write-Error ('Failed to load configuration from {0}: {1}' -f $configFilePath, $_.Exception.Message)
        exit
    }
}
else {
    Write-Error ('Configuration file not found: {0}' -f $configFilePath)
    exit
}

Write-Host ('Connecting to SharePoint Online tenant {0} at {1}' -f $tenantId, $adminUrl)

# Connect to the SharePoint Online tenant using PnP PowerShell
try {
    $connAdmin = Connect-PnPOnline -ReturnConnection -Tenant $tenantId -Url $adminUrl -ClientId $clientId -Thumbprint $certThumbprint -ErrorAction Stop
}
catch {
    Write-Error ('Failed to connect to {0}: {1}' -f $adminUrl, $_.Exception.Message)
    exit
}

# Get all tenant sites
# Note: The -IncludeOneDriveSites parameter is used to include OneDrive sites in the results.
# If you want to limit the number of sites returned, you can use the -First parameter.
# For example, to get the first 50 sites, you can use:
# $allTenantSites = Get-PnPTenantSite -Connection $connAdmin | Select-Object -First 50 | Sort-Object Url

try {
    if ($IncludeOneDriveSites) {
        Write-Host 'Including OneDrive sites in the report...'
        $allTenantSites = Get-PnPTenantSite -Connection $connAdmin -IncludeOneDriveSites -ErrorAction Stop | Sort-Object Url
        $onedriveSuffix = '-OD'
    }
    else {
        Write-Host 'Excluding OneDrive sites from the report...'
        $allTenantSites = Get-PnPTenantSite -Connection $connAdmin -ErrorAction Stop | Sort-Object Url
        $onedriveSuffix = ''
    }
}
catch {
    Write-Error ('Failed to retrieve tenant sites: {0}' -f $_.Exception.Message)
    exit
}

# variables required for the script
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$reportPath = Join-Path -Path $scriptDir -ChildPath $ReportDirectory
$entraUserCache = @{}

# Determine a readable tenant name for use in report file names.
# Uses the optional 'TenantDisplayName' property from the config file, falling back to a short name derived from the admin URL.
if ($config.TenantDisplayName) {
    $tenantDisplayName = $config.TenantDisplayName
}
else {
    $tenantDisplayName = $adminUrl -replace '-admin\.sharepoint\.com$', ''
}

# Remove characters that are not valid in file names
$invalidFileNameChars = [System.Text.RegularExpressions.Regex]::Escape([System.IO.Path]::GetInvalidFileNameChars() -join '')
$tenantDisplayName = $tenantDisplayName -replace "[$invalidFileNameChars]", '_'

# Check if the report directory exists, if not, create it
if (-not (Test-Path -Path $reportPath)) {
    $null = New-Item -Path $reportPath -ItemType Directory
}
else {
    # If the directory exists and ClearReports is specified, clear the reports
    if ($ClearReports -and $PSCmdlet.ShouldProcess($reportPath, 'Remove existing reports')) {
        Remove-Item -Path $reportPath -Force -Recurse -ErrorAction SilentlyContinue
        $null = New-Item -Path $reportPath -ItemType Directory
    }
}

#region Report Variables

$sitesReport = @()
$sitesCsvReportFilename = (Join-Path -Path $reportPath -ChildPath ('SitesReport_{0}{1}_{2}.csv' -f $tenantDisplayName, $onedriveSuffix, $timestamp))
$sitesHtmlReportFilename = (Join-Path -Path $reportPath -ChildPath ('SitesReport_{0}{1}_{2}.html' -f $tenantDisplayName, $onedriveSuffix, $timestamp))

$ownersReport = @()
$ownersCsvReportFilename = (Join-Path -Path $reportPath -ChildPath ('OwnersReport_{0}{1}_{2}.csv' -f $tenantDisplayName, $onedriveSuffix, $timestamp))
$ownersHtmlReportFilename = (Join-Path -Path $reportPath -ChildPath ('OwnersReport_{0}{1}_{2}.html' -f $tenantDisplayName, $onedriveSuffix, $timestamp))

$sitesErrorReport = @()
$errorsCsvReportFilename = (Join-Path -Path $reportPath -ChildPath ('SitesErrorReport_{0}{1}_{2}.csv' -f $tenantDisplayName, $onedriveSuffix, $timestamp))
$errorsHtmlReportFilename = (Join-Path -Path $reportPath -ChildPath ('SitesErrorReport_{0}{1}_{2}.html' -f $tenantDisplayName, $onedriveSuffix, $timestamp))

$unresolvedUpn = @()
$unresolvedUpnReportFilename = (Join-Path -Path $reportPath -ChildPath ('UnresolvedUPNReport_{0}{1}_{2}.txt' -f $tenantDisplayName, $onedriveSuffix, $timestamp))
#endregion

# nice counter for progress
Write-Host ('Found {0} sites in the tenant.' -f $allTenantSites.Count) -ForegroundColor Green
$i = 1

$redirectSiteCount = ($allTenantSites | Where-Object { $_.Template -eq 'RedirectSite#1' }).Count
if ($redirectSiteCount -gt 0) {
    Write-Host ('Excluding {0} Redirect sites from the report.' -f $redirectSiteCount) -ForegroundColor Yellow
}

$allTenantSites = $allTenantSites | Where-Object { $_.Template -ne 'RedirectSite#1' } # Exclude Redirect sites

foreach ($tenantSite in $allTenantSites) {

    Write-Progress -Activity 'Fetching Sites' -Status ('Site {0}/{1}: {2}' -f $i, $allTenantSites.Count, $tenantSite.Url) -PercentComplete ([math]::Round(($i / $allTenantSites.Count) * 100, 2))
    Write-Verbose ('URL: {0}' -f $tenantSite.Url)

    # Reset per site variables to avoid leaking state from the previous iteration
    $isGroupSite = $false
    $siteOwnerEmail = ''
    $siteOwnerDisplayName = ''
    $siteOwnersReport = @()
    $siteAdminCollection = @()

    try {
        # Connect to the site collection for further PnP actions
        $connSite = Connect-PnPOnline -ReturnConnection -Tenant $tenantId -Url $tenantSite.Url -ClientId $clientId -Thumbprint $certThumbprint -ErrorAction Stop

        # Connect to the site collection itself
        $site = Get-PnPSite -Connection $connSite -Includes RootWeb, GroupId, Owner

        if ($site.GroupId.Guid -eq '00000000-0000-0000-0000-000000000000') {
            # determine admins of a non-group site
            $isGroupSite = $false
            $ownerType = 'Site Collection Administrator'
            $siteAdmins = Get-PnPSiteCollectionAdmin -Connection $connSite | Where-Object { $_.PrincipalType -eq 'User' }

            foreach ($siteAdmin in $siteAdmins) {
                if (-not $siteAdmin.UserPrincipalName) {
                    # maybe guest user or group, fallback to login name
                    $userPrincipalName = $siteAdmin.LoginName
                }
                else {
                    $userPrincipalName = $siteAdmin.UserPrincipalName
                }

                # Add site admin to the collection for later use
                $siteAdminCollection += [PSCustomObject]@{
                    DisplayName       = $siteAdmin.Title
                    UserPrincipalName = $userPrincipalName
                    OwnerEmail        = $siteAdmin.Email
                    OwnerType         = $ownerType
                }
            }
        }
        else {
            # determine admins of a group site
            $isGroupSite = $true
            $ownerType = 'Group Owner'
            $siteAdmins = Get-PnPEntraIDGroupOwner -Connection $connAdmin -Identity $site.GroupId.Guid

            foreach ($siteAdmin in $siteAdmins) {
                # Add site admin to the collection for later use
                $siteAdminCollection += [PSCustomObject]@{
                    DisplayName       = $siteAdmin.DisplayName
                    UserPrincipalName = $siteAdmin.UserPrincipalName
                    OwnerEmail        = $siteAdmin.Mail
                    OwnerType         = $ownerType
                }
            }
        }

        foreach ($siteAdmin in $siteAdminCollection) {

            if ($siteAdmin.UserPrincipalName.StartsWith('i:0#.f|membership|')) {
                # remove the guest prefix for Entra ID user lookup
                $userPrincipalName = $siteAdmin.UserPrincipalName.Substring(18)
            }
            else {
                # use the UserPrincipalName as is
                $userPrincipalName = $siteAdmin.UserPrincipalName
            }

            $entraUser = Resolve-EntraIdOwner -Connection $connAdmin -UserPrincipalName $userPrincipalName -Cache $entraUserCache

            if (-not $entraUser) {
                Write-Warning ('Entra User not found: {0}' -f $siteAdmin.UserPrincipalName)
                $unresolvedUpn += $siteAdmin.UserPrincipalName
                continue
            }

            if ($null -eq $entraUser.Mail) {
                $siteOwnerEmail += $siteAdmin.UserPrincipalName + '; '
                $siteOwnerDisplayName += $siteAdmin.DisplayName + '; '
                $ownerEmail = $siteAdmin.UserPrincipalName
                $accountEnabled = $null
            }
            else {
                $siteOwnerEmail += $entraUser.Mail + '; '
                $siteOwnerDisplayName += $siteAdmin.DisplayName + '; '
                $ownerEmail = $entraUser.Mail
                $accountEnabled = $entraUser.AccountEnabled
            }

            # Add site admin to the report
            $siteOwnersReport += [PSCustomObject]@{
                SiteUrl        = $site.Url
                SiteTitle      = $site.RootWeb.Title
                IsGroupSite    = $isGroupSite
                OwnerEmail     = $ownerEmail
                OwnerName      = $siteAdmin.DisplayName
                OwnerType      = $ownerType
                AccountEnabled = $accountEnabled
            }
        }

        $ownersReport += $siteOwnersReport

        $sitesReport += [PSCustomObject]@{
            SiteUrl     = $site.Url
            SiteTitle   = $site.RootWeb.Title
            IsGroupSite = $isGroupSite
            OwnerEmail  = $siteOwnerEmail
            OwnerName   = $siteOwnerDisplayName
        }
    }
    catch {
        Write-Warning ('Error processing site: {0}' -f $tenantSite.Url)

        $sitesErrorReport += [PSCustomObject]@{
            SiteUrl      = $tenantSite.Url
            IsGroupSite  = $isGroupSite
            ErrorMessage = $_.Exception.Message
        }
        continue
    }

    $i++
}

# Define CSS styles for HTML reports
$header = @"
<style>
    :root {
        --header-bg: #1b3a57;
        --header-fg: #ffffff;
        --row-odd: #ffffff;
        --row-even: #f2f6fa;
        --row-hover: #e4edf5;
        --border-color: #d7dee5;
        --accent: #2f6fb0;
    }
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background-color: #f7f9fb;
        margin: 0;
        padding: 30px;
        color: #23303d;
    }
    h1 {
        font-weight: 600;
        color: var(--header-bg);
        border-bottom: 3px solid var(--accent);
        padding-bottom: 8px;
        margin-bottom: 20px;
    }
    table {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        border-collapse: collapse;
        width: 100%;
        background-color: #ffffff;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
    }
    th {
        background-color: var(--header-bg);
        color: var(--header-fg);
        text-align: left;
        padding: 10px 12px;
    }
    td {
        padding: 8px 12px;
        border-bottom: 1px solid var(--border-color);
    }
    tr:nth-child(odd) {background-color: var(--row-odd);}
    tr:nth-child(even) {background-color: var(--row-even);}
    tr:hover {background-color: var(--row-hover);}
    p {
        font-size: 0.95em;
        color: #4b5a68;
    }
    strong {
        color: var(--header-bg);
    }
</style>
"@

# Generate reports
Write-Host 'Generating reports...'

if ($sitesReport.Count -gt 0) {
    Write-Host ('Sites: {0}' -f $sitesReport.Count)
    if ($ExportCsv) {
        $sitesReport | Select-Object SiteUrl, SiteTitle, IsGroupSite, OwnerEmail, OwnerName | Export-Csv -Path $sitesCsvReportFilename -NoTypeInformation -Encoding UTF8 -Force
    }
    $sitesReport | Select-Object SiteUrl, SiteTitle, IsGroupSite, OwnerEmail, OwnerName | ConvertTo-Html -As Table -PreContent '<h1>Sites Report</h1>' -PostContent ('<p>Site Count: <strong>{0}</strong></p>' -f $sitesReport.Count) -Head $header | Out-File $sitesHtmlReportFilename -Encoding UTF8 -Force
}
else {
    Write-Host 'No sites found in the tenant.' -ForegroundColor Yellow
}

if ($ownersReport.Count -gt 0) {
    Write-Host ('Owners: {0}' -f $ownersReport.Count)
    if ($ExportCsv) {
        $ownersReport | Select-Object SiteUrl, SiteTitle, IsGroupSite, OwnerEmail, OwnerName, OwnerType, AccountEnabled | Export-Csv -Path $ownersCsvReportFilename -NoTypeInformation -Encoding UTF8 -Force
    }
    $ownersReport | Select-Object SiteUrl, SiteTitle, IsGroupSite, OwnerEmail, OwnerName, OwnerType, AccountEnabled | ConvertTo-Html -As Table -PreContent '<h1>Owners Report</h1>' -PostContent ('<p>Owner Count: <strong>{0}</strong></p>' -f $ownersReport.Count) -Head $header | Out-File $ownersHtmlReportFilename -Encoding UTF8 -Force
}
else {
    Write-Host 'No owners found in the tenant.' -ForegroundColor Yellow
}

if ($sitesErrorReport.Count -gt 0) {
    Write-Host ('Errors: {0}' -f $sitesErrorReport.Count)
    if ($ExportCsv) {
        $sitesErrorReport | Select-Object SiteUrl, IsGroupSite, ErrorMessage | Export-Csv -Path $errorsCsvReportFilename -NoTypeInformation -Encoding UTF8 -Force
    }
    $sitesErrorReport | Select-Object SiteUrl, IsGroupSite, ErrorMessage | ConvertTo-Html -As Table -PreContent '<h1>Sites Error Report</h1>' -PostContent ('<p>Error Count: <strong>{0}</strong></p>' -f $sitesErrorReport.Count) -Head $header | Out-File $errorsHtmlReportFilename -Encoding UTF8 -Force
}
else {
    Write-Host 'No sites with errors found.' -ForegroundColor Yellow
}

if ($unresolvedUpn.Count -gt 0) {
    # Always written regardless of -ExportCsv, since this is diagnostic information rather than a formatted report
    Write-Host ('Unresolved UPNs: {0}' -f $unresolvedUpn.Count)
    $unresolvedUpn | Sort-Object -Unique | Out-File -FilePath $unresolvedUpnReportFilename -Encoding UTF8 -Force
}
else {
    Write-Host 'No unresolved UPNs found.' -ForegroundColor Yellow
}

Write-Host ('Reports generated in: {0}' -f $reportPath)
Write-Host ('Sites Report: {0}' -f $sitesHtmlReportFilename)
Write-Host ('Owners Report: {0}' -f $ownersHtmlReportFilename)