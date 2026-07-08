<#
.SYNOPSIS
    Generates HTML reports for Enterprise Apps and App Registrations in a given Entra tenant.
    Highlights newly added items since the last run.

.DESCRIPTION
    - Queries Enterprise Apps (Service Principals) and App Registrations (Applications).
    - Outputs two HTML reports, sorted by Display Name.
    - Saves current data as JSON for change tracking.
    - Newly added items are highlighted in the HTML report.

.NOTES
    Requires: Microsoft.Graph PowerShell module

.PARAMETER TenantId
    The Tenant ID of the Entra tenant to query.

.PARAMETER EnterpriseAppFilter
    Filter for Enterprise Apps:
    - "All": Retrieve all Enterprise Apps.
    - "NonMicrosoftOnly": Retrieve only custom (non-Microsoft) Enterprise Apps. Default.

.PARAMETER Highlight3rdPartyOwners
    If set, highlights Enterprise Apps owned by third-party organizations.

.PARAMETER ExportCSV
    If set, exports the data to CSV files in addition to HTML reports.
#>
[CmdletBinding(DefaultParameterSetName = 'Manual')]
param (
    [Parameter(Mandatory, ParameterSetName = 'Manual')]
    [string]$TenantId,

    [Parameter()]
    [ValidateSet("All", "NonMicrosoftOnly")]
    [string]$EnterpriseAppFilter = "NonMicrosoftOnly",

    [switch]$Highlight3rdPartyOwners,
    [switch]$ExportCSV,
    [Parameter(Mandatory, ParameterSetName = 'Config')]     
    [string]$ConfigFile = 'dev-egxde.json',
    [switch]$DailyReport,
    [int]$ExpireInDays = 30
)

# Output directory for reports and JSON files
$OutputPath = Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "Reports"
if (-not (Test-Path $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory | Out-Null }

# Ensure Microsoft.Graph module is installed
<#
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
    #>

Write-Verbose "Importing Microsoft.Graph module..."
Import-Module Microsoft.Graph

# Load configuration from JSON file
$configFilePath = Join-Path -Path (Split-Path -Path $script:MyInvocation.MyCommand.Path) -ChildPath $ConfigFile

if (Test-Path -Path $configFilePath) {
    Write-Host ('Loading configuration from {0}' -f $configFilePath)
    try {
        $config = Get-Content -Path $configFilePath | ConvertFrom-Json
        # Extract configuration values
        $tenantId = $config.tenantid
        $clientid = $config.ClientId
        $certThumbprint = $config.CertThumbprint
    }
    catch {
        Write-Error ('Failed to load configuration from {0}: {1}' -f $configFilePath, $_.Exception.Message)
        exit
    }
}
else {
    # exit script if config file not found
    Write-Error ('Configuration file not found: {0}' -f $configFilePath)
    exit
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..."
if ( $clientid -and $certThumbprint) {
    Write-Host "Using certificate-based authentication with ClientId $clientid and TenantId $tenantId"
    Connect-MgGraph -ClientId $clientid -TenantId $tenantId -CertificateThumbprint $certThumbprint -NoWelcome
}
else {
    Connect-MgGraph -TenantId $TenantId -Scopes "Application.Read.All", "Directory.Read.All", "AuditLog.Read.All" -NoWelcome
}

#region Functions

# Helper function to get owners
function Get-ServicePrincipalOwners {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServicePrincipalId
    )

    try {
        # Retrieve all owners of the specified Service Principal
        $owners = Get-MgServicePrincipalOwner -ServicePrincipalId $ServicePrincipalId

        # Resolve DisplayNames via Get-MgUser
        $resolvedNames = @()

        foreach ($owner in $owners) {
            try {
                $user = Get-MgUser -UserId $owner.Id -errorAction SilentlyContinue
                $resolvedNames += $user.DisplayName
            }
            catch {
                Write-Warning "Could not resolve user with ID: $($owner.Id)"
            }
        }

        # Return comma-separated string of DisplayNames
        return ($resolvedNames -join ", ")

    }
    catch {
        # In case of an error, return a message or empty string
        Write-Warning "Failed to retrieve owners for ServicePrincipalId: $ServicePrincipalId"
        Write-Warning $_.Exception.Message
        return ""
    }
}

function Get-AppRegistrationOwnersAsString {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ApplicationId
    )

    try {
        # Retrieve all owners of the specified App Registration
        $owners = Get-MgApplicationOwner -ApplicationId $ApplicationId
        
        # Resolve DisplayNames via Get-MgUser
        $resolvedNames = @()
        foreach ($owner in $owners) {
            try {
                $resolvedNames += $owner.AdditionalProperties.displayName
            }
            catch {
                Write-Warning "Could not resolve user with ID: $($owner.Id)"
            }
        }
        # Return comma-separated string of DisplayNames
        return ($resolvedNames -join ", ")
    }
    catch {
        # In case of an error, return a message or empty string
        Write-Warning "Failed to retrieve owners for ApplicationId: $ApplicationId"
        Write-Warning $_.Exception.Message
        return ""
    }
}

function Get-ExpiredCertificates {
    param (
        [array]$KeyCredentials
    )

    $expirationDate = (Get-Date).AddDays($ExpireInDays)

    $expiredCerts = @()
        
    foreach ($cert in $KeyCredentials) {
        $remainingDays = ($cert.EndDateTime - (Get-Date)).Days

        if($remainingDays -le $ExpireInDays) {
            $expiredCerts += [PSCustomObject]@{
                'Certificate Name' = $cert.DisplayName
            }
        }
    }

    return $expiredCerts
}

function Get-ExpiredSecrets {
    param (
        [array]$PasswordCredentials
    )

    $expirationDate = (Get-Date).AddDays($ExpireInDays)

    $expiredSecrets = @()
        
    foreach ($secret in $PasswordCredentials) {
        $remainingDays = ($secret.EndDateTime - (Get-Date)).Days

        if($remainingDays -le $ExpireInDays) {
            $expiredSecrets += [PSCustomObject]@{
                'Secret Name' = $secret.DisplayName
            }
        }
    }

    return $expiredSecrets
}

# Helper function to get last sign-in date (if available)
function Get-LastSignIn {
    param($Id)
    try {
        $signIn = Get-MgAuditLogSignIn -Filter "AppId eq '$Id'" -Top 1 | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
        if ($signIn) { $signIn.CreatedDateTime } else { "" }
    }
    catch {
        ""
    }
}

function Get-TenantName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId
    )
    try {
        $org = Get-MgOrganization -OrganizationId $TenantId -ErrorAction Stop | Select-Object -First 1
        if ($org.DisplayName) { return $org.DisplayName }
        return $org.Id
    }
    catch {
        try {
            $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
            if ($org.DisplayName) { return $org.DisplayName }
            return $org.Id
        }
        catch {
            return "N/A"
        }
    }
}

function Get-TenantDisplayNameFromId {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    if($null -eq $TenantId -or $TenantId -eq ""){
        return ""
    }
    
    try {
        # Call the Microsoft Graph Beta API to get tenant information
        $tenantInfo = Find-MgBetaTenantRelationshipTenantInformationByTenantId -TenantId $TenantId -errorAction SilentlyContinue

        # Return the display name of the tenant
        return $tenantInfo.DisplayName
    }
    catch {
        # Handle errors gracefully
        Write-Warning "Could not retrieve tenant information for ID: $TenantId"
        Write-Warning $_.Exception.Message
        return ""
    }
}

#endregion  

# Query Enterprise Apps (Service Principals)
Write-Progress -Activity "Querying Enterprise Apps" -Status "Retrieving Service Principals..." -PercentComplete 10

if ($EnterpriseAppFilter -eq "NonMicrosoftOnly") {
    # Get only custom enterprise apps (not Microsoft or third-party apps)
    $spList = Get-MgBetaServicePrincipal -All -Filter "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')" | Select-Object Id, DisplayName, AppId, CreatedDateTime, Notes, Description, AdditionalProperties, AppOwnerOrganizationId, KeyCredentials
}
else {
    # Get all enterprise apps
    $spList = Get-MgServicePrincipal -All | Select-Object Id, DisplayName, AppId, CreatedDateTime, Notes, Description, AdditionalProperties, AppOwnerOrganizationId, KeyCredentials
}

Write-Progress -Activity "Processing Enterprise Apps" -Status "Building Enterprise App data objects..." -PercentComplete 20
$spData = foreach ($sp in $spList) {

    # Extract createdDateTime from AdditionalProperties if available
    $createdDate = if ($sp.AdditionalProperties.ContainsKey("createdDateTime")) { $sp.AdditionalProperties["createdDateTime"] } else { $null }

    # Build custom object
    [PSCustomObject]@{
        DisplayName          = $sp.DisplayName
        ApplicationId        = $sp.AppId
        Owner                = Get-ServicePrincipalOwners -ServicePrincipalId $sp.Id
        InternalNotes        = $sp.Notes
        Description          = $sp.Description
        ExpiredCertificates  = $expiredCerts
        CreatedDate          = if ($createdDate) { (Get-Date $createdDate).ToString('yyyy-MM-dd') } else { "N/A" }
        LastSignIn           = '' # Get-LastSignIn $sp.AppId
        Id                   = $sp.Id
        AppOwnerOrganization = if ($null -ne $sp.AppOwnerOrganizationId) { Get-TenantDisplayNameFromId -TenantId $sp.AppOwnerOrganizationId } else { "N/A" }
    }
}

# Query App Registrations (Applications)
Write-Progress -Activity "Querying App Registrations" -Status "Retrieving Applications..." -PercentComplete 30
$appList = Get-MgApplication -All | Select-Object Id, DisplayName, AppId, CreatedDateTime, Notes, Description, SignInAudience, PasswordCredentials, KeyCredentials

Write-Progress -Activity "Processing App Registrations" -Status "Building App Registration data objects..." -PercentComplete 40
$appData = foreach ($app in $appList) {
    [PSCustomObject]@{
        DisplayName    = $app.DisplayName
        ApplicationId  = $app.AppId
        Owner          = Get-AppRegistrationOwnersAsString -ApplicationId $app.Id
        InternalNotes  = $app.Notes
        Description    = $app.Description
        CreatedDate    = if ($app.CreatedDateTime) { (Get-Date $app.CreatedDateTime).ToString('yyyy-MM-dd') } else { "N/A" }
        LastSignIn     = '' #Get-LastSignIn $app.AppId
        SignInAudience = $app.SignInAudience
        ExpiredCertificates  = ((Get-ExpiredCertificates -KeyCredentials $app.KeyCredentials) | Measure-Object).Count
        ExpiredSecrets       = ((Get-ExpiredSecrets -PasswordCredentials $app.PasswordCredentials) | Measure-Object).Count
        Id             = $app.Id
    }
}

# Get tenant name for report titles and filenames
Write-Progress -Activity "Retrieving Tenant Name" -Status "Getting tenant display name..." -PercentComplete 45
$tenantName = Get-TenantName -TenantId $TenantId
$safeTenant = ($tenantName -replace '[^a-zA-Z0-9]', '')

# Sort alphabetically
Write-Progress -Activity "Sorting Data" -Status "Sorting Enterprise Apps and App Registrations..." -PercentComplete 50
$spDataSorted = $spData | Sort-Object DisplayName
$appDataSorted = $appData | Sort-Object DisplayName

# Load previous data for highlighting new items
Write-Progress -Activity "Loading Previous Data" -Status "Checking for previous JSON files..." -PercentComplete 60
$spJsonPath = Join-Path -Path $OutputPath -ChildPath ("{0}-EnterpriseApps.json" -f $safeTenant)
$appJsonPath = Join-Path -Path $OutputPath -ChildPath ("{0}-AppRegistrations.json" -f $safeTenant)

# Initialize arrays to hold previous IDs
$prevSpIds = @()
$prevAppIds = @()
$firstRunSpIds = $false
$firstRunAppIds = $false

if (Test-Path $spJsonPath) {
    $prevSpIds = (Get-Content $spJsonPath | ConvertFrom-Json).Id
}
else {
    Write-Host "No previous Enterprise Apps JSON file found. All items will be treated as new."
    $firstRunSpIds = $true
}
if (Test-Path $appJsonPath) {
    $prevAppIds = (Get-Content $appJsonPath | ConvertFrom-Json).Id
}
else {
    Write-Host "No previous App Registrations JSON file found. All items will be treated as new."
    $firstRunAppIds = $true
}

# Save current data as JSON
Write-Progress -Activity "Saving Data" -Status "Writing current data to JSON files..." -PercentComplete 70
$spDataSorted | ConvertTo-Json -Depth 4 | Set-Content $spJsonPath
$appDataSorted | ConvertTo-Json -Depth 4 | Set-Content $appJsonPath

# HTML style for highlighting new items
$highlightStyle = "font-weight:bold;background-color:#ffcccc;"
$generatedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$newSpCount = @($spDataSorted | Where-Object { -not (($prevSpIds -contains $_.Id) -or $firstRunSpIds) }).Count
$newAppCount = @($appDataSorted | Where-Object { -not (($prevAppIds -contains $_.Id) -or $firstRunAppIds) }).Count

# Generate HTML for Enterprise Apps
Write-Progress -Activity "Generating HTML" -Status "Building Enterprise Apps HTML report..." -PercentComplete 80
$spRows = foreach ($sp in $spDataSorted) {
    $isNew = -not (($prevSpIds -contains $sp.Id) -or $firstRunSpIds)
    $style = if ($isNew) { $highlightStyle } else { "" }
    "<tr>
        <td style='$style'>$($sp.DisplayName)</td>
        <td class='no-wrap'>$($sp.ApplicationId)</td>
        <td>$($sp.Owner)</td>
        $(
            if ($Highlight3rdPartyOwners -and $sp.AppOwnerOrganization -and ($sp.AppOwnerOrganization -ne $tenantName)) {
                "<td><b>$($sp.AppOwnerOrganization)</b></td>"
            } else {
                "<td>$($sp.AppOwnerOrganization)</td>"
            }
        )
        <td>$($sp.InternalNotes)</td>
        <td class='no-wrap'>$($sp.CreatedDate)</td>

    </tr>"
}

$spHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Enterprise Apps Report ($($tenantName))</title>
    <style>
        body { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; background: #f4f6f9; color: #222; margin: 0; padding: 20px; }
        h1 { color: #0078d4; margin-bottom: 4px; }
        .meta { color: #555; margin-bottom: 12px; font-size: 12px; }
        .filter-info { background: #e8f0fe; border-left: 4px solid #0078d4; padding: 8px 12px; margin-bottom: 16px; border-radius: 4px; }
        table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
        th { background: #0078d4; color: #fff; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }
        td { padding: 8px 12px; border-bottom: 1px solid #e5e9f0; vertical-align: top; }
        tr:last-child td { border-bottom: none; }
        tr:hover td { filter: brightness(0.96); }
        td.no-wrap { white-space: nowrap; }
        input[type=search] { padding: 7px 12px; width: 320px; border: 1px solid #ccc; border-radius: 20px; font-size: 13px; margin-bottom: 12px; outline: none; }
        input[type=search]:focus { border-color: #0078d4; box-shadow: 0 0 0 2px #c7e0f4; }
    </style>
</head>
<body>
    <h1>Enterprise Apps Report</h1>
    <div class="meta">Tenant: <strong>$tenantName</strong> &nbsp;|&nbsp; Generated: $generatedOn &nbsp;|&nbsp; Total Apps: <strong>$($spDataSorted.Count)</strong> &nbsp;|&nbsp; Newly Added: <strong>$newSpCount</strong> &nbsp;|&nbsp; Visible rows: <span id="spVisibleRowCount">$($spDataSorted.Count)</span></div>
    <div class="filter-info">$(if ($EnterpriseAppFilter -eq "NonMicrosoftOnly") { ("Filtered to show only custom (non-Microsoft) apps.") } else { ("Showing all enterprise apps.") })</div>
    <input type="search" id="spSearchBox" placeholder="Search enterprise apps..." oninput="filterSpTable()" />
    <table>
        <thead>
            <tr>
                <th>Display Name</th>
                <th><a href="https://granikos.eu/go/EntraAppId" target="_blank" title="Learn more about Application Id" style="color:#fff;">Application Id</a></th>
                <th>Owner</th>
                <th>App Owner Organization</th>
                <th>Internal Notes</th>
                <th>Created Date</th>
            </tr>
        </thead>
        <tbody id="spTableBody">
            $($spRows -join "`n")
        </tbody>
    </table>
    <p><b>Legend:</b> <span style='$highlightStyle'>Newly added since last report</span></p>
    <script>
        function filterSpTable() {
            var q = document.getElementById('spSearchBox').value.toLowerCase();
            var rows = document.getElementById('spTableBody').querySelectorAll('tr');
            var visible = 0;
            rows.forEach(function(row) {
                var text = row.innerText.toLowerCase();
                var show = text.indexOf(q) > -1;
                row.style.display = show ? '' : 'none';
                if (show) visible++;
            });
            document.getElementById('spVisibleRowCount').innerText = visible;
        }
    </script>
</body>
</html>
"@

if($DailyReport){
    # Save Enterprise Apps HTML with fixed name for daily report
    $entAppReportFilename = ("{1}-EnterpriseApps-{0}.html" -f (Get-Date -Format "yyyyMMdd"), $safeTenant)
}
else{
    # Save Enterprise Apps HTML with timestamp
    $entAppReportFilename = ("{1}-EnterpriseApps-{0}.html" -f (Get-Date -Format "yyyyMMdd-HHmmss"), $safeTenant)
}

$spHtmlPath = Join-Path -Path $OutputPath -ChildPath $entAppReportFilename
$spHtml | Set-Content -Path $spHtmlPath

# Generate HTML for App Registrations
Write-Progress -Activity "Generating HTML" -Status "Building App Registrations HTML report..." -PercentComplete 90
$appRows = foreach ($app in $appDataSorted) {
    $isNew = -not (($prevAppIds -contains $app.Id) -or $firstRunAppIds)
    $style = if ($isNew) { $highlightStyle } else { "" }
    
    # Format expired certs/secrets with bold red if > 0
    $certSecretDisplay = if ($app.ExpiredCertificates -gt 0 -or $app.ExpiredSecrets -gt 0) {
        "<span style='font-weight:bold;color:red;'>$($app.ExpiredCertificates)/$($app.ExpiredSecrets)</span>"
    } else {
        "$($app.ExpiredCertificates)/$($app.ExpiredSecrets)"
    }
    
    "<tr>
        <td style='$style'>$($app.DisplayName)</td>
        <td class='no-wrap'>$($app.ApplicationId)</td>
        <td>$($app.Owner)</td>
        <td>$($app.SignInAudience)</td>
        <td>$($app.InternalNotes)</td>
        <td class='desc-cell'>$($app.Description)</td>
        <td class='no-wrap'>$certSecretDisplay</td>
        <td class='no-wrap'>$($app.CreatedDate)</td>
       
    </tr>"
}

$appHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>App Registrations Report ($($tenantName))</title>
    <style>
        body { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; background: #f4f6f9; color: #222; margin: 0; padding: 20px; }
        h1 { color: #0078d4; margin-bottom: 4px; }
        .meta { color: #555; margin-bottom: 16px; font-size: 12px; }
        table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
        th { background: #0078d4; color: #fff; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }
        td { padding: 8px 12px; border-bottom: 1px solid #e5e9f0; vertical-align: top; }
        tr:last-child td { border-bottom: none; }
        tr:hover td { filter: brightness(0.96); }
        td.desc-cell { font-size: 0.9em; color: #444; }
        td.no-wrap { white-space: nowrap; }
        input[type=search] { padding: 7px 12px; width: 320px; border: 1px solid #ccc; border-radius: 20px; font-size: 13px; margin-bottom: 12px; outline: none; }
        input[type=search]:focus { border-color: #0078d4; box-shadow: 0 0 0 2px #c7e0f4; }
    </style>
</head>
<body>
    <h1>App Registrations Report</h1>
    <div class="meta">Tenant: <strong>$tenantName</strong> &nbsp;|&nbsp; Generated: $generatedOn &nbsp;|&nbsp; Total Apps: <strong>$($appDataSorted.Count)</strong> &nbsp;|&nbsp; Newly Added: <strong>$newAppCount</strong> &nbsp;|&nbsp; Visible rows: <span id="appVisibleRowCount">$($appDataSorted.Count)</span></div>
    <input type="search" id="appSearchBox" placeholder="Search app registrations..." oninput="filterAppTable()" />
    <table>
        <thead>
            <tr>
                <th>Display Name</th>
                <th>Application Id</th>
                <th>Owner</th>
                <th><a href="https://granikos.eu/go/vjKm" target="_blank" title="Learn more about Sign-In Audience" style="color:#fff;">Sign-In Audience</a></th>
                <th>Internal Notes</th>
                <th>Description</th>
                <th>Expired Certs/Secrets</th>
                <th>Created Date</th>
            </tr>
        </thead>
        <tbody id="appTableBody">
            $($appRows -join "`n")
        </tbody>
    </table>
    <p><b>Legend:</b> <span style='$highlightStyle'>Newly added since last report</span></p>
    <script>
        function filterAppTable() {
            var q = document.getElementById('appSearchBox').value.toLowerCase();
            var rows = document.getElementById('appTableBody').querySelectorAll('tr');
            var visible = 0;
            rows.forEach(function(row) {
                var text = row.innerText.toLowerCase();
                var show = text.indexOf(q) > -1;
                row.style.display = show ? '' : 'none';
                if (show) visible++;
            });
            document.getElementById('appVisibleRowCount').innerText = visible;
        }
    </script>
</body>
</html>
"@

if($DailyReport){
    # Save App Registrations HTML with fixed name for daily report
    $appRegReportFilename = ("{1}-AppRegistrations-{0}.html" -f (Get-Date -Format "yyyyMMdd"),$safeTenant)
}
else{
    # Save App Registrations HTML with timestamp
    $appRegReportFilename = ("{1}-AppRegistrations-{0}.html" -f (Get-Date -Format "yyyyMMdd-HHmmss"), $safeTenant)
}
$appHtmlPath = Join-Path -Path $OutputPath -ChildPath $appRegReportFilename
$appHtml | Set-Content -Path $appHtmlPath


if ($ExportCSV) {
    Write-Progress -Activity "Exporting CSV" -Status "Exporting Enterprise Apps and App Registrations to CSV..." -PercentComplete 95

    $entAppCsvFilename = "{1}-EnterpriseApps-{0}.csv" -f $timestamp, $safeTenant
    $appRegCsvFilename = "{1}-AppRegistrations-{0}.csv" -f $timestamp, $safeTenant

    $spCsvPath = Join-Path -Path $OutputPath -ChildPath $entAppCsvFilename
    $appCsvPath = Join-Path -Path $OutputPath -ChildPath $appRegCsvFilename

    # Select a clean set of properties for CSV export
    $spDataSorted |
    Select-Object DisplayName, ApplicationId, Owner, AppOwnerOrganization, InternalNotes, Description, CreatedDate, LastSignIn, Id |
    Export-Csv -Path $spCsvPath -NoTypeInformation -Encoding UTF8 -Force

    $appDataSorted |
    Select-Object DisplayName, ApplicationId, Owner, SignInAudience, InternalNotes, Description, CreatedDate, LastSignIn, Id |
    Export-Csv -Path $appCsvPath -NoTypeInformation -Encoding UTF8 -Force

    Write-Host "CSV exports generated:"
    Write-Host " - $spCsvPath"
    Write-Host " - $appCsvPath"
}
# Final progress update
Write-Progress -Activity "Completed" -Status "Reports generated." -PercentComplete 100 -Completed

Write-Host "Reports generated:"
Write-Host " - $spHtmlPath"
Write-Host " - $appHtmlPath"