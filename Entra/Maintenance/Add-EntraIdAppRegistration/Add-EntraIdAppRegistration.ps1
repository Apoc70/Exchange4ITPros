<#
    .SYNOPSIS
    Add a custom app registration to Entra ID for Microsoft Graph access

    Thomas Stensitzki

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.6, 2026-09-11

    Based on the work of Andres Bohren
    https://blog.icewolf.ch/archive/2022/12/02/create-azure-ad-app-registration-with-microsoft-graph-powershell

    .NOTES
    Requirements
    - PowerShell 7.1+
    - Account executing this script must be a role member of the Application Administrator or Global Administrator role

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial community release
    1.1     Parameters AppDescription and UseCertificate added
    1.2     Parameter AppLogoPath added
    1.3     Parameters OpenBrowser and PrivateBrowserSession added
    1.4     Required API permissions are now read from AppPermissions.json
    1.5     Parameter GrantAdminConsent added
    1.6     Parameter TenantId added, check for existing app registration added

    .PARAMETER TenantId

    The Tenant ID Guid of the Entra ID tenant to connect to.

    .PARAMETER AppName

    The display name of the application in Entra ID

    .PARAMETER AppDescription

    The description (notes) stored with the application in Entra ID

    .PARAMETER AppSecretName

    The name of the client secret in Entra ID

    .PARAMETER AppOwnerEmailAddress

    The email address of the application owner

    .PARAMETER AppSecretValidityInMonths

    The validity of the client secret in months

    .PARAMETER UseCertificate

    Switch to skip the creation of a client secret. The certificate must be uploaded manually after the application has been created.

    .PARAMETER AppLogoPath

    Optional path to an image file (.png, .jpg, .jpeg, or .gif) used as the application logo.
    The file must not exceed 256 KB, as required by Entra ID.

    .PARAMETER OpenBrowser

    Switch to open the Entra portal in a browser after the application has been created, to grant admin consent.

    .PARAMETER PrivateBrowserSession

    Switch to open the Entra portal in a private/incognito browser window. Requires OpenBrowser.

    .PARAMETER PermissionsConfigPath

    Path to the JSON file that defines the required API permissions of the application.
    Defaults to AppPermissions.json in the script folder.

    .PARAMETER GrantAdminConsent

    Switch to grant admin consent for the configured API permissions by code, instead of using the Entra portal.
    Requires the executing account to be a member of the Privileged Role Administrator or Global Administrator role.

    All permissions and IDs
    https://learn.microsoft.com/graph/permissions-reference#all-permissions-and-ids

#>
[CmdletBinding()]
param(
    [string]$TenantId = $null, #'TENANT GUID', # Adjust to your tenant ID Guid
    [string]$AppName = 'AppRegistration Report',
    [string]$AppDescription = 'Application for App-related MS Graph queries.',
    [string]$AppSecretName = 'AppClientSecret',
    [string]$AppOwnerEmailAddress = 'Admin@TENANT.onmicrosoft.com', #Adjust to your tenant and your admin user UPN
    [int]$AppSecretValidityInMonths = 12, # Validity of the client secret in months
    [switch]$UseCertificate,
    [string]$AppLogoPath,
    [switch]$OpenBrowser,
    [switch]$PrivateBrowserSession,
    [string]$PermissionsConfigPath = (Join-Path -Path $PSScriptRoot -ChildPath 'AppPermissions.json'),
    [switch]$GrantAdminConsent
)

# Maximum application logo file size accepted by Entra ID
$script:MaxAppLogoSizeInBytes = 256KB

# Load the required API permissions before any changes are made in the tenant
if (-not (Test-Path -LiteralPath $PermissionsConfigPath -PathType Leaf)) {
    Write-Warning -Message ('Permissions configuration file not found: {0}' -f $PermissionsConfigPath)
    exit
}

try {
    $permissionsConfig = Get-Content -LiteralPath $PermissionsConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}
catch {
    Write-Warning -Message ('Unable to read the permissions configuration file {0}: {1}' -f $PermissionsConfigPath, $_.Exception.Message)
    exit
}

$requiredResourceAccess = @(
    foreach ($resource in $permissionsConfig.requiredResourceAccess) {

        if ([string]::IsNullOrWhiteSpace($resource.resourceAppId) -or -not $resource.resourceAccess) {
            Write-Warning -Message 'Skipping an incomplete resource entry in the permissions configuration file.'
            continue
        }

        @{
            ResourceAppId  = $resource.resourceAppId
            ResourceAccess = @(
                foreach ($access in $resource.resourceAccess) {

                    if ([string]::IsNullOrWhiteSpace($access.id) -or [string]::IsNullOrWhiteSpace($access.type)) {
                        Write-Warning -Message ('Skipping an incomplete permission entry for resource {0}.' -f $resource.resourceAppId)
                        continue
                    }

                    @{
                        Id   = $access.id
                        Type = $access.type
                    }
                }
            )
        }
    }
)

if ($requiredResourceAccess.Count -eq 0) {
    Write-Warning -Message ('No valid API permissions found in {0}' -f $PermissionsConfigPath)
    exit
}

if ($null -ne (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable).Version) {
    Import-Module -Name Microsoft.Graph.Authentication
}
else {
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Authentication PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Microsoft.Graph'
    exit
}
if ($null -ne (Get-Module -Name Microsoft.Graph.Applications -ListAvailable).Version) {
    Import-Module -Name Microsoft.Graph.Applications
}
else {
    Write-Warning -Message 'Unable to load Import-Module Microsoft.Graph.Applications PowerShell module.'
    Write-Warning -Message 'Open an administrative PowerShell session and run Install-Module Microsoft.Graph'
    exit
}

# Connect to Microsoft Graph
$graphScopes = @('Application.Read.All', 'Application.ReadWrite.All', 'User.Read.All')

if ($GrantAdminConsent) {
    # Granting consent by code requires permission to create app role assignments and delegated permission grants
    $graphScopes += 'AppRoleAssignment.ReadWrite.All'
    $graphScopes += 'DelegatedPermissionGrant.ReadWrite.All'
}

$connectParams = @{
    Scopes    = $graphScopes
    NoWelcome = $true
}

if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
    $connectParams['TenantId'] = $TenantId
}

Connect-MgGraph @connectParams

# Check whether an application registration with the same name already exists
$escapedAppName = $AppName -replace "'", "''"
$existingApp = Get-MgApplication -Filter ("displayName eq '{0}'" -f $escapedAppName) -ErrorAction SilentlyContinue

if ($existingApp) {
    Write-Warning -Message ('An application registration with the name "{0}" already exists in the tenant (AppId: {1}, ObjectId: {2}).' -f $AppName, ($existingApp.AppId -join ', '), ($existingApp.Id -join ', '))
    exit
}

# Create a new application
$newApp= New-MgApplication -DisplayName $AppName -Notes $AppDescription
$appObjectId = $newApp.Id

# Set the owner of the application
$User = Get-MgUser -UserId $AppOwnerEmailAddress
$ObjectId = $User.ID
$NewOwner = @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/{$ObjectId}"
}
$null = New-MgApplicationOwnerByRef -ApplicationId $appObjectId -BodyParameter $NewOwner

# Upload the application logo, if requested
if (-not [string]::IsNullOrWhiteSpace($AppLogoPath)) {

    $logoFile = $null

    if (-not (Test-Path -LiteralPath $AppLogoPath -PathType Leaf)) {
        Write-Warning -Message ('Application logo file not found: {0}' -f $AppLogoPath)
    }
    else {
        $logoFile = Get-Item -LiteralPath $AppLogoPath

        if ($logoFile.Extension -notin @('.png', '.jpg', '.jpeg', '.gif')) {
            Write-Warning -Message ('Unsupported application logo file type "{0}". Supported types are .png, .jpg, .jpeg, and .gif.' -f $logoFile.Extension)
            $logoFile = $null
        }
        elseif ($logoFile.Length -eq 0) {
            Write-Warning -Message ('Application logo file is empty: {0}' -f $logoFile.FullName)
            $logoFile = $null
        }
        elseif ($logoFile.Length -gt $script:MaxAppLogoSizeInBytes) {
            Write-Warning -Message ('Application logo file exceeds the maximum size of {0} KB (file size: {1} KB): {2}' -f ($script:MaxAppLogoSizeInBytes / 1KB), [math]::Round($logoFile.Length / 1KB, 2), $logoFile.FullName)
            $logoFile = $null
        }
    }

    if ($null -ne $logoFile) {
        try {
            $contentType = switch ($logoFile.Extension) {
                '.png' { 'image/png' }
                '.gif' { 'image/gif' }
                default { 'image/jpeg' }
            }

            Set-MgApplicationLogo -ApplicationId $appObjectId -InFile $logoFile.FullName -ContentType $contentType -ErrorAction Stop

            Write-Host ('Application logo uploaded from {0}' -f $logoFile.FullName) -ForegroundColor Green
        }
        catch {
            Write-Warning -Message ('Failed to upload the application logo: {0}' -f $_.Exception.Message)
            Write-Warning -Message 'You can upload the logo manually using the Entra portal (Branding & properties).'
        }
    }
    else {
        Write-Warning -Message 'The application logo has not been uploaded.'
    }
}

if ($UseCertificate) {
    Write-Host 'No client secret has been created, as certificate-based authentication was selected.' -ForegroundColor Yellow
    Write-Host ('Upload your public certificate (.cer) manually to the app registration "{0}" in the Entra portal:' -f $AppName) -ForegroundColor Yellow
    Write-Host ('https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Credentials/appId/{0}' -f $newApp.AppId) -ForegroundColor Yellow
}
else {
    # Create a new client secret for the application
    $newAppSecret = @{
        "displayName" = $AppSecretName
        "endDateTime" = (Get-Date).AddMonths($AppSecretValidityInMonths)
    }
    $appSecret = Add-MgApplicationPassword -ApplicationId $appObjectId -PasswordCredential $newAppSecret

    Write-Host 'Copy the following information to your settings file' -ForegroundColor Green
    Write-Host ('ClientSecret: {0}' -f $appSecret.SecretText) -ForegroundColor Green
}

$params = @{
    RequiredResourceAccess = $requiredResourceAccess
}

# Add permissions to the application
$null = Update-MgApplication -ApplicationId $appObjectId -BodyParameter $params

# Return the application ID for the settings file
Write-Host ('ClientId (App ID): {0}' -f $newApp.AppId) -ForegroundColor Green

# Set the application as a public client with a redirect URI
$RedirectURI = @()
$RedirectURI += "https://login.microsoftonline.com/common/oauth2/nativeclient"

$params = @{
    RedirectUris = @($RedirectURI)
}

$null = Update-MgApplication -ApplicationId $appObjectId -IsFallbackPublicClient -PublicClient $params

if ($GrantAdminConsent) {

    Write-Host 'Granting admin consent for the configured API permissions.'

    try {
        # The service principal (enterprise application) is required to hold the consent
        $appServicePrincipal = Get-MgServicePrincipal -Filter ("appId eq '{0}'" -f $newApp.AppId) -ErrorAction Stop

        if ($null -eq $appServicePrincipal) {
            $appServicePrincipal = New-MgServicePrincipal -AppId $newApp.AppId -ErrorAction Stop
        }

        foreach ($resource in $requiredResourceAccess) {

            $resourceServicePrincipal = Get-MgServicePrincipal -Filter ("appId eq '{0}'" -f $resource.ResourceAppId) -ErrorAction Stop

            if ($null -eq $resourceServicePrincipal) {
                Write-Warning -Message ('No service principal found for resource {0}. Consent has not been granted for this resource.' -f $resource.ResourceAppId)
                continue
            }

            $delegatedScopes = @()

            foreach ($access in $resource.ResourceAccess) {

                if ($access.Type -eq 'Role') {
                    try {
                        $null = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appServicePrincipal.Id -PrincipalId $appServicePrincipal.Id -ResourceId $resourceServicePrincipal.Id -AppRoleId $access.Id -ErrorAction Stop
                        Write-Host ('Application permission {0} granted.' -f $access.Id) -ForegroundColor Green
                    }
                    catch {
                        Write-Warning -Message ('Unable to grant application permission {0}: {1}' -f $access.Id, $_.Exception.Message)
                    }
                }
                else {
                    $scope = $resourceServicePrincipal.Oauth2PermissionScopes | Where-Object { $_.Id -eq $access.Id }

                    if ($null -eq $scope) {
                        Write-Warning -Message ('Unable to resolve delegated permission {0} on resource {1}.' -f $access.Id, $resource.ResourceAppId)
                        continue
                    }

                    $delegatedScopes += $scope.Value
                }
            }

            if ($delegatedScopes.Count -gt 0) {
                try {
                    $grantBody = @{
                        clientId    = $appServicePrincipal.Id
                        consentType = 'AllPrincipals'
                        resourceId  = $resourceServicePrincipal.Id
                        scope       = ($delegatedScopes -join ' ')
                    }

                    $null = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/oauth2PermissionGrants' -Body $grantBody -ErrorAction Stop
                    Write-Host ('Delegated permissions granted: {0}' -f ($delegatedScopes -join ', ')) -ForegroundColor Green
                }
                catch {
                    Write-Warning -Message ('Unable to grant delegated permissions {0}: {1}' -f ($delegatedScopes -join ', '), $_.Exception.Message)
                }
            }
        }
    }
    catch {
        Write-Warning -Message ('Failed to grant admin consent: {0}' -f $_.Exception.Message)
        Write-Warning -Message 'Grant admin consent manually using the Entra portal.'
    }
}

# Portal URL to grant admin consent
$URL = ('https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/{0}' -f $newApp.AppId)

if ($OpenBrowser) {

    # Wait for 30 seconds to allow Entra ID to provision the application
    Write-Host 'Browser will open in 30 seconds.'
    Start-Sleep -Seconds 30

    if ($PrivateBrowserSession) {

        # Browsers are tried in order, as the private session argument is browser-specific
        $privateBrowsers = @(
            @{ Command = 'msedge.exe'; Argument = '--inprivate' }
            @{ Command = 'chrome.exe'; Argument = '--incognito' }
            @{ Command = 'firefox.exe'; Argument = '-private-window' }
        )

        $browserStarted = $false

        foreach ($browser in $privateBrowsers) {
            try {
                Start-Process -FilePath $browser.Command -ArgumentList $browser.Argument, $URL -ErrorAction Stop
                $browserStarted = $true
                break
            }
            catch {
                Write-Verbose -Message ('Unable to start {0}: {1}' -f $browser.Command, $_.Exception.Message)
            }
        }

        if (-not $browserStarted) {
            Write-Warning -Message 'No supported browser found for a private session. Open the following URL manually:'
            Write-Warning -Message $URL
        }
    }
    else {
        Start-Process $URL
    }
}
elseif (-not $GrantAdminConsent) {
    Write-Host 'Open the following URL to grant admin consent:' -ForegroundColor Green
    Write-Host $URL -ForegroundColor Green
}