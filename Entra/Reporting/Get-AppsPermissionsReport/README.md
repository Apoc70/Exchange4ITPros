# Get-AppsPermissionsReport

`Get-AppsPermissionsReport.ps1` generates a permissions inventory for Entra ID **app registrations**, **enterprise apps (service principals)**, or **both**, and exports the result to HTML and CSV.

## Version

- Current version: `2.1.1`
- The script stores the version in the `$ScriptVersion` variable.
- The generated HTML report footer includes the version and execution time.

The report includes:
- Application and delegated permissions for app registrations and/or enterprise apps
- Resource API/service principal names
- EWS-related permission detection
- Optional highlighting for high-privilege, Exchange, and SharePoint permissions
- Optional delivery by local files, email, and Teams webhook
- Optional "Object Type" column (visible only when reporting both sources)

## Requirements

- PowerShell 7+ (Windows PowerShell 5.1 also works in many environments)
- Microsoft Graph PowerShell SDK

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

The script uses Microsoft Graph cmdlets for:
- Applications
- Service principals
- App role assignments
- OAuth2 permission grants
- Organization details

## Authentication Modes

Use `-AuthMode` to select one of:

- `Interactive` (default)
- `AppCertificate`
- `AppSecret`

### Interactive (default)
Connects with delegated scopes:
- `Application.Read.All`
- `Directory.Read.All`
- `AppRoleAssignment.Read.All`
- `DelegatedPermissionGrant.Read.All`

### AppCertificate
Requires:
- `-TenantId`
- `-ClientId`
- `-CertificateThumbprint`

### AppSecret
Requires:
- `-TenantId`
- `-ClientId`
- `-ClientSecret` (`SecureString`)

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-ReportVariant` | `AppRegistrations` \| `EnterpriseApps` \| `Both` | `EnterpriseApps` | Select which objects to include in the report |
| `-IncludeFirstPartyApps` | Switch | off | Include Microsoft first-party enterprise apps. Default is third-party only |
| `-PredefinedSets` | `Exchange`, `SharePoint` | — | Filter by built-in permission groups |
| `-CustomPermissions` | String array | — | Additional permission names to include in filter |
| `-HighlightCategories` | `HighPrivilege`, `Exchange`, `SharePoint` | `HighPrivilege` | Controls HTML highlighting and badges |
| `-EwsFlagColor` | String | `#FFD700` | HTML color for app name cell when EWS permissions are detected |
| `-DeliveryOptions` | `FileSystem`, `Email`, `Teams` | `FileSystem` | Output destinations |
| `-OpenHtmlReport` | Switch | off | Opens the generated HTML report after completion |

Delivery-specific parameters:
- Email: `-EmailTo`, `-EmailFrom`, `-SmtpServer`, optional `-SmtpPort` (default `587`) and `-SmtpCredential`
- Teams: `-TeamsWebhookUrl`

## Output

Reports are saved under the local `Reports` folder next to the script.

Filename format:
- `<TenantName>_<ReportVariant>_<yyyyMMdd_HHmmss>.html`
- `<TenantName>_<ReportVariant>_<yyyyMMdd_HHmmss>.csv`

CSV columns:
- `AppName`
- `AppId`
- `AppSource` *(only present when `-ReportVariant Both` is used)*
- `ResourceName`
- `Permission`
- `PermissionType` (`Application` or `Delegated`)
- `IsHighPrivilege`
- `IsExchangePermission`
- `IsSharePointPermission`
- `HasEWS`

The HTML report includes:
- Report variant and enterprise scope in the title and metadata bar
- Grouped rows per app
- "Object Type" column visible only when `-ReportVariant Both`
- Search box filtering
- Legend and badges
- Summary stats (apps count, EWS app count, visible rows)


## NOTE

The Teams channel notification requires some code changes, because the direct webhook delivery has been deprecated by Microsoft. The connector format used in this script (MessageCard) is affected.

## Usage Examples

### 1) Enterprise apps only — third-party (default)

```powershell
.\Get-AppsPermissionsReport.ps1
```

### 2) Enterprise apps including Microsoft first-party apps

```powershell
.\Get-AppsPermissionsReport.ps1 -ReportVariant EnterpriseApps -IncludeFirstPartyApps
```

### 3) App registrations only

```powershell
.\Get-AppsPermissionsReport.ps1 -ReportVariant AppRegistrations
```

### 4) Both app registrations and enterprise apps (adds Object Type column)

```powershell
.\Get-AppsPermissionsReport.ps1 -ReportVariant Both
```

### 5) Filter to Exchange and SharePoint permissions

```powershell
.\Get-AppsPermissionsReport.ps1 -PredefinedSets Exchange,SharePoint
```

### 6) Filter by custom permissions

```powershell
.\Get-AppsPermissionsReport.ps1 -CustomPermissions "User.Export.All","Directory.Read.All"
```

### 7) Highlight all categories and open HTML automatically

```powershell
.\Get-AppsPermissionsReport.ps1 -HighlightCategories HighPrivilege,Exchange,SharePoint -OpenHtmlReport
```

### 8) Send by email

```powershell
.\Get-AppsPermissionsReport.ps1 `
    -ReportVariant Both `
    -PredefinedSets Exchange `
    -DeliveryOptions FileSystem,Email `
    -EmailTo admin@contoso.com `
    -EmailFrom noreply@contoso.com `
    -SmtpServer smtp.contoso.com
```

### 9) Use app-only auth with certificate

```powershell
.\Get-AppsPermissionsReport.ps1 `
    -AuthMode AppCertificate `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "00000000-0000-0000-0000-000000000000" `
    -CertificateThumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
```

### 10) Use app-only auth with client secret

```powershell
$secret = Read-Host "Client Secret" -AsSecureString
.\Get-AppsPermissionsReport.ps1 `
    -AuthMode AppSecret `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "00000000-0000-0000-0000-000000000000" `
    -ClientSecret $secret
```

## Notes

- Default `-ReportVariant` is `EnterpriseApps` with third-party apps only. Use `-IncludeFirstPartyApps` to add Microsoft-published apps.
- The "Object Type" column in the HTML/CSV only appears when `-ReportVariant Both` is selected.
- If `-PredefinedSets` and `-CustomPermissions` are both omitted, the script reports all discovered permissions.
- Enterprise app delegated permissions are read from OAuth2 permission grants; application permissions are read from app role assignments.
- If email or Teams settings are incomplete, the script skips that delivery path and continues.
- EWS-related permissions currently tracked: `full_access_as_app`, `full_access_as_user`, `EWS.AccessAsUser.All`, `Exchange.ManageAsApp`.
- The Teams channel notification (incoming webhook) may require updates as the connector format has been deprecated by Microsoft.
