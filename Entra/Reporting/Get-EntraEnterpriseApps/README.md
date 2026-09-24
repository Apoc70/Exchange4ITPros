# Get-EntraEnterpriseApps

`Get-EntraEnterpriseApps.ps1` generates HTML reports for Entra ID enterprise apps and app registrations, and highlights items that were newly added since the previous run.

## Overview

The script:
- Queries enterprise apps (service principals) and app registrations (applications)
- Exports HTML reports for both object types
- Saves the current data as JSON for change tracking
- Highlights newly added entries in the HTML report
- Optionally exports CSV output
- Supports filtering for custom or all enterprise apps

## Requirements

- PowerShell 7+ or Windows PowerShell 5.1+
- Microsoft.Graph PowerShell module

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

## Authentication

The script supports direct tenant-based authentication or certificate-based authentication through a JSON config file.

### Manual parameter mode

```powershell
./Get-EntraEnterpriseApps.ps1 -TenantId "<tenant-id>" -EnterpriseAppFilter NonMicrosoftOnly
```

### Config-file mode

The script looks for a configuration file in the same folder as the script (for example `dev-egxde.json`). It reads:
- `tenantid`
- `ClientId`
- `CertThumbprint`

```powershell
./Get-EntraEnterpriseApps.ps1 -ConfigFile "dev-egxde.json"
```

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-TenantId` | String | — | Entra tenant ID to query |
| `-EnterpriseAppFilter` | `All` \| `NonMicrosoftOnly` | `NonMicrosoftOnly` | Filter the enterprise apps set |
| `-Highlight3rdPartyOwners` | Switch | Off | Highlights enterprise apps owned by third-party organizations |
| `-ExportCSV` | Switch | Off | Exports CSV files in addition to HTML reports |
| `-ConfigFile` | String | `dev-egxde.json` | JSON config file for certificate-based authentication |
| `-DailyReport` | Switch | Off | Writes daily fixed-name HTML exports instead of timestamped names |
| `-ExpireInDays` | Int | `30` | Day threshold used for expiry warnings on certificates and secrets |

## Output

Reports are saved in a `Reports` subfolder next to the script.

### HTML output

The script creates two HTML files:
- `<TenantName>-EnterpriseApps-<timestamp>.html`
- `<TenantName>-AppRegistrations-<timestamp>.html`

When `-DailyReport` is used, the filenames use a fixed date-based naming pattern instead of a timestamp.

### JSON tracking files

The script stores previous reports as JSON in the same `Reports` folder so it can detect newly added apps between runs.

### CSV output

When `-ExportCSV` is used, the script also generates:
- `<TenantName>-EnterpriseApps-<timestamp>.csv`
- `<TenantName>-AppRegistrations-<timestamp>.csv`

## Usage Examples

### 1) Default behavior

```powershell
./Get-EntraEnterpriseApps.ps1 -TenantId "contoso.onmicrosoft.com"
```

### 2) Include all enterprise apps

```powershell
./Get-EntraEnterpriseApps.ps1 -TenantId "contoso.onmicrosoft.com" -EnterpriseAppFilter All
```

### 3) Highlight third-party app owners

```powershell
./Get-EntraEnterpriseApps.ps1 -TenantId "contoso.onmicrosoft.com" -Highlight3rdPartyOwners
```

### 4) Export CSV files as well

```powershell
./Get-EntraEnterpriseApps.ps1 -TenantId "contoso.onmicrosoft.com" -ExportCSV
```

### 5) Use the JSON config file for certificate auth

```powershell
./Get-EntraEnterpriseApps.ps1 -ConfigFile "customconfig.json"
```

### 6) Generate fixed daily report names

```powershell
./Get-EntraEnterpriseApps.ps1 -TenantId "contoso.onmicrosoft.com" -DailyReport
```

## Notes

- The default filter `NonMicrosoftOnly` limits enterprise apps to non-Microsoft/custom entries.
- Newly added apps are highlighted using a red background style in the HTML output.
- Expired certificate and secret counts are displayed for app registrations when the age threshold is reached.
- The script depends on Microsoft Graph connectivity and the correct permissions for reading service principals and applications.
