# Get-PnPSiteOwners

Retrieves all sites in a SharePoint Online tenant together with their owners and generates readable HTML (and optionally CSV) reports.

## What it does

The script connects to a SharePoint Online tenant using PnP PowerShell and certificate based app-only authentication. It enumerates all site collections (optionally including OneDrive sites), determines the owners for each site, either site collection administrators for classic sites or group owners for group connected sites, and resolves each owner against Entra ID to include display name, email, and account status.

Four outputs are generated in the report directory:

- **Sites Report**: one row per site, with the site URL, title, whether it is a group site, and the combined owner email(s)/name(s).
- **Owners Report**: one row per site/owner combination, including owner type and whether the account is enabled.
- **Sites Error Report**: sites that could not be processed, with the error message.
- **Unresolved UPN Report**: a plain text list of UserPrincipalNames that could not be resolved against Entra ID. Always written when unresolved entries exist, regardless of `-ExportCsv`.

Report file names include a readable tenant name and a timestamp, for example `SitesReport_Contoso_2026-07-08_14-30-00.html`.

## Prerequisites

- PowerShell 7 or Windows PowerShell 5.1
- `PnP.PowerShell` module installed
- An Entra ID app registration configured for certificate based authentication (`-ClientId` and `-Thumbprint`), with the necessary Microsoft Graph and SharePoint permissions to read sites, site collection admins, group owners, and Entra ID user details
- A configuration file, see below

## Configuration file

By default, the script looks for a configuration file named `Get-PnPSiteOwners.config` in the same directory as the script. A different file can be supplied through `-ConfigFile`.

```json
{
    "tenantid": "00000000-0000-0000-0000-000000000000",
    "AdminUrl": "contoso-admin.sharepoint.com",
    "ClientId": "00000000-0000-0000-0000-000000000000",
    "CertThumbprint": "0000000000000000000000000000000000000000",
    "TenantDisplayName": "Contoso"
}
```

| Property | Required | Description |
| --- | --- | --- |
| `tenantid` | Yes | Entra ID tenant ID |
| `AdminUrl` | Yes | SharePoint Online admin center URL |
| `ClientId` | Yes | Application (client) ID of the Entra ID app registration |
| `CertThumbprint` | Yes | Thumbprint of the certificate used for authentication |
| `TenantDisplayName` | No | Readable name used in report file names. If omitted, a short name is derived from `AdminUrl` |

## Parameters

| Parameter | Description |
| --- | --- |
| `-IncludeOneDriveSites` | Includes OneDrive sites in the report. Excluded by default. |
| `-ConfigFile` | Path to the configuration file. Defaults to `Get-PnPSiteOwners.config`. |
| `-ReportDirectory` | Directory where reports are saved. Defaults to `Reports`. |
| `-ExportCsv` | Also exports the Sites, Owners, and Sites Error reports as CSV. HTML is always generated. |
| `-ClearReports` | Clears the existing report directory before generating new reports. Supports `-WhatIf`/`-Confirm`. |

## Examples

```powershell
# Default run, excluding OneDrive sites, HTML reports only
.\Get-PnPSiteOwners.ps1

# Include OneDrive sites
.\Get-PnPSiteOwners.ps1 -IncludeOneDriveSites

# Use a specific configuration file, export CSV in addition to HTML, and clear previous reports first
.\Get-PnPSiteOwners.ps1 -ConfigFile 'contoso.config' -ExportCsv -ClearReports
```

## Notes

- Entra ID user lookups are cached during the script run, so owners that appear on multiple sites are only resolved once.
- Redirect sites (`RedirectSite#1` template) are automatically excluded from processing.

## Links

- Exchange for IT Pros website: https://exchangeforitpros.blog/