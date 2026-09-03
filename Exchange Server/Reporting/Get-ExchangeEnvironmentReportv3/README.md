# Exchange Environment Report

`Get-ExchangeEnvironmentReport.ps1` creates a self-contained HTML overview of an on-premises Microsoft Exchange environment.

The report is intended for Exchange 2007 through Exchange Server 2019 and includes limited support for older Exchange versions. It summarizes servers, roles, versions, update levels, mailbox counts, sites, namespaces, database availability groups, mailbox databases, database sizes, free disk space, backups, and circular logging.

## Requirements

- Windows PowerShell 5.1 or newer. PowerShell 1 is not supported.
- The Exchange Management Shell or Exchange management tools installed on the computer running the script.
- An account with permission to query the Exchange organization and mailbox databases.
- WMI and Remote Registry access from the computer running the script to the Exchange servers. These are needed for operating-system, disk-space, database-size, update-rollup, and legacy cluster information.
- The following files in the same folder as the script:
  - `EnvironmentReport.css`
  - `ExchangeVersionMappings.json`

Run the script from an Exchange Management Shell session, or allow the script to load the local Exchange management components automatically.

## Usage

```powershell
Set-Location 'C:\path\to\Get-ExchangeEnvironmentReportv3'
.\Get-ExchangeEnvironmentReport.ps1
```

The report is written relative to the script folder using a filename such as `Exchange Environment Report_2026-09-03_14-30.html`. The script shows progress while it collects data. Use `-OpenInBrowser` to open the report automatically when collection finishes. Supply `-HTMLReport` when a custom filename is required.

## Examples

Create a report with a custom filename and open it in the default browser:

```powershell
.\Get-ExchangeEnvironmentReport.ps1 -HTMLReport '.\ExchangeEnvironment.html' -OpenInBrowser
```

Analyze only servers whose names match a wildcard:

```powershell
.\Get-ExchangeEnvironmentReport.ps1 -HTMLReport '.\Northwind.html' -ServerFilter 'EXCH-NL-*'
```

Include database and log drive names, runtime, disconnected mailboxes, and provisioning status:

```powershell
.\Get-ExchangeEnvironmentReport.ps1 `
	-HTMLReport '.\DetailedReport.html' `
	-ShowDriveNames `
	-ShowRunTime `
	-ShowDisconnectedMailboxCount `
	-ShowProvisioningStatus
```

Display average mailbox sizes in GB instead of MB:

```powershell
.\Get-ExchangeEnvironmentReport.ps1 `
	-HTMLReport '.\ExchangeEnvironment.html' `
	-ShowAverageMailboxSizeInGB
```

Send the completed report as an HTML email with the generated report attached:

```powershell
.\Get-ExchangeEnvironmentReport.ps1 `
	-HTMLReport '.\ExchangeEnvironment.html' `
	-SendMail `
	-MailFrom 'exchange-report@example.com' `
	-MailTo 'messaging-team@example.com' `
	-MailServer 'smtp.example.com'
```

Use custom report support files:

```powershell
.\Get-ExchangeEnvironmentReport.ps1 `
	-HTMLReport '.\ExchangeEnvironment.html' `
	-CssFileName 'CustomReport.css' `
	-VersionMappingFileName 'CustomVersionMappings.json'
```

## Parameters

| Parameter | Description |
| --- | --- |
| `-HTMLReport` | Optional. File name for the generated HTML report. Defaults to `Exchange Environment Report_yyyy-MM-dd_HH-mm.html`. |
| `-ServerFilter` | Wildcard filter for Exchange server names. Defaults to `*`. |
| `-ViewEntireForest` | Controls whether Exchange queries view the entire forest. Defaults to `$true`. |
| `-OpenInBrowser` | Opens the generated report in the default browser. |
| `-SendMail` | Sends the report by SMTP and requires `-MailFrom`, `-MailTo`, and `-MailServer`. |
| `-MailFrom` | Sender address used for the report email. |
| `-MailTo` | Recipient address used for the report email. |
| `-MailServer` | SMTP server used to send the report. |
| `-ShowDriveNames` | Adds EDB and log drive names to database tables. |
| `-ShowAverageMailboxSizeInGB` | Displays average mailbox and archive mailbox sizes in GB instead of MB. |
| `-ShowRunTime` | Adds script runtime information to the report. |
| `-ShowDisconnectedMailboxCount` | Adds disconnected mailbox counts per database. |
| `-ShowProvisioningStatus` | Adds mailbox database provisioning exclusion status. |
| `-CssFileName` | CSS file used to format the report. Defaults to `EnvironmentReport.css`. |
| `-VersionMappingFileName` | JSON file used to translate Exchange build numbers into readable versions and security updates. Defaults to `ExchangeVersionMappings.json`. |

## Output

The HTML report contains:

- Report generation time and organization name.
- Server totals by Exchange version and role.
- Mailbox totals by version, site, and organization.
- Internal, external, and Client Access namespace information when available.
- Exchange version, service pack, cumulative/security update, operating system, and role details.
- Database availability group membership and database copy information.
- Mailbox and archive mailbox counts and average sizes.
- Database size, whitespace, disk free space, backup time, and circular logging status when available.

The script continues when some optional information cannot be collected and writes warnings. A server or database may be skipped if its required data cannot be queried.

## Limitations

- Public folder infrastructure is not reported.
- Exchange 2007 and 2003 CCR/SCC cluster details are not fully examined; clustered mailbox servers are identified where possible.
- Full results depend on remote WMI and Remote Registry access.
- The script does not authenticate to SMTP. The configured SMTP server must accept the connection from the host running the script.
- Run the script with the latest available Exchange management tools when reporting on newer Exchange versions.

## Updating version mappings

Exchange build-to-version and security-update mappings are maintained in `ExchangeVersionMappings.json`. Update that file when new Microsoft cumulative updates or security updates are released. The script reads the mapping at runtime, so mapping changes do not require editing the PowerShell script.

### A Small Versioning Trick

Exchange Server 2019 and Exchange Server Subscription Edition both report build number `15.2`, because apparently even version numbers enjoy identity crises. The script uses a small build-based trick to distinguish them and map them to the correct product name. Don't be surprised to see `15.3` in the JSON file.

## License

This script is provided under the MIT License. See [LICENSE](../../LICENSE).
