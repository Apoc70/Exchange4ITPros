# OWA Usage from IIS Logs

`Get-OwaUsageFromIisLogs.ps1` analyzes Exchange Front End W3C IIS logs and reports OWA requests by user and day. It reads the field order from each log file's `#Fields` header, so it supports logs with additional fields such as `X-Forwarded-For`.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7.
- Read access to the IIS log folders.
- IIS W3C log files, typically from `W3SVC1` (Default Web Site).

## Usage

Run the script with its defaults to analyze the last 30 days of logs in `C:\inetpub\logs\LogFiles\W3SVC1`:

```powershell
.\Get-OwaUsageFromIisLogs.ps1
```

Specify log folders from one or more Exchange servers and a custom date range:

```powershell
.\Get-OwaUsageFromIisLogs.ps1 `
    -LogPath '\\EX01\c$\inetpub\logs\LogFiles\W3SVC1', '\\EX02\c$\inetpub\logs\LogFiles\W3SVC1' `
    -StartDate (Get-Date).AddDays(-14) `
    -EndDate (Get-Date)
```

Group usernames such as `DOMAIN\user` and `user@domain` under the same name, and include failed HTTP requests:

```powershell
.\Get-OwaUsageFromIisLogs.ps1 -StripDomain -IncludeFailed
```

By default, the script counts `/owa/` requests and excludes HTTP status codes 400 and higher. To also count requests for another URI prefix, provide it with `-UriPrefix`:

```powershell
.\Get-OwaUsageFromIisLogs.ps1 -UriPrefix '/owa/', '/ecp/'
```

## Parameters

| Parameter | Description |
| --- | --- |
| `-LogPath` | One or more IIS log directories. Defaults to `$env:SystemDrive\inetpub\logs\LogFiles\W3SVC1`. |
| `-StartDate` | Start of the analysis period in local time. Defaults to midnight 30 days ago. |
| `-EndDate` | End of the analysis period in local time. Defaults to the current time. |
| `-UriPrefix` | URI prefixes to count. Defaults to `/owa/`. |
| `-ExcludeUserPattern` | Wildcard patterns for usernames to exclude. Defaults to common HealthMailbox, monitoring, and machine accounts. |
| `-ExcludeUserAgentPattern` | Wildcard patterns for user agents to exclude. Defaults to Exchange monitoring probes. |
| `-IncludeFailed` | Include requests with an HTTP status of 400 or higher. |
| `-StripDomain` | Remove the domain portion from `DOMAIN\user` and `user@domain` usernames before aggregation. |
| `-OutputFolder` | Destination for the CSV reports. Defaults to a timestamped `OWA-Report_yyyyMMdd_HHmm` folder in the current directory. |

## Output

The output folder contains two semicolon-delimited UTF-8 CSV files:

- `OWA_Users.csv` contains one row per user, with request counts, OWA service call counts, active days, first and last seen timestamps, client and server IPs, and the most frequent user agent.
- `OWA_DailyUniqueUsers.csv` contains one row per UTC date, with unique user and request counts.

The console summary includes processing counts and the 20 users with the most recent access. IIS timestamps are in UTC; the supplied start and end dates are interpreted in local time and converted to UTC for filtering.

`OwaServiceCalls` counts requests to `/owa/service.svc`. `RealOwaSession` is true when that count is greater than zero; users with other OWA requests but no service calls may have only generated a logon or redirect hit.
