# Email DNS Health Report

`Get-EmailDnsHealthReport.ps1` checks the DNS records that support email delivery and authentication for a configurable list of domains. It produces a CSV report and a styled HTML report, and can optionally email the results.

## Checks

The default report checks:

- MX and SPF records.
- DKIM selectors, including global selectors and per-domain selector additions.
- DMARC, Autodiscover CNAME and SRV, MTA-STS, TLS-RPT, and BIMI.
- SOA serial numbers and authoritative name servers.

Optional checks can include DNSSEC, CAA records, and Microsoft 365 domain verification TXT records.

The report compares current values with the previous run's state snapshot. Changed values and newly added domains are identified in the CSV and HTML output. Domain checks run concurrently through a PowerShell runspace pool, and results are sorted alphabetically by domain.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7 on Windows.
- The built-in Windows `DnsClient` module, which provides `Resolve-DnsName`.
- Network access to DNS servers able to resolve the configured domains.
- A JSON configuration file containing at least one domain.

## Configuration

Start with [`Config/DnsReportConfig.sample.json`](Config/DnsReportConfig.sample.json) and create a working configuration file. The configuration supports:

- `Domains`: Domains to check.
- `DkimSelectors.Default`: Global DKIM selector names.
- `DkimSelectors.PerDomainSelectorAdds`: Additional DKIM selectors for specific domains.
- `BimiSelector`: BIMI selector; defaults to `default` when omitted.
- `Email`: SMTP and recipient settings used with `-SendEmail`.
- `StateFileName`: Name of the state snapshot stored in the output directory.

Keep environment-specific SMTP settings and domain names out of shared sample configurations.

## Usage

Run from this folder:

```powershell
.\Get-EmailDnsHealthReport.ps1
```

The current script default uses `.\Config\DEV-DnsReportConfig.json` and writes output to `.\Reports`. Use explicit paths for a reusable configuration:

```powershell
.\Get-EmailDnsHealthReport.ps1 `
    -ConfigPath .\Config\DnsReportConfig.sample.json `
    -OutputPath .\Reports `
    -Verbose
```

Enable optional checks and open the HTML report:

```powershell
.\Get-EmailDnsHealthReport.ps1 `
    -IncludeDnssec `
    -IncludeCaa `
    -IncludeM365Verification `
    -OpenReport
```

Send the report using the `Email` settings in the configuration:

```powershell
.\Get-EmailDnsHealthReport.ps1 -SendEmail
```

Use `-SkipStateUpdate` for a comparison run that does not replace the saved state, or increase `-ThrottleLimit` for larger domain lists.

## Parameters

| Parameter | Description |
| --- | --- |
| `-ConfigPath` | JSON configuration path. Default: `.\Config\DEV-DnsReportConfig.json`. |
| `-OutputPath` | Directory for CSV, HTML, and state files. Default: `.\Reports`. |
| `-ThrottleLimit` | Maximum number of domains checked concurrently. Default: `5`; valid range: `1` to `64`. |
| `-IncludeDnssec` | Include DNSSEC status. |
| `-IncludeCaa` | Include CAA records. |
| `-IncludeM365Verification` | Include the Microsoft 365 `MS=` verification record. |
| `-OpenReport` | Open the generated HTML report in the default browser. |
| `-SendEmail` | Send the HTML report and CSV attachment using configured SMTP settings. |
| `-SkipStateUpdate` | Do not overwrite the previous state snapshot. |

## Output

Each run writes timestamped files to the output directory:

- `DnsReport-yyyyMMdd-HHmmss.csv`
- `DnsReport-yyyyMMdd-HHmmss.html`
- The configured state snapshot, normally `DnsReportState.json`.

Missing DNS values are shown as `N/A` in the HTML report. Multiple MX records are displayed on separate lines. Changed values, new domains, and missing values use colored rounded labels in the HTML report.
