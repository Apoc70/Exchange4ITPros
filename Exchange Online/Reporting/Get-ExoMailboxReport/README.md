# Get-ExoMailboxReport.ps1

This PowerShell script connects to Exchange Online and gathers detailed statistics for all mailboxes within the tenant. It collects information such as mailbox size, item count, last logon time, and other relevant attributes. The script then compiles the data into a structured HTML report, making it easy for administrators to review mailbox usage, identify inactive accounts, and monitor storage consumption. The report is suitable for auditing, capacity planning, and general Exchange Online management tasks. The script is designed to be run with appropriate permissions and may require the Exchange Online PowerShell module to be installed and imported prior to execution.

## Requirements 

- Exchange Online PowerShell module installed
- Entra app, with appropriate API permissions granted when using app authentication

## Parameters

### UseAppAuthentication

Use App Authentication to connect to Exchange Online. If not specified, the script assumes you run the script in an active Exchange Online Management Shell

### MailboxType

Specify the type of mailbox to report on. Valid values are 'User', 'Shared', and 'Room'. Default is 'All'.

### ReportType

Specify the type of report to generate. Valid values are 'All', 'Top10', and 'Below10Percent'. Default is 'All'.

### MailboxCount

Specify the number of mailboxes to process when MailboxType is set to 'Test'. Default is 10.    

### ConfigFile

Specify the configuration file to use when using App Authentication. Default is 'example.json'. This file should contain tenant ID, organization, client ID, and certificate thumbprint.

## EXAMPLE


```powershell
Get-ExoMailboxReport.ps1 -UseAppAuthentication -MailboxType 'All' -ReportType 'Top10' -ConfigFile 'varunagroup.json'
```

This command retrieves mailbox statistics for all mailboxes and generates a report showing the top 10 mailboxes by size, using app authentication.
