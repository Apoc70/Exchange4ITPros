<#
.SYNOPSIS
    Retrieves mailbox statistics from Exchange Online and generates a comprehensive HTML report.

.DESCRIPTION
    This PowerShell script connects to Exchange Online and gathers detailed statistics for all mailboxes within the tenant. It collects information such as mailbox size, item count, last logon time, and other relevant attributes. The script then compiles the data into a structured HTML report, making it easy for administrators to review mailbox usage, identify inactive accounts, and monitor storage consumption. The report is suitable for auditing, capacity planning, and general Exchange Online management tasks. The script is designed to be run with appropriate permissions and may require the Exchange Online PowerShell module to be installed and imported prior to execution.


    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.NOTES 
    Requirements 

    - Exchange Online PowerShell module installed
    - Entra app, with appropriate API permissions granted when using app authentication

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 

.LINK
    https://exchangeforitpros.blog/

.PARAMETER UseAppAuthentication
    Use App Authentication to connect to Exchange Online. If not specified, the script assumes you run the script in an active Exchange Online Management Shell

.PARAMETER MailboxType
    Specify the type of mailbox to report on. Valid values are 'User', 'Shared', and 'Room'. Default is 'All'.

.PARAMETER ReportType
    Specify the type of report to generate. Valid values are 'All', 'Top10', and 'Below10Percent'. Default is 'All'.

.PARAMETER MailboxCount
    Specify the number of mailboxes to process when MailboxType is set to 'Test'. Default is 10.    

.PARAMETER ConfigFile
    Specify the configuration file to use when using App Authentication. Default is 'example.json'. This file should contain tenant ID, organization, client ID, and certificate thumbprint.

.EXAMPLE

Get-ExoMailboxReport.ps1 -UseAppAuthentication -MailboxType 'All' -ReportType 'Top10' -ConfigFile 'varunagroup.json'

This command retrieves mailbox statistics for all mailboxes and generates a report showing the top 10 mailboxes by size, using app authentication.

#>

# Parameter section with examples
# Additional information parameters: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters
[CmdletBinding()]
param(
    [switch]$UseAppAuthentication,
    [ValidateSet('All', 'Test', 'User', 'Shared', 'Room')]
    [string]$MailboxType,
    [ValidateSet('All', 'Top10', 'Below10Percent')]
    [string]$ReportType,
    [int]$MailboxCount = 10,
    [string]$ConfigFile = 'example.json'
)

#region Initialize Script 

# Measure script running time
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

$script:ScriptPath = Split-Path $script:MyInvocation.MyCommand.Path
$script:ScriptName = $MyInvocation.MyCommand.Name

#endregion

#region Functions

function LoadScriptSettings {

    # Load configuration settings from JSON file
    $configFilePath = Join-Path -Path $script:ScriptPath -ChildPath $ConfigFile

    if (Test-Path -Path $configFilePath) {
        Write-Host ('Loading configuration from {0}' -f $configFilePath)    
        try {
            $config = Get-Content -Path $configFilePath | ConvertFrom-Json
            # Extract configuration values
            $tenantId = $config.tenantid
            $organization = $config.organization    
            $clientid = $config.clientId
            $certThumbprint = $config.certthumbprint
        }
        catch {
            Write-Error ('Failed to load configuration from {0}: {1}' -f $configFilePath, $_.Exception.Message)
            exit
        }
        
        Write-Verbose -Message 'Script settings loaded'
    }
    else { 
        Write-Error -Message 'Script settings file settings.xml missing. Please check documentation.'
        exit 99
    }
}

function Request-Choice {
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = "Provide a caption for the Y/N question.")]
        [string]$Caption
    )
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@('&Yes', '&No')
    [int]$defaultChoice = 1

    $choiceReturn = $Host.UI.PromptForChoice($Caption, '', $choices, $defaultChoice)

    return $choiceReturn   
}

function Process-Mailboxes {

    Write-Host ('Processing {0} mailbox(es)' -f ($script:Mailboxes | Measure-Object).count)

    $i = 1
    $totalCount = ($script:Mailboxes | Measure-Object).Count

    # Process each mailbox
    foreach ($mailbox in $script:Mailboxes) {

        Write-Progress -Activity ('Processing {0} mailboxes' -f $totalCount) -Status ('Processing {0} ({1}/{2})' -f $mailbox.PrimarySmtpAddress, $i, $totalCount) -PercentComplete (($i / $totalCount) * 100)

        if ($mailbox.DisplayName -like 'Discovery*') {
            Write-Verbose -Message ('Skipping mailbox {0} as it does not have a valid DisplayName' -f $mailbox.PrimarySmtpAddress)
            continue
        }

        $mailboxStat = $mailbox | Get-EXOMailboxStatistics
        # $mailboxArchiveStat = $mailbox | Get-EXOMailboxStatistics -Archive -ErrorAction SilentlyContinue
        # for an upcoming release

        $maxQuotaInMB = [int]([regex]::Match($mailbox.ProhibitSendReceiveQuota, '^([\d\.,]+)\s*GB').Groups[1].Value -replace ',', '.') * 1024

        $mailItemSizeInPercent = ( [math]::Round( ($mailboxStat.TotalItemSize.Value.ToMB() / $maxQuotaInMB ) * 100, 2 ) ) 

        # Create a custom object with DisplayName and MailboxSite
        $mailboxObject = [PSCustomObject]@{
            DisplayName              = $mailbox.DisplayName
            UserPrincipalName        = $mailbox.UserPrincipalName
            PrimarySmtpAddress       = $mailbox.PrimarySmtpAddress
            RecipientType            = $mailbox.RecipientType
            RecipientTypeDetails     = $mailbox.RecipientTypeDetails
            ProhibitSendReceiveQuota = ('{0} GB' -f [regex]::Match($mailbox.ProhibitSendReceiveQuota, '^([\d\.,]+)\s*GB').Groups[1].Value)
            MailboxItemSizeInMB      = $mailboxStat.TotalItemSize.Value.ToMB()
            MailboxItemSizeInGB      = $mailboxStat.TotalItemSize.Value.ToGB()
            MailboxItemSizeInPercent = $mailItemSizeInPercent
            FreeSizeInPercent        = 100 - $mailItemSizeInPercent
            # ArchiveItemSizeInGB      = if ($mailboxArchiveStat) { $mailboxArchiveStat.TotalItemSize.Value.ToGB() } else { 0 }
        }

        # Initialize the array if it doesn't exist
        if (-not $script:ProcessedMailboxes) {
            $script:ProcessedMailboxes = @()
        }

        # Add the object to the array
        $script:ProcessedMailboxes += $mailboxObject
        
        $i++
    }
}

function Create-HtmlReport {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        [string]$ReportTitle = 'Mailbox Report',
        [string]$SubReportTitle = '',
        [int]$MaxMailboxCount = 10
    )

    if ($MaxMailboxCount -gt 0) {

        $sortedMailboxes = $script:ProcessedMailboxes | Sort-Object -Property MailboxItemSizeInMB -Descending | Select-Object -First $MaxMailboxCount
    }
    else {
        $sortedMailboxes = $script:ProcessedMailboxes | Sort-Object -Property MailboxItemSizeInMB -Descending 
    }

    Write-Verbose -Message 'Creating HTML report...'

    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>$($ReportTitle)</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }
        table {
            font-family: Arial, sans-serif;
            border-collapse: collapse;
            # width: 100%;
        }
        th, td {
            border: 1px solid black;
            padding: 8px;
            text-align: left;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>
    <h1>$($ReportTitle)</h1><h4>$($SubReportTitle)</h4>
    <p><small>$(Get-Date -Format 'yyyy-MM-dd HH:mm')</small></p>
    <table>
        <tr>
            <th>#</th>
            <th>Display Name</th>
            <th>User Principal Name</th>
            <th>Primary SMTP Address</th>
            <th>Recipient Type Details</th>
            <th style="text-align: right;">Mailbox Size (MB)</th>
            <th style="text-align: right;">Mailbox Size (GB)</th>
            <th>Percent</th>
            <th>Prohibit Send/Receive Quota</th>
        </tr>
"@

    $index = 1
    foreach ($mailbox in $sortedMailboxes) {
        $htmlContent += @"
    <tr>
        <td>$index</td>
        <td>$($mailbox.DisplayName)</td>
        <td>$($mailbox.UserPrincipalName)</td>
        <td>$($mailbox.PrimarySmtpAddress)</td>
        <td>$($mailbox.RecipientTypeDetails)</td>
        <td style="text-align: right;">$($mailbox.MailboxItemSizeInMB)</td>
        <td style="text-align: right;">$($mailbox.MailboxItemSizeInGB)</td>
        <td style="text-align: right;">$('{0:N2}' -f $mailbox.MailboxItemSizeInPercent)</td>
        <td>$($mailbox.ProhibitSendReceiveQuota)</td>
    </tr>
"@
        $index++
    }

    # Close the HTML table
    $htmlContent += @"
    </table>
    <p>Total Mailboxes Processed: $($sortedMailboxes.Count)</p>
</body>
</html>
"@

    Set-Content -Path $OutputFile -Value $htmlContent -Force
    Write-Verbose -Message ('HTML report created at {0}' -f $OutputFile)
}

#endregion

#region MAIN

# 1. Load script settings
LoadScriptSettings

if ($UseAppAuthentication) {
    
    if (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue) {
        Write-Host "ExO already loaded!"
    }
    else {
        Write-Host "Loading ExO module..."

        Write-Verbose -Message ('Connecting to Exchange Online with AppId {0} and CertificateThumbprint {1}' -f $clientid, $certThumbprint)

        Connect-ExchangeOnline -AppId $clientid -CertificateThumbprint $certThumbprint -Organization $organization -ErrorAction Stop -ShowBanner:$false -Verbose:$false
        
        Write-Host "ExO module loaded successfully!"
    }
}

$mailboxes = $null


switch ($MailboxType) {
    'All' {
        $script:Mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties prohibitsendreceivequota
        Process-Mailboxes
    }
    'User' {
        $script:Mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox -Properties prohibitsendreceivequota
        Process-Mailboxes
    }
    'Shared' {
        $script:Mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -Properties prohibitsendreceivequota   
        Process-Mailboxes
    }
    'Room' {
        $script:Mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails RoomMailbox -Properties prohibitsendreceivequota
        Process-Mailboxes
    }
    'Test' {
        $script:Mailboxes = Get-EXOMailbox -ResultSize $MailboxCount -Properties prohibitsendreceivequota
        Process-Mailboxes
    }
}

switch ($ReportType) {
    'All' {
        Write-Verbose -Message 'Generating report for all mailboxes'
        $script:ProcessedMailboxes = $script:ProcessedMailboxes | Sort-Object -Property MailboxItemSizeInMB -Descending
        $subTitle = ('All Mailboxes (Filter: {0})' -f $MailboxType)
    }
    'Top10' {
        Write-Verbose -Message 'Generating report for top 10 mailboxes'
        $script:ProcessedMailboxes = $script:ProcessedMailboxes | Sort-Object -Property MailboxItemSizeInMB -Descending | Select-Object -First 10
        $subTitle = ('Top 10 Mailboxes by Size (Filter: {0})' -f $MailboxType) 
    }
    'Below10Percent' {
        Write-Verbose -Message 'Generating report for mailboxes below 10 percent free space'
        $script:ProcessedMailboxes = $script:ProcessedMailboxes | ? { $_.FreeSizeInPercent -le 10 } | Sort-Object -Property MailboxItemSizeInMB -Descending 
        $subTitle = ('All Mailboxes below 10% free space (Filter: {0})' -f $MailboxType)
    }
}

# 2. Create HTML report
Write-Verbose -Message 'Creating HTML report...'

Create-HtmlReport -OutputFile ('{0}\MailboxReport_{1:yyyyMMdd_HHmmss}_{3}_{2}.html' -f $script:ScriptPath, (Get-Date), $ReportType, $MailboxType ) -SubReportTitle $subTitle

#endregion

#region End Script

# Stop watch
$StopWatch.Stop()

# Write script runtime
Write-Verbose -Message ('It took {0:00}:{1:00}:{2:00} to run the script.' -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds)

return 0

#endregion