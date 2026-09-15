<#
    .SYNOPSIS
    Creates an HTML report describing the On-Premises Exchange environment.

    Based on the original 1.6.2 version by Steve Goodman

    Version 3.0.1, 2026-09-15

    .DESCRIPTION

    This script creates an HTML report showing the following information about an Exchange
    Subscription Edition, 2019, 2016, 2013, 2010, and, to a lesser extent, 2007 and 2003 environment.

    Requirements
    * Exchange Server Management Shell 2010 or newer
    * WMI and Remote Registry access from the computer running the script to all internal Exchange Servers
    * CSS file for HTML formatting
    * JSON file containing the Exchange version, service pack, and security update mapping

    The reports shows the following:

    * Report Generation Time
    * Total Servers per Exchange Version (2003 > 2010 or 2007 > 2019)
    * Total Mailboxes per Exchange Version, Office 365, and Organisation
    * Total Roles in the environment

    Then, per site:
    * Total Mailboxes per site
    * Internal, External and CAS Array Hostnames
    * Exchange Servers with:
      o Exchange Server Version
      o Service Pack and Security Update Level
      o Number of preferred and maximum active databases
      o Update Rollup and rollup version
      o Roles installed on server and mailbox counts
      o OS Version and Service Pack

    Then, per Database availability group (Exchange 2010/2013/2016/2019):
    * Total members per DAG
    * Member list
    * Databases, detailing:
      o Mailbox Count and Average Size
      o Archive Mailbox Count and Average Size (Only shown if DAG includes Archive Mailboxes)
      o Database Size and whitespace
      o Database and log disk free
      o Last Full Backup (Only shown if one or more DAG database has been backed up)
      o Circular Logging Enabled (Only shown if one or more DAG database has Circular Logging enabled)
      o Mailbox server hosting active copy
      o List of mailbox servers hosting copies and number of copies

    Finally, per Database (Non DAG DBs/Exchange 2007/Exchange 2003)
    * Databases, detailing:
      o Storage Group (if applicable) and DB name
      o Server hosting database
      o Mailbox Count and Average Size
      o Archive Mailbox Count and Average Size (Only shown if DAG includes Archive Mailboxes)
      o Database Size and whitespace
      o Database and log disk free
      o Last Full Backup (Only shown if one or more DAG database has been backed up)
      o Circular Logging Enabled (Only shown if one or more DAG database has Circular Logging enabled)

    This does not detail public folder infrastructure, or examine Exchange 2007/2003 CCR/SCC clusters
    (although it attempts to detect Clustered Exchange 2007/2003 servers, signified by ClusMBX).

    IMPORTANT NOTE: The script requires WMI and Remote Registry access to Exchange servers from the server
    it is run from to determine OS version, Update Rollup, Exchange 2007/2003 cluster and DB size information.

    .NOTES

    Revision History
    --------------------------------------------------------------------------------
    3.0.1   Added Markdown report export
    3.0     Major update with new features and improvements

    .LINK
    https://github.com/Apoc70/Exchange4ITPros

    .PARAMETER HTMLReport
    Filename to write HTML Report to. If omitted, a filename in the format
    'Exchange Environment Report_yyyy-MM-dd_HH-mm.html' is generated.

    .PARAMETER MarkdownReport
    Filename to write the Markdown report to. Only used when -OutputFormat is 'Markdown' or 'Both'.
    If omitted, a filename in the format 'Exchange Environment Report_yyyy-MM-dd_HH-mm.md' is generated.

    .PARAMETER OutputFormat
    The report output format(s) to generate. Valid values are 'HTML', 'Markdown', or 'Both'. Default: HTML
    The Markdown report is derived from the same report content as the HTML report, rendered as GitHub
    flavored Markdown tables instead of styled HTML tables.

    .PARAMETER SendMail
    Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory
    When -OutputFormat is 'Markdown', the Markdown file is attached and the mail body is a generic plain text
    notice. Otherwise the HTML report is used as the mail body, and, if -OutputFormat is 'Both', the Markdown
    file is attached in addition to the HTML file.

    .PARAMETER MailFrom
    Email address to send from. Passed directly to Send-MailMessage as -From

    .PARAMETER MailTo
    Email address to send to. Passed directly to Send-MailMessage as -To

    .PARAMETER MailServer
    SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer

    .PARAMETER ViewEntireForest
    By default, true. Set the option in Exchange 2007 or 2010 to view all Exchange servers and recipients in the forest.

    .PARAMETER ServerFilter
    Use a text based string to filter Exchange Servers by, e.g., NL-*
    Note the use of the wildcard (*) character to allow for multiple matches.

    .PARAMETER ShowDriveNames
    Include drive names of EDB file path and LOG file folder in database report table

    .PARAMETER ShowAverageMailboxSizeInGB
    List average mailbox and archive mailbox size in GB instead of MB

    .PARAMETER ShowRunTime
    Show script running time in HTML report

    .PARAMETER ShowDisconnectedMailboxCount
    Show the number of disconnected (soft-deleted) mailboxes per database

    .PARAMETER ShowProvisioningStatus
    Show  IsExludedFromProvisioning or IsExcludedFromProvisioningByOperator status in the report

    .PARAMETER OpenInBrowser
    Open the generated HTML report in the default browser after it has been written.

    .PARAMETER CssFileName
    The filename containing the Cascading Style Sheet (CSS) information for the HTML report
    Default: EnvironmentReport.css

    .PARAMETER VersionMappingFileName
    The filename containing the JSON based Exchange version, service pack, and security update mapping used
    to translate raw build numbers into human readable labels in the HTML report.
    Keeping this mapping in a separate file simplifies updates when Microsoft releases new Cumulative Updates
    or Security Updates, no script code changes are required.
    Default: ExchangeVersionMappings.json

    .EXAMPLE
    Generate the HTML report using the default timestamped filename
    .\Get-ExchangeEnvironmentReport.ps1

    .EXAMPLE
    Generate the HTML report with a custom filename
    .\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html

    .EXAMPLE
    Generate the HTML report and display average mailbox sizes in GB instead of MB
    .\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html -ShowAverageMailboxSizeInGB

    .EXAMPLE
    Generate the HTML report using a custom CSS file
    .\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html -CssFileName MyCustomCSSFile.css

    .EXAMPLE
    Generate am HTML report and send the result as HTML email with attachment to the specified recipient using a dedicated smart host
    .\Get-ExchangeEnvironmentReport.ps1 -HTMReport ExchangeEnvironment.html -SendMail -ViewEntireForet $true -MailFrom roaster@mcsmemail.de -MailTo grillmaster@mcsmemail.de -MailServer relay.mcsmemail.de

    .EXAMPLE
    Generate the HTML report using a custom Exchange version mapping file
    .\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html -VersionMappingFileName MyVersionMappings.json

    .EXAMPLE
    Generate the HTML report including EDB and LOG drive names
    .\Get-ExchangeEnvironmentReport.ps1 -ShowDriveNames -HTMLReport .\report.html

    .EXAMPLE
    Generate the HTML report and open it in the default browser when finished
    .\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html -OpenInBrowser

    .EXAMPLE
    Generate only a Markdown report
    .\Get-ExchangeEnvironmentReport.ps1 -OutputFormat Markdown -MarkdownReport .\report.md

    .EXAMPLE
    Generate both an HTML and a Markdown report and mail the HTML version with the Markdown file attached
    .\Get-ExchangeEnvironmentReport.ps1 -OutputFormat Both -SendMail -MailFrom roaster@mcsmemail.de -MailTo grillmaster@mcsmemail.de -MailServer relay.mcsmemail.de
#>
[CmdletBinding()]
param(
  [parameter(Position = 0, HelpMessage = 'Filename to write HTML report to')]
  [string]$HTMLReport = ('Exchange Environment Report_{0}.html' -f (Get-Date -Format 'yyyy-MM-dd_HH-mm')),
  [string]$MarkdownReport = ('Exchange Environment Report_{0}.md' -f (Get-Date -Format 'yyyy-MM-dd_HH-mm')),
  [ValidateSet('HTML', 'Markdown', 'Both')]
  [string]$OutputFormat = 'HTML',
  [switch]$SendMail,
  [switch]$OpenInBrowser,
  [string]$MailFrom = '',
  [string]$MailTo = '',
  [string]$MailServer = '',
  [bool]$ViewEntireForest = $true,
  [string]$ServerFilter = '*',
  [switch]$ShowDriveNames,
  [switch]$ShowAverageMailboxSizeInGB,
  [switch]$ShowRunTime,
  [switch]$ShowDisconnectedMailboxCount,
  [switch]$ShowProvisioningStatus,
  [string]$CssFileName = 'EnvironmentReport.css',
  [string]$VersionMappingFileName = 'ExchangeVersionMappings.json'
)

# Version
$ScriptVersion = '3.0.1'

# Start Stop Watch
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

# Warning Limits, adjust as needed
$MinFreeDiskspace = 30 # Mark free space less than this value (%) in red
$MaxDatabaseSize = 250 # Mark database larger than this value (GB) in red

# Default variables
$NotAvailable = 'N/A'
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$EdgeServerNote = 'If there are any Edge Transport Servers subscribed to the Exchange organization, the Exchange version information displayed in the report does not reflect the version currently installed.<br/>The version displayed is the version as of the subscription date.'
$ReportTitle = 'Exchange Environment Report'

# Set TLS version o TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

<#
    .SYNOPSIS
    Recursively converts a PSCustomObject (as returned by ConvertFrom-Json) into a hashtable.

    .DESCRIPTION
    Implemented manually because ConvertFrom-Json -AsHashtable is only available on
    PowerShell 6.0 and newer, and this script needs to remain compatible with Windows
    PowerShell 5.1. Walks nested objects and arrays recursively, arrays of objects are
    converted element by element.

    .PARAMETER InputObject
    The object to convert, typically the output of ConvertFrom-Json. May be $null, a
    PSCustomObject, an array/collection, or a scalar value.

    .EXAMPLE
    $Config = Get-Content -Path .\config.json -Raw | ConvertFrom-Json | ConvertTo-Hashtable

    .NOTES
    Author: Thomas Stensitzki
#>
function ConvertTo-Hashtable {
  [CmdletBinding()]
  [OutputType([hashtable])]
  param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [AllowNull()]
    $InputObject
  )

  process {
    if ($null -eq $InputObject) {
      return $null
    }

    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
      $Hashtable = @{}
      foreach ($Property in $InputObject.PSObject.Properties) {
        $Hashtable[$Property.Name] = ConvertTo-Hashtable -InputObject $Property.Value
      }
      return $Hashtable
    }
    elseif (($InputObject -is [System.Collections.IEnumerable]) -and ($InputObject -isnot [string])) {
      $Collection = foreach ($Item in $InputObject) { ConvertTo-Hashtable -InputObject $Item }
      return , $Collection
    }
    else {
      return $InputObject
    }
  }
}

<#
    .SYNOPSIS
    Strips HTML markup from a fragment and decodes HTML entities, turning it into plain text.

    .DESCRIPTION
    Used while converting the generated HTML report into Markdown. Converts <br> tags into
    newlines, strips any remaining tags, and decodes HTML entities (e.g. &nbsp;, &amp;, &uarr;)
    using System.Net.WebUtility.

    .PARAMETER Fragment
    The raw HTML fragment (inner content of a heading, paragraph, or table cell) to convert.

    .EXAMPLE
    ConvertFrom-HtmlFragment -Fragment '5.5&nbsp;GB<br/>Free'

    .NOTES
    Author: Thomas Stensitzki
#>
function ConvertFrom-HtmlFragment {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $false)]
    [AllowEmptyString()]
    [string]$Fragment = ''
  )

  $Text = $Fragment -replace '(?i)<br\s*/?>', "`n"
  $Text = $Text -replace '(?is)<[^>]+>', ''
  $Text = [System.Net.WebUtility]::HtmlDecode($Text)

  $Text.Trim()
}

<#
    .SYNOPSIS
    Converts the inner HTML of a single <table> element into a GitHub flavored Markdown table.

    .DESCRIPTION
    Parses <tr> rows and <th>/<td> cells out of the supplied HTML, honoring colspan by padding
    additional empty cells so column counts line up. The first row found is rendered as the
    Markdown table header; this collapses any secondary HTML header rows into regular data rows,
    since Markdown tables only support a single header row.

    .PARAMETER TableInnerHtml
    The HTML located between an opening and closing <table> tag.

    .EXAMPLE
    ConvertTo-MarkdownTable -TableInnerHtml $Match.Groups[1].Value

    .NOTES
    Author: Thomas Stensitzki
#>
function ConvertTo-MarkdownTable {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $false)]
    [AllowEmptyString()]
    [string]$TableInnerHtml = ''
  )

  $RowMatches = [regex]::Matches($TableInnerHtml, '(?is)<tr[^>]*>(.*?)</tr>')
  $Rows = @()

  foreach ($RowMatch in $RowMatches) {
    $CellMatches = [regex]::Matches($RowMatch.Groups[1].Value, '(?is)<(th|td)([^>]*)>(.*?)</\1>')
    $Cells = @()

    foreach ($CellMatch in $CellMatches) {
      $Colspan = 1
      if ($CellMatch.Groups[2].Value -match '(?i)colspan\s*=\s*"?''?(\d+)') {
        $Colspan = [int]$Matches[1]
      }

      $CellText = ConvertFrom-HtmlFragment -Fragment $CellMatch.Groups[3].Value
      $CellText = ($CellText -replace '\r?\n', '<br>') -replace '\|', '\|'

      $Cells += $CellText

      for ($i = 1; $i -lt $Colspan; $i++) {
        $Cells += ''
      }
    }

    if ($Cells.Count -gt 0) {
      $Rows += , $Cells
    }
  }

  if ($Rows.Count -eq 0) {
    return ''
  }

  $ColumnCount = ($Rows | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum

  $Lines = New-Object System.Collections.Generic.List[string]

  for ($r = 0; $r -lt $Rows.Count; $r++) {
    $Row = @($Rows[$r])

    while ($Row.Count -lt $ColumnCount) {
      $Row += ''
    }

    $Lines.Add('| ' + ($Row -join ' | ') + ' |')

    if ($r -eq 0) {
      $Separator = (1..$ColumnCount | ForEach-Object { '---' }) -join ' | '
      $Lines.Add('| ' + $Separator + ' |')
    }
  }

  [string]::Join("`n", $Lines)
}

<#
    .SYNOPSIS
    Converts the full generated HTML report into a Markdown document.

    .DESCRIPTION
    Strips the embedded stylesheet and document scaffolding, then walks the remaining HTML in
    document order, turning h2-h4 headings into Markdown headings, <table> elements into
    Markdown tables (see ConvertTo-MarkdownTable), and <p> elements into plain text lines.

    .PARAMETER Html
    The complete HTML report content, as produced for the HTML report / mail body.

    .EXAMPLE
    ConvertTo-MarkdownReport -Html $Output

    .NOTES
    Author: Thomas Stensitzki
#>
function ConvertTo-MarkdownReport {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Html
  )

  $Text = $Html
  $Text = $Text -replace '(?is)<style.*?</style>', ''
  $Text = $Text -replace '(?is)<!--.*?-->', ''
  $Text = $Text -replace '(?is)</?(html|body)[^>]*>', ''
  $Text = $Text -replace '(?is)<title>.*?</title>', ''

  $Markdown = New-Object System.Text.StringBuilder

  $Pattern = '(?is)<h([2-4])[^>]*>(.*?)</h\1>|<table[^>]*>(.*?)</table>|<p[^>]*>(.*?)</p>'

  foreach ($Match in [regex]::Matches($Text, $Pattern)) {
    if ($Match.Groups[1].Success) {
      $Level = [int]$Match.Groups[1].Value
      $HeadingText = ConvertFrom-HtmlFragment -Fragment $Match.Groups[2].Value
      [void]$Markdown.AppendLine(('{0} {1}' -f ('#' * $Level), $HeadingText))
      [void]$Markdown.AppendLine()
    }
    elseif ($Match.Groups[3].Success) {
      $TableMarkdown = ConvertTo-MarkdownTable -TableInnerHtml $Match.Groups[3].Value
      if ($TableMarkdown) {
        [void]$Markdown.AppendLine($TableMarkdown)
        [void]$Markdown.AppendLine()
      }
    }
    else {
      $ParagraphText = ConvertFrom-HtmlFragment -Fragment $Match.Groups[4].Value
      if ($ParagraphText) {
        [void]$Markdown.AppendLine($ParagraphText)
        [void]$Markdown.AppendLine()
      }
    }
  }

  $Markdown.ToString()
}

<#
    .SYNOPSIS
    Loads the Exchange version, service pack, and security update mapping from a JSON file.

    .DESCRIPTION
    Reads and validates the external version mapping configuration file, converts it from
    PSCustomObject to hashtable form for use with the existing lookup logic, and returns a
    structured hashtable with MajorVersions, ServicePackLevels, SecurityUpdates, and
    MaxCumulativeUpdate. Keeping this mapping in an external file allows Thomas to update
    monthly Security Update labels without touching the script code.

    .PARAMETER Path
    Full path to the JSON version mapping file, typically ExchangeVersionMappings.json
    located next to the script.

    .EXAMPLE
    Import-ExchangeVersionMapping -Path 'C:\Scripts\ExchangeVersionMappings.json'

    .NOTES
    Author: Thomas Stensitzki
#>
function Import-ExchangeVersionMapping {
  [CmdletBinding()]
  [OutputType([hashtable])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  try {
    $JsonContent = Get-Content -Path $Path -Raw -ErrorAction Stop
    $JsonObject = $JsonContent | ConvertFrom-Json -ErrorAction Stop
  }
  catch {
    throw ('Failed to read or parse version mapping file {0}: {1}' -f $Path, $_.Exception.Message)
  }

  $RequiredProperties = @('MajorVersions', 'ServicePackLevels', 'SecurityUpdates', 'MaxCumulativeUpdate')

  foreach ($RequiredProperty in $RequiredProperties) {
    if ($JsonObject.PSObject.Properties.Name -notcontains $RequiredProperty) {
      throw ('Version mapping file {0} is missing required property: {1}' -f $Path, $RequiredProperty)
    }
  }

  @{
    MajorVersions       = (ConvertTo-Hashtable -InputObject $JsonObject.MajorVersions)
    ServicePackLevels   = (ConvertTo-Hashtable -InputObject $JsonObject.ServicePackLevels)
    SecurityUpdates     = (ConvertTo-Hashtable -InputObject $JsonObject.SecurityUpdates)
    MaxCumulativeUpdate = [int]$JsonObject.MaxCumulativeUpdate
  }
}

<#
    .SYNOPSIS
    Builds a simplified information hashtable for a single Database Availability Group.

    .DESCRIPTION
    Extracts name, member count, member server names, and an empty Databases collection
    (populated later by the caller) from a DAG object as returned by Get-DatabaseAvailabilityGroup.

    .PARAMETER DAG
    A single DAG object as returned by Get-DatabaseAvailabilityGroup.
    Kept untyped on purpose, the underlying Exchange type differs across Exchange versions
    and management shells, and strict typing here would tie the script to a specific version.

    .EXAMPLE
    Get-DatabaseAvailabilityGroupInformation -DAG $DAG

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-DatabaseAvailabilityGroupInformation {
  [CmdletBinding()]
  [OutputType([hashtable])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    $DAG
  )

  try {
    @{Name        = $DAG.Name.ToUpper()
      MemberCount	= $DAG.Servers.Count
      Members     = [array]($DAG.Servers | ForEach-Object { $_.Name.ToUpper() })
      Databases   = @()
    }
  }
  catch {
    throw ('Failed to process Database Availability Group information: {0}' -f $_.Exception.Message)
  }
}

<#
    .SYNOPSIS
    Collects size, backup, disk space, and copy information for a single mailbox database.

    .DESCRIPTION
    Builds a hashtable describing a mailbox database, mailbox and archive mailbox counts and
    average sizes, database size and whitespace, DAG copy information, and disk free space.
    Handles both Exchange 2010+ (Database Availability Group aware) and legacy Exchange 2003/2007
    (WMI and ADSI based) code paths.

    .PARAMETER Database
    A single mailbox database object as returned by Get-MailboxDatabase.
    Kept untyped on purpose, the underlying Exchange type differs across Exchange versions.

    .PARAMETER ExchangeEnvironment
    The shared environment hashtable that accumulates server, site, and disk information
    collected earlier in the script run.

    .PARAMETER Mailboxes
    Array of mailbox objects as returned by Get-Mailbox, used to calculate mailbox counts and
    average sizes for the database.

    .PARAMETER ArchiveMailboxes
    Array of archive mailbox objects as returned by Get-Mailbox -Archive, used to calculate
    archive mailbox counts and average sizes. $null on Exchange versions without archive support.

    .PARAMETER E2010
    Indicates whether the script is running against Exchange 2010 or newer, this selects the
    Database Availability Group aware code path instead of the legacy WMI/ADSI based path.

    .EXAMPLE
    Get-DatabaseInformation -Database $Database -ExchangeEnvironment $ExchangeEnvironment -Mailboxes $Mailboxes -ArchiveMailboxes $ArchiveMailboxes -E2010 $true

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-DatabaseInformation {
  [CmdletBinding()]
  [OutputType([hashtable])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    $Database,
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [hashtable]$ExchangeEnvironment,
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Mailboxes,
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$ArchiveMailboxes,
    [Parameter(Mandatory = $true)]
    [bool]$E2010
  )

  # Circular Logging, Last Full Backup
  if ($Database.CircularLoggingEnabled) { $CircularLoggingEnabled = 'Yes' } else { $CircularLoggingEnabled = 'No' }
  if ($Database.LastFullBackup) { $LastFullBackup = $Database.LastFullBackup.ToString() } else { $LastFullBackup = 'Not Available' }

  # Drive Letter, GitHub issue #4
  $DriveNameEdb = ''
  try {
    $DriveNameEdb = $Database.EdbFilePath.DriveName
  }
  catch {
    $DriveNameEdb = $NotAvailable
  }

  $DriveNameLog = ''
  try {
    $DriveNameLog = $Database.LogFolderPath.DriveName
  }
  catch {
    $DriveNameLog = $NotAvailable
  }

  # Mailbox Average Sizes
  $MailboxStatistics = [array]($ExchangeEnvironment.Servers[$Database.Server.Name].MailboxStatistics | Where-Object { $_.Database -eq $Database.Identity })

  if ($MailboxStatistics) {
    [long]$MailboxItemSizeB = 0
    $MailboxStatistics | ForEach-Object { $MailboxItemSizeB += $_.TotalItemSizeB }
    [long]$MailboxAverageSize = $MailboxItemSizeB / $MailboxStatistics.Count
  }
  else {
    $MailboxAverageSize = 0
  }

  # Free Disk Space Percentage
  if ($ExchangeEnvironment.Servers[$Database.Server.Name].Disks) {

    foreach ($Disk in $ExchangeEnvironment.Servers[$Database.Server.Name].Disks) {
      if ($Database.EdbFilePath.PathName -like ('{0}*' -f $Disk.Name)) {
        $FreeDatabaseDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
      }
      if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14) {

        if ($Database.LogFolderPath.PathName -like ('{0}*' -f $Disk.Name)) {
          $FreeLogDiskSpace = ($Disk.FreeSpace / $Disk.Capacity) * 100
        }
      }
      else {
        try {
          $StorageGroupDN = $Database.DistinguishedName.Replace(('CN={0},' -f $Database.Name), '')
          $Adsi = [adsi]"LDAP://$($Database.OriginatingServer)/$($StorageGroupDN)"
          if ($Adsi.msExchESEParamLogFilePath -like ('{0}*' -f $Disk.Name)) {
            $FreeLogDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
          }
        }
        catch {
          Write-Warning -Message ('Cannot read storage group log file path via ADSI for database {0}: {1}' -f $Database.Name, $_.Exception.Message)
          $FreeLogDiskSpace = $null
        }
      }
    }
  }
  else {
    $FreeLogDiskSpace = $null
    $FreeDatabaseDiskSpace = $null
  }

  if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14 -and $E2010) {
    # Exchange 2010 Database Only
    $CopyCount = [int]$Database.Servers.Count

    if ($Database.MasterServerOrAvailabilityGroup.Name -ne $Database.Server.Name) {
      $Copies = [array]($Database.Servers | ForEach-Object { $_.Name.ToUpper() })
    }
    else {
      $Copies = @()
    }
    # Archive Info
    $ArchiveMailboxCount = [int]([array]($ArchiveMailboxes | Where-Object { $_.ArchiveDatabase -eq $Database.Name })).Count

    $ArchiveStatistics = [array]($ArchiveMailboxes | Where-Object { $_.ArchiveDatabase -eq $Database.Name } | Get-MailboxStatistics -Archive -ErrorAction SilentlyContinue )

    if ($ArchiveStatistics) {
      [long]$ArchiveItemSizeB = 0
      $ArchiveStatistics | ForEach-Object { $ArchiveItemSizeB += $_.TotalItemSize.Value.ToBytes() }
      [long]$ArchiveAverageSize = $ArchiveItemSizeB / $ArchiveStatistics.Count
    }
    else {
      $ArchiveAverageSize = 0
    }

    # Disconnected Mailbox, v2.7
    try {
      $DisconnectedMailboxCount = ( (Get-MailboxStatistics -Database $Database.Name -ErrorAction Stop | Where-Object { $_.DisconnectReason -eq 'SoftDeleted' }) | Measure-Object).Count
    }
    catch {
      Write-Warning -Message ('Cannot query disconnected mailbox count for database {0}: {1}' -f $Database.Name, $_.Exception.Message)
      $DisconnectedMailboxCount = 0
    }

    # Database Provisioning Status, v.2.7.1
    $IsExcludedFromProvisioning = $Database.IsExcludedFromProvisioning
    $IsExcludedFromProvisioningByOperator = $Database.IsExcludedFromProvisioningByOperator
    $IsExcludedFromProvisioningReason = $Database.IsExcludedFromProvisioningReason

    # DB Size / Whitespace Info
    [long]$Size = $Database.DatabaseSize.ToBytes()
    [long]$Whitespace = $Database.AvailableNewMailboxSpace.ToBytes()
    $StorageGroup = $null

  }
  else {
    $ArchiveMailboxCount = 0
    $CopyCount = 0
    $Copies = @()
    # 2003 & 2007, Use WMI (Based on code by Gary Siepser, http://bit.ly/kWWMb3)
    try {
      $Size = [long](Get-WmiObject -Class cim_datafile -ComputerName $Database.Server.Name -Filter ('name=''' + $Database.edbfilepath.pathname.replace('\', '\\') + '''') -ErrorAction Stop).filesize
    }
    catch {
      Write-Warning -Message ('Cannot query database file size via WMI for {0}: {1}' -f $Database.Server.Name, $_.Exception.Message)
      $Size = $null
    }

    if (!$Size) {
      Write-Warning -Message ('Cannot detect database size via WMI for {0}' -f $Database.Server.Name)
      [long]$Size = 0
      [long]$Whitespace = 0
    }
    else {
      [long]$MailboxDeletedItemSizeB = 0
      if ($MailboxStatistics) {
        $MailboxStatistics | ForEach-Object { $MailboxDeletedItemSizeB += $_.TotalDeletedItemSizeB }
      }

      # Calculate database whitespace
      $Whitespace = $Size - $MailboxItemSizeB - $MailboxDeletedItemSizeB
      if ($Whitespace -lt 0) { $Whitespace = 0 }
    }

    $StorageGroup = $Database.DistinguishedName.Split(',')[1].Replace('CN=', '')
  }

  @{
    Name                   = $Database.Name
    StorageGroup           = $StorageGroup
    ActiveOwner            = $Database.Server.Name.ToUpper()
    MailboxCount           = [long]([array]($Mailboxes | Where-Object { $_.Database -eq $Database.Identity })).Count
    MailboxAverageSize     = $MailboxAverageSize
    ArchiveMailboxCount    = $ArchiveMailboxCount
    ArchiveAverageSize     = $ArchiveAverageSize
    DisconnectedMailboxCount = $DisconnectedMailboxCount
    CircularLoggingEnabled = $CircularLoggingEnabled
    LastFullBackup         = $LastFullBackup
    Size                   = $Size
    Whitespace             = $Whitespace
    Copies                 = $Copies
    CopyCount              = $CopyCount
    FreeLogDiskSpace       = $FreeLogDiskSpace
    FreeDatabaseDiskSpace  = $FreeDatabaseDiskSpace
    DriveNameEdb           = $DriveNameEdb
    DriveNameLog           = $DriveNameLog
    IsExcludedFromProvisioning = $IsExcludedFromProvisioning
    IsExcludedFromProvisioningByOperator = $IsExcludedFromProvisioningByOperator
    IsExcludedFromProvisioningReason = $IsExcludedFromProvisioningReason

  }
}

<#
    .SYNOPSIS
    Counts the mailboxes hosted on a single Exchange server.

    .DESCRIPTION
    Determines the mailbox count for a server indirectly, by matching mailboxes to the databases
    assigned to that server, rather than relying on the mailbox ServerName property directly, which
    is not always populated reliably by Exchange.

    .PARAMETER Mailboxes
    Array of mailbox objects as returned by Get-Mailbox.

    .PARAMETER ExchangeServer
    A single Exchange server object as returned by Get-ExchangeServer.
    Kept untyped on purpose, the underlying Exchange type differs across Exchange versions.

    .PARAMETER Databases
    Array of mailbox database objects as returned by Get-MailboxDatabase.

    .EXAMPLE
    Get-ExchangeServerMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-ExchangeServerMailboxCount {
  [CmdletBinding()]
  [OutputType([int])]
  param(
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Mailboxes,
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    $ExchangeServer,
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Databases
  )

  $MailboxCount = 0

  foreach ($Database in [array]($Databases | Where-Object { $_.Server -eq $ExchangeServer.Name })) {
    $MailboxCount += ([array]($Mailboxes | Where-Object { $_.Database -eq $Database.Identity })).Count
  }

  $MailboxCount

}

<#
    .SYNOPSIS
    Counts the archive mailboxes hosted on a single Exchange server.

    .DESCRIPTION
    Determines the archive mailbox count for a server by matching mailboxes to the archive
    databases assigned to that server, mirroring the approach used by Get-ExchangeServerMailboxCount.

    .PARAMETER Mailboxes
    Array of mailbox objects as returned by Get-Mailbox.

    .PARAMETER ExchangeServer
    A single Exchange server object as returned by Get-ExchangeServer.
    Kept untyped on purpose, the underlying Exchange type differs across Exchange versions.

    .PARAMETER Databases
    Array of mailbox database objects as returned by Get-MailboxDatabase.

    .EXAMPLE
    Get-ExchangeServerArchiveMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-ExchangeServerArchiveMailboxCount {
  [CmdletBinding()]
  [OutputType([int])]
  param(
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Mailboxes,
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    $ExchangeServer,
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Databases
  )
  $ArchiveMailboxCount = 0

  foreach ($Database in [array]($Databases | Where-Object { $_.Server -eq $ExchangeServer.Name })) {
    $ArchiveMailboxCount += ([array]($Mailboxes | Where-Object { $_.ArchiveDatabase -eq $Database.Identity })).Count
  }

  $ArchiveMailboxCount

}

<#
    .SYNOPSIS
    Normalizes a virtual directory hostname value for display in the HTML report.

    .DESCRIPTION
    Some Exchange virtual directory properties (e.g. InternalUrl, ExternalUrl derived hostnames)
    can be $null or an empty string. This function returns a consistent 'None' placeholder in
    that case, and a trimmed string otherwise. Added for Issue #9, empty virtual directory
    hostname strings previously rendered inconsistently in the report.

    .PARAMETER VDirHost
    The virtual directory hostname value to normalize. May be $null.

    .EXAMPLE
    Test-vDirHost -VDirHost $OwaVirtualDirectory.InternalUrl.Host

    .NOTES
    Author: Thomas Stensitzki
#>
function Test-vDirHost {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    $VDirHost
  )

  [string]$Hostname = 'None'

  if ($null -ne $VDirHost) {
    $Hostname = ([string]$VDirHost).Trim()
  }

  $Hostname
}

<#
    .SYNOPSIS
    Collects detailed configuration and health information for a single Exchange server.

    .DESCRIPTION
    Gathers OS version, disk information, Exchange version, service pack/CU/security update
    level, update rollup level, installed roles, mailbox/archive mailbox counts, virtual
    directory hostnames, and CAS array membership for a single Exchange server. Supports
    Exchange 2000 through Exchange 2019/SE via WMI, remote registry, and ADSI as appropriate
    for the detected version.

    .PARAMETER E2010
    Indicates whether the script is running against Exchange 2010 or newer.

    .PARAMETER ExchangeServer
    A single Exchange server object as returned by Get-ExchangeServer.
    Kept untyped on purpose, the underlying Exchange type differs across Exchange versions.

    .PARAMETER Mailboxes
    Array of mailbox objects as returned by Get-Mailbox, used to calculate mailbox counts.

    .PARAMETER Databases
    Array of mailbox database objects as returned by Get-MailboxDatabase.

    .PARAMETER Hybrids
    Array of server names participating in hybrid mail flow, used to flag the server with an
    additional 'Hybrid' role in the report.

    .EXAMPLE
    Get-ExchangeServerInformation -E2010 $true -ExchangeServer $Server -Mailboxes $Mailboxes -Databases $Databases -Hybrids $HybridServers

    .NOTES
    Author: Thomas Stensitzki
    Requires WMI and Remote Registry access to the target Exchange server for full results.
#>
function Get-ExchangeServerInformation {
  [CmdletBinding()]
  [OutputType([hashtable])]
  param(
    [Parameter(Mandatory = $true)]
    [bool]$E2010,
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    $ExchangeServer,
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Mailboxes,
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Databases,
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Hybrids
  )

  # Set Basic Variables
  $MailboxCount = 0
  $ArchiveMailboxCount = 0 # new 3.0
  $RollupLevel = 0
  $RollupVersion = ''
  $ExtNames = @()
  $IntNames = @()
  $CASArrayName = ''

  # 2019-05-20 TST Added to handle max preferred/active databases per server
  $MaxPrefDatabases = 0
  $MaxActiveDatabases = 0
  $NotSet = '--'

  # Get WMI Information: Operatin System
  $tWMI = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue

  if ($tWMI) {
    $OSVersion = $tWMI.Caption.Replace('(R)', '').Replace('Microsoft ', '').Replace('Enterprise', 'Ent').Replace('Standard', 'Std').Replace(' Edition', '')
    $OSServicePack = $tWMI.CSDVersion
    $RealName = $tWMI.CSName.ToUpper()
  }
  else {
    Write-Warning -Message ('Cannot detect OS information via WMI for {0}' -f $ExchangeServer.Name)
    $OSVersion = $NotAvailable
    $OSServicePack = $NotAvailable
    $RealName = $ExchangeServer.Name.ToUpper()
  }

  # Get WMI Information: Disk Space
  $tWMI = Get-WmiObject -Query 'Select * from Win32_Volume' -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue

  if ($tWMI) {
    $Disks = $tWMI | Select-Object -Property Name, Capacity, FreeSpace | Sort-Object -Property Name
  }
  else {
    Write-Warning -Message ('Cannot detect OS information via WMI for {0}' -f $ExchangeServer.Name)
    $Disks = $null
  }

  # Get Exchange Version
  if ($ExchangeServer.AdminDisplayVersion.Major -eq 6) {
    $ExchangeMajorVersion = [double]('{0}.{1}' -f $ExchangeServer.AdminDisplayVersion.Major, $ExchangeServer.AdminDisplayVersion.Minor)
    $ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.FilePatchLevelDescription.Replace('Service Pack ', '')
  }
  elseif (($ExchangeServer.AdminDisplayVersion.Major -eq 15) -and ($ExchangeServer.AdminDisplayVersion.Minor -eq 2) -and ($ExchangeServer.AdminDisplayVersion.Build -ge 2562) ) {
    $ExchangeMajorVersion = [double]('{0}.{1}' -f $ExchangeServer.AdminDisplayVersion.Major, (([Double]($ExchangeServer.AdminDisplayVersion.Minor))+1) ) # dirty trick so separate 2019 fro SE while both are 15.2.x
    $ExchangeSPLevel = 0
  }
  elseif ($ExchangeServer.AdminDisplayVersion.Major -eq 15 -and $ExchangeServer.AdminDisplayVersion.Minor -ge 1) {
    $ExchangeMajorVersion = [double]('{0}.{1}' -f $ExchangeServer.AdminDisplayVersion.Major, $ExchangeServer.AdminDisplayVersion.Minor)
    $ExchangeSPLevel = 0
  }
  else {
    $ExchangeMajorVersion = $ExchangeServer.AdminDisplayVersion.Major
    $ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.Minor
  }

  #Write-host ('{0}.{1}.{2}' -f $ExchangeServer.AdminDisplayVersion.Major, $ExchangeServer.AdminDisplayVersion.Minor, $ExchangeServer.AdminDisplayVersion.Build)
  #Write-host ('ExchangeMajorVersion: {0}' -f $ExchangeMajorVersion)

  # Exchange 2007+
  if ($ExchangeMajorVersion -ge 8) {
    # Get Roles
    $MailboxStatistics = $null
    [array]$Roles = $ExchangeServer.ServerRole.ToString().Replace(' ', '').Split(',')

    # Add Hybrid "Role" for report
    if ($Hybrids -contains $ExchangeServer.Name) {
      $Roles += 'Hybrid'
    }

    if ($Roles -contains 'Mailbox') {

      $MailboxCount = Get-ExchangeServerMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases

      $ArchiveMailboxCount = Get-ExchangeServerArchiveMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases # new 3.0

      if ($ExchangeServer.Name.ToUpper() -ne $RealName) {
        $Roles = [array]($Roles | Where-Object { $_ -ne 'Mailbox' })
        $Roles += 'ClusteredMailbox'
      }

      # Get Mailbox Statistics the normal way, return in a consistent format
      # 2019-05-20 TST, try/catch added
      try {
        Write-Verbose -Message ('Fetching Mailbox Statistics - {0}' -f $ExchangeServer)
        $MailboxStatistics = Get-MailboxStatistics -Server $ExchangeServer -ErrorAction SilentlyContinue | Select-Object -Property DisplayName, @{Name = 'TotalItemSizeB'; Expression = { $_.TotalItemSize.Value.ToBytes() } }, @{Name = 'TotalDeletedItemSizeB'; Expression = { $_.TotalDeletedItemSize.Value.ToBytes() } }, Database
      }
      catch {
        $MailboxStatistics = $null
        Write-Warning -Message ('Cannot get mailbox statistics for server {0}' -f $ExchangeServer)
      }

      if ($ExchangeMajorVersion -ge 14) {
        try {
          $mailboxServer = Get-MailboxServer -Identity $($ExchangeServer.Name) -ErrorAction Stop
        }
        catch {
          Write-Warning -Message ('Cannot query mailbox server configuration for {0}: {1}' -f $ExchangeServer.Name, $_.Exception.Message)
          $mailboxServer = $null
        }

        # 2019-05-20 TST Gather max active/max preferred database config
        if ($ExchangeMajorVersion -lt 15) {
          # Exchange 2010
          $MaxActiveDatabases = $mailboxServer.MaximumActiveDatabases
        }
        else {
          # Exchange 2013+
          if ($null -ne $mailboxServer.MaximumPreferredActiveDatabases) {
            $MaxPrefDatabases = $mailboxServer.MaximumPreferredActiveDatabases
          }
          else {
            $MaxPrefDatabases = $NotSet
          }

          if ($null -ne $mailboxServer.MaximumActiveDatabases) {
            $MaxActiveDatabases = $mailboxServer.MaximumActiveDatabases
          }
          else {
            $MaxActiveDatabases = $NotSet
          }
        }
      }
    }

    # Get HTTPS Names (Exchange 2010 only due to time taken to retrieve data)
    # 2019-05-16 TST | Update to support 'Mailbox' role for gathering namespace information
    if (($Roles -contains 'ClientAccess' -and $E2010) -or ($Roles -contains 'Mailbox' -and $E2010)) {
      try {
        Get-OWAVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly -ErrorAction Stop | ForEach-Object { $ExtNames += (Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames += (Test-vDirHost -VDirHost $_.InternalURL.Host) }

        Get-WebServicesVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly -ErrorAction Stop | ForEach-Object { $ExtNames += (Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames += (Test-vDirHost -VDirHost $_.InternalURL.Host) }

        Get-OABVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly -ErrorAction Stop | ForEach-Object { $ExtNames += (Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames += (Test-vDirHost -VDirHost $_.InternalURL.Host) }

        Get-ActiveSyncVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly -ErrorAction Stop | ForEach-Object { $ExtNames += (Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames += (Test-vDirHost -VDirHost $_.InternalURL.Host) }

        if (Get-Command -Name Get-MAPIVirtualDirectory -ErrorAction SilentlyContinue) {
          Get-MAPIVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly -ErrorAction Stop | ForEach-Object { $ExtNames += (Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames += (Test-vDirHost -VDirHost $_.InternalURL.Host) }
        }

        if (Get-Command -Name Get-ClientAccessService -ErrorAction SilentlyContinue) {
          $IntNames += (Test-vDirHost -VDirHost (Get-ClientAccessService -Identity $ExchangeServer.Name -ErrorAction Stop).AutoDiscoverServiceInternalURI.Host)
        }
        else {
          # Fallback to use Get-ClientAccessServer cmdlet
          $IntNames += (Test-vDirHost -VDirHost (Get-ClientAccessServer -Identity $ExchangeServer.Name -ErrorAction Stop).AutoDiscoverServiceInternalURI.Host)
        }

        if ($ExchangeMajorVersion -ge 14) {
          Get-ECPVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly -ErrorAction Stop | ForEach-Object { $ExtNames += (Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames += (Test-vDirHost -VDirHost $_.InternalURL.Host); }
        }

        $CASArray = Get-ClientAccessArray -Site $ExchangeServer.Site.Name -ErrorAction Stop

        if ($CASArray) {
          $CASArrayName = $CASArray.Fqdn
        }
      }
      catch {
        Write-Warning -Message ('Cannot fully retrieve virtual directory or CAS array information for {0}: {1}' -f $ExchangeServer.Name, $_.Exception.Message)
      }

      $IntNames = $IntNames | Sort-Object -Unique
      $ExtNames = $ExtNames | Sort-Object -Unique
    }

    # Rollup Level / Versions
    # Thanks to Bhargav Shukla https://bhargavs.com/index.php/2009/12/14/how-do-i-check-update-rollup-version-on-exchange-20xx-server/
    switch ([string]$ExchangeMajorVersion) {
      # Exchange Server 2016 / 2019 / SE
      {'15.2','15.3'} { $RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\442189DC8B9EA5040962A6BED9EC1F1F\\Patches" }
      '15.1' { $RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\442189DC8B9EA5040962A6BED9EC1F1F\\Patches" }
      # Exchange Server 2010 / 2013
      '15' { $RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches" }
      '14' { $RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches" }
      # Exchange 2007
      default { $RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\461C2B4266EDEF444B864AD6D9E5B613\\Patches" }
    }

    # 2019-05-17 Thomas Stensitzki, try/catch added
    try {
      $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name)
    }
    catch {
      $RemoteRegistry = $null
    }

    if ($null -ne $RemoteRegistry) {

      $RUKeys = $RemoteRegistry.OpenSubKey($RegKey).GetSubKeyNames() | ForEach-Object { "$RegKey\\$_" }

      if ($RUKeys) {
        [array]($RUKeys | ForEach-Object { $RemoteRegistry.OpenSubKey($_).GetValue('DisplayName') }) | `
          ForEach-Object {
          if ($_ -like 'Update Rollup *') {
            $tRU = $_.Split(' ')[2]
            if ($tRU -like '*-*') { $tRUV = $tRU.Split('-')[1]; $tRU = $tRU.Split('-')[0] } else { $tRUV = '' }
            if ([int]$tRU -ge [int]$RollupLevel) { $RollupLevel = $tRU; $RollupVersion = $tRUV }
          }
        }
      }
    }
    else {
      Write-Warning -Message ('Cannot detect Rollup Version via Remote Registry for {0}' -f $ExchangeServer.Name)
    }

    # Exchange 2013+ CU or SP Level
    # 2023-12-28 TST, added Exchange SU support
    if ($ExchangeMajorVersion -ge 15) {
      $RegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Microsoft Exchange v15"
      $RegKeyBuildVersion = "SOFTWARE\\Microsoft\\ExchangeServer\\v15\\Setup"
      # 2019-05-17 Thomas Stensitzki, try/catch added
      try {
        $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name)
      }
      catch {
        $RemoteRegistry = $null
      }

      if ($RemoteRegistry) {
        $ExchangeSPLevel = $RemoteRegistry.OpenSubKey($RegKey).GetValue('DisplayName')

        if ($ExchangeSPLevel -like '*Service Pack*' -or $ExchangeSPLevel -like '*Cumulative Update*') {
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Microsoft Exchange Server 2013 ', '')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Microsoft Exchange Server 2016 ', '')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Microsoft Exchange Server 2019 ', '')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Microsoft Exchange Server Subscription Edition ', '')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Service Pack ', 'SP')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Cumulative Update ', 'CU')
        }
        else {
          $ExchangeSPLevel = 0
        }
      }
      else {
        Write-Warning -Message ('Cannot detect CU/SP via Remote Registry for {0}' -f $ExchangeServer.Name)
      }

      if ($RemoteRegistry){
        $ExchangeOwaVersion = $RemoteRegistry.OpenSubKey($RegKeyBuildVersion).GetValue('OwaVersion')

        $SUVersion = $ExchangeOwaVersion #$ExSUString.GetEnumerator() | Where-Object { $_.Key -eq $ExchangeOwaVersion } | Select-Object -ExpandProperty Value
      }
      else {
        $SUVersion = 'XX'
      }
    }
  }

  # Exchange 2003
  if ($ExchangeMajorVersion -eq 6.5) {

    # Mailbox Count
    $MailboxCount = Get-ExchangeServerMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases

    # Get Role via WMI
    $tWMI = Get-WMIObject -Class Exchange_Server -Namespace 'root\microsoftexchangev2' -ComputerName $ExchangeServer.Name -Filter "Name='$($ExchangeServer.Name)'"

    if ($tWMI) {
      if ($tWMI.IsFrontEndServer) { $Roles = @('FE') } else { $Roles = @('BE') }
    }
    else {
      Write-Warning -Message ('Cannot detect Front End/Back End Server information via WMI for {0}' -f $ExchangeServer.Name)
      $Roles += 'Unknown'
    }

    # Get Mailbox Statistics using WMI, return in a consistent format
    $tWMI = Get-WMIObject -class Exchange_Mailbox -Namespace ROOT\MicrosoftExchangev2 -ComputerName $ExchangeServer.Name -Filter ("ServerName='$($ExchangeServer.Name)'")
    if ($tWMI) {
      $MailboxStatistics = $tWMI | Select-Object -Property @{Name = 'DisplayName'; Expression = { $_.MailboxDisplayName } }, @{Name = 'TotalItemSizeB'; Expression = { $_.Size } }, @{Name = 'TotalDeletedItemSizeB'; Expression = { $_.DeletedMessageSizeExtended } }, @{Name = 'Database'; Expression = { ((Get-MailboxDatabase -Identity "$($_.ServerName)\$($_.StorageGroupName)\$($_.StoreName)").Identity) } }
    }
    else {
      Write-Warning -Message ('Cannot retrieve Mailbox Statistics via WMI for {0}' -f $ExchangeServer.Name)
      $MailboxStatistics = $null
    }
  }

  # Exchange 2000
  if ($ExchangeMajorVersion -eq '6.0') {
    # Mailbox Count
    $MailboxCount = Get-ExchangeServerMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases

    # Get Role via ADSI
    try {
      $tADSI = [ADSI]"LDAP://$($ExchangeServer.OriginatingServer)/$($ExchangeServer.DistinguishedName)"
    }
    catch {
      Write-Warning -Message ('Cannot bind to ADSI object for {0}: {1}' -f $ExchangeServer.Name, $_.Exception.Message)
      $tADSI = $null
    }

    if ($tADSI) {
      if ($tADSI.ServerRole -eq 1) { $Roles = @('FE') } else { $Roles = @('BE') }
    }
    else {
      Write-Warning -Message ('Cannot detect Front End/Back End Server information via ADSI for {0}' -f $ExchangeServer.Name)
      $Roles += 'Unknown'
    }
    $MailboxStatistics = $null
  }

  # Return Hashtable
  @{
    Name                      = $ExchangeServer.Name.ToUpper()
    RealName                  = $RealName
    ExchangeMajorVersion      = $ExchangeMajorVersion
    ExchangeSPLevel           = $ExchangeSPLevel
    Edition                   = $ExchangeServer.Edition
    Mailboxes                 = $MailboxCount
    ArchiveMailboxes          = $ArchiveMailboxCount
    OSVersion                 = $OSVersion
    OSServicePack             = $OSServicePack
    Roles                     = $Roles
    RollupLevel               = $RollupLevel
    RollupVersion             = $RollupVersion
    SuVersion                 = $SUVersion
    Site                      = $ExchangeServer.Site.Name
    MailboxStatistics         = $MailboxStatistics
    Disks                     = $Disks
    IntNames                  = $IntNames
    ExtNames                  = $ExtNames
    CASArrayName              = $CASArrayName
    MaximumPreferredDatabases = $MaxPrefDatabases
    MaximumActiveDatabases    = $MaxActiveDatabases
  }
}

<#
    .SYNOPSIS
    Aggregates server and mailbox counts per Exchange version across the environment.

    .DESCRIPTION
    Walks all collected sites and pre-Exchange 2007 servers and builds a hashtable keyed by
    "MajorVersion.SPLevel", each entry holding a ServerCount and MailboxCount total. Used to
    render the version summary table at the top of the HTML report.

    .PARAMETER ExchangeEnvironment
    The shared environment hashtable populated earlier in the script run, containing Sites
    and Pre2007 server collections.

    .EXAMPLE
    Get-TotalsByVersion -ExchangeEnvironment $ExchangeEnvironment

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-TotalsByVersion {
  [CmdletBinding()]
  [OutputType([hashtable])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [hashtable]$ExchangeEnvironment
  )

  # Create empty hash table
  $TotalMailboxesByVersion = @{}

  if ($ExchangeEnvironment.Sites) {
    foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator()) {
      foreach ($Server in $Site.Value) {
        if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"]) {
          $TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)", @{ServerCount = 1; MailboxCount = $Server.Mailboxes })
        }
        else {
          $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
          $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount += $Server.Mailboxes
        }
      }
    }
  }

  if ($ExchangeEnvironment.Pre2007) {
    foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator()) {
      foreach ($Server in $FakeSite.Value) {
        if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"]) {
          $TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)", @{ServerCount = 1; MailboxCount = $Server.Mailboxes })
        }
        else {
          $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
          $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount += $Server.Mailboxes
        }
      }
    }
  }
  $TotalMailboxesByVersion
}

<#
    .SYNOPSIS
    Aggregates server counts per Exchange role across the environment.

    .DESCRIPTION
    Walks all collected sites and pre-Exchange 2007 servers and counts how many servers hold
    each Exchange role (ClientAccess, HubTransport, UnifiedMessaging, Mailbox, Edge, and any
    additional roles encountered such as Hybrid or ClusteredMailbox). Used to render the role
    summary table at the top of the HTML report.

    .PARAMETER ExchangeEnvironment
    The shared environment hashtable populated earlier in the script run, containing Sites
    and Pre2007 server collections.

    .EXAMPLE
    Get-TotalsByRole -ExchangeEnvironment $ExchangeEnvironment

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-TotalsByRole {
  [CmdletBinding()]
  [OutputType([hashtable])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [hashtable]$ExchangeEnvironment
  )

  # Add Roles We Always Show
  $TotalServersByRole = @{
    'ClientAccess'     = 0
    'HubTransport'     = 0
    'UnifiedMessaging' = 0
    'Mailbox'          = 0
    'Edge'             = 0
  }

  if ($ExchangeEnvironment.Sites) {

    foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator()) {

      foreach ($Server in $Site.Value) {

        foreach ($Role in $Server.Roles) {
          if ($null -eq $TotalServersByRole[$Role]) {
            $TotalServersByRole.Add($Role, 1)
          }
          else {
            $TotalServersByRole[$Role]++
          }
        }
      }
    }
  }

  if ($ExchangeEnvironment.Pre2007['Pre 2007 Servers']) {

    foreach ($Server in $ExchangeEnvironment.Pre2007['Pre 2007 Servers']) {

      foreach ($Role in $Server.Roles) {
        if ($null -eq $TotalServersByRole[$Role]) {
          $TotalServersByRole.Add($Role, 1)
        }
        else {
          $TotalServersByRole[$Role]++
        }
      }
    }
  }

  $TotalServersByRole
}

<#
    .SYNOPSIS
    Renders the HTML overview table for a single Active Directory site or the pre-Exchange 2007
    server group.

    .DESCRIPTION
    Builds the site or pre-2007 header row, the internal/external namespace and CAS array
    summary line, and delegates to the per-server row rendering to produce the full HTML
    overview table for one site (or the pre-2007 pseudo site) shown in the report.

    .PARAMETER Servers
    A single dictionary entry (key/value pair) from ExchangeEnvironment.Sites or
    ExchangeEnvironment.Pre2007, where Key is the site name and Value is the array of servers.

    .PARAMETER ExchangeEnvironment
    The shared environment hashtable populated earlier in the script run.

    .PARAMETER ExRoleStrings
    Hashtable mapping internal role names to their display labels shown as column headers.

    .PARAMETER Pre2007
    Switch indicating whether this call renders the pre-Exchange 2007 server group instead of
    a regular Active Directory site.

    .EXAMPLE
    Get-HtmlOverview -Servers $Site -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-HtmlOverview {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    $Servers,
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [hashtable]$ExchangeEnvironment,
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    $ExRoleStrings,
    [switch]$Pre2007
  )


  if ($Pre2007) {
    $BGColHeader = '#880099'
    $BGColSubHeader = '#8800CC'
    $Prefix = ''
    $IntNamesText = ''
    $ExtNamesText = ''
    $CASArrayText = ''
    $ColClass = 'pre2007overviewcolheader'
    $ColSubClass = 'pre2007overviewcolsubheader'
  }
  else {
    $BGColHeader = '#000099'
    $BGColSubHeader = '#0000FF'
    $Prefix = 'Site:'
    $IntNamesText = ''
    $ExtNamesText = ''
    $CASArrayText = ''
    $IntNames = @()
    $ExtNames = @()
    $CASArrayName = ''
    $ColClass = 'overviewcolheader'
    $ColSubClass = 'overviewcolsubheader'

    foreach ($Server in $Servers.Value) {
      $IntNames += $Server.IntNames
      $ExtNames += $Server.ExtNames
      $CASArrayName = $Server.CASArrayName
    }

    $IntNames = $IntNames | Sort-Object -Unique
    $ExtNames = $ExtNames | Sort-Object -Unique

    $IntNames = [string]::Join(', ', $IntNames)
    $ExtNames = [string]::Join(', ', $ExtNames)
    $ExtNamesEmptyText = 'At least one of the analysed servers contains an empty ExternalUrl entry.'

    if ($IntNames) {

      $ExtNamesText = ('External Names: {0}<br />' -f $ExtNames)

      if ($ExtNames -notlike '*None*') {
        $IntNamesText = ('Internal Names: {0}' -f $IntNames)
      }
      else {
        $IntNamesText = ('Internal Names: {0}<br />{1}' -f $IntNames, $ExtNamesEmptyText)
      }
    }

    if ($CASArrayName) {
      $CASArrayText = "CAS Array: $($CASArrayName)"
    }
  }

  $Output = "<table class='overview'>
    <col width='20%'>
    <col width='20%'>
  <colgroup width='25%'>"

  $ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { $Output += "<col width='3%'>" }

  $Output += "</colgroup><col width='20%'><col width='20%'>
    <tr class=""$($ColClass)"">
    <th class='overview'>$($Prefix) $($Servers.Key)</th>
    <th class='overview' colspan=""$(($ExchangeEnvironment.TotalServersByRole.Count)+2)"" align=""left"">$($ExtNamesText)$($IntNamesText)</th>
    <th class='overview' align=""center"">$($CASArrayText)</th>
    <th colspan='2'>&nbsp;</th>
  </tr>"
  $TotalMailboxes = 0
  $Servers.Value | ForEach-Object { $TotalMailboxes += $_.Mailboxes }
  $Output += "<tr class=""$($ColSubClass)""><th class='overview'>Mailboxes: $($TotalMailboxes)</th>"
  $Output += "<th class='overview'>Exchange Version</th>"
  $ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { $Output += "<th class='overview'>$($ExRoleStrings[$_.Key].Short)</th>" }

  # 2016-04-19 Thomas Stensitzki Pref/Max Databases added
  $Output += "<th class='overview'>Databases Pref/Max</th><th class='overview'>OS Version</th><th class='overview'>OS Service Pack</th></tr>"

  $AlternateRow = 0

  foreach ($Server in $Servers.Value) {
    $Output += '<tr'

    if ($AlternateRow) {
      $Output += " class='alternaterow'"
      $AlternateRow = 0
    }
    else {
      $AlternateRow = 1
    }

    $Output += ('><td>{0}' -f $Server.Name)

    if ($Server.RealName -ne $Server.Name) {
      $Output += (' ({0})' -f $Server.RealName)
    }

    $Output += "</td><td>$($ExVersionStrings["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].Long) $($ExSUString[$Server.SuVersion])</td>"

    if ($Server.RollupLevel -gt 0) {
      $Output += (' UR{0}' -f $Server.RollupLevel)
      if ($Server.RollupVersion) {
        $Output += (' {0}' -f $Server.RollupVersion)
      }
    }

    $Output += '</td>'

    $ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort-Object -Property Name | ForEach-Object {
      $Output += '<td'
      if ($Server.Roles -contains $_.Key) {
        $Output += " class='roledata'>"
      }
      else {
        $Output += ' />'
      }

      if (($_.Key -eq 'ClusteredMailbox' -or $_.Key -eq 'Mailbox' -or $_.Key -eq 'BE') -and $Server.Roles -contains $_.Key) {
        # $Output += ('<div class="tooltip">{0}<span class="tooltiptext">{0} mailbox(es)</span></div>/{1}</td>' -f $Server.Mailboxes, $Server.ArchiveMailboxes) # add archive mailbox count
        $Output += ('{0}/{1}</td>' -f $Server.Mailboxes, $Server.ArchiveMailboxes) # add archive mailbox count
      }
    }

    # 2016-04-19 Thomas Stensitzki Max Databases added
    if ($Server.Roles -contains 'Mailbox') {
      $Output += "<td align=""center"">$($Server.MaximumPreferredDatabases) / $($Server.MaximumActiveDatabases)</td><td>$($Server.OSVersion)</td><td>$($Server.OSServicePack)</td></tr>"
    }
    else {
      $Output += ('<td align=""center"">Not Applicable</td><td>{0}</td><td>{1}</td></tr>' -f $Server.OSVersion, $Server.OSServicePack)
    }

    # $Output+="<td>$($Server.OSVersion)</td><td>$($Server.OSServicePack)</td></tr>";
  }

  $Output += '<tr />
  </table><br />'

  $Output
}

<#
    .SYNOPSIS
    Renders the HTML database information table for a DAG or the non-DAG database group.

    .DESCRIPTION
    Dynamically shows or hides optional columns (Storage Group, Archive Mailboxes, Copies,
    Last Full Backup, Circular Logging, free database/log disk space) depending on whether at
    least one database in the collection has relevant data, keeping the report compact for
    environments that do not use a given feature.

    .PARAMETER Databases
    Array of database information hashtables as returned by Get-DatabaseInformation.

    .EXAMPLE
    Get-HtmlDatabaseInformationTable -Databases $DAG.Databases

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-HtmlDatabaseInformationTable {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [array]$Databases
  )

  # Only Show Archive Mailbox Columns, Backup Columns and Circ Logging if at least one DB has an Archive mailbox, backed up or Cir Log enabled.
  $ShowArchiveDBs = $False
  $ShowLastFullBackup = $False
  $ShowCircularLogging = $False
  $ShowStorageGroups = $False
  $ShowCopies = $False
  $ShowFreeDatabaseSpace = $False
  $ShowFreeLogDiskSpace = $False

  foreach ($Database in $Databases) {
    if ($Database.ArchiveMailboxCount -gt 0) {
      $ShowArchiveDBs = $True
    }
    if ($Database.LastFullBackup -ne 'Not Available') {
      $ShowLastFullBackup = $True
    }
    if ($Database.CircularLoggingEnabled -eq 'Yes') {
      $ShowCircularLogging = $True
    }
    if ($Database.StorageGroup) {
      $ShowStorageGroups = $True
    }
    if ($Database.CopyCount -gt 0) {
      $ShowCopies = $True
    }
    if ($null -ne $Database.FreeDatabaseDiskSpace) {
      $ShowFreeDatabaseSpace = $true
    }
    if ($null -ne $Database.FreeLogDiskSpace) {
      $ShowFreeLogDiskSpace = $true
    }
  }

  $Output = "<table class='databases'>"

  #region database table header
  $Output += "<tr class='databases'>
    <th>Server</th>"

  if ($ShowStorageGroups) {
    $Output += '<th>Storage Group</th>'
  }

  $Output += '<th>Database Name</th>
    <th>Standard Mailboxes</th>
  <th>Av. Mailbox Size</th>'

  if ($ShowArchiveDBs) {
    $Output += '<th>Archive Mailboxes</th><th>Av. Archive Size</th>'
  }

  $Output += '<th>DB Size</th><th>DB Whitespace</th>'

  if ($ShowDisconnectedMailboxCount) {
    $Output += '<th>Disconnected Mailboxes</th>'
  }
  if ($ShowFreeDatabaseSpace) {
    $Output += '<th>Database Disk Free</th>'
  }
  if ($ShowFreeLogDiskSpace) {
    $Output += '<th>Log Disk Free</th>'
  }
  if ($ShowLastFullBackup) {
    $Output += '<th>Last Full Backup</th>'
  }
  if ($ShowCircularLogging) {
    $Output += '<th>Circular Logging</th>'
  }
  if ($ShowCopies) {
    $Output += '<th>DB Copies (n)</th>'
  }

  # Provisioning Status, v2.7.1
  if($ShowProvisioningStatus) {
    $Output += '<th>Privisioning Status</th>'
  }

  # Drive names, issue #4
  if ($ShowDriveNames) {
    $Output += '<th>EDB / LOG</th>'
  }

  $Output += '</tr>'
  #endregion

  $AlternateRow = 0

  foreach ($Database in $Databases) {
    $Output += '<tr'

    if ($AlternateRow) {
      $Output += " class='alternaterow'"
      $AlternateRow = 0
    }
    else {
      $AlternateRow = 1
    }

    # Close open <tr tag
    $Output += ('><td>{0}</td>' -f $Database.ActiveOwner)

    if ($ShowStorageGroups) {
      $Output += ('><td>{0}</td>' -f $Database.StorageGroup)
    }

    $Output += "<td>$($Database.Name)</td>
      <td class='center'>$($Database.MailboxCount)</td>"

    # Display average mailbox size in GB, if requested
    if($ShowAverageMailboxSizeInGB) {
      $Output+="<td class='center'>$("{0:N2}" -f ($Database.MailboxAverageSize/1GB)) GB</td>"
    }
    else {
      $Output+="<td class='center'>$("{0:N2}" -f ($Database.MailboxAverageSize/1MB)) MB</td>"
    }

    if ($ShowArchiveDBs) {
      $Output += "<td class=""center"">$($Database.ArchiveMailboxCount)</td>"

      # Display average archive mailbox size in GB, if requested
      if($ShowAverageMailboxSizeInGB) {
        $Output+="<td class='center'>$("{0:N2}" -f ($Database.ArchiveAverageSize/1GB)) GB</td>"
      }
      else {
        $Output+="<td class='center'>$("{0:N2}" -f ($Database.ArchiveAverageSize/1MB)) MB</td>"
      }
    }

    if ([double]($Database.Size / 1GB) -le $MaxDatabaseSize) {
      $Output += "<td class=""center"">$("{0:N2}" -f ($Database.Size/1GB)) GB </td>"
    }
    else {
      $Output += "<td class=""center alert"">$("{0:N2}" -f ($Database.Size/1GB)) GB &uarr;</td>"
    }

    $Output += "<td class='center'>$("{0:N2} GB" -f ($Database.Whitespace/1GB))</td>"

    # v2.7
    if($ShowDisconnectedMailboxCount) {
      $Output += "<td class='center'>$($Database.DisconnectedMailboxCount)</td>"
    }

    # $Output+="<td align=""center"">$("{0:N2}" -f ($Database.Size/1GB)) GB </td><td class='center'>$("{0:N2}" -f ($Database.Whitespace/1GB)) GB</td>"

    if ($ShowFreeDatabaseSpace) {
      if ([double]($Database.FreeDatabaseDiskSpace) -gt $MinFreeDiskspace) {
        $Output += "<td class='center'>$("{0:N1}" -f $Database.FreeDatabaseDiskSpace)%</td>"
      }
      else {
        $Output += "<td class='center alert'>$("{0:N1}" -f $Database.FreeDatabaseDiskSpace)% &darr;</td>"
      }
    }
    if ($ShowFreeLogDiskSpace) {
      if ([double]($Database.FreeLogDiskSpace) -gt $MinFreeDiskspace) {
        $Output += "<td class='center'>$("{0:N1}" -f $Database.FreeLogDiskSpace)%</td>"
      }
      else {
        $Output += "<td class='center alert'>$("{0:N1}" -f $Database.FreeLogDiskSpace)% &darr;</td>"
      }
    }
    if ($ShowLastFullBackup) {
      $Output += "<td class='center'>$($Database.LastFullBackup)</td>"
    }
    if ($ShowCircularLogging) {
      $Output += "<td class='center'>$($Database.CircularLoggingEnabled)</td>"
    }
    if ($ShowCopies) {
      $Output += "<td>$($Database.Copies | ForEach-Object{$_}) ($($Database.CopyCount))</td>"
    }

    # Database Provisioning Status
    if($ShowProvisioningStatus) {
      $ProvisioningStatus = @()
      if($Database.IsExcludedFromProvisioning) {
        $ProvisioningStatus += 'IsExcluded'
      }
      if($Database.IsExcludedFromProvisioningByOperator) {
        $ProvisioningStatus += 'IsExludedByOperator'
      }

      if(-not ([string]::IsNullOrEmpty($Database.IsExcludedFromProvisioningReason) ) ) {
        $ProvisioningReason = ('Reason: {0}' -f $Database.IsExcludedFromProvisioningReason)
      }
      else {
        $ProvisioningReason = ''
      }

      if($Database.IsExcludedFromProvisioning -or $Database.IsExcludedFromProvisioningByOperator) {
        $Output += "<td class='center'>$("{0}<br/>{1}" -f ($ProvisioningStatus -Join ', ') , $ProvisioningReason)</td>"
      }
      else {
        $Output += "<td>&nbsp;</td>"
      }
    }

    # Drive names, issue #4
    if ($ShowDriveNames) {
      $Output += "<td class='center'>$("{0} / {1}" -f $Database.DriveNameEdb, $Database.DriveNameLog)</td>"
    }
    $Output += '</tr>'
  }

  $Output += '</table><br />'

  $Output += '<p class="dagtablefooter">Explanation</p>'
  $Output += ("<p class='dagtablefooter'>Maximum mailbox database size: {0} GB<br/>Minimum free disk space: {1}%</p>" -f $MaxDatabaseSize, $MinFreeDiskspace)
  $Output += ("<p class='dagtablefooter'>NOTE<br/>{0}</p>" -f $EdgeServerNote)

  $StopWatch.Stop()

  if($ShowRunTime) {
    $Output += ("<p class='dagtablefooter'>Run time: {0:0}m {1:0}s</p>" -f $StopWatch.Elapsed.TotalMinutes, ( [math]::Round($StopWatch.Elapsed.TotalSeconds)-([math]::Round($StopWatch.Elapsed.TotalMinutes)*60) ) )
  }

  $Output
}

<#
    .SYNOPSIS
    Renders the HTML document head, embedded stylesheet, and the top level summary table.

    .DESCRIPTION
    Emits the HTML/body opening tags, embeds the CSS file content inline as a style block, and
    renders the header table showing total servers, total mailboxes, and total roles per
    Exchange version detected in the environment.

    .PARAMETER ExchangeEnvironment
    The shared environment hashtable populated earlier in the script run.

    .PARAMETER Path
    Full path to the running script, used to resolve the CSS file location relative to the
    script directory.

    .EXAMPLE
    Get-HtmlReportHeader -ExchangeEnvironment $ExchangeEnvironment -Path $MyInvocation.MyCommand.Path

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-HtmlReportHeader {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [hashtable]$ExchangeEnvironment,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )

  # Labels and stuff
  $LabelTotalServers = 'Total Servers'
  $LabelTotalMailboxes = 'Total Mailboxes'
  $LabelTotalRoles = 'Total Roles'
  $DateLabelFormat = 'yyyy-MM-dd HH:mm'
  $UseCss = $true

  # Header
  $Output = "
    <html>
    <body>
  <title>$($ReportTitle)</title>"

  if ($UseCss -and (Test-Path -Path (Join-Path -Path (Split-Path -Parent $Path) -ChildPath $CssFileName))) {
    $Output += "<style type=""text/css"">$(Get-Content -Path (Join-Path -Path (Split-Path -Parent $Path) -ChildPath $CssFileName))</style>"
  }

  $Output += ("<h2 align=""center"">{5}</h2><h3 align=""center"">Organization: {3}</h3>
      <h4 align=""center"">Generated {0} - V{4}</h4>
      <table class='header'>
      <tr class='header'>
  <th colspan=""{1}"" class='header header-totalservers'>{2}</th>" -f (Get-Date -Format $DateLabelFormat), $ExchangeEnvironment.TotalMailboxesByVersion.Count, $LabelTotalServers, $ExchangeEnvironment.OrganizationName, $ScriptVersion, $ReportTitle)

  if ($ExchangeEnvironment.RemoteMailboxes) {
    $Output += ("<th colspan=""{0}"" class='header header-totalmailboxes'>{1}</th>" -f ($ExchangeEnvironment.TotalMailboxesByVersion.Count + 2), $LabelTotalMailboxes)
  }
  else {
    $Output += ("<th colspan=""{0}"" class='header header-totalmailboxes'>{1}</th>" -f ($ExchangeEnvironment.TotalMailboxesByVersion.Count + 1), $LabelTotalMailboxes)
  }

  $Output += ("<th colspan=""{0}"" class='header header-totalroles'>{1}</th></tr>
  <tr class='subheader'>" -f $ExchangeEnvironment.TotalServersByRole.Count, $LabelTotalRoles)

  # Show Column Headings based on the Exchange versions we have
  # Total Servers
  $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { $Output += "<th class='subheader subheader-totalservers'>$($ExVersionStrings[$_.Key].Short)</th>" }
  # By Exchange Version
  $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { $Output += "<th class='subheader subheader-totalmailboxes'>$($ExVersionStrings[$_.Key].Short)</th>" }

  if ($ExchangeEnvironment.RemoteMailboxes) {
    $Output += "<th class='subheader subheader-totalmailboxes'>Office 365</th>"
  }

  $Output += "<th class='subheader subheader-totalmailboxes'>Org</th>"

  # Exchange Server Roles
  $ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { $Output += "<th class='subheader subheader-totalroles'>$($ExRoleStrings[$_.Key].Short)</th>" }

  $Output += '</tr>'

  $Output += "<tr class='headerdata'>"

  $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { $Output += "<td class='headerdata'>$($_.Value.ServerCount)</td>" }
  $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { $Output += "<td class='headerdata'>$($_.Value.MailboxCount)</td>" }

  if ($RemoteMailboxes) {
    $Output += "<td class='headerdata'>$($ExchangeEnvironment.RemoteMailboxes)</td>"
  }

  $Output += "<td class='headerdata'>$($ExchangeEnvironment.TotalMailboxes)</td>"

  $ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { $Output += "<td class='headerdata'>$($_.Value)</td>" }

  #$Output+="</tr><tr><tr></table><br>"
  $Output += '</tr></table><!-- End --><br />'

  $Output
}

<#
    .SYNOPSIS
    Renders the HTML summary header row for a single Database Availability Group.

    .DESCRIPTION
    Emits a small summary table showing the DAG name, member count, database count, and the
    sorted list of member server names, shown above the DAG's database information table.

    .PARAMETER DAG
    A single DAG information hashtable as returned by Get-DatabaseAvailabilityGroupInformation.

    .EXAMPLE
    Get-HtmlDagHeader -DAG $DAG

    .NOTES
    Author: Thomas Stensitzki
#>
function Get-HtmlDagHeader {
  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [hashtable]$DAG
  )

  # Database Availability Group Header
  $Output += "<table class=""dagsummary"">
    <col width='20%'><col width='10%'><col width='10%'><col width='70%''>
    <tr class=""dagsummary""><th>Database Availability Group Name</th><th>Member Count</th><th># Databases</th>
    <th>Database Availability Group Members</th></tr>
  <tr><td>$($DAG.Name)</td><td>$($DAG.MemberCount)</td><td>$(($DAG.Databases | Measure-Object).Count)</td><td>"

  # 2.7.2 ordered list of member servers
  $DAG.Members | Sort-Object | ForEach-Object { $Output += ('{0} ' -f $_) }

  $Output += '</td></tr></table><br />'

  $Output
}

<#
    .SYNOPSIS
    Updates the PowerShell progress bar for the overall script run.

    .DESCRIPTION
    Maps a percent-complete value within a given processing stage onto the overall 5 stage
    script run, so the displayed progress bar advances smoothly across the whole run instead
    of resetting to 0% at the start of every stage.

    .PARAMETER PercentComplete
    Percent complete (0-100) within the current stage.

    .PARAMETER Status
    Status text shown next to the progress bar.

    .PARAMETER Stage
    The current overall processing stage (1-5).

    .EXAMPLE
    Show-ProgressBar -PercentComplete 50 -Status 'Getting Databases' -Stage 1

    .NOTES
    Author: Thomas Stensitzki
#>
function Show-ProgressBar {
  [CmdletBinding()]
  [OutputType([void])]
  param(
    [Parameter(Mandatory = $true)]
    [int]$PercentComplete,
    [Parameter(Mandatory = $true)]
    [string]$Status,
    [Parameter(Mandatory = $true)]
    [int]$Stage
  )

  $TotalStages = 5
  Write-Progress -Id 1 -Activity 'Get-ExchangeEnvironmentReport' -Status $Status -PercentComplete (($PercentComplete / $TotalStages) + (1 / $TotalStages * $Stage * 100))
}

# 1. Initial Startup

# 1.0 Check Powershell Version
if ((Get-Host).Version.Major -eq 1) {
  throw 'Powershell Version 1 not supported'
}

# 1.1 Check Exchange Management Shell, attempt to load
if (!(Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue)) {
  # 2019-05-17 Thomas Stensitzki, Support for Exchange Scripts located in non-default locations
  # Use $env:ExchangeInstallPath for Exchange 2010/2013+ installations
  $ExchangeInstallPath = $env:ExchangeInstallPath

  if (($ExchangeInstallPath -eq '') -or ($null -eq $ExchangeInstallPath)) {
    # $env:ExchangeInstallPath not available on Exchange Server 2007 Setups
    try {
      $ExchangeInstallPath = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Exchange\Setup).MsiInstallPath
    }
    catch {}
  }

  Write-Verbose -Message ('Exchange Install Path: {0}' -f $ExchangeInstallPath)

  $RemoteExchangePath = Join-Path -Path $ExchangeInstallPath -ChildPath 'bin\RemoteExchange.ps1'
  $LocalExchangePath = Join-Path -Path $ExchangeInstallPath -ChildPath 'bin\Exchange.ps1'

  if (Test-Path -Path $RemoteExchangePath) {
    . $RemoteExchangePath
    Connect-ExchangeServer -auto
  }
  elseif (Test-Path -Path $LocalExchangePath) {
    Add-PSSnapIn -Name Microsoft.Exchange.Management.PowerShell.Admin
    . $LocalExchangePath
  }
  else {
    throw 'Exchange Management Shell cannot be loaded'
  }
}

# 1.1.1 Check if CSS file is present
# Issue #6
# CSS is only required when an HTML report is generated
if ($OutputFormat -ne 'Markdown') {
  if (Test-Path -Path (Join-Path -Path $ScriptDir -ChildPath $CssFileName)) {
    Write-Verbose ('Using {0} as CSS file for HTML report.' -f $CssFileName )
  }
  else {
    throw ('CSS file {0} is missing. It is required for a proper HTML report. Please see the GitHub repository for more information.' -f $CssFileName)
  }
}

# 1.1.2 Check if the Exchange version mapping JSON file is present
$VersionMappingFilePath = Join-Path -Path $ScriptDir -ChildPath $VersionMappingFileName

if (Test-Path -Path $VersionMappingFilePath) {
  Write-Verbose -Message ('Using {0} as Exchange version mapping file.' -f $VersionMappingFileName)
}
else {
  throw ('Version mapping file {0} is missing. It is required to translate Exchange build numbers into readable labels. Please see the GitHub repository for more information.' -f $VersionMappingFileName)
}

# 1.2 Check if -SendMail parameter set and if so check -MailFrom, -MailTo and -MailServer are set
if ($SendMail) {
  if (!$MailFrom -or !$MailTo -or !$MailServer) {
    throw 'If -SendMail specified, you must also specify -MailFrom, -MailTo and -MailServer'
  }
}

# 1.3 Check Exchange Management Shell Version
if ((Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue)) {
  $E2010 = $false;
  if (Get-ExchangeServer | Where-Object { $_.AdminDisplayVersion.Major -gt 14 }) {
    Write-Warning -Message "Exchange 2010 or higher detected. You'll get better results if you run this script from the latest management shell"
  }
}
else {

  $E2010 = $true

  # 2019-05-17 Thomas Stensitzki, Support for Exchange 2013+ servers with installed management tools
  $localversion = $localserver = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiProductMajor

  if ($localversion -eq 15) { $E2013 = $true }
}

# 1.4 Check view entire forest if set (by default, true)
if ($E2010) {
  Set-ADServerSettings -ViewEntireForest:$ViewEntireForest
}
else {
  $global:AdminSessionADSettings.ViewEntireForest = $ViewEntireForest
}

# 1.5 Initial Variables

# 1.5.1 Hashtable to update with environment data
$ExchangeEnvironment = @{
  Sites            = @{}
  Pre2007          = @{}
  Servers          = @{}
  DAGs             = @()
  NonDAGDatabases  = @()
  OrganizationName = ''
}

# 1.5.7 Exchange Major Version, Service Pack, and Security Update String Mapping
# Loaded from an external JSON file to simplify updates whenever Microsoft releases a new
# Cumulative Update or Security Update, no script code changes are required for that case.
$VersionMapping = Import-ExchangeVersionMapping -Path $VersionMappingFilePath

$ExMajorVersionStrings = $VersionMapping.MajorVersions
$ExSPLevelStrings = $VersionMapping.ServicePackLevels
$ExSUString = $VersionMapping.SecurityUpdates

# Add CUx entries up to the configured maximum, this range rarely changes and therefore stays code driven
for ($i = 1; $i -le $VersionMapping.MaxCumulativeUpdate; $i++) {
  $ExSPLevelStrings.Add("CU$($i)", "CU$($i)")
}

# 1.5.9 Populate Full Mapping using above info
$ExVersionStrings = @{}

foreach ($Major in $ExMajorVersionStrings.GetEnumerator()) {
  foreach ($Minor in $ExSPLevelStrings.GetEnumerator()) {
    $ExVersionStrings.Add("$($Major.Key).$($Minor.Key)", @{Long = "$($Major.Value.Long) $($Minor.Value)"; Short = "$($Major.Value.Short)$($Minor.Value)" })
  }
}

Write-Verbose $ExMajorVersionStrings

# 1.5.10 Exchange Role String Mapping
$ExRoleStrings = @{
  'ClusteredMailbox'                  = @{Short = 'ClusMBX'; Long = 'CCR/SCC Clustered Mailbox' }
  'Mailbox'                           = @{Short = 'MBX'; Long = 'Mailbox' }
  'ClientAccess'                      = @{Short = 'CAS'; Long = 'Client Access' }
  'HubTransport'                      = @{Short = 'HUB'; Long = 'Hub Transport' }
  'UnifiedMessaging'                  = @{Short = 'UM'; Long = 'Unified Messaging' }
  'Edge'                              = @{Short = 'EDGE'; Long = 'Edge Transport' }
  'FE'                                = @{Short = 'FE'; Long = 'Frontend' }
  'BE'                                = @{Short = 'BE'; Long = 'Backend' }
  'Hybrid'                            = @{Short = 'HYB'; Long = 'Hybrid' }
  'Coexistence'                       = @{Short = 'COEX'; Long = 'Coexistence' } #2019-05-17 TST Coexistence added
  'Unknown'                           = @{Short = 'Unknown'; Long = 'Unknown' }
}

# 2 Get Relevant Exchange Information Up-Front

# 2.1 Get Server, Exchange and Mailbox Information
Show-ProgressBar -PercentComplete 1 -Status 'Getting Exchange Server List' -Stage 1

try {
  $ExchangeServers = [array](Get-ExchangeServer $ServerFilter -ErrorAction Stop | Sort-Object Name)
}
catch {
  throw ('Failed to query Exchange servers using -ServerFilter {0}: {1}' -f $ServerFilter, $_.Exception.Message)
}

if (!$ExchangeServers) {
  throw ('No Exchange Servers matched by -ServerFilter {0}' -f $ServerFilter)
}

$HybridServers = @()
if (Get-Command -Name Get-HybridConfiguration -ErrorAction SilentlyContinue) {
  try {
    $HybridConfig = Get-HybridConfiguration -ErrorAction Stop
    $HybridConfig.ReceivingTransportServers | ForEach-Object { $HybridServers += $_.Name }
    $HybridConfig.SendingTransportServers | ForEach-Object { $HybridServers += $_.Name }
    $HybridServers = $HybridServers | Sort-Object -Unique
  }
  catch {
    Write-Warning -Message ('Cannot query hybrid configuration, continuing without hybrid server detection: {0}' -f $_.Exception.Message)
  }
}

Show-ProgressBar -PercentComplete 10 -Status 'Getting Mailboxes' -Stage 1

try {
  $Mailboxes = [array](Get-Mailbox -ResultSize Unlimited -ErrorAction Stop) | Where-Object { $_.ServerName -like $ServerFilter }
}
catch {
  throw ('Failed to query mailboxes: {0}' -f $_.Exception.Message)
}

if ($E2010) {

  Show-ProgressBar -PercentComplete 60 -Status 'Getting Archive Mailboxes' -Stage 1

  try {
    $ArchiveMailboxes = [array](Get-Mailbox -Archive -ResultSize Unlimited -ErrorAction Stop) | Where-Object { $_.ServerName -like $ServerFilter }
  }
  catch {
    Write-Warning -Message ('Cannot query archive mailboxes, continuing without archive mailbox data: {0}' -f $_.Exception.Message)
    $ArchiveMailboxes = @()
  }

  Show-ProgressBar -PercentComplete 70 -Status 'Getting Remote Mailboxes' -Stage 1

  try {
    $RemoteMailboxes = [array](Get-RemoteMailbox -ResultSize Unlimited -ErrorAction Stop)
  }
  catch {
    Write-Warning -Message ('Cannot query remote mailboxes, continuing without remote mailbox data: {0}' -f $_.Exception.Message)
    $RemoteMailboxes = @()
  }
  $ExchangeEnvironment.Add('RemoteMailboxes', $RemoteMailboxes.Count)

  Show-ProgressBar -PercentComplete 90 -Status 'Getting Databases' -Stage 1

  try {
    if ($E2013) {
      # 2019-05-17 TST Sorting added
      $Databases = [array](Get-MailboxDatabase -IncludePreExchange2013 -Status -ErrorAction Stop) | Sort-Object -Property Name | Where-Object { $_.Server -like $ServerFilter }
    }
    elseif ($E2010) {
      # 2019-05-17 TST Sorting added
      $Databases = [array](Get-MailboxDatabase -IncludePreExchange2010 -Status -ErrorAction Stop) | Sort-Object -Property Name | Where-Object { $_.Server -like $ServerFilter }
    }
  }
  catch {
    throw ('Failed to query mailbox databases: {0}' -f $_.Exception.Message)
  }

  try {
    $DAGs = [array](Get-DatabaseAvailabilityGroup -ErrorAction Stop) | Where-Object { $_.Servers -like $ServerFilter }
  }
  catch {
    Write-Warning -Message ('Cannot query Database Availability Groups, continuing without DAG data: {0}' -f $_.Exception.Message)
    $DAGs = @()
  }
}
else {
  $ArchiveMailboxes = $null
  $ArchiveMailboxStats = $null
  $DAGs = $null

  Show-ProgressBar -PercentComplete 90 -Status 'Getting Databases' -Stage 1
  try {
    $Databases = [array](Get-MailboxDatabase -IncludePreExchange2007 -Status -ErrorAction Stop) | Where-Object { $_.Server -like $ServerFilter }
  }
  catch {
    throw ('Failed to query mailbox databases: {0}' -f $_.Exception.Message)
  }
  $ExchangeEnvironment.Add('RemoteMailboxes', 0)
}

# 2.3 Populate Information we know
$ExchangeEnvironment.Add('TotalMailboxes', $Mailboxes.Count + $ExchangeEnvironment.RemoteMailboxes)

# 2.4 Organizational Info

try {
  $ExchangeEnvironment.OrganizationName = (Get-OrganizationConfig -ErrorAction Stop).Name
}
catch {
  Write-Warning -Message ('Cannot query organization configuration, organization name will be blank: {0}' -f $_.Exception.Message)
  $ExchangeEnvironment.OrganizationName = $NotAvailable
}

# 3 Process High-Level Exchange Information

# 3.1 Collect Exchange Server Information
for ($i = 0; $i -lt $ExchangeServers.Count; $i++) {
  Show-ProgressBar -PercentComplete ($i / $ExchangeServers.Count * 100) -Status 'Getting Exchange Server Information' -Stage 2

  # Get Exchange Info
  try {
    $ExSvr = Get-ExchangeServerInformation -E2010 $E2010 -ExchangeServer $ExchangeServers[$i] -Mailboxes $Mailboxes -Databases $Databases -Hybrids $HybridServers
  }
  catch {
    Write-Warning -Message ('Skipping server {0}, failed to collect Exchange server information: {1}' -f $ExchangeServers[$i].Name, $_.Exception.Message)
    continue
  }

  # Add to site or pre-Exchange 2007 list
  if ($ExSvr.Site) {
    # Exchange 2007 or higher
    if (!$ExchangeEnvironment.Sites[$ExSvr.Site]) {
      $ExchangeEnvironment.Sites.Add($ExSvr.Site, @($ExSvr))
    }
    else {
      $ExchangeEnvironment.Sites[$ExSvr.Site] += $ExSvr
    }
  }
  else {
    # Exchange 2003 or lower
    if (!$ExchangeEnvironment.Pre2007['Pre 2007 Servers']) {
      $ExchangeEnvironment.Pre2007.Add('Pre 2007 Servers', @($ExSvr))
    }
    else {
      $ExchangeEnvironment.Pre2007['Pre 2007 Servers'] += $ExSvr
    }
  }

  # Add to Servers List
  $ExchangeEnvironment.Servers.Add($ExSvr.Name, $ExSvr)
}

# 3.2 Calculate Environment Totals for Version/Role using collected data
Show-ProgressBar -PercentComplete 1 -Status 'Getting Totals' -Stage 3

$ExchangeEnvironment.Add('TotalMailboxesByVersion', (Get-TotalsByVersion -ExchangeEnvironment $ExchangeEnvironment))
$ExchangeEnvironment.Add('TotalServersByRole', (Get-TotalsByRole -ExchangeEnvironment $ExchangeEnvironment))

# 3.4 Populate Environment DAGs
Show-ProgressBar -PercentComplete 5 -Status 'Getting DAG Info' -Stage 3

if ($DAGs) {
  foreach ($DAG in $DAGs) {
    $ExchangeEnvironment.DAGs += (Get-DatabaseAvailabilityGroupInformation -DAG $DAG)
  }
}

# 3.5 Get Database information
Show-ProgressBar -PercentComplete 60 -Status 'Getting Database Info' -Stage 3

for ($i = 0; $i -lt $Databases.Count; $i++) {
  try {
    $Database = Get-DatabaseInformation -Database $Databases[$i] -ExchangeEnvironment $ExchangeEnvironment -Mailboxes $Mailboxes -ArchiveMailboxes $ArchiveMailboxes -E2010 $E2010
  }
  catch {
    Write-Warning -Message ('Skipping database {0}, failed to collect database information: {1}' -f $Databases[$i].Name, $_.Exception.Message)
    continue
  }
  $DAGDB = $false
  for ($j = 0; $j -lt $ExchangeEnvironment.DAGs.Count; $j++) {
    if ($ExchangeEnvironment.DAGs[$j].Members -contains $Database.ActiveOwner) {
      $DAGDB = $true
      $ExchangeEnvironment.DAGs[$j].Databases += $Database
    }
  }
  if (!$DAGDB) {
    $ExchangeEnvironment.NonDAGDatabases += $Database
  }
}

# 4 Write Information
Show-ProgressBar -PercentComplete 5 -Status 'Writing HTML Report Header' -Stage 4

$Output = Get-HtmlReportHeader -ExchangeEnvironment $ExchangeEnvironment -Path $MyInvocation.MyCommand.Path

# Sites and Servers
Show-ProgressBar -PercentComplete 20 -Status 'Writing HTML Site Information' -Stage 4

foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator()) {
  $Output += Get-HtmlOverview -Servers $Site -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings
}

Show-ProgressBar -PercentComplete 40 -Status 'Writing HTML Pre-2007 Information' -Stage 4

foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator()) {
  $Output += Get-HtmlOverview -Servers $FakeSite -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings -Pre2007:$true
}

Show-ProgressBar -PercentComplete 60 -Status 'Writing HTML DAG Information' -Stage 4

foreach ($DAG in $ExchangeEnvironment.DAGs) {

  if ($DAG.MemberCount -gt 0) {

    # Get DAG Header
    $Output += Get-HtmlDagHeader -DAG $DAG

    # Get Table HTML for DAG databases
    $Output += Get-HtmlDatabaseInformationTable -Databases $DAG.Databases
  }
}

if ($ExchangeEnvironment.NonDAGDatabases.Count) {

  Show-ProgressBar -PercentComplete 80 -Status 'Writing HTML Non-DAG Database Information' -Stage 4

  $Output += '<table class="dagsummary">
  <tr class="dagsummarynondag"><th>Mailbox Databases (Non-DAG)</th></table>'

  # Get Table HTML for non-DAG databases
  $Output += Get-HtmlDatabaseInformationTable -Databases $ExchangeEnvironment.NonDAGDatabases
}

# End
Show-ProgressBar -PercentComplete 90 -Status 'Finishing off..' -Stage 4

$Output += '</body></html>'

# 2019-05-20 TST Updated to ensure script path as storage location
$HtmlReportFullPath = Join-Path -Path (Split-Path -Path $script:MyInvocation.MyCommand.Path) -ChildPath $HTMLReport
$MarkdownReportFullPath = Join-Path -Path (Split-Path -Path $script:MyInvocation.MyCommand.Path) -ChildPath $MarkdownReport

if ($OutputFormat -in 'HTML', 'Both') {
  try {
    $Output | Out-File -FilePath $HtmlReportFullPath -Force -Encoding utf8 -ErrorAction Stop
  }
  catch {
    throw ('Failed to write HTML report to {0}: {1}' -f $HtmlReportFullPath, $_.Exception.Message)
  }

  if ($OpenInBrowser) {
    try {
      Start-Process -FilePath $HtmlReportFullPath -ErrorAction Stop
    }
    catch {
      throw ('Failed to open report in the default browser: {0}' -f $_.Exception.Message)
    }
  }
}

if ($OutputFormat -in 'Markdown', 'Both') {
  try {
    ConvertTo-MarkdownReport -Html $Output | Out-File -FilePath $MarkdownReportFullPath -Force -Encoding utf8 -ErrorAction Stop
  }
  catch {
    throw ('Failed to write Markdown report to {0}: {1}' -f $MarkdownReportFullPath, $_.Exception.Message)
  }
}

# 5 Send Mail
if ($SendMail) {
  Show-ProgressBar -PercentComplete 95 -Status 'Sending mail message...' -Stage 4

  # 2019-05-17 TST, Changed to .NET send method to work as scheduled job

  try {
    $smtpMail = New-Object Net.Mail.SmtpClient($MailServer)

    $smtpMessage = New-Object System.Net.Mail.MailMessage $MailFrom, $MailTo

    if ($OutputFormat -eq 'Markdown') {
      # Markdown only: attach the Markdown file and send a generic plain text body
      if (Test-Path -Path $MarkdownReportFullPath) {
        $smtpAttachment = New-Object Net.Mail.Attachment($MarkdownReportFullPath, 'text/markdown')
        $smtpMessage.Attachments.Add($smtpAttachment)
      }

      $smtpMessage.Subject = $ReportTitle
      $smtpMessage.Body = ('The {0} has been generated and is attached as a Markdown file.' -f $ReportTitle)
      $smtpMessage.IsBodyHtml = $false
    }
    else {
      if (Test-Path -Path $HtmlReportFullPath) {
        $smtpAttachment = New-Object Net.Mail.Attachment($HtmlReportFullPath, 'text/plain')
        $smtpMessage.Attachments.Add($smtpAttachment)
      }

      if ($OutputFormat -eq 'Both' -and (Test-Path -Path $MarkdownReportFullPath)) {
        $smtpAttachment = New-Object Net.Mail.Attachment($MarkdownReportFullPath, 'text/markdown')
        $smtpMessage.Attachments.Add($smtpAttachment)
      }

      $smtpMessage.Subject = $ReportTitle
      $smtpMessage.Body = $Output
      $smtpMessage.IsBodyHtml = $true
    }

    $smtpMail.Send($smtpMessage)
  }
  catch {
    throw ('Failed to send report by mail via {0}: {1}' -f $MailServer, $_.Exception.Message)
  }

  Return 0
}