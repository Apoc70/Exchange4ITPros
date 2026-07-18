#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Sends password expiration reminder emails to Active Directory users based on configurable
    notification thresholds.

.DESCRIPTION
    This script queries on-premises Active Directory for enabled user accounts whose password
    is about to expire. It uses the constructed attribute 'msDS-UserPasswordExpiryTimeComputed'
    to determine the effective expiration date, which automatically accounts for Fine Grained
    Password Policies (FGPP) as well as the default domain password policy.

    Only user accounts that have either a local Exchange mailbox or an Exchange Online
    mailbox are considered. This is determined via the 'msExchRecipientTypeDetails'
    attribute (UserMailbox for on-premises mailboxes, RemoteUserMailbox for hybrid/Exchange
    Online mailboxes, by default). Disabled accounts and accounts without a matching
    recipient type are ignored.

    For every user whose remaining days until password expiration matches one of the configured
    thresholds (e.g. 14, 7, 5, 3, 2, 1 days), a notification email is sent using either a plain
    text or an HTML template. All behavior, including thresholds, sender information, SMTP server
    settings, recipient address resolution, and template selection, is controlled through an
    external JSON configuration file.

    A local tracking file is used to prevent duplicate notifications if the script is executed
    more than once on the same day.

    The script is intended to be executed once per day via Windows Task Scheduler on an
    Exchange Server (or any domain-joined member server with the ActiveDirectory PowerShell
    module installed).

.PARAMETER ConfigurationFilePath
    Path to the JSON configuration file. Defaults to 'PasswordExpirationReminder.config.json'
    in the same directory as the script.

.PARAMETER LogFilePath
    Optional path to a script execution log file. If not specified, the path defined in the
    configuration file (Logging.ScriptLogPath) is used, if present.

.PARAMETER ReportOnly
    Runs the script in an interactive, read-only verification mode. Active Directory is
    queried using the same filters and thresholds as a normal run, and the resulting list of
    users (SamAccountName, DisplayName, resolved recipient address, days remaining, and
    expiration date) is written to the console as a table. No email is sent and the
    notification tracking file is neither read nor updated. Intended for an administrator to
    interactively verify that the correct users are being matched before relying on the
    scheduled task.

.PARAMETER IncludeExpiredPasswordAccounts
    When used together with -ReportOnly, includes users whose passwords have already expired
    in the Active Directory query and writes those accounts to the log file only. Expired
    accounts are not added to the console report.

.PARAMETER TestEmail
    Sends a test email to the recipient specified by -TestEmailRecipient. The test mode uses
    the configured SMTP settings and sends one message rendered from the HTML template and one
    message rendered from the text template. Active Directory is not queried and the
    notification tracking file is not read or updated.

.PARAMETER TestEmailRecipient
    The recipient email address used when -TestEmail is specified.

.PARAMETER UseTextTemplate
    Sends notifications using the text template instead of the HTML template.
    Html is used by default when this switch is not specified.

.EXAMPLE
    .\Send-PasswordExpirationReminder.ps1

    Runs the script using the default configuration file located next to the script.

.EXAMPLE
    .\Send-PasswordExpirationReminder.ps1 -ConfigurationFilePath 'D:\Scripts\Config\ContosoConfig.json' -Verbose

    Runs the script using a specific configuration file and shows verbose progress information.

.EXAMPLE
    .\Send-PasswordExpirationReminder.ps1 -WhatIf

    Performs a dry run. Active Directory is queried and matching users are identified and logged,
    but no emails are actually sent.

.EXAMPLE
    .\Send-PasswordExpirationReminder.ps1 -ReportOnly

    Interactively lists all users currently matching the configured thresholds and mailbox
    filter, without sending any email and without touching the notification tracking file.

    .EXAMPLE
    .\Send-PasswordExpirationReminder.ps1 -ReportOnly -IncludeExpiredPasswordAccounts

    Interactively lists all users currently matching the configured thresholds and mailbox
    filter, and writes already-expired accounts to the log file only.

    .EXAMPLE
        .\Send-PasswordExpirationReminder.ps1 -TestEmail -TestEmailRecipient 'admin@contoso.com'

        Sends two test emails to the specified recipient, one using the HTML template and one
        using the text template, without querying Active Directory or updating the tracking file.

.NOTES
    Author: Thomas Stensitzki
    Target platform: Exchange Server (on-premises), any domain-joined member server
    Required modules: ActiveDirectory (RSAT-AD-PowerShell)
    PowerShell compatibility: Windows PowerShell 5.1

    Required permissions for the executing account:
    - Read access to Active Directory user objects, including the constructed attribute
      'msDS-UserPasswordExpiryTimeComputed'. This attribute has been readable by
      'Authenticated Users' by default since the Windows Server 2008 R2 domain functional
      level. If the domain functional level is lower, or if the default ACL has been
      hardened, explicit read permission on this attribute may need to be granted via
      dsacls.exe or ADSI Edit.
    - No write access to Active Directory is required; the script is read-only with
      respect to AD.
    - Read access to the 'msExchRecipientTypeDetails' attribute, used to filter for users
      with a local Exchange mailbox or an Exchange Online (hybrid) mailbox. This attribute
      is only populated once the Exchange schema extensions are present in the forest and
      requires no additional permissions beyond standard AD read access.
    - Network access to the configured SMTP server/relay on the configured port. If the
      SMTP server does not allow anonymous relay from this server's IP address, SMTP
      credentials must be supplied via the configuration.
    - Local read/write access to the configured log file, notification tracking file,
      and template files.
    - If executed as a scheduled task under a dedicated service account (rather than as
      the SYSTEM account), that account requires the 'Log on as a batch job'
      (SeBatchLogonRight) user right on the local server.

    Common 'msExchRecipientTypeDetails' values for reference when adjusting
    ActiveDirectory.MailboxRecipientTypeDetails in the configuration file:
    1              UserMailbox (on-premises)
    4              SharedMailbox (on-premises)
    16             RoomMailbox (on-premises)
    32             EquipmentMailbox (on-premises)
    2147483648     RemoteUserMailbox (hybrid, mailbox hosted in Exchange Online)
    137438953472   RemoteSharedMailbox (hybrid)

    Change log:
    1.0.0 - Initial version
    1.1.0 - Added ReportOnly switch for interactive, read-only verification.
            Added mailbox recipient type filtering (on-premises and Exchange Online
            mailboxes only) via msExchRecipientTypeDetails.
    1.2.0 - Added IncludeExpiredPasswordAccounts switch for ReportOnly mode.
    1.3.0 - Added TestEmail switch for SMTP and template validation.
    1.4.0 - Added UseTextTemplate switch to select the email template type.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationFilePath = (Join-Path -Path $PSScriptRoot -ChildPath 'PasswordExpirationReminder.config.json'),

    [Parameter(Mandatory = $false)]
    [string]$LogFilePath,

    [Parameter(Mandatory = $false)]
    [switch]$ReportOnly,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeExpiredPasswordAccounts,

    [Parameter(Mandatory = $false)]
    [switch]$TestEmail,

    [Parameter(Mandatory = $false)]
    [switch]$UseTextTemplate,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TestEmailRecipient
)

$script:LogFilePath = $LogFilePath

function Write-ScriptLog {
    <#
    .SYNOPSIS
        Writes a log message to Verbose/Warning/Error streams and optionally to a log file.

    .PARAMETER Message
        The message text to log.

    .PARAMETER Level
        The severity level of the message. Valid values: Info, Warning, Error.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $logLine = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'Warning' { Write-Warning -Message $Message }
        'Error'   { Write-Error -Message $Message -ErrorAction Continue }
        default   { Write-Verbose -Message $Message }
    }

    if (-not [string]::IsNullOrWhiteSpace($script:LogFilePath)) {
        try {
            $logDirectory = Split-Path -Path $script:LogFilePath -Parent
            if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -Path $logDirectory)) {
                New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
            }
            Add-Content -Path $script:LogFilePath -Value $logLine -Encoding UTF8
        }
        catch {
            Write-Warning -Message "Failed to write to log file '$($script:LogFilePath)': $($_.Exception.Message)"
        }
    }
}

function Write-LogFileEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    if ([string]::IsNullOrWhiteSpace($script:LogFilePath)) {
        return
    }

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $logLine = "[$timestamp] [$Level] $Message"

    try {
        $logDirectory = Split-Path -Path $script:LogFilePath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -Path $logDirectory)) {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        }

        Add-Content -Path $script:LogFilePath -Value $logLine -Encoding UTF8
    }
    catch {
    }
}

function Get-ScriptConfiguration {
    <#
    .SYNOPSIS
        Loads and validates the JSON configuration file.

    .PARAMETER ConfigurationFilePath
        Path to the JSON configuration file.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigurationFilePath
    )

    if (-not (Test-Path -Path $ConfigurationFilePath)) {
        throw "Configuration file not found at path '$ConfigurationFilePath'."
    }

    try {
        $configuration = Get-Content -Path $ConfigurationFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "Failed to parse configuration file '$ConfigurationFilePath': $($_.Exception.Message)"
    }

    $requiredProperties = @('Thresholds', 'Sender', 'Smtp', 'Recipient', 'Templates', 'Logging', 'Subject')

    foreach ($property in $requiredProperties) {
        if (-not ($configuration.PSObject.Properties.Name -contains $property)) {
            throw "Configuration file is missing required top level property '$property'."
        }
    }

    if (-not $configuration.Thresholds -or $configuration.Thresholds.Count -eq 0) {
        throw "Configuration property 'Thresholds' must contain at least one value."
    }

    return $configuration
}

function Get-ExpiringPasswordUser {
    <#
    .SYNOPSIS
        Queries Active Directory for enabled users whose password expires within the
        configured notification thresholds.

    .PARAMETER SearchBase
        Optional distinguished name to limit the search to a specific OU.

    .PARAMETER Thresholds
        Array of integers representing the number of days before expiration at which a
        notification should be sent.

    .PARAMETER MailboxRecipientTypeDetails
        Array of 'msExchRecipientTypeDetails' values that qualify as having a mailbox.
        Defaults to 1 (UserMailbox, on-premises) and 2147483648 (RemoteUserMailbox,
        hybrid/Exchange Online). Users without a matching recipient type are excluded.

    .PARAMETER IncludeExpiredPasswords
        When set, users whose passwords have already expired are included in the output.
    #>
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SearchBase,

        [Parameter(Mandatory = $true)]
        [int[]]$Thresholds,

        [Parameter(Mandatory = $false)]
        [int64[]]$MailboxRecipientTypeDetails = @(1, 2147483648),

        [Parameter(Mandatory = $false)]
        [switch]$IncludeExpiredPasswords
    )

    Write-ScriptLog -Message 'Querying Active Directory for users with upcoming password expiration.' -Level 'Info'

    # LDAP filter excludes disabled accounts (userAccountControl bit 2) and accounts with
    # the PASSWORD_NEVER_EXPIRES flag set (userAccountControl bit 65536). Filtering by
    # recipient type is done afterwards in PowerShell, since msExchRecipientTypeDetails is
    # a bitmask style attribute that is impractical to match reliably via LDAP filter alone.
    $ldapFilter = '(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(userAccountControl:1.2.840.113556.1.4.803:=65536)))'

    $adUserParameters = @{
        LDAPFilter  = $ldapFilter
        Properties  = @('DisplayName', 'mail', 'proxyAddresses', 'PasswordNeverExpires', 'msDS-UserPasswordExpiryTimeComputed', 'msExchRecipientTypeDetails', 'SamAccountName')
        ErrorAction = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
        $adUserParameters['SearchBase'] = $SearchBase
    }

    try {
        $adUsers = Get-ADUser @adUserParameters
    }
    catch {
        Write-ScriptLog -Message "Failed to query Active Directory: $($_.Exception.Message)" -Level 'Error'
        throw
    }

    $today = (Get-Date).Date
    $results = New-Object -TypeName System.Collections.Generic.List[PSObject]

    foreach ($user in $adUsers) {

        if ($user.PasswordNeverExpires) {
            continue
        }

        $recipientTypeDetails = $user.msExchRecipientTypeDetails

        # Skip users without a mailbox (e.g. plain AD accounts, mail contacts, mail users)
        # or without any Exchange attributes at all.
        if ($null -eq $recipientTypeDetails -or $recipientTypeDetails -notin $MailboxRecipientTypeDetails) {
            continue
        }

        $expiryFileTime = $user.'msDS-UserPasswordExpiryTimeComputed'

        # A value of $null, 0, or Int64.MaxValue indicates that no computed expiration
        # exists (e.g. password never expires), even though PasswordNeverExpires was
        # already excluded via the LDAP filter above. This is an additional safety check.
        if ($null -eq $expiryFileTime -or $expiryFileTime -eq 0 -or $expiryFileTime -eq [int64]::MaxValue) {
            continue
        }

        $expiryDate = [datetime]::FromFileTime($expiryFileTime)
        $daysRemaining = ($expiryDate.Date - $today).Days

        if ($daysRemaining -lt 0 -and -not $IncludeExpiredPasswords) {
            continue
        }

        if ($daysRemaining -in $Thresholds -or ($IncludeExpiredPasswords -and $daysRemaining -lt 0)) {
            $results.Add([PSCustomObject]@{
                SamAccountName = $user.SamAccountName
                DisplayName    = $user.DisplayName
                Mail           = $user.mail
                ProxyAddresses = $user.proxyAddresses
                ExpiryDate     = $expiryDate
                DaysRemaining  = $daysRemaining
                PasswordStatus = if ($daysRemaining -lt 0) { 'Expired' } else { 'Expiring' }
            })
        }
    }

    $thresholdList = ($Thresholds | Sort-Object -Unique) -join ', '
    $mailboxRecipientTypeList = if ($MailboxRecipientTypeDetails) { ($MailboxRecipientTypeDetails | Sort-Object -Unique) -join ', ' } else { '<not configured>' }
    Write-ScriptLog -Message "Report criteria: Thresholds=[$thresholdList]; MailboxRecipientTypeDetails=[$mailboxRecipientTypeList]; IncludeExpiredPasswords=$IncludeExpiredPasswords." -Level 'Info'
    Write-ScriptLog -Message "Found $($results.Count) user(s) matching configured report criteria." -Level 'Info'

    return $results
}

function Get-PrimarySmtpAddress {
    <#
    .SYNOPSIS
        Resolves the recipient email address for a user based on the configured address source.

    .PARAMETER User
        The user object as returned by Get-ExpiringPasswordUser.

    .PARAMETER AddressSource
        Either 'PrimarySmtpAddress' (derived from proxyAddresses) or 'WindowsMailAddress'
        (the 'mail' attribute directly).

    .PARAMETER FallbackToWindowsMailAttribute
        If $true and AddressSource is 'PrimarySmtpAddress' but no primary SMTP proxy
        address is found, the 'mail' attribute is used as a fallback.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$User,

        [Parameter(Mandatory = $true)]
        [ValidateSet('PrimarySmtpAddress', 'WindowsMailAddress')]
        [string]$AddressSource,

        [Parameter(Mandatory = $false)]
        [bool]$FallbackToWindowsMailAttribute = $true
    )

    $address = $null

    if ($AddressSource -eq 'PrimarySmtpAddress') {
        # Primary SMTP proxy address is prefixed with uppercase 'SMTP:' by convention.
        $primaryProxy = $User.ProxyAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1

        if ($primaryProxy) {
            $address = $primaryProxy -replace '^SMTP:', ''
        }
        elseif ($FallbackToWindowsMailAttribute) {
            $address = $User.Mail
        }
    }
    else {
        $address = $User.Mail
    }

    return $address
}

function Test-NotificationAlreadySent {
    <#
    .SYNOPSIS
        Checks whether a notification for a given user, threshold, and date has already
        been recorded in the tracking file.

    .PARAMETER TrackingFilePath
        Path to the JSON tracking file.

    .PARAMETER SamAccountName
        SamAccountName of the user.

    .PARAMETER DaysRemaining
        The matched threshold value.

    .PARAMETER Date
        The date to check against.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TrackingFilePath,

        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,

        [Parameter(Mandatory = $true)]
        [int]$DaysRemaining,

        [Parameter(Mandatory = $true)]
        [datetime]$Date
    )

    if (-not (Test-Path -Path $TrackingFilePath)) {
        return $false
    }

    try {
        $entries = @(Get-Content -Path $TrackingFilePath -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    catch {
        Write-ScriptLog -Message "Could not parse notification tracking file '$TrackingFilePath'. Treating as no prior notifications." -Level 'Warning'
        return $false
    }

    $dateString = $Date.ToString('yyyy-MM-dd')

    $match = $entries | Where-Object {
        $_.SamAccountName -eq $SamAccountName -and
        $_.DaysRemaining -eq $DaysRemaining -and
        $_.Date -eq $dateString
    }

    return [bool]$match
}

function Add-NotificationLogEntry {
    <#
    .SYNOPSIS
        Records a sent notification in the tracking file and prunes entries older than
        60 days.

    .PARAMETER TrackingFilePath
        Path to the JSON tracking file.

    .PARAMETER SamAccountName
        SamAccountName of the user the notification was sent to.

    .PARAMETER DaysRemaining
        The matched threshold value.

    .PARAMETER Date
        The date the notification was sent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TrackingFilePath,

        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,

        [Parameter(Mandatory = $true)]
        [int]$DaysRemaining,

        [Parameter(Mandatory = $true)]
        [datetime]$Date
    )

    $entries = @()

    if (Test-Path -Path $TrackingFilePath) {
        try {
            $entries = @(Get-Content -Path $TrackingFilePath -Raw -Encoding UTF8 | ConvertFrom-Json)
        }
        catch {
            Write-ScriptLog -Message "Could not parse existing notification tracking file '$TrackingFilePath'. Starting a new one." -Level 'Warning'
            $entries = @()
        }
    }

    $entries += [PSCustomObject]@{
        SamAccountName = $SamAccountName
        DaysRemaining  = $DaysRemaining
        Date           = $Date.ToString('yyyy-MM-dd')
        SentUtc        = (Get-Date).ToUniversalTime().ToString('o')
    }

    # Prune entries older than 60 days to keep the tracking file from growing indefinitely.
    $cutoffDate = (Get-Date).AddDays(-60)
    $entries = $entries | Where-Object { [datetime]::Parse($_.Date) -ge $cutoffDate }

    $trackingDirectory = Split-Path -Path $TrackingFilePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($trackingDirectory) -and -not (Test-Path -Path $trackingDirectory)) {
        New-Item -Path $trackingDirectory -ItemType Directory -Force | Out-Null
    }

    $entries | ConvertTo-Json -Depth 3 | Set-Content -Path $TrackingFilePath -Encoding UTF8
}

function Get-NotificationTemplateContent {
    <#
    .SYNOPSIS
        Reads the raw content of a notification template file.

    .PARAMETER TemplatePath
        Path to the template file (.txt or .html).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath
    )

    if (-not (Test-Path -Path $TemplatePath)) {
        throw "Template file not found at path '$TemplatePath'."
    }

    return Get-Content -Path $TemplatePath -Raw -Encoding UTF8
}

function Format-NotificationContent {
    <#
    .SYNOPSIS
        Replaces placeholders in a template with user specific values.

    .DESCRIPTION
        Supported placeholders: {DisplayName}, {DaysRemaining}, {ExpiryDateDe}, {ExpiryDateEn}

    .PARAMETER TemplateContent
        The raw template content.

    .PARAMETER DisplayName
        The user's display name.

    .PARAMETER DaysRemaining
        Number of days until password expiration.

    .PARAMETER ExpiryDate
        The computed password expiration date.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplateContent,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [int]$DaysRemaining,

        [Parameter(Mandatory = $true)]
        [datetime]$ExpiryDate
    )

    # In .NET regex replacement strings, '$' has special meaning (e.g. backreferences).
    # Escaping '$' as '$$' in the replacement value prevents display names containing
    # a dollar sign from breaking the substitution.
    $safeDisplayName = $DisplayName -replace '\$', '$$$$'
    $cultureDe = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    $cultureEn = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')

    $formattedExpiryDateDe = $ExpiryDate.ToString('dd.MM.yyyy HH:mm', $cultureDe)
    $formattedExpiryDateEn = $ExpiryDate.ToString('MMMM d, yyyy h:mm tt', $cultureEn)

    $formatted = $TemplateContent
    $formatted = $formatted -replace '\{DisplayName\}', $safeDisplayName
    $formatted = $formatted -replace '\{DaysRemaining\}', $DaysRemaining
    $formatted = $formatted -replace '\{ExpiryDateDe\}', $formattedExpiryDateDe
    $formatted = $formatted -replace '\{ExpiryDateEn\}', $formattedExpiryDateEn

    return $formatted
}

function Send-NotificationMail {
    <#
    .SYNOPSIS
        Sends a notification email via SMTP using System.Net.Mail.

    .PARAMETER SmtpServer
        The SMTP server hostname or IP address.

    .PARAMETER SmtpPort
        The SMTP server port.

    .PARAMETER UseSsl
        Whether to use SSL/TLS for the SMTP connection.

    .PARAMETER Credential
        Optional PSCredential for SMTP authentication.

    .PARAMETER SenderDisplayName
        The display name to use as the sender.

    .PARAMETER SenderAddress
        The email address to use as the sender.

    .PARAMETER RecipientAddress
        The recipient's email address.

    .PARAMETER Subject
        The email subject line.

    .PARAMETER Body
        The email body content.

    .PARAMETER IsBodyHtml
        Whether the body content is HTML formatted.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SmtpServer,

        [Parameter(Mandatory = $true)]
        [int]$SmtpPort,

        [Parameter(Mandatory = $false)]
        [bool]$UseSsl = $false,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory = $true)]
        [string]$SenderDisplayName,

        [Parameter(Mandatory = $true)]
        [string]$SenderAddress,

        [Parameter(Mandatory = $true)]
        [string]$RecipientAddress,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$Body,

        [Parameter(Mandatory = $true)]
        [bool]$IsBodyHtml
    )

    if (-not $PSCmdlet.ShouldProcess($RecipientAddress, 'Send password expiration notification email')) {
        return
    }

    $mailMessage = $null
    $smtpClient = $null

    try {
        $mailMessage = New-Object -TypeName System.Net.Mail.MailMessage
        $mailMessage.From = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList $SenderAddress, $SenderDisplayName
        $mailMessage.To.Add($RecipientAddress)
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $Body
        $mailMessage.IsBodyHtml = $IsBodyHtml
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8

        $smtpClient = New-Object -TypeName System.Net.Mail.SmtpClient -ArgumentList $SmtpServer, $SmtpPort
        $smtpClient.EnableSsl = $UseSsl

        if ($Credential) {
            $smtpClient.Credentials = New-Object -TypeName System.Net.NetworkCredential -ArgumentList $Credential.UserName, $Credential.GetNetworkCredential().Password
        }

        try {
            $smtpClient.Send($mailMessage)
        }
        catch [System.Net.Mail.SmtpException] {
            Write-ScriptLog -Message "SMTP error while sending notification email to '$RecipientAddress': $($_.Exception.Message)" -Level 'Error'
            throw
        }

        Write-ScriptLog -Message "Notification email sent to '$RecipientAddress'." -Level 'Info'
    }
    catch {
        Write-ScriptLog -Message "Failed to send notification email to '$RecipientAddress': $($_.Exception.Message)" -Level 'Error'
        throw
    }
    finally {
        if ($mailMessage) { $mailMessage.Dispose() }
        if ($smtpClient) { $smtpClient.Dispose() }
    }
}

function Send-TestNotificationMail {
    <#
    .SYNOPSIS
        Sends a test notification email using both template formats.

    .PARAMETER Configuration
        The loaded script configuration object.

    .PARAMETER RecipientAddress
        The recipient email address used for the test message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$RecipientAddress
    )

    $sampleDisplayName = 'Test Recipient'
    $sampleDaysRemaining = 14
    $sampleExpiryDate = (Get-Date).Date.AddDays($sampleDaysRemaining)
    $subjectBase = $Configuration.Subject -replace '\{DaysRemaining\}', $sampleDaysRemaining

    $testTemplates = @(
        [PSCustomObject]@{
            Name   = 'HTML'
            Path   = $Configuration.Templates.HtmlTemplatePath
            IsHtml = $true
        },
        [PSCustomObject]@{
            Name   = 'Text'
            Path   = $Configuration.Templates.TextTemplatePath
            IsHtml = $false
        }
    )

    foreach ($template in $testTemplates) {
        $templateContent = Get-NotificationTemplateContent -TemplatePath $template.Path
        $body = Format-NotificationContent -TemplateContent $templateContent -DisplayName $sampleDisplayName -DaysRemaining $sampleDaysRemaining -ExpiryDate $sampleExpiryDate
        $templateType = if ($template.IsHtml) { 'Html' } else { 'Text' }
        $subject = "Test: $subjectBase [$templateType]"

        Send-NotificationMail -SmtpServer $Configuration.Smtp.Server `
            -SmtpPort $Configuration.Smtp.Port `
            -UseSsl $Configuration.Smtp.UseSsl `
            -SenderDisplayName $Configuration.Sender.DisplayName `
            -SenderAddress $Configuration.Sender.EmailAddress `
            -RecipientAddress $RecipientAddress `
            -Subject $subject `
            -Body $body `
            -IsBodyHtml $template.IsHtml
    }
}

# Main execution

try {
    Write-ScriptLog -Message 'Starting password expiration reminder run.' -Level 'Info'

    $configuration = Get-ScriptConfiguration -ConfigurationFilePath $ConfigurationFilePath

    if ([string]::IsNullOrWhiteSpace($script:LogFilePath) -and $configuration.Logging.PSObject.Properties.Name -contains 'ScriptLogPath') {
        $script:LogFilePath = $configuration.Logging.ScriptLogPath
    }

    $getExpiringUserParameters = @{
        SearchBase = $configuration.ActiveDirectory.SearchBase
        Thresholds = $configuration.Thresholds
    }

    if ($configuration.ActiveDirectory.PSObject.Properties.Name -contains 'MailboxRecipientTypeDetails' -and $configuration.ActiveDirectory.MailboxRecipientTypeDetails) {
        $getExpiringUserParameters['MailboxRecipientTypeDetails'] = [int64[]]$configuration.ActiveDirectory.MailboxRecipientTypeDetails
    }

    if ($ReportOnly -and $IncludeExpiredPasswordAccounts) {
        $getExpiringUserParameters['IncludeExpiredPasswords'] = $true
    }

    if ($TestEmail) {
        if ($ReportOnly -or $IncludeExpiredPasswordAccounts) {
            throw 'TestEmail cannot be combined with ReportOnly or IncludeExpiredPasswordAccounts.'
        }

        if ([string]::IsNullOrWhiteSpace($TestEmailRecipient)) {
            throw 'TestEmailRecipient must be specified when TestEmail is used.'
        }

        Write-ScriptLog -Message "Running in TestEmail mode. Sending HTML and text template test messages to '$TestEmailRecipient'." -Level 'Info'
        Send-TestNotificationMail -Configuration $configuration -RecipientAddress $TestEmailRecipient
        Write-ScriptLog -Message 'TestEmail mode completed.' -Level 'Info'
        return
    }

    $expiringUsers = Get-ExpiringPasswordUser @getExpiringUserParameters

    if ($ReportOnly) {

        Write-ScriptLog -Message 'Running in ReportOnly mode. No email will be sent and the notification tracking file will not be read or updated.' -Level 'Info'

        $expiredPasswordUsers = @()
        $reportUsers = $expiringUsers

        if ($IncludeExpiredPasswordAccounts) {
            $expiredPasswordUsers = @($expiringUsers | Where-Object { $_.DaysRemaining -lt 0 })
            $reportUsers = @($expiringUsers | Where-Object { $_.DaysRemaining -ge 0 })

            foreach ($expiredUser in $expiredPasswordUsers) {
                $recipientAddress = Get-PrimarySmtpAddress -User $expiredUser -AddressSource $configuration.Recipient.AddressSource -FallbackToWindowsMailAttribute $configuration.Recipient.FallbackToWindowsMailAttribute
                $resolvedRecipientAddress = if ([string]::IsNullOrWhiteSpace($recipientAddress)) { '<no address found>' } else { $recipientAddress }

                Write-LogFileEntry -Message ("ReportOnly expired password account: SamAccountName='{0}', DisplayName='{1}', RecipientAddress='{2}', ExpiryDate='{3}', DaysRemaining={4}" -f $expiredUser.SamAccountName, $expiredUser.DisplayName, $resolvedRecipientAddress, $expiredUser.ExpiryDate.ToString('yyyy-MM-dd'), $expiredUser.DaysRemaining)
            }
        }

        if ($reportUsers.Count -eq 0) {
            Write-ScriptLog -Message 'No users matched the configured notification thresholds today.' -Level 'Info'

            if ($IncludeExpiredPasswordAccounts -and $expiredPasswordUsers.Count -gt 0) {
                Write-LogFileEntry -Message "ReportOnly mode: logged $($expiredPasswordUsers.Count) expired password account(s)."
            }
        }
        else {
            $report = foreach ($user in $reportUsers) {

                $recipientAddress = Get-PrimarySmtpAddress -User $user -AddressSource $configuration.Recipient.AddressSource -FallbackToWindowsMailAttribute $configuration.Recipient.FallbackToWindowsMailAttribute

                [PSCustomObject]@{
                    SamAccountName    = $user.SamAccountName
                    DisplayName       = $user.DisplayName
                    RecipientAddress  = if ([string]::IsNullOrWhiteSpace($recipientAddress)) { '<no address found>' } else { $recipientAddress }
                    DaysRemaining     = $user.DaysRemaining
                    ExpiryDate        = $user.ExpiryDate.ToString('yyyy-MM-dd')
                }
            }

            $report | Sort-Object -Property DaysRemaining | Format-Table -AutoSize | Out-Host

            Write-ScriptLog -Message "ReportOnly mode: displayed $($report.Count) matching user(s)." -Level 'Info'

            if ($IncludeExpiredPasswordAccounts -and $expiredPasswordUsers.Count -gt 0) {
                Write-LogFileEntry -Message "ReportOnly mode: logged $($expiredPasswordUsers.Count) expired password account(s)."
            }
        }
    }
    else {

        if ($expiringUsers.Count -eq 0) {
            Write-ScriptLog -Message 'No users matched the configured notification thresholds today.' -Level 'Info'
        }
        else {
            $isHtml = -not $UseTextTemplate
            $templatePath = if ($isHtml) { $configuration.Templates.HtmlTemplatePath } else { $configuration.Templates.TextTemplatePath }
            $templateContent = Get-NotificationTemplateContent -TemplatePath $templatePath

            foreach ($user in $expiringUsers) {

                $recipientAddress = Get-PrimarySmtpAddress -User $user -AddressSource $configuration.Recipient.AddressSource -FallbackToWindowsMailAttribute $configuration.Recipient.FallbackToWindowsMailAttribute

                if ([string]::IsNullOrWhiteSpace($recipientAddress)) {
                    Write-ScriptLog -Message "No email address found for user '$($user.SamAccountName)'. Skipping." -Level 'Warning'
                    continue
                }

                $alreadySent = Test-NotificationAlreadySent -TrackingFilePath $configuration.Logging.NotificationTrackingFilePath -SamAccountName $user.SamAccountName -DaysRemaining $user.DaysRemaining -Date (Get-Date)

                if ($alreadySent) {
                    Write-ScriptLog -Message "Notification already sent today for user '$($user.SamAccountName)' at threshold '$($user.DaysRemaining)' day(s). Skipping." -Level 'Info'
                    continue
                }

                $subject = $configuration.Subject -replace '\{DaysRemaining\}', $user.DaysRemaining
                $body = Format-NotificationContent -TemplateContent $templateContent -DisplayName $user.DisplayName -DaysRemaining $user.DaysRemaining -ExpiryDate $user.ExpiryDate

                Send-NotificationMail -SmtpServer $configuration.Smtp.Server `
                    -SmtpPort $configuration.Smtp.Port `
                    -UseSsl $configuration.Smtp.UseSsl `
                    -SenderDisplayName $configuration.Sender.DisplayName `
                    -SenderAddress $configuration.Sender.EmailAddress `
                    -RecipientAddress $recipientAddress `
                    -Subject $subject `
                    -Body $body `
                    -IsBodyHtml $isHtml

                if ($PSCmdlet.ShouldProcess($recipientAddress, 'Record notification in tracking file')) {
                    Add-NotificationLogEntry -TrackingFilePath $configuration.Logging.NotificationTrackingFilePath -SamAccountName $user.SamAccountName -DaysRemaining $user.DaysRemaining -Date (Get-Date)
                }
            }
        }
    }

    Write-ScriptLog -Message 'Password expiration reminder run completed.' -Level 'Info'
}
catch {
    Write-ScriptLog -Message "Script execution failed: $($_.Exception.Message)" -Level 'Error'
    throw
}