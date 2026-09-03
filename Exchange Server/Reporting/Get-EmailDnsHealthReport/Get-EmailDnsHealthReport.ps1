#Requires -Version 5.1
<#
.SYNOPSIS
    Generates an email deliverability DNS health report for a configurable list of domains.

.DESCRIPTION
    Queries all DNS records relevant for email deliverability and security (MX, SPF, DKIM,
    DMARC, Autodiscover, MTA-STS, TLS-RPT, BIMI, SOA, NS, and optionally DNSSEC, CAA, and
    Microsoft 365 domain verification TXT records) for each domain listed in a JSON
    configuration file.

    The current result set is compared against the state saved during the previous run.
    Changed records and newly added domains are flagged so they can be highlighted in the
    generated HTML report. Both a CSV export and a styled HTML report are produced. The
    HTML report can optionally be sent by email.

    DKIM selector names are configurable through the JSON configuration file, both as a
    global default list and as optional per-domain overrides.

    Domain checks are executed concurrently using a runspace pool to reduce total runtime
    for larger domain lists. Regardless of completion order, the CSV and HTML reports are
    always sorted alphabetically by domain name.

.PARAMETER ConfigPath
    Path to the JSON configuration file describing the domain list, DKIM selectors, and
    email settings. Defaults to '.\Config\DnsReportConfig.json'.

.PARAMETER OutputPath
    Directory the CSV report, HTML report, and the state snapshot file are written to.
    The directory is created if it does not exist. Defaults to '.\Reports'.

.PARAMETER ThrottleLimit
    Maximum number of domains resolved concurrently through the runspace pool. Defaults to 5.

.PARAMETER IncludeDnssec
    Includes DNSSEC status (presence of DNSKEY/DS records) in the report.

.PARAMETER IncludeCaa
    Includes CAA record information in the report.

.PARAMETER IncludeM365Verification
    Includes the Microsoft 365 domain verification TXT record (MS=...) in the report.

.PARAMETER OpenReport
    Opens the generated HTML report in the default browser for immediate review after
    it has been written to disk.

.PARAMETER SendEmail
    Sends the generated HTML report by email using the SMTP settings from the
    configuration file. Without this switch, reports are only written to disk.

.PARAMETER SkipStateUpdate
    Generates the report and performs the change comparison against the existing state
    file, but does not overwrite the state file afterwards. Useful for test runs.

.EXAMPLE
    .\Get-EmailDnsHealthReport.ps1 -ConfigPath .\Config\DnsReportConfig.json -Verbose

    Runs the core checks (MX, SPF, DKIM, DMARC, Autodiscover, MTA-STS, TLS-RPT, BIMI, SOA, NS)
    for all configured domains and writes CSV and HTML reports to .\Reports.

.EXAMPLE
    .\Get-EmailDnsHealthReport.ps1 -IncludeDnssec -IncludeCaa -IncludeM365Verification -SendEmail

    Runs all checks including the optional ones and emails the resulting HTML report.

.EXAMPLE
    .\Get-EmailDnsHealthReport.ps1 -ThrottleLimit 10

    Resolves up to 10 domains concurrently, useful for larger domain lists.

.NOTES
    Author: Thomas Stensitzki
    Target platform: Generic Windows PowerShell / PowerShell 7 (Windows), relevant for
                      Exchange Server, Exchange Online and Microsoft 365 email domains
    Required modules: DnsClient (built into Windows, provides Resolve-DnsName)
    Change log:
        1.0.0 - Initial version
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigPath = '.\Config\DEV-DnsReportConfig.json',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = '.\Reports',

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 5,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDnssec,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeCaa,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeM365Verification,

    [Parameter(Mandatory = $false)]
    [switch]$OpenReport,

    [Parameter(Mandatory = $false)]
    [switch]$SendEmail,

    [Parameter(Mandatory = $false)]
    [switch]$SkipStateUpdate
)

#region Configuration and state handling

function Import-ReportConfiguration {
    <#
    .SYNOPSIS
        Loads and validates the JSON configuration file for the DNS health report.

    .PARAMETER Path
        Path to the JSON configuration file.

    .EXAMPLE
        $config = Import-ReportConfiguration -Path '.\Config\DnsReportConfig.json'
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if (-not (Test-Path -Path $Path)) {
            throw "Configuration file not found at path '$Path'."
        }

        $configContent = Get-Content -Path $Path -Raw -ErrorAction Stop
        $config = $configContent | ConvertFrom-Json -ErrorAction Stop

        if (-not $config.Domains -or $config.Domains.Count -eq 0) {
            throw 'Configuration file does not contain any domains.'
        }

        Write-Verbose -Message ('Loaded configuration with {0} domain(s).' -f $config.Domains.Count)
        return $config
    }
    catch {
        Write-Error -Message "Failed to load configuration file '$Path'. $($_.Exception.Message)" -ErrorAction Stop
    }
}

function Get-DkimSelectorList {
    <#
    .SYNOPSIS
        Resolves the effective DKIM selector list for a given domain.

    .DESCRIPTION
        Combines the global default selector list with any per-domain selectors.
        Per-domain overrides are treated as additional selectors to validate instead
        of replacing the default list for the domain.

    .PARAMETER Domain
        The domain to resolve the selector list for.

    .PARAMETER Configuration
        The loaded report configuration object.

    .EXAMPLE
        Get-DkimSelectorList -Domain 'fabrikam.com' -Configuration $config
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Configuration
    )

    $selectors = New-Object System.Collections.Generic.List[string]

    foreach ($selector in @($Configuration.DkimSelectors.Default)) {
        if (-not [string]::IsNullOrWhiteSpace($selector) -and -not $selectors.Contains($selector)) {
            [void]$selectors.Add($selector)
        }
    }

    $overrides = $Configuration.DkimSelectors.PerDomainSelectorAdds
    if ($overrides -and ($overrides.PSObject.Properties.Name -contains $Domain)) {
        foreach ($selector in @($overrides.$Domain)) {
            if (-not [string]::IsNullOrWhiteSpace($selector) -and -not $selectors.Contains($selector)) {
                [void]$selectors.Add($selector)
            }
        }
    }

    return @($selectors)
}

function Import-PreviousReportState {
    <#
    .SYNOPSIS
        Loads the DNS report state snapshot saved during the previous run.

    .PARAMETER Path
        Path to the state JSON file. If the file does not exist, an empty state is returned.

    .EXAMPLE
        $previousState = Import-PreviousReportState -Path '.\Reports\DnsReportState.json'
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $state = @{}

    if (Test-Path -Path $Path) {
        try {
            $rawState = Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            foreach ($domainProperty in $rawState.PSObject.Properties) {
                $state[$domainProperty.Name] = $domainProperty.Value
            }
            Write-Verbose -Message ('Loaded previous state for {0} domain(s) from {1}.' -f $state.Count, $Path)
        }
        catch {
            Write-Warning -Message "Could not read previous state file '$Path'. Treating as first run. $($_.Exception.Message)"
        }
    }
    else {
        Write-Verbose -Message 'No previous state file found. Treating this as the first run.'
    }

    return $state
}

function Save-CurrentReportState {
    <#
    .SYNOPSIS
        Persists the current DNS report results as the new state snapshot for the next run.

    .PARAMETER DomainResults
        Array of domain result objects produced by Get-DomainDnsStatus.

    .PARAMETER Path
        Path the state JSON file should be written to.

    .EXAMPLE
        Save-CurrentReportState -DomainResults $results -Path '.\Reports\DnsReportState.json'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$DomainResults,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $stateObject = [ordered]@{}
        foreach ($domainResult in $DomainResults) {
            $stateObject[$domainResult.Domain] = $domainResult.Fields
        }

        $stateObject | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding utf8 -Force -ErrorAction Stop
        Write-Verbose -Message "Saved current state to '$Path'."
    }
    catch {
        Write-Error -Message "Failed to save state file '$Path'. $($_.Exception.Message)"
    }
}

#endregion

#region DNS resolution helpers

function Resolve-DnsRecordSafe {
    <#
    .SYNOPSIS
        Wraps Resolve-DnsName with consistent error handling for the report.

    .PARAMETER Name
        The DNS name to query.

    .PARAMETER Type
        The DNS record type to query.

    .EXAMPLE
        Resolve-DnsRecordSafe -Name 'contoso.com' -Type TXT
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Type
    )

    try {
        return Resolve-DnsName -Name $Name -Type $Type -ErrorAction Stop
    }
    catch {
        Write-Verbose -Message ('No {0} record found for {1}. {2}' -f $Type, $Name, $_.Exception.Message)
        return $null
    }
}

function Get-TxtRecordValue {
    <#
    .SYNOPSIS
        Returns the first TXT record value at a given name matching an optional prefix filter.

    .PARAMETER Name
        The DNS name to query for TXT records.

    .PARAMETER Prefix
        Optional prefix (e.g. 'v=spf1') the TXT record content must start with.

    .EXAMPLE
        Get-TxtRecordValue -Name 'contoso.com' -Prefix 'v=spf1'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$Prefix
    )

    $txtRecords = Resolve-DnsRecordSafe -Name $Name -Type TXT
    if (-not $txtRecords) {
        return $null
    }

    foreach ($record in $txtRecords) {
        if (-not $record.Strings) { continue }
        $value = ($record.Strings -join '')

        if ([string]::IsNullOrWhiteSpace($Prefix) -or $value -like "$Prefix*") {
            return $value
        }
    }

    return $null
}

function Get-MxRecordInfo {
    <#
    .SYNOPSIS
        Retrieves the MX records for a domain as a normalized, sorted string.

    .PARAMETER Domain
        The domain to query.

    .EXAMPLE
        Get-MxRecordInfo -Domain 'contoso.com'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $mxRecords = Resolve-DnsRecordSafe -Name $Domain -Type MX
    if (-not $mxRecords) { return $null }

    $entries = $mxRecords |
        Sort-Object -Property Preference |
        ForEach-Object { '{0} ({1})' -f $_.NameExchange.TrimEnd('.'), $_.Preference }

    return ($entries -join '; ')
}

function Get-DkimRecordInfo {
    <#
    .SYNOPSIS
        Checks a list of DKIM selectors for a domain and returns the selectors with a valid key.

    .PARAMETER Domain
        The domain to query.

    .PARAMETER Selectors
        The list of DKIM selector names to test.

    .EXAMPLE
        Get-DkimRecordInfo -Domain 'contoso.com' -Selectors @('selector1','selector2')
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [Parameter(Mandatory = $true)]
        [string[]]$Selectors
    )

    $foundSelectors = New-Object System.Collections.Generic.List[string]

    foreach ($selector in $Selectors) {
        $name = '{0}._domainkey.{1}' -f $selector, $Domain
        $value = Get-TxtRecordValue -Name $name -Prefix 'v=DKIM1'
        if ($value) {
            $foundSelectors.Add($selector)
        }
    }

    if ($foundSelectors.Count -eq 0) { return $null }
    return ($foundSelectors -join ', ')
}

function Get-AutodiscoverInfo {
    <#
    .SYNOPSIS
        Retrieves the Autodiscover CNAME and SRV record information for a domain.

    .PARAMETER Domain
        The domain to query.

    .EXAMPLE
        Get-AutodiscoverInfo -Domain 'contoso.com'
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $cnameTarget = $null
    $cnameRecords = Resolve-DnsRecordSafe -Name "autodiscover.$Domain" -Type CNAME
    if ($cnameRecords) {
        $cnameTarget = ($cnameRecords | Select-Object -First 1).NameHost.TrimEnd('.')
    }

    $srvTarget = $null
    $srvRecords = Resolve-DnsRecordSafe -Name "_autodiscover._tcp.$Domain" -Type SRV
    if ($srvRecords) {
        $srvTarget = ($srvRecords | Select-Object -First 1).NameTarget.TrimEnd('.')
    }

    return [PSCustomObject]@{
        Cname = $cnameTarget
        Srv   = $srvTarget
    }
}

function Get-SoaInfo {
    <#
    .SYNOPSIS
        Retrieves the SOA serial number for a domain, used as a fast zone change indicator.

    .PARAMETER Domain
        The domain to query.

    .EXAMPLE
        Get-SoaInfo -Domain 'contoso.com'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $soaRecord = Resolve-DnsRecordSafe -Name $Domain -Type SOA | Select-Object -First 1
    if (-not $soaRecord) { return $null }
    return [string]$soaRecord.SerialNumber
}

function Get-NsInfo {
    <#
    .SYNOPSIS
        Retrieves the sorted, comma-separated list of authoritative name servers for a domain.

    .PARAMETER Domain
        The domain to query.

    .EXAMPLE
        Get-NsInfo -Domain 'contoso.com'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $nsRecords = Resolve-DnsRecordSafe -Name $Domain -Type NS
    if (-not $nsRecords) { return $null }

    $names = $nsRecords | ForEach-Object { $_.NameHost.TrimEnd('.') } | Sort-Object
    return ($names -join ', ')
}

function Get-DnssecInfo {
    <#
    .SYNOPSIS
        Determines whether a domain publishes DNSSEC signing keys.

    .DESCRIPTION
        Checks for the presence of DNSKEY records at the domain apex. This indicates the
        zone is signed. It does not perform full chain-of-trust validation.

    .PARAMETER Domain
        The domain to query.

    .EXAMPLE
        Get-DnssecInfo -Domain 'contoso.com'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $dnsKeyRecords = Resolve-DnsRecordSafe -Name $Domain -Type DNSKEY
    if ($dnsKeyRecords) {
        return 'Enabled ({0} key(s))' -f ($dnsKeyRecords | Measure-Object).Count
    }

    return 'Not enabled'
}

function Get-CaaInfo {
    <#
    .SYNOPSIS
        Retrieves CAA records for a domain.

    .PARAMETER Domain
        The domain to query.

    .EXAMPLE
        Get-CaaInfo -Domain 'contoso.com'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $caaRecords = Resolve-DnsRecordSafe -Name $Domain -Type CAA
    if (-not $caaRecords) { return $null }

    $entries = $caaRecords | ForEach-Object { '{0} {1} "{2}"' -f $_.Flags, $_.Tag, $_.Value }
    return ($entries -join '; ')
}

function Get-M365VerificationInfo {
    <#
    .SYNOPSIS
        Retrieves the Microsoft 365 domain verification TXT record (MS=...) for a domain.

    .PARAMETER Domain
        The domain to query.

    .EXAMPLE
        Get-M365VerificationInfo -Domain 'contoso.com'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    return Get-TxtRecordValue -Name $Domain -Prefix 'MS='
}

function Get-DomainDnsStatus {
    <#
    .SYNOPSIS
        Aggregates all configured DNS checks for a single domain into one result object.

    .PARAMETER Domain
        The domain to check.

    .PARAMETER Configuration
        The loaded report configuration object, used to resolve DKIM selectors and the
        BIMI selector.

    .PARAMETER IncludeDnssec
        Includes the DNSSEC check.

    .PARAMETER IncludeCaa
        Includes the CAA check.

    .PARAMETER IncludeM365Verification
        Includes the Microsoft 365 domain verification TXT check.

    .EXAMPLE
        Get-DomainDnsStatus -Domain 'contoso.com' -Configuration $config -IncludeDnssec -IncludeCaa
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Configuration,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeDnssec,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeCaa,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeM365Verification
    )

    Write-Verbose -Message "Checking DNS records for domain '$Domain'."

    $selectors = Get-DkimSelectorList -Domain $Domain -Configuration $Configuration
    $bimiSelector = if ($Configuration.BimiSelector) { $Configuration.BimiSelector } else { 'default' }
    $autodiscover = Get-AutodiscoverInfo -Domain $Domain

    $fields = [ordered]@{
        Mx            = Get-MxRecordInfo -Domain $Domain
        Spf           = Get-TxtRecordValue -Name $Domain -Prefix 'v=spf1'
        Dkim          = Get-DkimRecordInfo -Domain $Domain -Selectors $selectors
        Dmarc         = Get-TxtRecordValue -Name "_dmarc.$Domain" -Prefix 'v=DMARC1'
        AutodiscoverCname = $autodiscover.Cname
        AutodiscoverSrv   = $autodiscover.Srv
        MtaSts        = Get-TxtRecordValue -Name "_mta-sts.$Domain" -Prefix 'v=STSv1'
        TlsRpt        = Get-TxtRecordValue -Name "_smtp._tls.$Domain" -Prefix 'v=TLSRPTv1'
        Bimi          = Get-TxtRecordValue -Name "$bimiSelector._bimi.$Domain" -Prefix 'v=BIMI1'
        Soa           = Get-SoaInfo -Domain $Domain
        Ns            = Get-NsInfo -Domain $Domain
    }

    if ($IncludeDnssec) { $fields['Dnssec'] = Get-DnssecInfo -Domain $Domain }
    if ($IncludeCaa) { $fields['Caa'] = Get-CaaInfo -Domain $Domain }
    if ($IncludeM365Verification) { $fields['M365Verification'] = Get-M365VerificationInfo -Domain $Domain }

    return [PSCustomObject]@{
        Domain = $Domain
        Fields = $fields
    }
}

#endregion

#region Parallel execution

function Invoke-DomainDnsStatusParallel {
    <#
    .SYNOPSIS
        Runs Get-DomainDnsStatus for multiple domains concurrently using a runspace pool.

    .DESCRIPTION
        Creates a runspace pool and dispatches one runspace per domain so DNS resolution
        for the configured domain list runs concurrently instead of sequentially. This
        significantly reduces total runtime for larger domain lists. A runspace pool is
        used instead of ForEach-Object -Parallel so the script keeps working on both
        Windows PowerShell 5.1 and PowerShell 7.

        Results are always returned sorted alphabetically by domain name, regardless of
        the order in which the individual runspaces complete.

    .PARAMETER Domains
        The list of domains to check.

    .PARAMETER Configuration
        The loaded report configuration object, passed through to each runspace.

    .PARAMETER ThrottleLimit
        Maximum number of domains resolved concurrently.

    .PARAMETER IncludeDnssec
        Includes the DNSSEC check for each domain.

    .PARAMETER IncludeCaa
        Includes the CAA check for each domain.

    .PARAMETER IncludeM365Verification
        Includes the Microsoft 365 domain verification TXT check for each domain.

    .EXAMPLE
        Invoke-DomainDnsStatusParallel -Domains $config.Domains -Configuration $config -ThrottleLimit 8
    #>
    [CmdletBinding()]
    [OutputType([array])]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Domains,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Configuration,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 64)]
        [int]$ThrottleLimit = 5,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeDnssec,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeCaa,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeM365Verification
    )

    # Functions the runspaces need access to. Get-DomainDnsStatus is the entry point and
    # pulls in everything it calls internally.
    $functionsToImport = @(
        'Get-DkimSelectorList',
        'Resolve-DnsRecordSafe',
        'Get-TxtRecordValue',
        'Get-MxRecordInfo',
        'Get-DkimRecordInfo',
        'Get-AutodiscoverInfo',
        'Get-SoaInfo',
        'Get-NsInfo',
        'Get-DnssecInfo',
        'Get-CaaInfo',
        'Get-M365VerificationInfo',
        'Get-DomainDnsStatus'
    )

    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    foreach ($functionName in $functionsToImport) {
        $functionCommand = Get-Command -Name $functionName -CommandType Function -ErrorAction Stop
        $sessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry(
            $functionName, $functionCommand.Definition
        )
        $initialSessionState.Commands.Add($sessionStateFunction)
    }

    $runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $initialSessionState, $Host)
    $runspacePool.Open()

    $pendingJobs = New-Object System.Collections.Generic.List[object]

    try {
        foreach ($domain in $Domains) {
            $powershellInstance = [System.Management.Automation.PowerShell]::Create()
            $powershellInstance.RunspacePool = $runspacePool

            [void]$powershellInstance.AddScript({
                param($Domain, $Configuration, $IncludeDnssec, $IncludeCaa, $IncludeM365Verification)
                Get-DomainDnsStatus -Domain $Domain -Configuration $Configuration `
                    -IncludeDnssec:$IncludeDnssec -IncludeCaa:$IncludeCaa -IncludeM365Verification:$IncludeM365Verification
            })
            [void]$powershellInstance.AddParameters(@{
                Domain                  = $domain
                Configuration           = $Configuration
                IncludeDnssec           = $IncludeDnssec.IsPresent
                IncludeCaa              = $IncludeCaa.IsPresent
                IncludeM365Verification = $IncludeM365Verification.IsPresent
            })

            $asyncResult = $powershellInstance.BeginInvoke()

            $pendingJobs.Add([PSCustomObject]@{
                Domain      = $domain
                PowerShell  = $powershellInstance
                AsyncResult = $asyncResult
            })
        }

        $domainResults = foreach ($job in $pendingJobs) {
            try {
                $job.PowerShell.EndInvoke($job.AsyncResult)
            }
            catch {
                Write-Warning -Message "DNS check for domain '$($job.Domain)' failed. $($_.Exception.Message)"
            }
            finally {
                $job.PowerShell.Dispose()
            }
        }

        # Sort alphabetically regardless of completion order.
        return @($domainResults) | Sort-Object -Property Domain
    }
    finally {
        $runspacePool.Close()
        $runspacePool.Dispose()
    }
}

#endregion

#region Change comparison

function Compare-DomainDnsStatus {
    <#
    .SYNOPSIS
        Compares a domain's current DNS field values against the previous run's state.

    .PARAMETER DomainResult
        The current domain result object as produced by Get-DomainDnsStatus.

    .PARAMETER PreviousState
        Hashtable containing the previous run's state, keyed by domain name.

    .EXAMPLE
        Compare-DomainDnsStatus -DomainResult $result -PreviousState $previousState
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$DomainResult,

        [Parameter(Mandatory = $true)]
        [hashtable]$PreviousState
    )

    $isNewDomain = -not $PreviousState.ContainsKey($DomainResult.Domain)
    $changedFields = New-Object System.Collections.Generic.List[string]

    if (-not $isNewDomain) {
        $previousFields = $PreviousState[$DomainResult.Domain]

        foreach ($fieldName in $DomainResult.Fields.Keys) {
            $currentValue = [string]$DomainResult.Fields[$fieldName]
            $previousValue = $null

            if ($previousFields.PSObject -and ($previousFields.PSObject.Properties.Name -contains $fieldName)) {
                $previousValue = [string]$previousFields.$fieldName
            }

            if ($currentValue -ne $previousValue) {
                $changedFields.Add($fieldName)
            }
        }
    }

    return [PSCustomObject]@{
        Domain       = $DomainResult.Domain
        Fields       = $DomainResult.Fields
        IsNewDomain  = $isNewDomain
        ChangedFields = $changedFields
    }
}

#endregion

#region Report export

function ConvertTo-DnsFieldLabel {
    <#
    .SYNOPSIS
        Returns a readable display label for a DNS field name used in report output.

    .PARAMETER FieldName
        The internal field name as used in the result object.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FieldName
    )

    switch ($FieldName) {
        'Mx' { return 'MX' }
        'Spf' { return 'SPF' }
        'Dkim' { return 'DKIM' }
        'Dmarc' { return 'DMARC' }
        'AutodiscoverCname' { return 'AutoDiscover CNAME' }
        'AutodiscoverSrv' { return 'AutoDiscover SRV' }
        'MtaSts' { return 'MTA-STS' }
        'TlsRpt' { return 'TLS-RPT' }
        'Bimi' { return 'BIMI' }
        'Soa' { return 'SOA' }
        'Ns' { return 'NS' }
        'Dnssec' { return 'DNSSEC' }
        'Caa' { return 'CAA' }
        'M365Verification' { return 'M365 Verification' }
        default { return $FieldName }
    }
}

function Export-DnsReportCsv {
    <#
    .SYNOPSIS
        Exports the DNS report results to a CSV file.

    .PARAMETER ComparedResults
        Array of compared domain result objects produced by Compare-DomainDnsStatus.

    .PARAMETER Path
        Path the CSV file should be written to.

    .EXAMPLE
        Export-DnsReportCsv -ComparedResults $results -Path '.\Reports\DnsReport.csv'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$ComparedResults,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $rows = foreach ($result in $ComparedResults) {
            $row = [ordered]@{
                Domain          = $result.Domain
                'Is New Domain' = $result.IsNewDomain
                'Changed Fields' = ($result.ChangedFields -join '; ')
            }
            foreach ($fieldName in $result.Fields.Keys) {
                $label = ConvertTo-DnsFieldLabel -FieldName $fieldName
                $row[$label] = $result.Fields[$fieldName]
            }
            [PSCustomObject]$row
        }

        $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8 -Force -ErrorAction Stop
        Write-Verbose -Message "CSV report written to '$Path'."
    }
    catch {
        Write-Error -Message "Failed to write CSV report to '$Path'. $($_.Exception.Message)"
    }
}

function New-DnsReportHtml {
    <#
    .SYNOPSIS
        Builds the HTML body of the DNS health report, highlighting changes and new domains.

    .PARAMETER ComparedResults
        Array of compared domain result objects produced by Compare-DomainDnsStatus.

    .EXAMPLE
        New-DnsReportHtml -ComparedResults $results
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [array]$ComparedResults
    )

    $fieldNames = $ComparedResults[0].Fields.Keys
    $headerLabels = @($fieldNames | ForEach-Object { ConvertTo-DnsFieldLabel -FieldName $_ })

    $style = @'
<style>
    body { font-family: Segoe UI, Arial, sans-serif; background-color: #f4f6f8; color: #1f2933; margin: 0; padding: 24px; }
    h1 { font-size: 20px; color: #102a43; }
    .meta { color: #627d98; font-size: 13px; margin-bottom: 20px; }
    table { border-collapse: collapse; width: 100%; background-color: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
    th, td { border: 1px solid #d9e2ec; padding: 8px 10px; text-align: left; font-size: 12px; vertical-align: top; }
    th { background-color: #102a43; color: #ffffff; position: sticky; top: 0; }
    tbody tr:nth-child(odd) td { background-color: #ffffff; }
    tbody tr:nth-child(even) td { background-color: #f0f4f8; }
    td.changed { font-weight: 600; }
    td.missing { color: #cf1124; font-style: italic; }
    .status-label { display: inline-block; padding: 2px 7px; border: 1px solid currentColor; border-radius: 999px; font-size: 11px; font-weight: 700; letter-spacing: 0; line-height: 1.2; white-space: nowrap; }
    .changed-label { color: #b45309; background-color: #fff7ed; }
    .new-domain-label { color: #1976d2; background-color: #eff6ff; }

    .missing-label { display: inline-block; padding: 2px 7px; border: 1px solid #f5b5b5; border-radius: 999px; color: #cf1124; background-color: #fff1f2; font-style: normal; font-weight: 700; }
    .legend { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 12px; width: 100%; box-sizing: border-box; margin-top: 16px; font-size: 12px; color: #627d98; }
    .legend-item { display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }
</style>
'@

    $headerCells = "<th>Domain</th>" + (($headerLabels | ForEach-Object { "<th>$_</th>" }) -join '')

    $rowsHtml = foreach ($result in $ComparedResults) {
        $rowClass = if ($result.IsNewDomain) { 'new-domain' } else { '' }
        $domainCell = if ($result.IsNewDomain) {
            "<td>$($result.Domain) <span class='status-label new-domain-label'>NEW DOMAIN</span></td>"
        }
        else {
            "<td>$($result.Domain)</td>"
        }

        $dataCells = foreach ($fieldName in $fieldNames) {
            $value = $result.Fields[$fieldName]
            $isMissing = [string]::IsNullOrWhiteSpace([string]$value)
            if ($isMissing) {
                $displayValue = "<span class='missing-label'>N/A</span>"
            }
            elseif ($fieldName -eq 'Mx') {
                $displayValue = (([string]$value -split ';\s*' | ForEach-Object {
                    [System.Net.WebUtility]::HtmlEncode($_)
                }) -join '<br />')
            }
            else {
                $displayValue = [System.Net.WebUtility]::HtmlEncode([string]$value)
            }

            $classes = New-Object System.Collections.Generic.List[string]
            $changeLabel = ''
            if ($result.ChangedFields -contains $fieldName) {
                $classes.Add('changed')
                $changeLabel = "<span class='status-label changed-label'>CHANGED</span>"
            }
            if ($isMissing) { $classes.Add('missing') }

            $classAttribute = if ($classes.Count -gt 0) { " class='$($classes -join ' ')'" } else { '' }
            "<td$classAttribute>$changeLabel$displayValue</td>"
        }

        "<tr class='$rowClass'>$domainCell$($dataCells -join '')</tr>"
    }

    $generatedOn = (Get-Date).ToString('yyyy-MM-dd HH:mm')

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Email DNS Health Report</title>
$style
</head>
<body>
<h1>Email DNS Health Report</h1>
<div class="meta">Generated on $generatedOn UTC offset local time &middot; $($ComparedResults.Count) domain(s) checked</div>
<table>
<thead><tr>$headerCells</tr></thead>
<tbody>
$($rowsHtml -join "`n")
</tbody>
</table>
<div class="legend">
    <span class="legend-item"><span class="status-label changed-label">CHANGED</span> Value changed since last run</span>
    <span class="legend-item"><span class="status-label new-domain-label">NEW DOMAIN</span> Domain added since last run</span>
    <span class="legend-item"><span class="missing-label">N/A</span> Missing value</span>
</div>
</body>
</html>
"@
}

function Export-DnsReportHtml {
    <#
    .SYNOPSIS
        Writes the generated HTML report to disk.

    .PARAMETER ComparedResults
        Array of compared domain result objects produced by Compare-DomainDnsStatus.

    .PARAMETER Path
        Path the HTML file should be written to.

    .EXAMPLE
        Export-DnsReportHtml -ComparedResults $results -Path '.\Reports\DnsReport.html'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [array]$ComparedResults,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $html = New-DnsReportHtml -ComparedResults $ComparedResults
        $html | Out-File -FilePath $Path -Encoding utf8 -Force -ErrorAction Stop
        Write-Verbose -Message "HTML report written to '$Path'."
        return $html
    }
    catch {
        Write-Error -Message "Failed to write HTML report to '$Path'. $($_.Exception.Message)"
    }
}

#endregion

#region Email dispatch

function Send-DnsReportEmail {
    <#
    .SYNOPSIS
        Sends the generated HTML report by email using System.Net.Mail.SmtpClient.

    .DESCRIPTION
        SMTP authentication is intentionally not implemented here. If the SMTP server
        requires authentication, set $SmtpClient.Credentials before calling Send, or
        extend this function accordingly.

    .PARAMETER Configuration
        The loaded report configuration object, providing SMTP and recipient settings.

    .PARAMETER HtmlBody
        The HTML report body to send.

    .PARAMETER CsvPath
        Path to the CSV report to attach.

    .PARAMETER HtmlPath
        Path to the HTML report to attach.

    .EXAMPLE
        Send-DnsReportEmail -Configuration $config -HtmlBody $html -CsvPath $csvPath -HtmlPath $htmlPath
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,

        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string]$HtmlPath
    )

    try {
        $emailConfig = $Configuration.Email
        if (-not $emailConfig) {
            throw "Configuration file does not contain an 'Email' section."
        }

        $smtpClient = New-Object System.Net.Mail.SmtpClient($emailConfig.SmtpServer, $emailConfig.SmtpPort)
        $smtpClient.EnableSsl = [bool]$emailConfig.UseSsl

        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $emailConfig.From
        foreach ($recipient in $emailConfig.To) {
            $mailMessage.To.Add($recipient)
        }
        $mailMessage.Subject = $emailConfig.Subject
        $mailMessage.Body = $HtmlBody
        $mailMessage.IsBodyHtml = $true

        $csvAttachment = New-Object System.Net.Mail.Attachment($CsvPath)
        $htmlAttachment = New-Object System.Net.Mail.Attachment($HtmlPath)
        $mailMessage.Attachments.Add($csvAttachment)
        $mailMessage.Attachments.Add($htmlAttachment)

        $smtpClient.Send($mailMessage)
        Write-Verbose -Message ('Report email sent to {0}.' -f ($emailConfig.To -join ', '))

        $csvAttachment.Dispose()
        $htmlAttachment.Dispose()
        $mailMessage.Dispose()
        $smtpClient.Dispose()
    }
    catch {
        Write-Error -Message "Failed to send report email. $($_.Exception.Message)"
    }
}

#endregion

#region Main execution

try {
    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $configuration = Import-ReportConfiguration -Path $ConfigPath
    $stateFileName = if ($configuration.StateFileName) { $configuration.StateFileName } else { 'DnsReportState.json' }
    $stateFilePath = Join-Path -Path $OutputPath -ChildPath $stateFileName
    $previousState = Import-PreviousReportState -Path $stateFilePath

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $csvPath = Join-Path -Path $OutputPath -ChildPath "DnsReport-$timestamp.csv"
    $htmlPath = Join-Path -Path $OutputPath -ChildPath "DnsReport-$timestamp.html"

    # Domain checks run concurrently; results come back sorted alphabetically by domain name.
    $domainResults = Invoke-DomainDnsStatusParallel -Domains $configuration.Domains -Configuration $configuration `
        -ThrottleLimit $ThrottleLimit -IncludeDnssec:$IncludeDnssec -IncludeCaa:$IncludeCaa -IncludeM365Verification:$IncludeM365Verification

    $comparedResults = foreach ($domainResult in $domainResults) {
        Compare-DomainDnsStatus -DomainResult $domainResult -PreviousState $previousState
    }

    Export-DnsReportCsv -ComparedResults $comparedResults -Path $csvPath
    $htmlBody = Export-DnsReportHtml -ComparedResults $comparedResults -Path $htmlPath

    if (-not $SkipStateUpdate) {
        Save-CurrentReportState -DomainResults $domainResults -Path $stateFilePath
    }

    if ($OpenReport) {
        try {
            Start-Process -FilePath $htmlPath -ErrorAction Stop
            Write-Verbose -Message "Opened HTML report in the default browser: '$htmlPath'."
        }
        catch {
            Write-Warning -Message "Failed to open report in the default browser. $($_.Exception.Message)"
        }
    }

    if ($SendEmail) {
        Send-DnsReportEmail -Configuration $configuration -HtmlBody $htmlBody -CsvPath $csvPath -HtmlPath $htmlPath
    }

    Write-Output 'DNS health report completed.'
    Write-Output "CSV: $csvPath"
    Write-Output "HTML: $htmlPath"
}
catch {
    Write-Error -Message "DNS health report failed. $($_.Exception.Message)"
}

#endregion
