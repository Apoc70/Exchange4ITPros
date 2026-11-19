<#
.SYNOPSIS
    Analyzes Exchange IIS logs to identify users who still access OWA.

.DESCRIPTION
    Parses W3C IIS logs from the Exchange Front End site (Default Web Site, usually W3SVC1).
    Field order is read dynamically from the #Fields header, so the script also works with
    customized log fields (for example, X-Forwarded-For behind a load balancer).

    Output:
      - OWA_Users.csv              One record per user (requests, time range, client IPs, user agent, ...)
      - OWA_DailyUniqueUsers.csv   Unique users and requests per day
      - Summary in the console

    Note: IIS logs timestamps in UTC. StartDate and EndDate are converted to UTC internally.

    "OwaServiceCalls" column: Number of calls to /owa/service.svc. These occur only in a
    fully loaded OWA session. Users with Requests > 0 but OwaServiceCalls = 0 typically have
    only logon or redirect hits (for example, already migrated mailboxes or TargetOwaURL).

.PARAMETER LogPath
    One or more log directories. For multiple servers, specify UNC paths, for example:
    \\EX01\c$\inetpub\logs\LogFiles\W3SVC1, \\EX02\c$\inetpub\logs\LogFiles\W3SVC1

.PARAMETER StartDate / EndDate
    Analysis time range (local time). Default: last 30 days.

.PARAMETER UriPrefix
    URI prefixes to count. Default: /owa/. Optionally, add /ecp/, for example.

.PARAMETER IncludeFailed
    Also count requests with HTTP status >= 400.

.PARAMETER StripDomain
    Normalize DOMAIN\user and user@domain to "user" so mixed formats are grouped together.

.EXAMPLE
    .\Get-OwaUsageFromIisLogs.ps1

.EXAMPLE
    .\Get-OwaUsageFromIisLogs.ps1 -LogPath '\\EX01\c$\inetpub\logs\LogFiles\W3SVC1','\\EX02\c$\inetpub\logs\LogFiles\W3SVC1' -StartDate (Get-Date).AddDays(-14) -StripDomain

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7.
#>
[CmdletBinding()]
param(
    [string[]]$LogPath = @("$env:SystemDrive\inetpub\logs\LogFiles\W3SVC1"),

    [datetime]$StartDate = (Get-Date).Date.AddDays(-30),

    [datetime]$EndDate = (Get-Date),

    [string[]]$UriPrefix = @('/owa/'),

    [string[]]$ExcludeUserPattern = @('HealthMailbox*', '*\HealthMailbox*', 'SM_*', '*\SM_*', '*$'),

    [string[]]$ExcludeUserAgentPattern = @('*AMProbe*', '*ACTIVEMONITORING*', '*MSEXCHMON*', '*HealthService*'),

    [switch]$IncludeFailed,

    [switch]$StripDomain,

    [string]$OutputFolder = (Join-Path -Path (Get-Location) -ChildPath ("OWA-Report_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmm')))
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$startUtc = $StartDate.ToUniversalTime()
$endUtc   = $EndDate.ToUniversalTime()
$inv      = [System.Globalization.CultureInfo]::InvariantCulture
$dtStyle  = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
$cmpIC    = [System.StringComparison]::OrdinalIgnoreCase

# Aggregation structures
$users          = @{}   # user -> statistics object
$daily          = @{}   # yyyy-MM-dd -> @{ Users = HashSet; Requests = int }
$excludeCache   = @{}   # user -> bool
$uaExcludeCache = @{}   # useragent -> bool

$stats = [ordered]@{
    Files          = 0
    LinesTotal     = 0
    LinesMatched   = 0
    SkippedNoUser  = 0
    SkippedExclude = 0
    SkippedStatus  = 0
}

function Test-Like {
    param([string]$Value, [string[]]$Patterns)
    foreach ($pat in $Patterns) {
        if ($Value -like $pat) { return $true }
    }
    return $false
}

# Find log files (filenames u_exYYMMDD.log or u_exYYMMDDHH.log)
$files = foreach ($path in $LogPath) {
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Warning "Path not found: $path"
        continue
    }
    Get-ChildItem -LiteralPath $path -Filter '*.log' -File | Where-Object {
        if ($_.LastWriteTime -lt $StartDate) { return $false }
        if ($_.BaseName -match '^u_ex(\d{6})') {
            $fileDate = [datetime]::ParseExact($Matches[1], 'yyMMdd', $inv)
            return ($fileDate -le $EndDate.Date.AddDays(1))
        }
        return $true
    }
}
$files = @($files | Sort-Object FullName)

if ($files.Count -eq 0) {
    Write-Warning 'No matching log files found in the specified time range.'
    return
}

Write-Host ("{0} log file(s) in the time range {1:yyyy-MM-dd HH:mm} to {2:yyyy-MM-dd HH:mm} (local)" -f $files.Count, $StartDate, $EndDate) -ForegroundColor Cyan

$fileCounter = 0
foreach ($file in $files) {
    $fileCounter++
    Write-Progress -Activity 'Analyzing IIS logs' -Status $file.FullName -PercentComplete (($fileCounter / $files.Count) * 100)
    $stats.Files++

    $idx = $null
    $fs = $null
    $reader = $null
    try {
        # FileShare ReadWrite allows reading the file currently being written by IIS.
        $fs = [System.IO.File]::Open($file.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $reader = New-Object System.IO.StreamReader($fs)

        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line.Length -eq 0) { continue }

            if ($line[0] -eq '#') {
                if ($line.StartsWith('#Fields:')) {
                    $fieldNames = $line.Substring(8).Trim().Split(' ')
                    $idx = @{}
                    for ($i = 0; $i -lt $fieldNames.Count; $i++) { $idx[$fieldNames[$i]] = $i }

                    foreach ($req in 'date', 'time', 'cs-uri-stem', 'cs-username', 'c-ip', 'sc-status') {
                        if (-not $idx.ContainsKey($req)) {
                            Write-Warning "Field '$req' is missing from $($file.Name). Skipping file."
                            $idx = $null
                            break
                        }
                    }
                    if ($null -ne $idx) {
                        $iDate   = $idx['date']
                        $iTime   = $idx['time']
                        $iUri    = $idx['cs-uri-stem']
                        $iUser   = $idx['cs-username']
                        $iCip    = $idx['c-ip']
                        $iStatus = $idx['sc-status']
                        $iUa     = if ($idx.ContainsKey('cs(User-Agent)')) { $idx['cs(User-Agent)'] } else { -1 }
                        $iXff    = if ($idx.ContainsKey('X-Forwarded-For')) { $idx['X-Forwarded-For'] } else { -1 }
                        $iSip    = if ($idx.ContainsKey('s-ip')) { $idx['s-ip'] } else { -1 }
                    }
                }
                continue
            }

            if ($null -eq $idx) { continue }
            $stats.LinesTotal++

            $p = $line.Split(' ')
            if ($p.Count -le $iStatus) { continue }

            # URI filter
            $uri = $p[$iUri]
            $uriMatch = $false
            foreach ($prefix in $UriPrefix) {
                if ($uri.StartsWith($prefix, $cmpIC)) { $uriMatch = $true; break }
            }
            if (-not $uriMatch) { continue }

            # User
            $user = $p[$iUser]
            if ($user -eq '-' -or [string]::IsNullOrEmpty($user)) { $stats.SkippedNoUser++; continue }
            $user = [System.Uri]::UnescapeDataString($user).ToLowerInvariant()

            if (-not $excludeCache.ContainsKey($user)) {
                $excludeCache[$user] = Test-Like -Value $user -Patterns $ExcludeUserPattern
            }
            if ($excludeCache[$user]) { $stats.SkippedExclude++; continue }

            if ($StripDomain) {
                if ($user.Contains('\')) { $user = $user.Substring($user.IndexOf('\') + 1) }
                if ($user.Contains('@')) { $user = $user.Substring(0, $user.IndexOf('@')) }
            }

            # Status
            $status = 0
            [void][int]::TryParse($p[$iStatus], [ref]$status)
            if (-not $IncludeFailed -and $status -ge 400) { $stats.SkippedStatus++; continue }

            # User agent (exclude monitoring probes)
            $ua = '-'
            if ($iUa -ge 0) {
                $ua = $p[$iUa].Replace('+', ' ')
                if (-not $uaExcludeCache.ContainsKey($ua)) {
                    $uaExcludeCache[$ua] = Test-Like -Value $ua -Patterns $ExcludeUserAgentPattern
                }
                if ($uaExcludeCache[$ua]) { $stats.SkippedExclude++; continue }
            }

            # Timestamp (UTC)
            $ts = [datetime]::ParseExact(('{0} {1}' -f $p[$iDate], $p[$iTime]), 'yyyy-MM-dd HH:mm:ss', $inv, $dtStyle)
            if ($ts -lt $startUtc -or $ts -gt $endUtc) { continue }

            # Client IP (prefer X-Forwarded-For if logged)
            $clientIp = $p[$iCip]
            if ($iXff -ge 0 -and $p[$iXff] -ne '-') { $clientIp = $p[$iXff].Split(',')[0].Trim() }

            $stats.LinesMatched++

            # Aggregate by user
            if (-not $users.ContainsKey($user)) {
                $users[$user] = [pscustomobject]@{
                    Requests        = 0
                    OwaServiceCalls = 0
                    FirstSeen       = $ts
                    LastSeen        = $ts
                    Days            = New-Object 'System.Collections.Generic.HashSet[string]'
                    ClientIPs       = New-Object 'System.Collections.Generic.HashSet[string]'
                    ServerIPs       = New-Object 'System.Collections.Generic.HashSet[string]'
                    UserAgents      = @{}
                }
            }
            $u = $users[$user]
            $u.Requests++
            if ($uri.StartsWith('/owa/service.svc', $cmpIC)) { $u.OwaServiceCalls++ }
            if ($ts -lt $u.FirstSeen) { $u.FirstSeen = $ts }
            if ($ts -gt $u.LastSeen)  { $u.LastSeen  = $ts }
            $day = $p[$iDate]
            [void]$u.Days.Add($day)
            [void]$u.ClientIPs.Add($clientIp)
            if ($iSip -ge 0) { [void]$u.ServerIPs.Add($p[$iSip]) }
            if ($u.UserAgents.ContainsKey($ua)) { $u.UserAgents[$ua]++ } else { $u.UserAgents[$ua] = 1 }

            # Aggregate by day
            if (-not $daily.ContainsKey($day)) {
                $daily[$day] = @{ Users = (New-Object 'System.Collections.Generic.HashSet[string]'); Requests = 0 }
            }
            [void]$daily[$day].Users.Add($user)
            $daily[$day].Requests++
        }
    }
    catch {
        Write-Warning "Error reading $($file.FullName): $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $fs) { $fs.Dispose() }
    }
}
Write-Progress -Activity 'Analyzing IIS logs' -Completed

# Prepare results
$userReport = foreach ($key in $users.Keys) {
    $u = $users[$key]
    $topUa = $u.UserAgents.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
    [pscustomobject]@{
        User            = $key
        Requests        = $u.Requests
        OwaServiceCalls = $u.OwaServiceCalls
        RealOwaSession  = ($u.OwaServiceCalls -gt 0)
        ActiveDays      = $u.Days.Count
        FirstSeenUtc    = $u.FirstSeen.ToString('yyyy-MM-dd HH:mm:ss')
        LastSeenUtc     = $u.LastSeen.ToString('yyyy-MM-dd HH:mm:ss')
        ClientIPCount   = $u.ClientIPs.Count
        ClientIPs       = (@($u.ClientIPs) | Select-Object -First 10) -join ';'
        ServerIPs       = @($u.ServerIPs) -join ';'
        TopUserAgent    = if ($topUa) { $topUa.Key } else { '' }
    }
}
$userReport = @($userReport | Sort-Object LastSeenUtc -Descending)

$dailyReport = foreach ($d in ($daily.Keys | Sort-Object)) {
    [pscustomobject]@{
        DateUtc     = $d
        UniqueUsers = $daily[$d].Users.Count
        Requests    = $daily[$d].Requests
    }
}
$dailyReport = @($dailyReport)

# Export results
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
}
$userCsv  = Join-Path $OutputFolder 'OWA_Users.csv'
$dailyCsv = Join-Path $OutputFolder 'OWA_DailyUniqueUsers.csv'
$userReport  | Export-Csv -Path $userCsv  -NoTypeInformation -Delimiter ';' -Encoding UTF8
$dailyReport | Export-Csv -Path $dailyCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8

# Summary
$realUsers = @($userReport | Where-Object { $_.RealOwaSession }).Count
Write-Host ''
Write-Host '===== OWA usage =====' -ForegroundColor Green
Write-Host ("Log files processed         : {0}" -f $stats.Files)
Write-Host ("Total log lines             : {0}" -f $stats.LinesTotal)
Write-Host ("OWA requests counted        : {0}" -f $stats.LinesMatched)
Write-Host ("Skipped (no user)           : {0}" -f $stats.SkippedNoUser)
Write-Host ("Skipped (excluded)          : {0}" -f $stats.SkippedExclude)
Write-Host ("Skipped (status >= 400)     : {0}" -f $stats.SkippedStatus)
Write-Host ''
Write-Host ("Unique users                : {0}" -f $userReport.Count) -ForegroundColor Yellow
Write-Host ("With a real OWA session     : {0}" -f $realUsers) -ForegroundColor Yellow
Write-Host ''
Write-Host "Top 20 by most recent access:"
$userReport | Select-Object -First 20 User, Requests, OwaServiceCalls, ActiveDays, LastSeenUtc | Format-Table -AutoSize
Write-Host "Exported: $userCsv"
Write-Host "Exported: $dailyCsv"
