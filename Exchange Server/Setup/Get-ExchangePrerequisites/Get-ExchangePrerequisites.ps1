<#
.SYNOPSIS
    Downloads the prerequisites for the Exchange Server SE Mailbox role and, optionally,
    the Exchange Server SE ISO and the latest Security Update.

.DESCRIPTION
    This script downloads .NET Framework (depending on the Windows Server version),
    the Visual C++ 2012 and 2013 Redistributables (x64), and the IIS URL Rewrite Module
    to a local installation directory. The filenames match the examples in the
    knowledge base article so that its installation commands work unchanged.

    With -IncludeExchangeSE, the script also downloads the Exchange Server SE RTM ISO
    and the latest Exchange Server SE Security Update (SU). The SU metadata (KB number,
    Download Center ID, SHA256 hashes) is maintained in the $ExchangeSE block below and
    must be updated after each Patch Tuesday. The direct download URL of the SU is
    resolved from the Download Center page at runtime.

    After downloading, the script checks each executable's Authenticode signature and
    displays its file version. Where a SHA256 hash is known (ISO and SU), the file hash
    is verified as well. This lets you verify whether the Visual C++ packages match the
    latest builds (2012: 11.0.61030, 2013: 12.0.40664).

    UCMA 4.0 is not downloaded. Install the runtime with UCMARunTimeSetup.exe from the
    UCMARedist folder on the Exchange ISO. The script creates a placeholder text file
    (UCMA-4.0-NOT-DOWNLOADED.txt) in the destination directory as a reminder.

.PARAMETER Path
    Destination directory for downloads. Default: C:\Install

.PARAMETER TargetOS
    Windows Server version of the Exchange server: 2025, 2022, or 2019.
    If omitted, the local system's version is used. Specify this value when
    downloading packages on an admin system for a different server.

.PARAMETER Proxy
    Optional proxy server, for example http://proxy.squeakylobster.dev:8080

.PARAMETER Force
    Downloads files again, even if they already exist.

.PARAMETER IncludeExchangeSE
    Additionally downloads the Exchange Server SE RTM ISO (about 6 GB) and the
    latest Exchange Server SE Security Update.

.PARAMETER Language
    Language of the Security Update package. Use the language of the Exchange
    installation. Default: en

.EXAMPLE
    .\Get-ExchangePrerequisites.ps1

    Downloads the packages for the local Windows Server version to C:\Install.

.EXAMPLE
    .\Get-ExchangePrerequisites.ps1 -TargetOS 2022 -Path D:\Install -Proxy http://10.1.50.31:8080

    Downloads the packages for Windows Server 2022 through a proxy to D:\Install.

.EXAMPLE
    .\Get-ExchangePrerequisites.ps1 -TargetOS 2025 -IncludeExchangeSE

    Downloads the prerequisites for Windows Server 2025, the Exchange Server SE ISO,
    and the latest English Security Update to C:\Install.

.NOTES
    Revision: 2026-11-10 - Initial release
    Revision: 2026-11-11 - Added -IncludeExchangeSE and -Language (Exchange SE ISO and latest SU)
    Revision: 2026-11-11 - Added UCMA placeholder text file
#>
[CmdletBinding()]
param(
    [string]$Path = 'C:\INSTALL_',

    [ValidateSet('2025', '2022', '2019')]
    [string]$TargetOS,

    [string]$Proxy,

    [switch]$Force,

    [switch]$IncludeExchangeSE,

    [ValidateSet('de', 'en', 'es', 'fr', 'it', 'ja', 'ko', 'pt-br', 'ru', 'zh-hans', 'zh-hant')]
    [string]$Language = 'en'
)

# Exchange Server SE build information. Update the SU block after each Patch Tuesday.
$ExchangeSE = @{
    IsoName      = 'Exchange Server SE RTM ISO (15.02.2562.017)'
    IsoUrl       = 'https://download.microsoft.com/download/ad9b1e53-f28e-4a82-987d-6c88803795b2/ExchangeServerSE-x64.iso'
    IsoFileName  = 'ExchangeServerSE-x64.iso'
    IsoSha256    = 'CD2B13F2C297187776AF4CFF3541B4BE3C677CF907CCA69D85AB0E2B70377BD1'

    SuName       = 'Exchange Server SE RTM SU12, November 2026'
    SuKB         = 'KB5133453'
    SuDownloadId = '108908'
    # Known SHA256 hashes per language. Missing languages fall back to the signature check only.
    SuSha256     = @{
        'en' = 'C0703944B9E2E7B8980DAEF0A582761C2BBA9AD10DA99E2C79E1759DC9897D34'
        'fr' = '029F48734E8B6320EA0677A51F0563A105CB5C44DC531A3BE1264498976AD2AC'
        'it' = 'B4E6078074701EBDC5194C79D1EE77B5588BF59C9EF3D67EDD54A711E4EF4C81'
        'ja' = '42DE1CFBF88681DCCFE0DD6549A07FECDD03848807A61B3A5ADE1289A933D638'
        'ru' = '11EF5272C8854DA2EEC7AF48D01C6506393C5BFDEC876B84314BC64C25FCF2BB'
    }
}

# Windows PowerShell 5.1 does not use TLS 1.2 by default.
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# The progress display significantly slows down Invoke-WebRequest in Windows PowerShell.
$ProgressPreference = 'SilentlyContinue'

function Resolve-DownloadCenterUrl {
    # Reads a Download Center details page and returns the direct download URL for a file name.
    param(
        [Parameter(Mandatory)][string]$DownloadId,
        [Parameter(Mandatory)][string]$FileName,
        [hashtable]$WebParams = @{}
    )

    $Page = Invoke-WebRequest -Uri "https://www.microsoft.com/download/details.aspx?id=$DownloadId" @WebParams -ErrorAction Stop
    # Embedded JSON may contain escaped slashes.
    $Content = $Page.Content -replace '\\/', '/'
    $Pattern = "https://download\.microsoft\.com/download/[^""'\s<>]*/$([regex]::Escape($FileName))"
    $Match = [regex]::Match($Content, $Pattern, 'IgnoreCase')

    if (-not $Match.Success) {
        throw "No download link for $FileName found on Download Center page $DownloadId. Check the language or the SU metadata."
    }
    $Match.Value
}

# Determine the Windows Server version if none was specified.
if (-not $TargetOS) {
    $Build = [int](Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber
    $TargetOS = switch ($Build) {
        { $_ -ge 26100 } { '2025'; break }
        { $_ -ge 20348 } { '2022'; break }
        { $_ -ge 17763 } { '2019'; break }
        default { throw "Build $Build is not a supported operating system for Exchange Server SE. Specify -TargetOS." }
    }
    Write-Host "Detected target version: Windows Server $TargetOS (build $Build)"
}

# Package list
$Packages = @(
    [pscustomobject]@{
        Name       = 'Visual C++ Redistributable 2012 Update 4 (x64)'
        Url        = 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe'
        DownloadId = $null
        FileName   = 'vcredist_x64_2012.exe'
        Expected   = '11.0.61030'
        Sha256     = $null
    },
    [pscustomobject]@{
        Name       = 'Visual C++ Redistributable 2013 (x64)'
        Url        = 'https://aka.ms/highdpimfc2013x64enu'
        DownloadId = $null
        FileName   = 'vcredist_x64_2013.exe'
        Expected   = '12.0.40664'
        Sha256     = $null
    },
    [pscustomobject]@{
        Name       = 'IIS URL Rewrite Module 2.1 (x64)'
        Url        = 'https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi'
        DownloadId = $null
        FileName   = 'rewrite_amd64_en-US.msi'
        Expected   = $null
        Sha256     = $null
    }
)

switch ($TargetOS) {
    '2022' {
        $Packages += [pscustomobject]@{
            Name       = '.NET Framework 4.8.1 (Offline-Installer)'
            Url        = 'https://go.microsoft.com/fwlink/?linkid=2203305'
            DownloadId = $null
            FileName   = 'NDP481-x86-x64-AllOS-ENU.exe'
            Expected   = $null
            Sha256     = $null
        }
    }
    '2019' {
        $Packages += [pscustomobject]@{
            Name       = '.NET Framework 4.8 (Offline-Installer)'
            Url        = 'https://go.microsoft.com/fwlink/?linkid=2088631'
            DownloadId = $null
            FileName   = 'ndp48-x86-x64-allos-enu.exe'
            Expected   = $null
            Sha256     = $null
        }
    }
    '2025' {
        Write-Host '.NET Framework 4.8.1 is included in Windows Server 2025; no download is needed.'
    }
}

if ($IncludeExchangeSE) {
    $Packages += [pscustomobject]@{
        Name       = $ExchangeSE.IsoName
        Url        = $ExchangeSE.IsoUrl
        DownloadId = $null
        FileName   = $ExchangeSE.IsoFileName
        Expected   = $null
        Sha256     = $ExchangeSE.IsoSha256
    }
    $Packages += [pscustomobject]@{
        Name       = "$($ExchangeSE.SuName) ($($ExchangeSE.SuKB), $Language)"
        Url        = $null
        DownloadId = $ExchangeSE.SuDownloadId
        FileName   = "ExchangeSubscriptionEdition-$($ExchangeSE.SuKB)-x64-$Language.exe"
        Expected   = $null
        Sha256     = $ExchangeSE.SuSha256[$Language]
    }
}

# Create the destination directory.
if (-not (Test-Path -Path $Path)) {
    New-Item -Path $Path -ItemType Directory | Out-Null
}

# The ISO alone needs about 6 GB.
if ($IncludeExchangeSE) {
    $Drive = (Get-Item -Path $Path).PSDrive
    if ($Drive -and $Drive.Free -lt 7GB) {
        Write-Warning "Less than 7 GB free on drive $($Drive.Name). The Exchange ISO download may fail."
    }
}

$WebParams = @{ UseBasicParsing = $true }
if ($Proxy) {
    $WebParams.Proxy = $Proxy
    $WebParams.ProxyUseDefaultCredentials = $true
}

$Results = foreach ($Package in $Packages) {
    $Target = Join-Path -Path $Path -ChildPath $Package.FileName

    if ((Test-Path -Path $Target) -and -not $Force) {
        Write-Host "Already present: $($Package.FileName)"
    }
    else {
        Write-Host "Downloading:   $($Package.Name)"
        try {
            $Url = $Package.Url
            if (-not $Url -and $Package.DownloadId) {
                $Url = Resolve-DownloadCenterUrl -DownloadId $Package.DownloadId -FileName $Package.FileName -WebParams $WebParams
            }
            Invoke-WebRequest -Uri $Url -OutFile $Target @WebParams -ErrorAction Stop
        }
        catch {
            Write-Warning "Download failed: $($Package.Name). $($_.Exception.Message)"
            # Do not leave a partial file behind; it would be skipped on the next run.
            Remove-Item -Path $Target -ErrorAction SilentlyContinue
            [pscustomobject]@{
                File      = $Package.FileName
                Version   = $null
                Signature = 'Download failed'
                Note      = if ($Url) { $Url } else { "Download Center ID $($Package.DownloadId)" }
            }
            continue
        }
    }

    Unblock-File -Path $Target

    $Version = $null
    $SignatureStatus = 'n/a'
    $Hints = @()

    # Authenticode applies to executables and MSI packages, not to the ISO.
    if ($Target -match '\.(exe|msi)$') {
        $Signature = Get-AuthenticodeSignature -FilePath $Target
        $SignatureStatus = $Signature.Status
        $Version = (Get-Item -Path $Target).VersionInfo.ProductVersion

        if ($Signature.Status -ne 'Valid') {
            $Hints += 'Invalid signature; do not use this file.'
        }
        elseif ($Signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
            $Hints += 'Not signed by Microsoft; verify this file.'
        }
        elseif ($Package.Expected -and $Version -and -not $Version.StartsWith($Package.Expected)) {
            $Hints += "Expected $($Package.Expected); verify the download."
        }
    }

    # Verify the SHA256 hash where Microsoft publishes one.
    if ($Package.Sha256) {
        Write-Host "Hashing:       $($Package.FileName)"
        $Hash = (Get-FileHash -Path $Target -Algorithm SHA256).Hash
        if ($Hash -ne $Package.Sha256) {
            $Hints += 'SHA256 mismatch; do not use this file.'
        }
        elseif (-not $Hints) {
            $Hints += 'SHA256 verified.'
        }
    }

    [pscustomobject]@{
        File      = $Package.FileName
        Version   = $Version
        Signature = $SignatureStatus
        Note      = $Hints -join ' '
    }
}

# Create a placeholder for UCMA 4.0, which is installed from the Exchange ISO.
$UcmaFileName = 'UCMA-4.0-NOT-DOWNLOADED.txt'
$UcmaPlaceholder = Join-Path -Path $Path -ChildPath $UcmaFileName
$UcmaText = @"
Unified Communications Managed API (UCMA) 4.0 Core Runtime
==========================================================

This script does not download the UCMA 4.0 setup.

UCMA 4.0 is a prerequisite for the Exchange Server SE Mailbox role. It is
recommended to install it with UCMARunTimeSetup.exe from the UCMARedist
folder of the Exchange Server SE ISO. This package matches the Exchange
build and is the source Microsoft refers to in the Exchange prerequisites.

Example:

    `$Iso   = Mount-DiskImage -ImagePath '$(Join-Path -Path $Path -ChildPath 'ExchangeServerSE-x64.iso')' -PassThru
    `$Drive = (`$Iso | Get-Volume).DriveLetter
    & "`$(`$Drive):\UCMARedist\UCMARunTimeSetup.exe" -q

Install UCMA 4.0 after the Visual C++ Redistributables and the IIS URL
Rewrite Module, and before you start Exchange setup.

Created by Get-ExchangePrerequisites.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm')
"@

try {
    Set-Content -Path $UcmaPlaceholder -Value $UcmaText -Encoding UTF8 -ErrorAction Stop
    $Results = @($Results) + [pscustomobject]@{
        File      = $UcmaFileName
        Version   = $null
        Signature = 'n/a'
        Note      = 'Placeholder; install UCMA from the Exchange ISO.'
    }
}
catch {
    Write-Warning "Could not create $UcmaFileName. $($_.Exception.Message)"
}

$Results | Format-Table -AutoSize

Write-Host 'UCMA 4.0: Install the runtime with UCMARunTimeSetup.exe from the UCMARedist folder on the Exchange ISO.'
if ($IncludeExchangeSE) {
    Write-Host "Install $($ExchangeSE.SuKB) after setup completes, then run the Exchange Server Health Checker."
}
