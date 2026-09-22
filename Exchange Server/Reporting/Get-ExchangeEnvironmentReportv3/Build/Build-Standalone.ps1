<#
.SYNOPSIS
Builds a standalone version of Get-ExchangeEnvironmentReport.ps1
 
.DESCRIPTION
Reads the source script, CSS file, and JSON mapping file
from the Source folder and embeds CSS + JSON directly
into the PowerShell script.
 
The resulting standalone script is written to the
repository root folder.
 
Author: Thomas Stensitzki
#>

[CmdletBinding()]
param()

# ---------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------

$RepoRoot = Split-Path -Path $PSScriptRoot -Parent

$SourceScript = Join-Path $RepoRoot 'Source\Get-ExchangeEnvironmentReport.ps1'
$CssFile = Join-Path $RepoRoot 'Source\EnvironmentReport.css'
$JsonFile = Join-Path $RepoRoot 'Source\ExchangeVersionMappings.json'

$TargetScript = Join-Path $RepoRoot 'Get-ExchangeEnvironmentReport.ps1'

# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------

foreach ($File in @(
        $SourceScript,
        $CssFile,
        $JsonFile
    )) {
    if (-not (Test-Path $File)) {
        throw "Required file not found: $File"
    }
}

# ---------------------------------------------------------------------
# Load files
# ---------------------------------------------------------------------

Write-Verbose "Loading source files"

$ScriptContent = Get-Content $SourceScript -Raw -Encoding UTF8
$CssContent = Get-Content $CssFile -Raw -Encoding UTF8
$JsonContent = Get-Content $JsonFile -Raw -Encoding UTF8

# ---------------------------------------------------------------------
# Read mapping metadata
# ---------------------------------------------------------------------

$Mapping = $JsonContent | ConvertFrom-Json

$MappingVersion = $Mapping.MappingVersion
$MappingDate = $Mapping.LastUpdated

# ---------------------------------------------------------------------
# Extract script version
# ---------------------------------------------------------------------

$ScriptVersion = 'Unknown'

if ($ScriptContent -match "\`$ScriptVersion\s*=\s*'([^']+)'") {
    $ScriptVersion = $Matches[1]
}

# ---------------------------------------------------------------------
# Build resource blocks
# ---------------------------------------------------------------------

$CssBlock = @"
`$EmbeddedCss = @'
$CssContent
'@
"@

$JsonBlock = @"
`$EmbeddedVersionMappingJson = @'
$JsonContent
'@
"@

# ---------------------------------------------------------------------
# Replace placeholders
# ---------------------------------------------------------------------
$StartMarker = '# BUILD-REPLACE-START'
$EndMarker = '# BUILD-REPLACE-END'

$StartIndex = $ScriptContent.IndexOf($StartMarker)
$EndIndex = $ScriptContent.IndexOf($EndMarker)

if ($StartIndex -lt 0) {
    throw "Start marker not found."
}

if ($EndIndex -lt 0) {
    throw "End marker not found."
}

$EndIndex += $EndMarker.Length

$ScriptContent = $ScriptContent.Substring(0, $StartIndex) + "`r`n" + $CssBlock + "`r`n`r`n" + $JsonBlock + "`r`n" + $ScriptContent.Substring($EndIndex)

# ---------------------------------------------------------------------
# Undo regex escaping
# ---------------------------------------------------------------------

# $ScriptContent = [Regex]::Unescape($ScriptContent)

# ---------------------------------------------------------------------
# Build header
# ---------------------------------------------------------------------

$Header = @"
<#
==================== GENERATED FILE =================================

Exchange Environment Report

Script Version  : $ScriptVersion
Mapping Version : $MappingVersion
Mapping Updated : $MappingDate
Build Date : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

Do not edit manually.
Edit the files in the Source folder instead.
=====================================================================
#>

"@

# ---------------------------------------------------------------------
# Write output
# ---------------------------------------------------------------------

$FinalScript = $Header + $ScriptContent

Set-Content `
    -Path $TargetScript `
    -Value $FinalScript `
    -Encoding UTF8

Write-Host ''
Write-Host 'Build completed successfully'
Write-Host "Output : $TargetScript"
Write-Host "Script : $ScriptVersion"
Write-Host "Mapping: $MappingVersion"