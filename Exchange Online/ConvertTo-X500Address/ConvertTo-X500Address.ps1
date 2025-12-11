<#
.SYNOPSIS
Converts an IMCEAEX address to X500 format.

.DESCRIPTION
This script takes an IMCEAEX address (used in Exchange for legacy reasons) and converts it to the X500 address format.
It can optionally copy the resulting X500 address to the clipboard for easy use.

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.LINK 
https://exchangeforitpros.blog  

.PARAMETER IMCAEX
The IMCEAEX address to be converted.

.PARAMETER ToClipboard
If specified, the resulting X500 address will be copied to the clipboard.   

.EXAMPLE
.\ConvertTo-X500Address.ps1 -IMCAEX "IMCEAEX-_O=EXAMPLE_OU=EXAMPLE_CN=RECIPIENT"
Converts the given IMCEAEX address to X500 format and outputs it.

#>
param(
    [Parameter(Mandatory=$true, HelpMessage="IMCAEX address to convert")]
    [string]$IMCAEX,
    
    [switch]$ToClipboard
)

# Decode the IMCEAEX address
$cleanAddress = ([System.Uri]::UnescapeDataString($IMCEAEX.replace("_","/").replace("IMCEAEX-","")))

# Build X500 format
$X500Address = "X500:$cleanAddress"

# Output result
Write-Output $X500Address

# Copy to clipboard if switch is used
if ($ToClipboard) {
    $X500Address | Set-Clipboard
    Write-Host "✓ Address copied to clipboard" -ForegroundColor Green
}