## Description
Converts IMCEAEX addresses (used in Exchange for legacy reasons) to X500 address format. Optionally copies the result to clipboard.

## Syntax
```powershell
ConvertTo-X500Address.ps1 [-IMCAEX] <string> [-ToClipboard]
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| IMCAEX | String | True | The IMCEAEX address to be converted |
| ToClipboard | Switch | False | If specified, copies the X500 address to clipboard |

## Examples

```powershell
.\ConvertTo-X500Address.ps1 -IMCAEX "IMCEAEX-_O=EXAMPLE_OU=EXAMPLE_CN=RECIPIENT"
```

Copy to clipboard:
```powershell
.\ConvertTo-X500Address.ps1 -IMCAEX "IMCEAEX-_O=EXAMPLE_OU=EXAMPLE_CN=RECIPIENT" -ToClipboard
```

## Output
Returns the X500 address string in the format `X500:<decoded address>`

## Requirements
- PowerShell 5.0 or later

## Notes
- IMCEAEX addresses use URL encoding with underscores that must be decoded
- X500 addresses are legacy Exchange identifiers used for mailbox migrations and coexistence

