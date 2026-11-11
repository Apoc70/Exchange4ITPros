# Get-ExchangePrerequisites

Downloads the installation prerequisites for the Exchange Server Subscription Edition (SE) Mailbox role to a local directory.

The script downloads the Visual C++ Redistributables 2012 and 2013 (x64), the IIS URL Rewrite Module 2.1, and the .NET Framework offline installer appropriate for the selected Windows Server version. It checks each downloaded or existing file's Authenticode signature and reports its product version.

## Requirements

- Windows PowerShell 5.1
- Network access to the Microsoft download URLs, or access through a proxy
- Permissions to create and write files in the destination directory

No PowerShell modules are required.

## Supported Windows Server versions

| Version | .NET Framework handling |
|---|---|
| Windows Server 2025 | .NET Framework 4.8.1 is included; no installer is downloaded |
| Windows Server 2022 | Downloads the .NET Framework 4.8.1 offline installer |
| Windows Server 2019 | Downloads the .NET Framework 4.8 offline installer |

If `-TargetOS` is not specified, the script detects the local Windows Server version from its build number. Specify `-TargetOS` when downloading prerequisites on a system other than the target server.

UCMA 4.0 is not downloaded by the script. Install it separately from the `UCMARedist` folder on the Exchange ISO.

## Usage

Run the script from PowerShell:

```powershell
.\Get-ExchangePrerequisites.ps1
```

By default, files are downloaded to `C:\Install`. To select a target version and destination:

```powershell
.\Get-ExchangePrerequisites.ps1 -TargetOS 2022 -Path 'D:\Install'
```

To use a proxy:

```powershell
.\Get-ExchangePrerequisites.ps1 -TargetOS 2022 -Path 'D:\Install' -Proxy 'http://proxy.example.com:8080'
```

The proxy uses the current user's default credentials. Use `-Force` to download files again when they already exist:

```powershell
.\Get-ExchangePrerequisites.ps1 -TargetOS 2022 -Force
```

## Parameters

| Parameter | Description |
|---|---|
| `-Path` | Destination directory. Defaults to `C:\Install`. |
| `-TargetOS` | Target Windows Server version: `2025`, `2022`, or `2019`. Defaults to automatic detection on the local system. |
| `-Proxy` | Optional proxy server URL. The current user's default credentials are used. |
| `-Force` | Downloads files again even if they already exist. |

## Results

The script displays a table with each file's name, product version, signature status, and any verification note. Review the signature status and notes before using the downloaded installers. Download failures are reported as warnings and included in the results.

## Script

[Get-ExchangePrerequisites.ps1](./Get-ExchangePrerequisites.ps1)
