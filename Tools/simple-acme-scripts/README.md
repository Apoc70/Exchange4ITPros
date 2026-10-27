# ImportExchange.v3.ps1

Installation script for [simple-acme](https://github.com/simple-acme/simple-acme) (formerly win-acme / WACS) that imports an ACME certificate and enables it for Exchange Server services. Version 3 is built for unattended operation: it writes no console output, returns defined exit codes, and logs every action to a file.

## Features

- Enables the certificate for IIS, SMTP, POP, IMAP and, on Exchange Server 2016, UM and UMCallRouter.
- Optionally distributes the certificate to all Exchange mailbox servers of the organization. The certificate is only imported on servers where it does not exist yet, so reusing a certificate that was issued minutes ago works without errors.
- Optionally restores the transport certificate that was configured on each server before the ACME certificate was enabled.
- Optionally removes older certificates with the same subject.
- Silent by default. Console output only with `-DebugOn`.
- Defined exit codes for the calling automation.
- Daily log file in the script directory.

## Requirements

| Requirement | Details |
|---|---|
| Exchange Server | 2016, 2019 or Subscription Edition |
| PowerShell | 5.1 (enforced with `#Requires -Version 5.1`) |
| simple-acme | Installation step of type "Script" |
| Execution | On an Exchange server with the Mailbox role |

## Permissions

### Exchange RBAC

Use a dedicated role group with the management role "Exchange Server Certificates" and limit it to the servers the script manages.

For the local server only:

```powershell
New-ManagementScope -Name 'ACME EX01' -ServerRestrictionFilter { Name -eq 'EX01' }
```

For all mailbox servers (`DeployToAllMailboxServers = 1`):

```powershell
New-ManagementScope -Name 'ACME Mailbox Servers' -ServerRestrictionFilter { ServerRole -like '*Mailbox*' }
```

Then create the role group:

```powershell
New-RoleGroup -Name 'ACME Certificate Management' `
  -Roles 'Exchange Server Certificates' `
  -CustomConfigWriteScope 'ACME Mailbox Servers' `
  -Members svc-acme
```

`RestoreTransportCertificate = 1` additionally requires `Get-TransportService`. Check whether the role contains the cmdlet and add a read role such as "View-Only Configuration" if it does not:

```powershell
Get-ManagementRoleEntry 'Exchange Server Certificates\Get-TransportService'
```

The built-in role group "Server Management" works as well. Avoid "Organization Management" for automated tasks.

### Additional requirements

- `RemotePowerShellEnabled` must be `$true` for the executing account.
- The account must be a member of the local Administrators group. simple-acme needs this to write to `LocalMachine\My` and to set private key ACLs.
- The account needs write access to the script directory for the log file.
- The simple-acme scheduled task runs as SYSTEM by default. The Exchange remote session then authenticates as the computer account, which needs the RBAC permissions described above. Alternatively, run the task as a dedicated service account.

## Parameters

All parameters are positional, because simple-acme passes the script arguments as a single string.

| Position | Name | Default | Description |
|---|---|---|---|
| 0 | `NewCertThumbprint` | | Thumbprint of the certificate to enable. |
| 1 | `ExchangeServices` | | Comma-separated list: `IIS`, `SMTP`, `POP`, `IMAP`, `UM`, `UMCallRouter`. |
| 2 | `LeaveOldExchangeCerts` | `1` | `1` keeps older certificates with the same subject, `0` removes them. |
| 3 | `CertificatePath` | | Full path to the PFX file provided by simple-acme. |
| 4 | `PfxPassword` | | Password of the PFX file. Required together with `CertificatePath`. |
| 5 | `FriendlyName` | | Optional friendly name for the imported certificate. |
| 6 | `DeployToAllMailboxServers` | `0` | `0` local server only, `1` all Exchange mailbox servers. |
| 7 | `RestoreTransportCertificate` | `0` | `1` restores the previous transport certificate on each server. |
| | `-DebugOn` | | Writes all log entries and error details to the console. |

To skip a position and set a later one, pass an empty string, for example `''` for positions 3 to 5.

## simple-acme configuration

simple-acme calls the script directly as installation step of type "Script". The arguments are passed with `--scriptparameters` as a single string. simple-acme replaces the placeholders in curly braces before the call.

```powershell
$sa = "$env:ProgramFiles\simple-acme"
$script = "$sa\scripts\ImportExchange.v3.ps1"

& "$sa\wacs.exe" --source manual `
  --host mail.squeakylobster.dev,autodiscover.squeakylobster.dev `
  --validation cloudflare --cloudflareapitoken '<CloudflareApiToken>' `
  --store certificatestore --certificatestore My `
  --acl-fullcontrol "network service,administrators" `
  --installation script `
  --script $script `
  --scriptparameters "'{CertThumbprint}' 'IIS,SMTP' 1 '{CacheFile}' '{CachePassword}' '{CertFriendlyName}' 1 1" `
  --accepttos --emailaddress admin@squeakylobster.dev
```

This example enables the certificate for IIS and SMTP on all mailbox servers, restores the previous transport certificate on each server, and keeps older certificates. On the first run, simple-acme also offers to create a scheduled task for automatic renewal. Each renewal calls the script again with the same arguments.

| Placeholder | Script parameter |
|---|---|
| `{CertThumbprint}` | `NewCertThumbprint` |
| `{CacheFile}` | `CertificatePath` |
| `{CachePassword}` | `PfxPassword` |
| `{CertFriendlyName}` | `FriendlyName` |

Put every value in single quotes and use `''` for positions you do not use. Add `--test` while you set up the renewal to run simple-acme in test mode.

> [!IMPORTANT]
> Never store the Cloudflare API token in a script or repository. Use a token that is limited to `Zone:DNS:Edit` for the required zone.

> [!TIP]
> Verify once that the exit code of the script reaches simple-acme. A run with an invalid service name must result in exit code 2.

## Examples

Enable a certificate that already exists in the local store:

```powershell
.\ImportExchange.v3.ps1 <Thumbprint> IIS,SMTP
```

Import from PFX, distribute to all mailbox servers, restore the transport certificate, and remove older certificates:

```powershell
.\ImportExchange.v3.ps1 <Thumbprint> IIS,SMTP 0 C:\Certs\cert.pfx <PfxPassword> 'Exchange ACME' 1 1
```

Test run with console output:

```powershell
.\ImportExchange.v3.ps1 <Thumbprint> IIS,SMTP 1 '' '' '' 1 0 -DebugOn
$LASTEXITCODE
```

## Exit codes

| Code | Meaning |
|---|---|
| 0 | Success |
| 1 | Unexpected error |
| 2 | Invalid parameters |
| 3 | Certificate store unavailable |
| 4 | Certificate not found or without private key |
| 5 | PFX import failed |
| 6 | Exchange session or permissions |
| 7 | `Enable-ExchangeCertificate` failed on the local server |
| 8 | Cleanup of old certificates failed |
| 9 | Deployment to at least one additional mailbox server failed |
| 10 | Transport certificate could not be read or restored |

When several partial failures occur, the script returns the most severe one in the order 9, 10, 8.

Exit codes 3 to 7 stop the script on the local server before or during the first change. Exit codes 8 to 10 are partial failures: the script continues with the remaining servers and reports the failure at the end.

## Logging

The script writes to `ImportExchange_yyyyMMdd.log` in the script directory. Each line has a timestamp and one of the levels INFO, WARN or ERROR. Server names are always logged in upper case. The PFX password is never written to the log.

Example of a run with distribution and restore of the transport certificate:

```text
2026-10-27 12:45:02 [INFO] Run started on VM-EXCH-001.
2026-10-27 12:45:02 [INFO] Thumbprint: D52E4100.... Services: IIS,SMTP. Scope: all mailbox servers. Remove old certificates: False. Restore transport certificate: True.
2026-10-27 12:45:03 [INFO] Certificate D52E4100... already present in the local store on VM-EXCH-001. Import skipped.
2026-10-27 12:45:08 [INFO] Current transport certificate on VM-EXCH-001: 8A1F....
2026-10-27 12:45:11 [INFO] Certificate D52E4100... enabled on VM-EXCH-001 for services: IIS,SMTP.
2026-10-27 12:45:13 [INFO] Transport certificate 8A1F... restored on VM-EXCH-001.
2026-10-27 12:45:15 [INFO] Certificate D52E4100... already present on VM-EXCH-003. Import skipped.
2026-10-27 12:45:22 [INFO] Run finished with exit code 0.
```

Warnings of `Enable-ExchangeCertificate`, for example about connectors whose FQDN does not match the certificate, are written to the log with level WARN.

## Restoring the transport certificate

`Enable-ExchangeCertificate -Services SMTP -Force` replaces the default certificate of the transport service (`InternalTransportCertificateThumbprint`) without asking. This is usually not what you want, because the default certificate is used for internal transport between Exchange servers.

With `RestoreTransportCertificate = 1`, the script reads the current transport certificate of each server before the ACME certificate is enabled and enables it again afterwards. The ACME certificate keeps the SMTP flag and is still selected for TLS through FQDN matching on the connectors.

The restore is skipped when SMTP is not part of `ExchangeServices`, when no transport certificate was configured, or when the ACME certificate already was the transport certificate. A restored transport certificate is never removed by the cleanup of older certificates.

## Distribution without PFX file

Without `CertificatePath` and `PfxPassword`, the script exports the certificate temporarily from the local store to distribute it. This requires an exportable private key, which simple-acme does not set by default. Pass the PFX parameters when you use `DeployToAllMailboxServers = 1`. The PFX data is only prepared when at least one server needs an import.

## Known issues

- For remote servers, `Get-ExchangeCertificate -Server` does not reliably report `HasPrivateKey` through the Exchange remote session. The script therefore only checks whether the certificate exists on a remote server. A missing private key is reported by `Enable-ExchangeCertificate` and results in exit code 9.
- If PowerShell 5.1 is not available, the script stops with exit code 1 before any action and without a log entry.

## Migration from v2

Version 3 is not a drop-in replacement. Review the simple-acme configuration before switching.

| Topic | v2 | v3 |
|---|---|---|
| Script name | `ImportExchange.v2.ps1` | `ImportExchange.v3.ps1` |
| `RenewalId` | Documented parameter | Removed |
| `CertificatePath` | Documented as directory | Full path to the PFX file |
| Console output | `Write-Host` and `Write-Error` | None, unless `-DebugOn` is used |
| Error handling | Rethrows the exception | Exit codes 0 to 10 |
| Mandatory parameters | Interactive prompt when missing | Exit code 2 |
| Positions 6 and 7 | Not available | `DeployToAllMailboxServers`, `RestoreTransportCertificate` |
| Logging | None | Daily log file in the script directory |

The positions 0 to 5 are unchanged, so an existing v2 argument string works with v3. The certificate is then enabled on the local server only, and the transport certificate is not restored.

## Version history

| Version | Changes |
|---|---|
| 3.0 | Exit codes and silent operation, error handling for the certificate provider, optional distribution to all mailbox servers, import only where the certificate is missing, optional restore of the transport certificate, log file, PowerShell 5.1 required, Windows Server 2008 R2 check and `RenewalId` removed. |
| 2.1 | Corrected PFX parameter binding and import, validation, safer cleanup, session cleanup. |
| 2.0 | Initial release. |

## Links

- [Enable-ExchangeCertificate](https://learn.microsoft.com/powershell/module/exchangepowershell/enable-exchangecertificate?WT.mc_id=M365-MVP-5003086)
- [Import-ExchangeCertificate](https://learn.microsoft.com/powershell/module/exchangepowershell/import-exchangecertificate?WT.mc_id=M365-MVP-5003086)
- [Get-TransportService](https://learn.microsoft.com/powershell/module/exchangepowershell/get-transportservice?WT.mc_id=M365-MVP-5003086)
- [simple-acme: Install Script](https://github.com/simple-acme/simple-acme/wiki/Install-Script)
- [simple-acme: Example Scripts](https://github.com/simple-acme/simple-acme/wiki/Example-Scripts)
- [Exchange für IT Pros](https://exchangeforitpros.blog/)