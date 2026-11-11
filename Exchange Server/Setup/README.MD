# Exchange Server SE - Setup

This folder contains scripts and tools for preparing and configuring Exchange Server Subscription Edition (SE).

The scripts are organized into subfolders. See each script's README for detailed usage and requirements where available.

## Scripts

| Script | Purpose | Requirements |
|---|---|---|
| [Get-ExchangePrerequisites](./Get-ExchangePrerequisites/README.MD) | Downloads Exchange Server SE Mailbox role prerequisites, checks their Authenticode signatures, and reports file versions. | Windows PowerShell 5.1; no additional modules required. Supports Windows Server 2019, 2022, and 2025. |
| [Update-ClientAccessServerSCP.ps1](./Update-ClientAccessServerSCP/Update-ClientAccessServerSCP.ps1) | Intended to update the Client Access Server Service Connection Point (SCP). | The script file is currently empty; implementation, usage, and requirements are not yet available. |

## Notes

- Refer to each script's README and `.NOTES` section for target platform, required modules, and revision history.
- Additional scripts and updates may be added to this folder over time.
