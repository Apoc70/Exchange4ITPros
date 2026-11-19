# Exchange Server Maintenance

This section contains scripts and operational checklists for maintaining on-premises Microsoft Exchange Server environments.

## Scripts

| Script | Purpose | Required Modules | Minimum Version |
|---|---|---|---|
| [Export-MessageQueue](Export-MessageQueue/README.md) | Suspends and exports messages from an Exchange transport queue; optionally deletes suspended messages after a successful export | Exchange Management Shell, GlobalFunctions | Exchange Server 2016+, Windows Server 2016+ |
| [Get-PublicFolderCustomFormItems](Get-PublicFolderCustomFormItems/README.md) | Counts items of a specified type in public folders and exports the results to CSV | Exchange Management Shell or Exchange Online PowerShell module | Not specified |

## Checklists

- [Maintenance checklists](checklists/README.md) includes post-maintenance checklists in English and German.

## Notes

- Refer to each script's README for usage examples and additional requirements.
- Copy checklist templates to your own tracking location before completing them; do not edit the repository templates.
- Additional scripts and maintenance resources will be added to this folder over time.
