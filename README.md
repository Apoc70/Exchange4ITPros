# Exchange for IT Pros

Central repository for **Exchange for IT Pros**, containing PowerShell scripts and code samples for everyday challenges when working with Exchange Server, Exchange Hybrid, and Exchange Online.

## About this repository

This repository collects community scripts used and shared as part of the Exchange for IT Pros project. Scripts are organized by Microsoft 365 workload rather than by the underlying technology, so you can quickly find what applies to your environment.

## Repository structure

| Folder | Scope |
| --- | --- |
| [Exchange Server](./Exchange%20Server/) | Scripts for on-premises Exchange Server and Exchange Hybrid, using the Exchange Management Shell |
| [Exchange Online](./Exchange%20Online/) | Scripts for Exchange Online, using the Exchange Online PowerShell module |
| [Entra](./Entra/) | Scripts for Entra ID, mostly using the Microsoft Graph PowerShell SDK |
| [SharePoint Online](./SharePoint%20Online/) | Scripts for SharePoint Online, using PnP PowerShell |

Each workload folder contains its own README with an overview of the available categories and scripts. Each script lives in its own subfolder together with a dedicated README describing purpose, parameters, and prerequisites.

## Prerequisites

Requirements vary per script. Every script includes a comment based help block stating the target platform and required PowerShell modules. Check the script header before running it.

## Contributing

Feedback, issues, and pull requests are welcome. Please keep contributions consistent with the existing folder structure and documentation style.

## License

This project is licensed under the MIT License. See the [LICENSE](./LICENSE) file for details.

## Links

- Exchange for IT Pros website: https://exchangeforitpros.blog/