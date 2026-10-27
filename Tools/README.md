# Tools

Scripts and utilities for Exchange for IT Pros helper scenarios.

## Categories

| Category | Description |
| --- | --- |
| Automation | Scripts that automate recurring operational tasks, such as password expiration reminders and scheduled task registration |
| Certificate management | Installation scripts for simple-acme that import ACME certificates and enable them for Exchange Server services |

## Solutions

| Folder | Description |
| --- | --- |
| [Send-PasswordExpirationReminder](Send-PasswordExpirationReminder/) | Sends password expiration reminder emails based on Active Directory password age and configurable notification thresholds. Contains the reminder script, the scheduled task registration helper, the configuration file, and supporting templates. |
| [simple-acme-scripts](simple-acme-scripts/) | Installation script for simple-acme (formerly win-acme / WACS) that imports an ACME certificate and enables it for Exchange Server services. Supports unattended operation with defined exit codes, optional distribution to all Exchange mailbox servers, optional restore of the transport certificate, and a daily log file. |

Each subfolder contains its own README with prerequisites, parameters, and usage examples.

## Requirements

Scripts in this folder generally require Windows PowerShell 5.1. Additional prerequisites depend on the solution.

| Solution | Additional requirements |
| --- | --- |
| Send-PasswordExpirationReminder | `ActiveDirectory` module and access to a configured SMTP server |
| simple-acme-scripts | Exchange Server 2016, 2019 or Subscription Edition, simple-acme, and Exchange RBAC permissions for certificate management |

Specific prerequisites, parameter details, and usage examples are documented in each script header and in the README of each subfolder.

## Links

- [Send-PasswordExpirationReminder.ps1](Send-PasswordExpirationReminder/Send-PasswordExpirationReminder.ps1)
- [Register-PasswordExpirationScheduledTask.ps1](Send-PasswordExpirationReminder/Register-PasswordExpirationScheduledTask.ps1)
- [ImportExchange.v3.ps1](simple-acme-scripts/ImportExchange.v3.ps1)
- simple-acme: https://github.com/simple-acme/simple-acme
- Exchange for IT Pros website: https://exchangeforitpros.blog/