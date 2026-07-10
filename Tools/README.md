# Tools

Scripts and utilities for Exchange for IT Pros helper scenarios.

## Categories

| Category | Description |
| --- | --- |
| Automation | Scripts that automate recurring operational tasks, such as password expiration reminders and scheduled task registration |

The current toolset in this folder is organized under the password expiration reminder solution. The subfolder contains the reminder script, the scheduled task registration helper, the configuration file, and supporting templates.

## Requirements

Scripts in this folder generally require Windows PowerShell 5.1. The password expiration reminder solution also requires the `ActiveDirectory` module and access to a configured SMTP server. Specific prerequisites, parameter details, and usage examples are documented in each script header.

## Links

- [Send-PasswordExpirationReminder.ps1](Send-PasswordExpirationReminder/Send-PasswordExpirationReminder.ps1)
- [Register-PasswordExpirationScheduledTask.ps1](Send-PasswordExpirationReminder/Register-PasswordExpirationScheduledTask.ps1)
- Exchange for IT Pros website: https://exchangeforitpros.blog/
