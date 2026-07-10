# Password Expiration Reminder

Scripts for sending password expiration reminder emails based on Active Directory password age and configurable notification thresholds.

## Scripts

| Script | Description |
| --- | --- |
| [Send-PasswordExpirationReminder.ps1](Send-PasswordExpirationReminder.ps1) | Sends reminder emails for users whose passwords are approaching expiration |
| [Register-PasswordExpirationScheduledTask.ps1](Register-PasswordExpirationScheduledTask.ps1) | Registers a Windows Scheduled Task that runs the reminder script daily |

## Requirements

The reminder script generally requires Windows PowerShell 5.1 and the `ActiveDirectory` module. It also needs access to a configured SMTP server and read/write access to the local log, tracking, and template files. Specific configuration values, parameter details, and examples are documented in each script header.

## Supporting Files

- [PasswordExpirationReminder.config.json](PasswordExpirationReminder.config.json) - Sample configuration used by the reminder script
- [Templates/PasswordExpirationReminder-Notification.html](Templates/PasswordExpirationReminder-Notification.html) - HTML notification template
- [Templates/PasswordExpirationReminder-Notification.txt](Templates/PasswordExpirationReminder-Notification.txt) - Plain text notification template

## Links

- Exchange for IT Pros website: https://exchangeforitpros.blog/
