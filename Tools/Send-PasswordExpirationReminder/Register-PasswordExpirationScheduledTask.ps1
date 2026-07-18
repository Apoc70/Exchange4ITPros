#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Registers a Windows Scheduled Task that runs the password expiration reminder script
    once per day.

.DESCRIPTION
    Creates a scheduled task on the local server that executes
    Send-PasswordExpirationReminder.ps1 daily at a configurable time. The task can run
    under the local SYSTEM account or under a dedicated service account.

    If a dedicated service account is used (via -Credential), that account must be granted
    the 'Log on as a batch job' (SeBatchLogonRight) user right on the local server. This can
    be assigned via Local Security Policy (secpol.msc) under
    Local Policies -> User Rights Assignment, or via Group Policy in larger environments.

    This script only registers the task. It does not itself send any notification emails.

.PARAMETER ScriptPath
    Full path to Send-PasswordExpirationReminder.ps1. Defaults to a file with that name
    located in the same directory as this script.

.PARAMETER ConfigurationFilePath
    Full path to the JSON configuration file to be passed to the reminder script. Defaults
    to 'PasswordExpirationReminder.config.json' located in the same directory as this script.

.PARAMETER TaskName
    Name of the scheduled task to create. Defaults to 'Password Expiration Reminder'.

.PARAMETER TriggerTime
    The daily time at which the task should run. Defaults to 07:00.

.PARAMETER Credential
    Optional credential of a dedicated service account to run the task as. If omitted, the
    task is registered to run under the local SYSTEM account.

.EXAMPLE
    .\Register-PasswordExpirationScheduledTask.ps1

    Registers the scheduled task to run daily at 07:00 under the local SYSTEM account, using
    the default script and configuration file paths.

.EXAMPLE
    .\Register-PasswordExpirationScheduledTask.ps1 -TriggerTime '06:30' -Credential (Get-Credential)

    Registers the scheduled task to run daily at 06:30 under a dedicated service account.

.EXAMPLE
    .\Register-PasswordExpirationScheduledTask.ps1 -WhatIf

    Shows what would happen without actually registering the scheduled task.

.NOTES
    Author: Thomas Stensitzki
    Target platform: Exchange Server (on-premises), any domain-joined member server
    Required modules: ScheduledTasks (built-in on Windows Server 2012 and later)
    PowerShell compatibility: Windows PowerShell 5.1

    Required permissions for the executing account (running this setup script):
    - Local Administrator rights on the server, since creating a Scheduled Task requires
      administrative privileges.

    Required permissions for the account the task runs as (SYSTEM or service account):
    - See the .NOTES section of Send-PasswordExpirationReminder.ps1 for the permissions
      required to execute the actual reminder logic (Active Directory read access, SMTP
      access, local file access).
    - If a dedicated service account is used instead of SYSTEM, that account additionally
      requires the 'Log on as a batch job' (SeBatchLogonRight) user right on this server.

    Change log:
    1.0.0 - Initial version
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ScriptPath = (Join-Path -Path $PSScriptRoot -ChildPath 'Send-PasswordExpirationReminder.ps1'),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationFilePath = (Join-Path -Path $PSScriptRoot -ChildPath 'PasswordExpirationReminder.config.json'),

    [Parameter(Mandatory = $false)]
    [string]$TaskName = 'Password Expiration Reminder',

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^([01]?[0-9]|2[0-3]):[0-5][0-9]$')]
    [string]$TriggerTime = '07:00',

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential
)

try {
    if (-not (Test-Path -Path $ScriptPath)) {
        throw "Reminder script not found at path '$ScriptPath'."
    }

    if (-not (Test-Path -Path $ConfigurationFilePath)) {
        throw "Configuration file not found at path '$ConfigurationFilePath'."
    }

    $powerShellExecutable = Join-Path -Path $PSHOME -ChildPath 'powershell.exe'

    $taskArguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -ConfigurationFilePath "{1}"' -f $ScriptPath, $ConfigurationFilePath

    $taskAction = New-ScheduledTaskAction -Execute $powerShellExecutable -Argument $taskArguments

    $triggerDateTime = [datetime]::ParseExact($TriggerTime, 'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
    $taskTrigger = New-ScheduledTaskTrigger -Daily -At $triggerDateTime

    $taskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Hours 1)

    if ($Credential) {
        $taskPrincipal = New-ScheduledTaskPrincipal -UserId $Credential.UserName -LogonType Password -RunLevel Highest
    }
    else {
        $taskPrincipal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    }

    $registerParameters = @{
        TaskName    = $TaskName
        Action      = $taskAction
        Trigger     = $taskTrigger
        Settings    = $taskSettings
        Principal   = $taskPrincipal
        Description = 'Sends password expiration reminder emails based on configured thresholds. Managed by Send-PasswordExpirationReminder.ps1.'
        ErrorAction = 'Stop'
    }

    if ($PSCmdlet.ShouldProcess($TaskName, 'Register scheduled task')) {

        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Write-Verbose -Message "Scheduled task '$TaskName' already exists and will be updated."
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        }

        if ($Credential) {
            Register-ScheduledTask @registerParameters -Password $Credential.GetNetworkCredential().Password | Out-Null
        }
        else {
            Register-ScheduledTask @registerParameters | Out-Null
        }

        Write-Output -InputObject "Scheduled task '$TaskName' registered successfully. Daily trigger at $TriggerTime."
    }
}
catch {
    Write-Error -Message "Failed to register scheduled task: $($_.Exception.Message)"
    throw
}