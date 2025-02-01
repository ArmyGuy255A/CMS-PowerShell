# CMS Meeting Room Management Scripts

This PowerShell script collection is designed to manage Cisco Meeting Server (CMS) meeting rooms for an organization. The main script, `Manage-CMSMeetings.ps1`, automates meeting room scheduling and ensures compliance with organizational policies.

## Prerequisites

1. **PowerShell 5.1 or later** (Windows built-in)
2. **An authentication key (`authkey.key`)** for secure access
3. **A properly configured `config.json`** file
4. **CMS API access** with the appropriate credentials

## Installation

1. Clone or download this repository to your local machine.
2. Ensure that `config.json` is customized for your unit.
3. Generate a new authentication key before running the script.

## Generating a New Authentication Key

The authentication key is required to securely authenticate requests to CMS. Follow these steps to generate a new key:

1. Open PowerShell in the script directory.
2. Run the following command to import helper functions:
   ```powershell
   . .\CMS.functions.ps1
   ```
3. Generate the authentication key:
   ```powershell
   New-AuthKey
   ```
4. A new `authkey.key` file will be created in the script directory. Keep this file secure.

## Configuring `config.json`

The `config.json` file contains key settings for managing CMS meetings. Below is a breakdown of the configurable parameters:

### **AdminMailboxes**
A list of administrator email addresses that will receive notifications.
```json
"AdminMailboxes": [
    "admin1@contoso.com",
    "group-mailbox@contoso.com"
]
```

### **CiscoMeetingServer Settings**
These parameters control meeting behaviors and server connections.

- **`AdditionalNotificationEmails`**: Email recipients for meeting notifications.
- **`Servers`**: CMS server URLs (modify for your environment).
- **`BannerText` & `EndingBannerText`**: Text displayed during meetings.
- **`MeetingParticipantLimit`**: Maximum participants allowed in a meeting.
- **`MeetingPreparationTimeInMinutes`**: Time allocated before a meeting starts.
- **`StandardMeetingDurationInMinutes`**: Default meeting duration.
- **`ExtendedMeetingDurationInMinutes`** & **`LongMeetingDurationInMinutes`**: Lengths for special meetings.
- **`MeetingWarningThresholdInMinutes`**: Time before a warning is displayed.
- **`ExtendedMeetings`**, **`LongMeetings`**, and **`ExemptMeetings`**: Rooms assigned to different meeting durations or exempt from limits.

#### Example:
```json
"Automation": {
    "CiscoMeetingServer": {
        "AdditionalNotificationEmails": [
            "meeting-managers@contoso.com"
        ],
        "Servers": [
            "https://cms-server1.company.com:445",
            "https://cms-server2.company.com:445"
        ],
        "BannerText": "CLASSIFICATION // RELEASABILITY",
        "EndingBannerText": "ENDING SOON",
        "MeetingParticipantLimit": 50,
        "MeetingPreparationTimeInMinutes": 20,
        "StandardMeetingDurationInMinutes": 60,
        "ExtendedMeetingDurationInMinutes": 120,
        "LongMeetingDurationInMinutes": 180,
        "MeetingWarningThresholdInMinutes": 10,
        "ExtendedMeetings": [
            "EXECUTIVE ROOM",
            "BOARD ROOM"
        ],
        "LongMeetings": [
            "CONFERENCE ROOM A",
            "CONFERENCE ROOM B"
        ],
        "ExemptMeetings": [
            "WAR ROOM",
            "SECURITY BRIEFING ROOM"
        ]
    }
}
```

Modify these values based on your organization’s meeting policies.

## Running the Script

Once the `config.json` file is set up and the authentication key is generated, execute the main script:

```powershell
.\Manage-CMSMeetings.ps1
```

To automate execution, you can register the script as a scheduled task.

## Setting Up a Scheduled Task

To run `Manage-CMSMeetings.ps1` every minute as a scheduled task:

1. Open PowerShell as Administrator.
2. Run the following command:

```powershell
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"$PSScriptRoot\Manage-CMSMeetings.ps1`""
$Trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 1) -Once -At (Get-Date)
$Principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount
Register-ScheduledTask -TaskName "ManageCMSMeetings" -Action $Action -Trigger $Trigger -Principal $Principal -Description "Automates CMS Meeting Room Management"
```

This creates a scheduled task named **ManageCMSMeetings**, which runs every minute.

## Troubleshooting

- If the script fails to run, check:
  - PowerShell execution policies: `Get-ExecutionPolicy`
  - CMS server connectivity: `Test-NetConnection <cms-server-url>`
  - Authentication key validity
  - The format of `config.json`

For additional support, contact your CMS administrator.

---
**Author**: Phil Dieppa
**Version**: 1.0  
**Last Updated**: 15 December 2024
