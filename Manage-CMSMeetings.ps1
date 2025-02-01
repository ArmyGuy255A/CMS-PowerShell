#requires -version 5

#Allow the script to accept $configData from any other script in the toolkit
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    $configData
)

Set-Location $PSScriptRoot

function Get-ToolkitFile {
    [CmdletBinding()]
    Param (
        [Parameter()]
        [string]$Directory = ".",
        [Parameter()]
        [string]$File,
        [Parameter()]
        [switch]$RecurseUp
    )

    # Check for the file in the current directory
    $tkFile = Get-ChildItem -Path $Directory -Filter $File -Depth 1 -ErrorAction Ignore

    if ($null -ne $tkFile) {
        return $tkFile
    }
    elseif ($RecurseUp) {
        $currentPath = Get-Item -Path $Directory
        # Stop recursion if we are at the root directory
        if ($currentPath.PSParentPath -eq $null) {
            Write-Warning "Reached the root directory. File not found: $File"
            return $null
        }
        return Get-ToolkitFile -Directory $currentPath.Parent.FullName -File $File -RecurseUp
    }

    # If the file wasn't found and recursion is not enabled, return $null
    Write-Warning "File not found: $File in $Directory"
    return $null
}

# Note, ensure RecurseUp is enabled if this function is called below the root directory
if ($null -eq $configData) {
    $configData = Get-ToolkitFile -File "config.json" -RecurseUp 
    if ($null -eq $configData) {
        return -1
    }
    $configData = $configData | Get-Content -Encoding utf8 | ConvertFrom-Json 
}

#This imports the common libraries for use throughout every script.
# $stCommon = Get-ToolkitFile -File "Libraries/STCommon.ps1" -RecurseUp
# . $stCommon.FullName
. .\CMS.functions.ps1

# $logFileBaseName = $MyInvocation.MyCommand.Name.Replace(".ps1", "")
# if ([string]::IsNullOrEmpty($logFileBaseName)) {
#     $logFileBaseName = "Default"
# }

# $LogFile = ("{0}\{1}\{2}.log" -f $configData.LogDirectory, (Get-Location | Get-Item).Name, $logFileBaseName)
# Write-STAlert -Message ("Running the script" -f $adObjects.Count, $mapping.ADObjectType) -OutFile $LogFile

#Stop the script if something fails
$ErrorActionPreference = 'Stop'
#endregion


function Send-AutomationMessage($Body) {
    $mailServer = $configData.ExchangeServers | Where-Object DefaultServer | Select-Object -ExpandProperty Name
    $title = "AUTOMATION - Manage-CMSMeetings.ps1 complete"
    $message = "The automation script Manage-CMSMeetings.ps1 completed running on $($env:COMPUTERNAME). <br /> <br />"
    $message += $Body
    $adminMailboxes = $configData.AdminMailboxes + $configData.Automation.CiscoMeetingServer.AdditionalNotificationEmails
    Send-MailMessage -To $adminMailboxes -From ("{0}@{1}" -f $env:COMPUTERNAME, $env:USERDNSDOMAIN) -Subject $title -Body $message -SmtpServer $mailServer -BodyAsHtml    
    
}

# EMAIL SETTINGS
$warnings = @()
$disconnectedCalls = @()
$updatedCoSpaces = @()

# CMS Settings
$cmsAuthKey = Read-AuthKey
$cmsHeaders = Get-CMSHeaders -authKey $cmsAuthKey

# Update the call profile settings if they need to be changed
$callProfileSettings = New-CallProfileSettingsObject
$callProfileSettings.messageBannerText = $configData.Automation.CiscoMeetingServer.BannerText
$callProfileSettings.participantLimit = $configData.Automation.CiscoMeetingServer.MeetingParticipantLimit


# Loop through each CMS server and manage the meetings
foreach ($cmsServer in $configData.Automation.CiscoMeetingServer.Servers) {

    # Check if the cmsServer is reachable
    
    $result = Test-NetConnection $($cmsServer.Replace("//", "").split(":")[1]) -port 445
    if (!$result.TcpTestSucceeded) {
        Write-Host "$cmsServer Unreachable" -ForegroundColor Red
        continue
    }

    $callProfileId = Get-CallProfileIds -cmsUrl $cmsServer -headers $cmsHeaders

    # BUG: We're only using a single Call Profile, but we're iterating over all of them
    $callProfile = Get-CallProfile -cmsUrl $cmsServer -headers $cmsHeaders -callProfileId $callProfileId

    if ($callProfile.messageBannerText -ne $callProfileSettings.messageBannerText -or $callProfile.participantLimit -ne $callProfileSettings.participantLimit) {
        Write-Host "Updating Call Profile ID : $callProfileId for $cmsServer" -ForegroundColor Yellow
        Update-CallProfile -cmsUrl $cmsServer -headers $cmsHeaders -callProfileId $callProfileId -callProfileSettings $callProfileSettings.ToBodyString()
    }

    # Ensure all CoSpaces are using the correct call profile
    $coSpaces = Get-CoSpaces -cmsUrl $cmsServer -headers $cmsHeaders
    if ($null -ne $coSpaces) {
        $coSpaceIds = $coSpaces | Select-Object -ExpandProperty id
        $coSpaceDetails = $coSpaceIds | Foreach-Object {Get-CoSpace -cmsUrl $cmsServer -headers $cmsHeaders -coSpaceId $_}
    
        $coSpacesRequireFixing = $coSpaceDetails | Where-Object { $_.callProfile -ne $callProfileId }
    
        if ($coSpacesRequireFixing.Count -gt 0) {
            Write-Host "Updating CoSpaces for $cmsServer" -ForegroundColor Yellow
            
            # Update-CoSpaceCallProfileIdAsync -cmsUrl $cmsServer -headers $cmsHeaders -coSpaceIds $coSpaceIdsRequireFixing -callProfileId $callProfileId

            foreach ($coSpace in $coSpacesRequireFixing) {
                Write-Host "Updating $($coSpace.name) Call Profile" -ForegroundColor DarkYellow
                Update-CoSpaceCallProfileId -cmsUrl $cmsServer -headers $cmsHeaders -coSpaceId $coSpace.id -callProfileId $callProfileId
                $updatedCoSpaces += @{
                    Server    = $cmsServer
                    CoSpace   = $coSpace.name -join ", "
                    CallProfileId = $callProfileId -join ", "
                }
            }
        }
    }

    # Get all calls and check if they need a Disconnect Warning Message
    $callIds = Get-Calls -cmsUrl $cmsServer -headers $cmsHeaders
    # $calls = Ge
    # $calls = Get-CallAsync -cmsUrl $cmsServer -headers $cmsHeaders -callIds $callIds

    foreach ($callId in $callIds) {
        $callInfo = Get-Call -cmsUrl $cmsServer -headers $cmsHeaders -callId $callId.id
        
        $coSpace = $coSpaces | Where-Object {$_.id -eq $callInfo.coSpace}

        if ($configData.Automation.CiscoMeetingServer.ExemptMeetings -contains $callInfo.name) {
            Write-Host "$($callInfo.name) is exempt from all rules." -ForegroundColor Cyan
            # Just update the meeting banner with the expected end time
            $callExpirationSettings = New-CallExpirationSettingsObject
            $callExpirationSettings.messageDuration = 0
            $callExpirationSettings.messageBannerText = "{0} :: {1} :: End DTG :: Never" -f `
                $configData.Automation.CiscoMeetingServer.BannerText, `
                (Replace-SpecialChars -source $callInfo.name.ToUpper())

            if ($callInfo.messageBannerText -ne $callExpirationSettings.messageBannerText) {
                Write-Host "Updating the Message Banner for $($callInfo.name) on $cmsServer"
                Update-Call -cmsUrl $cmsServer -headers $cmsHeaders -callId $callInfo.id -callSettings $callExpirationSettings.ToBodyString()
            }
            continue
        }

        $maxMeetingDuration = $configData.Automation.CiscoMeetingServer.MaxMeetingDurationInMinutes
        if ($configData.Automation.CiscoMeetingServer.ExtendedMeetings -contains $callInfo.name) {
            $maxMeetingDuration = $configData.Automation.CiscoMeetingServer.ExtendedMeetingDurationInMinutes
        }
        elseif ($configData.Automation.CiscoMeetingServer.LongMeetings -contains $callInfo.name) {
            $maxMeetingDuration = $configData.Automation.CiscoMeetingServer.LongMeetingDurationInMinutes
        }
        else {
            $maxMeetingDuration = $configData.Automation.CiscoMeetingServer.StandardMeetingDurationInMinutes
        }

        $currentTimeUtc = $(Get-Date).ToUniversalTime()
        $callDurationMinutes = [int]($callInfo.durationSeconds / 60)
        $terminationTime = Get-TerminationTime `
            -callDurationMinutes $callDurationMinutes  `
            -currentTimeUtc $currentTimeUtc `
            -maxCallDurationMinutes $maxMeetingDuration `
            -gracePeriodMinutes $configData.Automation.CiscoMeetingServer.MeetingPreparationTimeInMinutes

        $minutesUntilTermination = [int]($terminationTime.EstimatedEndTime - $currentTimeUtc).TotalMinutes

        # Disconnect the call if it's past the expected termination time
        if ($minutesUntilTermination -le 0) {
            Write-Host "Disconnecting $($callInfo.name) on $cmsServer" -ForegroundColor Red
            Disconnect-Call -cmsUrl $cmsServer -headers $cmsHeaders -callId $callInfo.id
        
            $disconnectedCalls += @{
                Server               = $cmsServer
                CallName             = $callInfo.name
                CallId               = $coSpace.callId
                EstimatedStartTime   = To-MilitaryDTG -date $terminationTime.EstimatedStartTime
                EstimatedEndTime     = To-MilitaryDTG -date $terminationTime.EstimatedEndTime
                TotalDurationMinutes = ($terminationTime.EstimatedEndTime - $terminationTime.EstimatedStartTime).TotalMinutes
            }
            continue
        }

        # Check if the meeting is in the warning window
        $callExpirationSettings = New-CallExpirationSettingsObject
        $callExpirationSettings.messageDuration = $configData.Automation.CiscoMeetingServer.MeetingWarningThresholdInMinutes
        $callExpirationSettings.messagePosition = "middle"
        $callExpirationSettings.messageText = "This meeting will end in $minutesUntilTermination minutes"
        $callExpirationSettings.messageBannerText = "{0} !! {1} !! {2} !! {0}" -f `
            $configData.Automation.CiscoMeetingServer.EndingBannerText, `
            $configData.Automation.CiscoMeetingServer.BannerText, `
            (Replace-SpecialChars -source $callInfo.name.ToUpper())
        
        if ($minutesUntilTermination -le $configData.Automation.CiscoMeetingServer.MeetingWarningThresholdInMinutes) {
            Write-Host "Sending warning message to $($callInfo.name) on $cmsServer" -ForegroundColor Yellow
            Update-Call -cmsUrl $cmsServer -headers $cmsHeaders -callId $callInfo.id -callSettings $callExpirationSettings.ToBodyString()
        
            $warnings += @{
                Server             = $cmsServer
                CallName           = $callInfo.name
                CallId             = $coSpace.callId
                MinutesLeft        = $minutesUntilTermination
                EstimatedStartTime = To-MilitaryDTG -date $terminationTime.EstimatedStartTime
                EstimatedEndTime   = To-MilitaryDTG -date $terminationTime.EstimatedEndTime
                TotalDurationMinutes = ($terminationTime.EstimatedEndTime - $terminationTime.EstimatedStartTime).TotalMinutes
            }
            continue
        }

        # Just update the meeting banner with the expected end time
        $callExpirationSettings = New-CallExpirationSettingsObject
        $callExpirationSettings.messageDuration = 0
        $callExpirationSettings.messageBannerText = "{0} :: {1} :: End DTG :: {2}" -f `
            $configData.Automation.CiscoMeetingServer.BannerText, `
            (Replace-SpecialChars -source $callInfo.name.ToUpper()), `
        $(To-MilitaryDTG -date $terminationTime.EstimatedEndTime)
        
        if ($callInfo.messageBannerText -ne $callExpirationSettings.messageBannerText) {
            Write-Host "Updating the Message Banner for $($callInfo.name) on $cmsServer"
            Update-Call -cmsUrl $cmsServer -headers $cmsHeaders -callId $callInfo.id -callSettings $callExpirationSettings.ToBodyString()
        }
    }
}


if ($warnings.Length -eq 0 -and $disconnectedCalls.Length -eq 0 -and $updatedCoSpaces.Length -eq 0) {
    Write-Host "No actions taken." -ForegroundColor Green
    return
}

# Compile the email message
$emailBody = "<br />"

if ($warnings.Length -gt 0) {

    $emailBody += "The following warnings were sent:<br />"
    foreach ($warning in $warnings) {
        $emailBody += "$($warning.CallName) on $($warning.Server.Replace(":445", "/en-US/meeting/") + $warning.CallId) will end in $($warning.MinutesLeft) minutes. <br />
        Estimated Start Time: $($warning.EstimatedStartTime). <br />
        Estimated End Time: $($warning.EstimatedEndTime) <br />
        Current Meeting Duration: $($warning.TotalDurationMinutes) <br />"

    }
    $emailBody += "<br />"
}

if ($disconnectedCalls.Length -gt 0) {
    $emailBody += "The following calls were disconnected:<br />"
    foreach ($disconnectedCall in $disconnectedCalls) {
        $emailBody += "$($disconnectedCall.CallName) on $($disconnectedCall.Server.Replace(":445", "/en-US/meeting/") + $disconnectedCall.CallId). <br />
        Estimated Start Time: $($disconnectedCall.EstimatedStartTime) <br />
        Estimated End Time: $($disconnectedCall.EstimatedEndTime) <br />
        Current Meeting Duration: $($disconnectedCall.TotalDurationMinutes) <br />"
    }
    $emailBody += "<br />"
}

if ($updatedCoSpaces.Length -gt 0) {
    $emailBody += "The following CoSpaces were updated to use CallProfile: $($updatedCoSpace.CallProfileId)<br />"
    foreach ($updatedCoSpace in $updatedCoSpaces) {
        $emailBody += "$($updatedCoSpace.CoSpace) on $($updatedCoSpace.Server) <br />"
    }
}

Send-AutomationMessage -Body $emailBody