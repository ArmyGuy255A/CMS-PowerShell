
function Get-CMSHeaders{
    [CmdletBinding()]
    Param (
        [string]$authKey
    )

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", ("Basic {0}" -f $authKey))
    $headers.Add("Content-Type", "application/x-www-form-urlencoded")
    return $headers
}

function New-AuthKey {
    [CmdletBinding()]
    Param (
        [string]$fileName = "authkey.key"
    )

    if ($fileName.length -eq 0) {
        $fileName = "authkey.key"
    }

    $cred = Get-Credential
    $authKey = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(("{0}:{1}" -f $cred.GetNetworkCredential().UserName, $cred.GetNetworkCredential().Password)))
    if ($PSScriptRoot.Length -eq 0) {
        $authKey > $filename
    }
    else {
        $authKey > ($PSScriptRoot + "\$fileName")
    }
    return $authKey
}

function Read-AuthKey{
    [CmdletBinding()]
    Param (
        [string]$fileName = "authkey.key"
    )

    if ($fileName.length -eq 0) {
        $fileName = "authkey.key"
    }

    $authKey = $(Get-Content $fileName).trim()
    return $authKey
}

# Helper Functions

function Invoke-CMSRestMethod {
    [CmdletBinding()]
    Param (
        [string]$url,
        [ValidateSet("GET", "POST", "PUT", "DELETE")]
        [string]$method,
        [hashtable]$headers,
        [string]$body = $null
    )

    if ($Verbose) {
        Write-Host "$method : $url" -ForegroundColor Yellow
        if ($body) {
            Write-Host "Request Body: $body" -ForegroundColor DarkYellow
        }
    }
    

    try {
        if ($method -eq "GET") {
            $response = Invoke-RestMethod -Uri $url -Method $method -Headers $headers 
            
        } else {
            $response = Invoke-RestMethod -Uri $url -Method $method -Headers $headers -Body $body 
        }
        if (-not $response -and $method -eq "GET") {
            Write-Warning "Empty response received from $url"
        }
        return $response
    } catch {
        Write-Warning "Error during REST call to $url : $_"
        return $null
    }
}

function Replace-SpecialChars($source) {
    $chars = @("%", "&")
    $newString = $source
    foreach ($char in $chars) {
        $newString = $source.Replace($char, " and ").Trim()
    }
    return $newString
}

function Add-ToBodyStringMethod {
    [CmdletBinding()]
    Param (
        [PSCustomObject]$settingsObject
    )

    if (-not $settingsObject) {
        throw "No object provided to add ToBodyString method."
    }

    $settingsObject | Add-Member -MemberType ScriptMethod -Name ToBodyString -Value {
        ($this.PSObject.Properties |
            Where-Object { $null -ne $_.Value  } | # Skip null values
            ForEach-Object { "$($_.Name)=$($_.Value)" }) -join "&"
    }

    return $settingsObject
}

function New-SettingsObject {
    Param (
        [hashtable]$defaultProperties
    )

    $settings = [PSCustomObject]$defaultProperties
    return Add-ToBodyStringMethod -settingsObject $settings
}

function New-CallExpirationSettingsObject {
    $defaultProperties = @{
        messageBannerText = $null
        messageDuration   = $null
        messagePosition   = $null
        messageText       = $null
    }
    return New-SettingsObject -defaultProperties $defaultProperties
}

function New-CallProfileSettingsObject {
    $defaultProperties = @{
        messageBannerText = $null
        participantLimit  = $null
        logoFileName      = $null
    }
    return New-SettingsObject -defaultProperties $defaultProperties
}

function Get-PaginatedItems {
    Param (
        [string]$baseUrl,
        [hashtable]$headers,
        [int]$numItems,
        [string]$keyPropertyName
    )

    $items = [System.Collections.ArrayList]::new()
    $skipRemaining = $false
    for ($i = 0; $i -lt 200; $i += $numItems) {
        if ($skipRemaining) {
            break
        }
        try {
            $url = "{0}?offset={1}" -f $baseUrl, $i
            $key = $keyPropertyName.split(".")[0]
            $property = $keyPropertyName.split(".")[1]
            $result = Invoke-CMSRestMethod -url $url -method 'GET' -headers $headers
            if ($null -eq $result.$key.$property) {
                $skipRemaining = $true
            } else {
                $result.$key.$property | ForEach-Object {$items.add($_) | Out-Null}
                if ($result.$key.$property.length -lt $numItems) {
                    $skipRemaining = $true
                }
            }
        } catch {
            Write-Warning "Error fetching items: $_"
            break
        }
    }
    return @($items)
}


# GET Methods
function Get-Calls {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers
    )
    return @(Get-PaginatedItems -baseUrl "$cmsUrl/api/v1/calls" -headers $headers -numItems 10 -keyPropertyName "calls.call")
}

function Get-CallProfileIds {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers
    )
    # BUG: This doesn't always return an array of strings.
    return @(Get-PaginatedItems -baseUrl "$cmsUrl/api/v1/callProfiles" -headers $headers -numItems 10 -keyPropertyName "callProfiles.callProfile" | Select-Object -ExpandProperty id)
}

function Get-CoSpaceTemplates {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers
    )
    return @(Get-PaginatedItems -baseUrl "$cmsUrl/api/v1/coSpaceTemplates" -headers $headers -numItems 10 -keyPropertyName "coSpaceTemplates.coSpaceTemplate")
}

function Get-CoSpaces {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers
    )
    return @(Get-PaginatedItems -baseUrl "$cmsUrl/api/v1/coSpaces" -headers $headers -numItems 20 -keyPropertyName "coSpaces.coSpace")
}

function Get-CallProfile {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers,
        [string]$callProfileId
    )
    $result = Invoke-CMSRestMethod "$cmsUrl/api/v1/callProfiles/$callProfileId" -Method 'GET' -Headers $headers
    return $result.callprofile
}

function Get-CoSpace {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers,
        [string]$coSpaceId
    )
    $result = Invoke-CMSRestMethod "$cmsUrl/api/v1/coSpaces/$coSpaceId" -Method 'GET' -Headers $headers

    return $result.cospace
}
function Get-Call {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers,
        [string]$callId
    )
    $result = Invoke-CMSRestMethod "$cmsUrl/api/v1/calls/$callId" -Method 'GET' -Headers $headers
    return $result.call
}
# PUT Methods
function Update-CallProfile {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers,
        [string]$callProfileId,
        [string]$callProfileSettings
    )
    Invoke-CMSRestMethod "$cmsUrl/api/v1/callProfiles/$callProfileId" -Method 'PUT' -Headers $headers -Body $callProfileSettings
}

function Update-CoSpaceCallProfileId {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers,
        [string]$coSpaceId,
        [string]$callProfileId
    )
    $settings = "callProfile=$callProfileId"
    Invoke-CMSRestMethod "$cmsUrl/api/v1/coSpaces/$coSpaceId" -Method 'PUT' -Headers $headers -Body $settings
}

function Update-Call {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers,
        [string]$callId,
        [string]$callSettings
    )

    $url = "$cmsUrl/api/v1/calls/$callId"
    Write-Verbose "Updating call with ID: $callId"

    return Invoke-CMSRestMethod -url $url -method 'PUT' -headers $headers -body $callSettings
}

# DELETE Methods
function Disconnect-Call {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers,
        [string]$callId
    )

    $result = Invoke-RestMethod "$cmsUrl/api/v1/calls/$callId" -Method 'DELETE' -Headers $headers 
    return $result
}

# POST Methods

function New-CallProfile {
    [CmdletBinding()]
    Param (
        [string]$cmsUrl,
        [hashtable]$headers,
        [string]$callProfileSettings
    )

    $result = Invoke-CMSRestMethod "$cmsUrl/api/v1/callProfiles" -Method 'POST' -Headers $headers -Body $callProfileSettings
    return $result.callProfile
}

# Async Methods

function Get-CoSpaceAsync {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$cmsUrl,
        [Parameter(Mandatory)]
        [hashtable]$headers,
        [Parameter(Mandatory)]
        [string[]]$coSpaceIds
    )

    # Array to store job objects
    #$jobs = @()

    # Start a job for each CoSpace
    foreach ($coSpaceId in $coSpaceIds) {
        Invoke-Command {
            Set-Location $PSScriptRoot
            . .\CMS.functions.ps1
            $result = Get-CoSpace -cmsUrl $cmsUrl -headers $headers -coSpaceId $coSpaceId
            return $result
        } -ArgumentList $cmsUrl, $headers, $coSpaceId -AsJob

        #$jobs += Invokk -ScriptBlock {
            # Import required functions and variables into the job
            #Param ($cmsUrl, $headers, $coSpaceId)

            # Call the Get-CoSpace function
            . .\CMS.functions.ps1
            #$result = Get-CoSpace -cmsUrl $using:cmsUrl -headers $using:headers -coSpaceId $using:coSpaceId

            #return $result.coSpace
        #} -ArgumentList $cmsUrl, $headers, $coSpaceId 
    }

    # Wait for all jobs to complete
    Write-Host "Waiting for all jobs to complete..." -ForegroundColor Yellow
    $jobs | Wait-Job

    # Collect the results from the jobs
    $results = $jobs | Receive-Job

    # Remove completed jobs
    $jobs | Remove-Job -Force

    # Return the combined results
    return $results
}

function Update-CoSpaceCallProfileIdAsync {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$cmsUrl,
        [Parameter(Mandatory)]
        [hashtable]$headers,
        [Parameter(Mandatory)]
        [string[]]$coSpaceIds,
        [Parameter(Mandatory)]
        [string]$callProfileId
    )

    # Array to store job objects
    $jobs = @()

    # Start a job for each CoSpace
    foreach ($coSpaceId in $coSpaceIds) {
        $jobs += Start-Job -ScriptBlock {
            # Import required functions and variables into the job
            Param ($cmsUrl, $headers, $coSpaceId, $callProfileId)

            # Call the Update-CoSpaceCallProfileId function
            $result = Update-CoSpaceCallProfileId -cmsUrl $cmsUrl -headers $headers -coSpaceId $coSpaceId -callProfileId $callProfileId
            return $result
        } -ArgumentList $cmsUrl, $headers, $coSpaceId, $callProfileId
    }

    # Wait for all jobs to complete
    Write-Host "Waiting for all jobs to complete..." -ForegroundColor Yellow
    $jobs | Wait-Job

    # Collect the results from the jobs
    $results = $jobs | Receive-Job

    # Remove completed jobs
    $jobs | Remove-Job -Force

    # Return the combined results
    return $results
}


# Utility Methods

function Get-ExpiredCallInfo($cmsUrl, $headers, $expirationDurationInSeconds) {
    $calls = Get-Calls -cmsUrl $cmsUrl -headers $headers

    $items = [System.Collections.ArrayList]::new()
    foreach ($call in $calls) {
        $callInfo = Get-Call -cmsUrl $cmsUrl -headers $headers -callId $call.id
        write-host $callInfo.name " : Current Duration : " $callInfo.durationSeconds
        if ([int]$callInfo.durationSeconds -gt $expirationDurationInSeconds) {
            $items.Add($callInfo) | Out-Null
        }
    }
    return $items
}

function Should-DisconnectCall {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [int]$callDurationSeconds, # Call duration in seconds
        [Parameter(Mandatory)]
        [string]$currentTimeUtc,   # Current time as HH:MM in UTC
        [int]$gracePeriodMinutes = 5 # Grace period in minutes
    )

    # Parse the current UTC time
    $now = [datetime]::ParseExact($currentTimeUtc, "HH:mm", $null)

    # Calculate the estimated start time
    $estimatedStartTime = $now.AddSeconds(-$callDurationSeconds)

    # Calculate the next quarter-hour
    $nextQuarter = $now.AddMinutes((15 - ($now.Minute % 15))).AddSeconds(-$now.Second)

    # Calculate the grace period window
    $gracePeriodStart = $nextQuarter.AddMinutes(-$gracePeriodMinutes)
    $gracePeriodEnd = $nextQuarter.AddMinutes($gracePeriodMinutes)

    # Determine if the current time is within the grace period
    $withinGracePeriod = ($now -ge $gracePeriodStart -and $now -le $gracePeriodEnd)

    # Return the decision and debugging information
    return @{
        ShouldDisconnect   = $withinGracePeriod
        EstimatedStartTime = $estimatedStartTime
        NextQuarterHour    = $nextQuarter
        GracePeriodStart   = $gracePeriodStart
        GracePeriodEnd     = $gracePeriodEnd
    }
}

function Get-TerminationTime {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [int]$callDurationMinutes,   # Call duration in seconds
        [Parameter(Mandatory)]
        [datetime]$currentTimeUtc,     # Current time as HH:mm in UTC
        [Parameter(Mandatory)]
        [int]$maxCallDurationMinutes, # Maximum allowed call duration in minutes
        [int]$gracePeriodMinutes = 15 # Grace period in minutes
    )

    $callDurationSeconds = $callDurationMinutes * 60

    # Convert the call duration to seconds
    $maxCallDurationSeconds = $maxCallDurationMinutes * 60

    # Calculate the estimated start time
    $estimatedStartTime = $currentTimeUtc.AddSeconds(-$callDurationSeconds)

    # Round the estimated start time to the next whole or half hour
    $estimatedStartTime = $estimatedStartTime.AddMinutes((30 - ($estimatedStartTime.Minute % 30))).AddSeconds(-$estimatedStartTime.Second)

    # Estimate the meeting end time
    $estimatedEndTime = $estimatedStartTime.AddSeconds($maxCallDurationSeconds)

    # Round the estimated end time to the grace period
    $estimatedEndTime = $estimatedEndTime.AddMinutes(($gracePeriodMinutes - ($estimatedEndTime.Minute % $gracePeriodMinutes))).AddSeconds(-$estimatedEndTime.Second)

    # Return all relevant information
    return @{
        EstimatedStartTime = $estimatedStartTime
        EstimatedEndTime    = $estimatedEndTime
    }
}

function To-MilitaryDTG([datetime]$date) {
    return $date.ToUniversalTime().ToString("ddHHmmZMMMyy").toupper()
}

return
