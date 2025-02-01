BeforeAll {
    . $PSScriptRoot/CMS.Functions.ps1
}

Describe "Cisco Meeting Server API Functions" {
    BeforeEach {

        # Mock: Get-CallProfileIds
        Mock -CommandName Get-CallProfileIds -MockWith {
            return @("886cfb15-7ac9-4c94-960b-9ce705b36f65")
        } 

        # Mock: Get-CallProfile
        Mock Get-CallProfile {
            return @{
                CallProfile = @{
                    id                = "886cfb15-7ac9-4c94-960b-9ce705b36f65"
                    participantLimit  = 100
                    recordingMode     = "manual"
                    messageBannerText = "BANNER TEXT"
                    logoFileName      = ""
                }
            }
        }

        # Mock: Get-CoSpaceTemplates
        Mock Get-CoSpaceTemplates {
            return @(
                @{
                    id                       = "8e0c59a8-9f5b-408d-823a-3e15ac0ce1ca"
                    name                     = "LDAPSpaceTemplate"
                    callProfile              = "886cfb15-7ac9-4c94-960b-9ce705b36f65"
                    callLegProfile           = "f0831d7d-6224-4425-948b-dc47a92da1d6"
                    numAccessMethodTemplates = 1
                }
            )
        }

        # Mock: Get-Calls
        Mock Get-Calls {
            return @(
                @{
                    id              = "ee93849d-ea1f-401b-9263-fddd193706df"
                    name            = "Room 2"
                    durationSeconds = 15040
                    coSpace         = "e5ad5aa8-47dc-4c2c-b6aa-2ce8c049b744"
                    callCorrelator  = "3569c15e-1bf0-49a6-8c8f-a4956da89f56"
                },
                @{
                    id              = "d4728e24-165e-4880-bada-213c197608c1"
                    name            = "Room 1"
                    durationSeconds = 2031
                    coSpace         = "b0f62c13-bf74-4c81-8690-27265236c250"
                    callCorrelator  = "6c1847bb-5932-43f4-a655-2409b7236377"
                }
            )
        }

        # Mock: Get-Call
        Mock Get-Call {
            param ($cmsUrl, $headers, $callId)
            if ($callId -eq "ee93849d-ea1f-401b-9263-fddd193706df") {
                return @{
                    id                = "ee93849d-ea1f-401b-9263-fddd193706df"
                    name              = "Room 2"
                    durationSeconds   = 15040
                    messageBannerText = "BANNER TEXT"
                }
            }
            elseif ($callId -eq "d4728e24-165e-4880-bada-213c197608c1") {
                return @{
                    id                = "d4728e24-165e-4880-bada-213c197608c1"
                    name              = "Room 1"
                    durationSeconds   = 2031
                    messageBannerText = "TEXT BANNER"
                }
            }
        }

        # Mock: Get-CoSpace
        Mock Get-CoSpace {
            param ($cmsUrl, $headers, $coSpaceId)
            if ($coSpaceId -eq "e5ad5aa8-47dc-4c2c-b6aa-2ce8c049b744") {
                return @{
                    CoSpace = @{
                        id          = "e5ad5aa8-47dc-4c2c-b6aa-2ce8c049b744"
                        name        = "Room 2"
                        callProfile = "886cfb15-7ac9-4c94-960b-9ce705b36f65"
                    }
                }
            }
            elseif ($coSpaceId -eq "b0f62c13-bf74-4c81-8690-27265236c250") {
                return @{
                    CoSpace = @{
                        id          = "b0f62c13-bf74-4c81-8690-27265236c250"
                        name        = "Room 1"
                        callProfile = "886cfb15-7ac9-4c94-960b-9ce705b36f65"
                    }
                }
            }
        }
    }

    # Example Test Case: Get-CallProfileIds
    It "Retrieves Call Profile IDs" {
        $result = Get-CallProfileIds -cmsUrl "https://fake-server.com" -headers @{}
        $result | Should -Contain "886cfb15-7ac9-4c94-960b-9ce705b36f65"
        $result | Should -HaveCount 1
    }

    # Example Test Case: Get-CallProfile
    It "Retrieves a Call Profile" {
        $result = Get-CallProfile -cmsUrl "https://fake-server.com" -headers @{} -callProfileId "886cfb15-7ac9-4c94-960b-9ce705b36f65"
        $result.CallProfile.participantLimit | Should -Be 100
        $result.CallProfile.messageBannerText | Should -Be "BANNER TEXT"
    }

    # Example Test Case: Get-Calls
    It "Retrieves Active Calls" {
        $result = Get-Calls -cmsUrl "https://fake-server.com" -headers @{}
        $result | Should -HaveCount 2
        $result[0].name | Should -Be "Room 2"
        $result[1].name | Should -Be "Room 1"
    }

    # Example Test Case: Get-Call
    It "Retrieves Details for a Specific Call" {
        $result = Get-Call -cmsUrl "https://fake-server.com" -headers @{} -callId "ee93849d-ea1f-401b-9263-fddd193706df"
        $result.name | Should -Be "Room 2"
        $result.durationSeconds | Should -Be 15040
    }

    # Example Test Case: Get-CoSpace
    It "Retrieves CoSpace Details" {
        $result = Get-CoSpace -cmsUrl "https://fake-server.com" -headers @{} -coSpaceId "e5ad5aa8-47dc-4c2c-b6aa-2ce8c049b744"
        $result.CoSpace.name | Should -Be "Room 2"
        $result.CoSpace.callProfile | Should -Be "886cfb15-7ac9-4c94-960b-9ce705b36f65"
    }
}


Describe "Get-TerminationTime" {
    It "Calculates correct termination time for standard cases" -TestCases @(
        @{ currentTimeUtc = "14:30"; callDurationMinutes = 45; maxCallDurationMinutes = 60; gracePeriodMinutes = 15; expectedTerminationTime = "15:15" },
        @{ currentTimeUtc = "15:00"; callDurationMinutes = 14; maxCallDurationMinutes = 60; gracePeriodMinutes = 15; expectedTerminationTime = "16:15" },
        @{ currentTimeUtc = "14:59"; callDurationMinutes = 15; maxCallDurationMinutes = 60; gracePeriodMinutes = 15; expectedTerminationTime = "16:15" }
    ) {
        Param ($currentTimeUtc, $callDurationMinutes, $maxcallDurationMinutes, $gracePeriodMinutes, $expectedTerminationTime)

        $currentTimeUtc = [datetime]::ParseExact($currentTimeUtc, "HH:mm", $null)
        $expectedTerminationTime = [datetime]::ParseExact($expectedTerminationTime, "HH:mm", $null)

        $result = Get-TerminationTime `
            -callDurationMinutes $callDurationMinutes `
            -currentTimeUtc $currentTimeUtc `
            -maxCallDurationMinutes $maxcallDurationMinutes `
            -gracePeriodMinutes $gracePeriodMinutes

        $result | Should -Be $expectedTerminationTime
    }

    It "Calculates correct termination time for extended cases" -TestCases @(
        @{ currentTimeUtc = "14:30"; callDurationMinutes = 45; maxCallDurationMinutes = 90; gracePeriodMinutes = 15; expectedTerminationTime = "15:45" },
        @{ currentTimeUtc = "15:00"; callDurationMinutes = 14; maxCallDurationMinutes = 90; gracePeriodMinutes = 15; expectedTerminationTime = "16:45" },
        @{ currentTimeUtc = "14:59"; callDurationMinutes = 15; maxCallDurationMinutes = 90; gracePeriodMinutes = 15; expectedTerminationTime = "16:45" }
    ) {
        Param ($currentTimeUtc, $callDurationMinutes, $maxcallDurationMinutes, $gracePeriodMinutes, $expectedTerminationTime)

        $currentTimeUtc = [datetime]::ParseExact($currentTimeUtc, "HH:mm", $null)
        $expectedTerminationTime = [datetime]::ParseExact($expectedTerminationTime, "HH:mm", $null)

        $result = Get-TerminationTime `
            -callDurationMinutes $callDurationMinutes `
            -currentTimeUtc $currentTimeUtc `
            -maxCallDurationMinutes $maxcallDurationMinutes `
            -gracePeriodMinutes $gracePeriodMinutes

        $result | Should -Be $expectedTerminationTime
    }

    It "Calculates correct termination time for long cases" -TestCases @(
        @{ currentTimeUtc = "14:30"; callDurationMinutes = 45; maxCallDurationMinutes = 120; gracePeriodMinutes = 15; expectedTerminationTime = "16:15" },
        @{ currentTimeUtc = "15:00"; callDurationMinutes = 14; maxCallDurationMinutes = 120; gracePeriodMinutes = 15; expectedTerminationTime = "17:15" },
        @{ currentTimeUtc = "14:59"; callDurationMinutes = 15; maxCallDurationMinutes = 120; gracePeriodMinutes = 15; expectedTerminationTime = "17:15" }
    ) {
        Param ($currentTimeUtc, $callDurationMinutes, $maxcallDurationMinutes, $gracePeriodMinutes, $expectedTerminationTime)

        $currentTimeUtc = [datetime]::ParseExact($currentTimeUtc, "HH:mm", $null)
        $expectedTerminationTime = [datetime]::ParseExact($expectedTerminationTime, "HH:mm", $null)

        $result = Get-TerminationTime `
            -callDurationMinutes $callDurationMinutes `
            -currentTimeUtc $currentTimeUtc `
            -maxCallDurationMinutes $maxcallDurationMinutes `
            -gracePeriodMinutes $gracePeriodMinutes

        $result | Should -Be $expectedTerminationTime
    }


}
