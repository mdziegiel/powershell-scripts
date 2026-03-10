<#
.Synopsis
v3.5.0.0 12-October-2023
.Description
This script Uninstalls and Cleans up Dell SupportAssist for PCs if installed on the box.
.FileName
SupportAssistUninstall_Cleanup.ps1
#>
<#
Return Codes: 
	Exit 0 : Uninstall/Cleanup Successful.
	Exit 1 : Uninstall/Cleanup Successful with reboot required.
	Exit X : Uninstall/Cleanup Failed.
EventID :
	Information : 0
	Error : 11725
#>

$ProgressPreference = "SilentlyContinue"

# Define the name of the application
$AppName = "Dell SupportAssist"
New-EventLog -LogName Application -Source "SupportAssistUninstall_Cleanup" -ErrorAction SilentlyContinue

# Define the registry path for uninstall information
$UninstallPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"

# Function to find the product code for the given application
function Find-ProductCode {
    param (
        [string]$appName
    )

    # Get a list of all subkeys in the Uninstall registry path
    $uninstallKeys = Get-ChildItem $UninstallPath

    foreach ($key in $uninstallKeys) {
        $appInfo = Get-ItemProperty $key.PSPath
		
        # Check if DisplayName matches the application name
        if ($appInfo.DisplayName -eq $appName) 
		{
			#Write-Host "Dell SupportAssist Application found"
			$SupportAssistType = Get-ItemProperty -Path 'HKLM:\SOFTWARE\DELL\SupportAssistAgent' | Select-Object -ExpandProperty 'Type'
			if($SupportAssistType = "Consumer")
			{
				#Write-Host "Consumer found"
				# Return the ProductCode
				return $appInfo.PSChildName	
			}
        }
    }

    # If the application is not found, return $null
    return $null
}

# Function to uninstall the application using msiexec and log exit codes
function Uninstall-Application {
    param (
        [string]$productCode
    )
	
	#Write-Host $productCode
	if ($productCode -eq "")	
	{
        Write-EventLog -LogName "Application" -Source "SupportAssistUninstall_Cleanup" -EventID 0 -EntryType Information -Message "SupportAssist for Home PCs not found. SupportAssistUninstall_Cleanup script execution was completed successfully." -Category 0
		[System.Environment]::Exit(0)
    }

    # Define the parameters for silent uninstallation
    $msiParams = "/x $productCode /qn /log uninstall.log"

    # Uninstall the application using msiexec
    $result = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiParams -Wait -PassThru

    # Log exit code and description
    $logEntry = "Exit Code: $($result.ExitCode) - Description: $($result.ExitCodeMessage)"
    Add-Content -Path "uninstall.log" -Value $logEntry

    switch ($result.ExitCode) {
        0 {
            #Write-Host "Uninstallation successful."
			Write-EventLog -LogName "Application" -Source "SupportAssistUninstall_Cleanup" -EventID 0 -EntryType Information -Message "SupportAssist for Home PCs uninstallation successful. SupportAssistUninstall_Cleanup script execution was completed successfully." -Category 0
			[System.Environment]::Exit(0)
        }
        1641 {
            #Write-Host "Uninstallation successful, but a reboot is required."
			Write-EventLog -LogName "Application" -Source "SupportAssistUninstall_Cleanup" -EventID 0 -EntryType Information -Message "Restart Required.  SupportAssist for Home PCs uninstallation successful. SupportAssistUninstall_Cleanup script execution Completed." -Category 0
			[System.Environment]::Exit(1)
        }
		3010 {
            #Write-Host "Uninstallation successful, but a reboot is required."
			Write-EventLog -LogName "Application" -Source "SupportAssistUninstall_Cleanup" -EventID 0 -EntryType Information -Message "Restart Required. SupportAssistUninstall_Cleanup script execution completed." -Category 0
			[System.Environment]::Exit(1)
        }
        default {
            #Write-Host "Uninstallation failed. Exit Code: $($result.ExitCode)"
			Write-EventLog -LogName "Application" -Source "SupportAssistUninstall_Cleanup" -EventID 0 -EntryType Error -Message "SupportAssistUninstall_Cleanup script failed to remove SupportAssist for Home PCs.
			" -Category 0
			[System.Environment]::Exit(2)
        }
    }
}

# Main Entry point
try {
    # Find the product code for the application
    $productCode = Find-ProductCode -appName $AppName

    # Uninstall the application using the product code
    Uninstall-Application -productCode $productCode
} catch {
    #Write-Host "An error occurred: $_"
	Write-EventLog -LogName "Application" -Source "SupportAssistUninstall_Cleanup" -EventID 0 -EntryType Information -Message "SupportAssistUninstall_Cleanup script failed with error: $_" -Category 0
	[System.Environment]::Exit(2)
}

# End of script
# SIG # Begin signature block
# MIIrWwYJKoZIhvcNAQcCoIIrTDCCK0gCAQExDzANBglghkgBZQMEAgMFADCBmwYK
# KwYBBAGCNwIBBKCBjDCBiTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63
# JNLGKX7zUQIBAAIBAAIBAAIBAAIBADBRMA0GCWCGSAFlAwQCAwUABEDmSPVZvJBk
# 01T7PipOYDAad8rUiIrGmIZQCy45hish3TDyFlyYhFXeCgNX3PgqHm9AEWJdkXHv
# RFx+u8pL3DE2oIISjjCCBd8wggTHoAMCAQICEE5A5DdU7eaMAAAAAFHTlH8wDQYJ
# KoZIhvcNAQELBQAwgb4xCzAJBgNVBAYTAlVTMRYwFAYDVQQKEw1FbnRydXN0LCBJ
# bmMuMSgwJgYDVQQLEx9TZWUgd3d3LmVudHJ1c3QubmV0L2xlZ2FsLXRlcm1zMTkw
# NwYDVQQLEzAoYykgMjAwOSBFbnRydXN0LCBJbmMuIC0gZm9yIGF1dGhvcml6ZWQg
# dXNlIG9ubHkxMjAwBgNVBAMTKUVudHJ1c3QgUm9vdCBDZXJ0aWZpY2F0aW9uIEF1
# dGhvcml0eSAtIEcyMB4XDTIxMDUwNzE1NDM0NVoXDTMwMTEwNzE2MTM0NVowaTEL
# MAkGA1UEBhMCVVMxFjAUBgNVBAoMDUVudHJ1c3QsIEluYy4xQjBABgNVBAMMOUVu
# dHJ1c3QgQ29kZSBTaWduaW5nIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkg
# LSBDU0JSMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKeBj/cURbZi
# Q/LYrtMlXkhPUb/FfZ9QHDXR1n5hKpQZbSdGpKYaXfdUUWqAIsaoZnVNVIPJXmgb
# q/ZbZLCtrSC9VO9Ga20C50WudfaOirkyLou4dxxSTXmIX6U6GMlQLJcnLb/aAH1j
# f+8y7EaHY9uan8NaITZ7+ZvVyqBuciz84fGecE0IVhVvkKv7SLq518GCeIVlLn+1
# ycDiFLc3EUEG4orgqPblfrZ4BQHDYO1PB0EuChNJ45Cbf929+qy/ZFHRXJu09Vzn
# XP87m6WgGtd9CbLCt/9uHLzIfebpK/xysxTpSlUShJxEJXUd9irwT6UgPWgl62GX
# fA/ltj3zrsPBEbwbjszgRzBeQgCGceNYrAbKZR97lKZLV2cMfl6teGdbVeNe68fY
# 7Exuhsvz3Pifh6pyWBIPfab4+EI5Ozws5DJNSYzg4QDCOKCc+oQ+QdxuVq7GGlv0
# Z2gFAc0bv66HvJ1T9i7otmvkmd7FT4dYqNJlHsgf1XJu7lkcVzsJcp3XyreQxs17
# RZKRQgNMfT/K8qq4wg6G8xCfRi6kZoZoWmgYcCk4EYBga4pDo3Ns47NrN//mnWcB
# kobfL0jR+1Bg1Vz+IdMBQmP+73C0F8CPqO7TwUtfEur9/S4Oh0Rg46n0whij4/3O
# DIQiDfOneNqT89s4z7kvM8b/BzxevkXTAgMBAAGjggErMIIBJzAOBgNVHQ8BAf8E
# BAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBATAdBgNVHSUEFjAUBggrBgEFBQcDAwYI
# KwYBBQUHAwgwOwYDVR0gBDQwMjAwBgRVHSAAMCgwJgYIKwYBBQUHAgEWGmh0dHA6
# Ly93d3cuZW50cnVzdC5uZXQvcnBhMDMGCCsGAQUFBwEBBCcwJTAjBggrBgEFBQcw
# AYYXaHR0cDovL29jc3AuZW50cnVzdC5uZXQwMAYDVR0fBCkwJzAloCOgIYYfaHR0
# cDovL2NybC5lbnRydXN0Lm5ldC9nMmNhLmNybDAdBgNVHQ4EFgQUgrrWPZfOn89x
# 6JI3r/2ztWk1V88wHwYDVR0jBBgwFoAUanImetAe733nO2lR1GyNn5ASZqswDQYJ
# KoZIhvcNAQELBQADggEBAB9eQQS2g3AkUyxVcx1lOsDstHsEmF5ZOBMJpFmUQl5Q
# v09sbiUgkJNYQA31GbRi7iRewgFYFQIdEAlvqNT7kn43OD4vFH2PHUM2ZLNmE18U
# zKVx91shS8aXvtyV/HB9ERzTId3QJDkpxf4KGqXPe3nuOm/e3L/pEd0WgwjTLI1/
# TagUeS8FYVI462DzFGh9y7KKrcCUXOQmDiyK3UbDzuRWUcVW44W4TZtFcosH8Yr7
# Sbhf0fKWgV1pUiTxCCPS1iMP64vXfovBk2v68WJ7WOlQm5duF4gN4cZDmNeBYbaF
# nUfssZ6uPyA7Q53Yohzg1HwIwq92BvhiZnq29/rIrzUwggYzMIIEG6ADAgECAhBN
# 5PZMB6ZZLJUliCaOZum4MA0GCSqGSIb3DQEBDQUAME8xCzAJBgNVBAYTAlVTMRYw
# FAYDVQQKEw1FbnRydXN0LCBJbmMuMSgwJgYDVQQDEx9FbnRydXN0IENvZGUgU2ln
# bmluZyBDQSAtIE9WQ1MyMB4XDTIzMDExMjE2MjY1MloXDTI0MDEyMzE2MjY1MFow
# dzELMAkGA1UEBhMCVVMxDjAMBgNVBAgTBVRleGFzMRMwEQYDVQQHEwpSb3VuZCBS
# b2NrMREwDwYDVQQKEwhEZWxsIEluYzEdMBsGA1UECxMUU3VwcG9ydEFzc2lzdCBD
# bGllbnQxETAPBgNVBAMTCERlbGwgSW5jMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEA3FrIe3esozZzu317OAhH3QI8D3BgyVKQrmHPDx7FJ+m8MwVMWE5x
# VmEDPzwnUURfgSuJ1nffnyBNbTyCl+KowGBSoZx1JrkF4sb0gyJ0K+6fbdW2jIF6
# KiLRGpyHER2Ww1cEopwFAQLIV0FcIeBYKDfHU+dG9iCCNfo8PLdc4eHLqsRtsgXH
# bk3G8/J3nWWYWWXQX1LJfWewhJJN8L1UvK7LPF81H4xIxw3zcYzQfdkiSzVTJnI0
# KwyL0jj1dU6H+MTZ8nEW2p8DHZ4qoNfAHSqXKQfQLhXMtHBCPZPXgvpO9mzJf4Jj
# /6+yAnq1RE2qsgcIQilASeB0ZOE9zct6gBLPeUiZwSXCjeriYHBbSz0vmvZMbYMT
# HfCKQtbKCOby+6bfWLe/F+pNyg9hNYPcPeqWm7HJC6E4LUMUTxDoj5kewMc5WjUg
# 4m99C7Soxc7cveSA9gfFJXQNVxPCqr9T1+v2GC9pnmFW+3rM58K1wfHsDqXnVVqs
# DmhCidSJng71AgMBAAGjggFhMIIBXTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBSI
# b3bsiIPEHkf1H1ypAU63X4NT/zAfBgNVHSMEGDAWgBTvn7p5sHPyJR54nANSnBtT
# hN6N7TBnBggrBgEFBQcBAQRbMFkwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLmVu
# dHJ1c3QubmV0MDIGCCsGAQUFBzAChiZodHRwOi8vYWlhLmVudHJ1c3QubmV0L292
# Y3MyLWNoYWluLnA3YzAxBgNVHR8EKjAoMCagJKAihiBodHRwOi8vY3JsLmVudHJ1
# c3QubmV0L292Y3MyLmNybDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwTAYDVR0gBEUwQzA3BgpghkgBhvpsCgEDMCkwJwYIKwYBBQUHAgEWG2h0
# dHBzOi8vd3d3LmVudHJ1c3QubmV0L3JwYTAIBgZngQwBBAEwDQYJKoZIhvcNAQEN
# BQADggIBAGhhmZbbANmfLcbP5GF9YM/5GexgliIMhIEWRr+WSw3RXu4uUVukJol9
# U2XIoUFmskeCwIUcmA2ChC6bgzvzebxXTi7QPPs2S20P+dMlNWlTUZaWJZ6XJaV6
# jJEWMXoWchyJWLqDnO/zbycxjHBAxYnX46ltB3jQRnGwRdt2bWnQsHSUMYn04g0b
# DB3AAJE1X520mYXw631lhCWek+b/eto3kT6oxQjrS76IaBQW8wksitUyCE7wX2AT
# WitnVdMw0xeFAaP99vYFCG/DQPKqeNpx4y9QZu12UirPMGRqjfOYPoK8s9IwVbCe
# xJ+XMav6fh2vnd5jEFIVH5QGLwOpwpuDCB1kqsUb0ZjdlGehWfqLX/5yURCywE8Q
# pBRY3k9VOFKzR6d7ApzmpJBin3s3tfmuOtgFjHbCmIhUP5TzqJlrEiGPWAfMon62
# YfAm8LlSyVfwRzbhgZgfcaCkFyUTuoGQvKGggURQCbAruHW10koFd/aJDmbGenRj
# cH2jzC8DkNFhweczwfVimygri44mqrFLpKnyqswmhSgj4wmgZA92sFMzshItMO3Q
# e8DZ9Xu+AeBmRFU85Ye4Y3e5TTM0H7fFGheCzs3Sx8Wj6j8n613+GnZiHV2TuOsz
# qJa6wxBcmZ7c9CmI00h0fE1w75WCg7sc52o8rNtwH3HhVCzgQI9fMIIGcDCCBFig
# AwIBAgIQce9VdK81VMNaLGn2b0trzTANBgkqhkiG9w0BAQ0FADBpMQswCQYDVQQG
# EwJVUzEWMBQGA1UECgwNRW50cnVzdCwgSW5jLjFCMEAGA1UEAww5RW50cnVzdCBD
# b2RlIFNpZ25pbmcgUm9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSAtIENTQlIx
# MB4XDTIxMDUwNzE5MjA0NVoXDTQwMTIyOTIzNTkwMFowTzELMAkGA1UEBhMCVVMx
# FjAUBgNVBAoTDUVudHJ1c3QsIEluYy4xKDAmBgNVBAMTH0VudHJ1c3QgQ29kZSBT
# aWduaW5nIENBIC0gT1ZDUzIwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQCemXYXGp5WFwhjLJNNg2GEMzQCttlioN7CDrkgTMhXnQ/dVFsNDNYB3S9I4ZEJ
# 4dvIFQSCtnvw2NYwOxlxcPuoppf2KV2kDKn0Uz5X2wxObvx2218k6apfQ+OT5w7P
# yiW8xEwwC1oP5gb05W4MmWZYT4NhwnN8XCJvAUXFD/dAT2RL0BcKqQ4eAi+hj0zy
# Z1DbPuSfwk8/dOsxpNCU0Jm8MJIJasskzaLYdlLQTnWYT2Ra0l6D9FjAXWp1xNg/
# ZDqLFA3YduHquWvnEXBJEThjE27xxvq9EEU1B+Z2FdB1FqrCQ1f+q/5jc0YioLjz
# 5MdwRgn5qTdBmrNLbB9wcqMH9jWSdBFkbvkC1cCSlfGXWX4N7qIl8nFVuJuNv83u
# rt37DOeuMk5QjaHf0XO/wc5/ddqrv9CtgjjF54jtom06hhG317DhqIs7DEEXml/k
# W5jInQCf93PSw+mfBYd5IYPWC+3RzAif4PHFyVi6U1/Uh7GLWajSXs1p0D76xDkJ
# r7S17ec8+iKH1nP5F5Vqwxz1VXhf1PoLwFs/jHgVDlpMOm7lJpjQJ8wg38CGO3qN
# ZUZ+2WFeqfSuPtT8r0XHOrOFBEqLyAlds3sCKFnjhn2AolhAZmLgOFWDq58pQSa6
# u+nYZPi2uyhzzRVK155z42ZMsVGdgSOLyIZ3srYsNyJwIQIDAQABo4IBLDCCASgw
# EgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU75+6ebBz8iUeeJwDUpwbU4Te
# je0wHwYDVR0jBBgwFoAUgrrWPZfOn89x6JI3r/2ztWk1V88wMwYIKwYBBQUHAQEE
# JzAlMCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5lbnRydXN0Lm5ldDAxBgNVHR8E
# KjAoMCagJKAihiBodHRwOi8vY3JsLmVudHJ1c3QubmV0L2NzYnIxLmNybDAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwRQYDVR0gBD4wPDAwBgRV
# HSAAMCgwJgYIKwYBBQUHAgEWGmh0dHA6Ly93d3cuZW50cnVzdC5uZXQvcnBhMAgG
# BmeBDAEEATANBgkqhkiG9w0BAQ0FAAOCAgEAXvOGmTXBee7wEK/XkkPShdBb4Jig
# 4HFRyRTLUJpgDrAEJkmxz+m6mwih2kNd1G8jorn4QMdH/k0BC0iQP8jcarQ+UzUo
# vkBKR4VqHndAzIB/YbQ8T3mo5qOmoH5EhnG/EhuVgXL3DaXQ3mefxqK48Wr5/P50
# ZsZk5nk9agNhTksfzCBiywIY7GPtfnE/lroLXmgiZ+wfwNIFFmaxsqTq/MWVo40S
# pfWN7xsgzZn35zLzWXEf3ZTmeeVSIxBWKvxZOL+/eSWSasf9q2d3cbEEfTWtFME+
# qPwjF1YIGHzXeiJrkWrMNUVtTzudQ50FuJ3z/DQhXAQYMlc4NMHKgyNGpogjIcZ+
# FICrse+7C6wJP+5TkTGz4lREqrV9MDwsI5zoP6NY6kAIF6MgX3rADNuq/wMWAw10
# ZCKalF4wNXYT9dPh4+AHytnqRYhGnFTVEOLzMglAtudcFzL+zK/rbc9gPHXz7lxg
# QFUbtVmvciNoTZx0BAwQya9QW6cNZg+W5ZqV4CCiGtCw7jhJnipnnpGWbJjbxBBt
# YHwebkjntn6vMwcSce+9lTu+qYPUQn23pzTXX4aRta9WWNpVfRe927zNZEEVjTFR
# Bk+0LrKLPZzzTeNYA1TMrIj4UjxOS0YJJRn/FeenmEYufbrq4+N8//m5GZW+drkN
# ebICURpKyJ+IwkMxghgAMIIX/AIBATBjME8xCzAJBgNVBAYTAlVTMRYwFAYDVQQK
# Ew1FbnRydXN0LCBJbmMuMSgwJgYDVQQDEx9FbnRydXN0IENvZGUgU2lnbmluZyBD
# QSAtIE9WQ1MyAhBN5PZMB6ZZLJUliCaOZum4MA0GCWCGSAFlAwQCAwUAoIGcMBAG
# CisGAQQBgjcCAQwxAjAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVME8GCSqGSIb3DQEJBDFCBECoQGTn3xek
# nwDrnNj2haN4lebxi7lKGXqWXmv6jHm7F2zFjeyFHJcB7PA0OZzvc38KD7a92ktp
# bve7OEKVQKUIMA0GCSqGSIb3DQEBAQUABIIBgKSGf11CBZS15fX6QyQjQob+2oLN
# pEzx9INLTAIoDYW2evCFb9Z11j9OwenFHq8sOiL0Xkgy2br+2ZJWQH2lbhTWpiwE
# 6HrClhO8Z2Q9e9hvEkYha7/bclGVsSLu+/sGa+EeN/hoeNEZFlWEeAASlMX/hfQz
# 5SuiMugA4aYQFNBzvdhkonC09PzaamjUnPdkBCgAdgLoiEAyhGSLClpPiMnMoCbt
# Sx0B3UG4leGTFQmDnCLQQNKVIhMjvPIBTmLU80zHQYaR47z03dDSM+xB+tZx6XjB
# g8kRoGr/eeFAmKb4pt1w20VthmOex+RL1Rz0OstABWnVKh30yIcQ69tW6BdBLDT8
# cQiScX3DThDRdogCpFjQFWjy0PXvt0QxlmtkTT0TBwH533homaV1R7ztblD9KthN
# 86KqHoGxoGtyyTMPxXCmZXQeRAVhyXzlHiODc4Ry5AjEJLpkboLaYi2AxydkIOTw
# 78QHdUZBZE/M0FOW9WX6ClQ1enX2gBTRpgKZeqGCFU8wghVLBgorBgEEAYI3AwMB
# MYIVOzCCFTcGCSqGSIb3DQEHAqCCFSgwghUkAgEDMQ0wCwYJYIZIAWUDBAIBMIIB
# FgYLKoZIhvcNAQkQAQSgggEFBIIBATCB/gIBAQYKYIZIAYb6bAoDBTBRMA0GCWCG
# SAFlAwQCAwUABEAG1d6HpXcepy0zIW0256jaWVjmOsbTX9s70NFlCWk9ABcPTU+X
# iFECy6zj96s8LaOR/Sz+ls/Tczu0UAQNPvX4AgkA0vcyljBKgUAYDzIwMjMxMDEz
# MDczOTUwWjADAgEBoHmkdzB1MQswCQYDVQQGEwJDQTEQMA4GA1UECBMHT250YXJp
# bzEPMA0GA1UEBxMGT3R0YXdhMRYwFAYDVQQKEw1FbnRydXN0LCBJbmMuMSswKQYD
# VQQDEyJFbnRydXN0IFRpbWVzdGFtcCBBdXRob3JpdHkgLSBUU0ExoIIPWDCCBCow
# ggMSoAMCAQICBDhj3vgwDQYJKoZIhvcNAQEFBQAwgbQxFDASBgNVBAoTC0VudHJ1
# c3QubmV0MUAwPgYDVQQLFDd3d3cuZW50cnVzdC5uZXQvQ1BTXzIwNDggaW5jb3Jw
# LiBieSByZWYuIChsaW1pdHMgbGlhYi4pMSUwIwYDVQQLExwoYykgMTk5OSBFbnRy
# dXN0Lm5ldCBMaW1pdGVkMTMwMQYDVQQDEypFbnRydXN0Lm5ldCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSAoMjA0OCkwHhcNOTkxMjI0MTc1MDUxWhcNMjkwNzI0MTQx
# NTEyWjCBtDEUMBIGA1UEChMLRW50cnVzdC5uZXQxQDA+BgNVBAsUN3d3dy5lbnRy
# dXN0Lm5ldC9DUFNfMjA0OCBpbmNvcnAuIGJ5IHJlZi4gKGxpbWl0cyBsaWFiLikx
# JTAjBgNVBAsTHChjKSAxOTk5IEVudHJ1c3QubmV0IExpbWl0ZWQxMzAxBgNVBAMT
# KkVudHJ1c3QubmV0IENlcnRpZmljYXRpb24gQXV0aG9yaXR5ICgyMDQ4KTCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAK1NS6kShrLqoyAHFRZkKitL0b8L
# Sk2O7YB2pWe3eEDAc0LIaMDbUyvdXrh2mDWTixqdfBM6Dh9btx7P5SQUHrGBqY19
# uMxrSwPxAgzcq6VAJAB/dJShnQgps4gL9Yd3nVXN5MN+12pkq4UUhpVblzJQbz3I
# umYM4/y9uEnBdolJGf3AqL2Jo2cvxp+8cRlguC3pLMmQdmZ7lOKveNZlU1081pyy
# zykD+S+kULLUSM4FMlWK/bJkTA7kmAd123/fuQhVYIUwKfl7SKRphuM1Px6GXXp6
# Fb3vAI4VIlQXAJAmk7wOSWiRv/hH052VQsEOTd9vJs/DGCFiZkNw1tXAB+ECAwEA
# AaNCMEAwDgYDVR0PAQH/BAQDAgEGMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FFXkgdERgL7YibkIozH5oSQJFrlwMA0GCSqGSIb3DQEBBQUAA4IBAQA7m49WmzDn
# U5l8enmnTZfXGZWQ+wYfyjN8RmOPlmYk+kAbISfK5nJz8k/+MZn9yAxMaFPGgIIT
# mPq2rdpdPfHObvYVEZSCDO4/la8Rqw/XL94fA49XLB7Ju5oaRJXrGE+mH819VxAv
# mwQJWoS1btgdOuHWntFseV55HBTF49BMkztlPO3fPb6m5ZUaw7UZw71eW7v/I+9o
# GcsSkydcAy1vMNAethqs3lr30aqoJ6b+eYHEeZkzV7oSsKngQmyTylbe/m2ECwiL
# fo3q15ghxvPnPHkvXpzRTBWN4ewiN8yaQwuX3ICQjbNnm29ICBVWz7/xK3xemnbp
# WZDFfIM1EWVRMIIFEzCCA/ugAwIBAgIMWNoT/wAAAABRzg33MA0GCSqGSIb3DQEB
# CwUAMIG0MRQwEgYDVQQKEwtFbnRydXN0Lm5ldDFAMD4GA1UECxQ3d3d3LmVudHJ1
# c3QubmV0L0NQU18yMDQ4IGluY29ycC4gYnkgcmVmLiAobGltaXRzIGxpYWIuKTEl
# MCMGA1UECxMcKGMpIDE5OTkgRW50cnVzdC5uZXQgTGltaXRlZDEzMDEGA1UEAxMq
# RW50cnVzdC5uZXQgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkgKDIwNDgpMB4XDTE1
# MDcyMjE5MDI1NFoXDTI5MDYyMjE5MzI1NFowgbIxCzAJBgNVBAYTAlVTMRYwFAYD
# VQQKEw1FbnRydXN0LCBJbmMuMSgwJgYDVQQLEx9TZWUgd3d3LmVudHJ1c3QubmV0
# L2xlZ2FsLXRlcm1zMTkwNwYDVQQLEzAoYykgMjAxNSBFbnRydXN0LCBJbmMuIC0g
# Zm9yIGF1dGhvcml6ZWQgdXNlIG9ubHkxJjAkBgNVBAMTHUVudHJ1c3QgVGltZXN0
# YW1waW5nIENBIC0gVFMxMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# 2SPmFKTofEuFcVj7+IHmcotdRsOIAB840Irh1m5WMOWv2mRQfcITOfu9ZrTahPuD
# 0Cgfy3boYFBpm/POTxPiwT7B3xLLMqP4XkQiDsw66Y1JuWB0yN5UPUFeQ18oRqmm
# t8oQKyK8W01bjBdlEob9LHfVxaCMysKD4EdXfOdwrmJFJzEYCtTApBhVUvdgxgRL
# s91oMm4QHzQRuBJ4ZPHuqeD347EijzRaZcuK9OFFUHTfk5emNObQTDufN0lSp1NO
# ny5nXO2W/KW/dFGI46qOvdmxL19QMBb0UWAia5nL/+FUO7n7RDilCDkjm2lH+jzE
# 0Oeq30ay7PKKGawpsjiVdQIDAQABo4IBIzCCAR8wEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAQYwOwYDVR0gBDQwMjAwBgRVHSAAMCgwJgYIKwYBBQUH
# AgEWGmh0dHA6Ly93d3cuZW50cnVzdC5uZXQvcnBhMDMGCCsGAQUFBwEBBCcwJTAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3AuZW50cnVzdC5uZXQwMgYDVR0fBCswKTAn
# oCWgI4YhaHR0cDovL2NybC5lbnRydXN0Lm5ldC8yMDQ4Y2EuY3JsMBMGA1UdJQQM
# MAoGCCsGAQUFBwMIMB0GA1UdDgQWBBTDwnHSe9doBa47OZs0JQxiA8dXaDAfBgNV
# HSMEGDAWgBRV5IHREYC+2Im5CKMx+aEkCRa5cDANBgkqhkiG9w0BAQsFAAOCAQEA
# HSTnmnRbqnD8sQ4xRdcsAH9mOiugmjSqrGNtifmf3w13/SQj/E+ct2+P8/QftsH9
# 1hzEjIhmwWONuld307gaHshRrcxgNhqHaijqEWXezDwsjHS36FBD08wo6BVsESqf
# FJUpyQVXtWc26Dypg+9BwSEW0373LRFHZnZgghJpjHZVcw/fL0td6Wwj+Af2tX3W
# aUWcWH1hLvx4S0NOiZFGRCygU6hFofYWWLuRE/JLxd8LwOeuKXq9RbPncDDnNI7r
# evbTtdHeaxOZRrOL0k2TdbXxb7/cACjCJb+856NlNOw/DR2XjPqqiCKkGDXbBY52
# 4xDIKY9j0K6sGNnaxJ9REjCCBg8wggT3oAMCAQICEFarlXUonKWfDhfUC+oFwx8w
# DQYJKoZIhvcNAQELBQAwgbIxCzAJBgNVBAYTAlVTMRYwFAYDVQQKEw1FbnRydXN0
# LCBJbmMuMSgwJgYDVQQLEx9TZWUgd3d3LmVudHJ1c3QubmV0L2xlZ2FsLXRlcm1z
# MTkwNwYDVQQLEzAoYykgMjAxNSBFbnRydXN0LCBJbmMuIC0gZm9yIGF1dGhvcml6
# ZWQgdXNlIG9ubHkxJjAkBgNVBAMTHUVudHJ1c3QgVGltZXN0YW1waW5nIENBIC0g
# VFMxMB4XDTIyMTAwNDE3MjEwM1oXDTI5MDEwMTAwMDAwMFowdTELMAkGA1UEBhMC
# Q0ExEDAOBgNVBAgTB09udGFyaW8xDzANBgNVBAcTBk90dGF3YTEWMBQGA1UEChMN
# RW50cnVzdCwgSW5jLjErMCkGA1UEAxMiRW50cnVzdCBUaW1lc3RhbXAgQXV0aG9y
# aXR5IC0gVFNBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAMHqoyWO
# RTQXyvWzxRnMCLMLE06gcvbupHZQWA9G2BfpuHdUaYBxuopfWGc+Z1jvTlDN+sbZ
# bhMzj63Drkb2Z05tQzeLid34F53T0un5YkP+pKRU6cq3muBKKvliEdCtZ1xY/hpM
# SWpfFh+0ZZQeEBaUm3ft850bUl8ZGXXb7s9nWa3LLkpjDrGCd1pvY03ouCLsLwrb
# TJslnoMrZBm2t3Slw7pMDU2vA+YWbeTwGLOjZ8VXrO6zL69+fCuGem2yydoUUjK2
# MVAm6vZPMAqpOBXDqop0SC9Gd/8FhMin5GCF+bzZhiCqDNypoz/8k5RLlazbVLWv
# w/19LkHjoyCGUEVB1C42MSj2Ci/3wTrNBhQ++qdJwVRKusB+50kHhEpcma5nAgRI
# qbTlxR3tjx2xPyTCBwdiHEvFNulU6otrnXRL+qoBJydPuvp9J83tZo0neBcOvIdz
# zm19yY6WOKJLFexXQHimwIHZyQNs2+Mv8cqCq0B9rJpwGuLfPptYZCtrbw9oHiVn
# PnrdAb5jlFHTWoR0so8p2i8kurbPbkvNDb+B3pG3NeaShS0DwZ8UwsVc33koOgh4
# Y+4hqzJ1gN/1BA7fRtzGqbOazXbbc1cfDA0w/W91CbyfizO78GQ8/LndYoOyvJes
# lzHoOcXMsku/l2SY3S4q/sQn8FcVTVFPhK0DAgMBAAGjggFbMIIBVzAMBgNVHRMB
# Af8EAjAAMB0GA1UdDgQWBBRKDtGm6oLNjldj7P2+AzdAkCSj2DAfBgNVHSMEGDAW
# gBTDwnHSe9doBa47OZs0JQxiA8dXaDBoBggrBgEFBQcBAQRcMFowIwYIKwYBBQUH
# MAGGF2h0dHA6Ly9vY3NwLmVudHJ1c3QubmV0MDMGCCsGAQUFBzAChidodHRwOi8v
# YWlhLmVudHJ1c3QubmV0L3RzMS1jaGFpbjI1Ni5jZXIwMQYDVR0fBCowKDAmoCSg
# IoYgaHR0cDovL2NybC5lbnRydXN0Lm5ldC90czFjYS5jcmwwDgYDVR0PAQH/BAQD
# AgeAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIGA1UdIAQ7MDkwNwYKYIZIAYb6
# bAoBBzApMCcGCCsGAQUFBwIBFhtodHRwczovL3d3dy5lbnRydXN0Lm5ldC9ycGEw
# DQYJKoZIhvcNAQELBQADggEBALHAWjNiHGozMuYkMD5vxuawUjQdqUg7Bw3ukGTE
# qsqtexpHb4H3zzHEwvoGKLkXvC6EDtJHNf6ueOdjFeDaJQSxY9MLw+ujjEEpEeXq
# mO7I3cUV9sdUGv5Ujbso9gUbBIKoM4aWyI+0wiAFRAsCW6MmWuxMtGm+2PxARF3F
# x0FuPze1eBPmlRTt3s7KhLPQ288roPzus/6IAF+msHnZmg8/XZfln0IPHOfvsOmv
# cw24eZoCV5prbEsUKs6uVabPseuxoWDebdwvJlPhPXMusQiccocpyIVQiMzqXwic
# kHqqyrCb6upAM+IiFdLRd36/GJgivr3ySD/MAtLStuubSaQxggSYMIIElAIBATCB
# xzCBsjELMAkGA1UEBhMCVVMxFjAUBgNVBAoTDUVudHJ1c3QsIEluYy4xKDAmBgNV
# BAsTH1NlZSB3d3cuZW50cnVzdC5uZXQvbGVnYWwtdGVybXMxOTA3BgNVBAsTMChj
# KSAyMDE1IEVudHJ1c3QsIEluYy4gLSBmb3IgYXV0aG9yaXplZCB1c2Ugb25seTEm
# MCQGA1UEAxMdRW50cnVzdCBUaW1lc3RhbXBpbmcgQ0EgLSBUUzECEFarlXUonKWf
# DhfUC+oFwx8wCwYJYIZIAWUDBAIBoIIBpTAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTIzMTAxMzA3Mzk1MFowKQYJKoZIhvcNAQk0
# MRwwGjALBglghkgBZQMEAgGhCwYJKoZIhvcNAQELMC8GCSqGSIb3DQEJBDEiBCBk
# Lmxlan6NVw5BOgCzzuY30fdfk2OaGMWk/vfxEKx61TCCAQsGCyqGSIb3DQEJEAIv
# MYH7MIH4MIH1MIHyBCDuYRmrup3Fu4wsgb75vu6ODiX92F6z8aY1NSaB2baLrzCB
# zTCBuKSBtTCBsjELMAkGA1UEBhMCVVMxFjAUBgNVBAoTDUVudHJ1c3QsIEluYy4x
# KDAmBgNVBAsTH1NlZSB3d3cuZW50cnVzdC5uZXQvbGVnYWwtdGVybXMxOTA3BgNV
# BAsTMChjKSAyMDE1IEVudHJ1c3QsIEluYy4gLSBmb3IgYXV0aG9yaXplZCB1c2Ug
# b25seTEmMCQGA1UEAxMdRW50cnVzdCBUaW1lc3RhbXBpbmcgQ0EgLSBUUzECEFar
# lXUonKWfDhfUC+oFwx8wCwYJKoZIhvcNAQELBIICAD+8Ks7COzICyLnq+RM0Di+S
# GocoV1eDtYTzvv/Di80jU1GZVKj6E3W5HXV3LK7k+IhDEH1ziLgSXZ2zqjvNC8Xv
# oDKmsNNrd5XHznxwEAYeNDpf9lfC0xiWA3aOLFmolA+/PdaHXkMxfaIHH9UBhpk0
# HtybXoDVRYOm14CoDes6/egxpzij9FoYVtuCAUwmWjGSs6K9zaLmk7kqFIzs5NRM
# HVS45QhXxLnEsQsRY/oAlxiVgOSZHcSxLoYqfzmXaArcjVnFOtVAXsBAMBLJVWw+
# sa338IK6lGMUx6OPJ7F38Vv3p4ts/J6FBEoX/vQ540v/HzV3h1r6YggRwTgby+bc
# cTtlWFN9bPxQfxLcakrKMTDqoN37IrnsgaQna1oLyq1IKzD82MRCx0Gzl69pMBU5
# GbN/+2KGdGY0K7gNpKFZuc4jMquseg36MAhEbzlAT8w8HbXxHxlYDeBOiJRrdv/5
# zTjIPS+VhveFzwZhQdG1k7JtxT2rQQKCJN78NA0JadDAp41RufoSUXMTHDjS0Fvq
# wL/AcKq4twx8aKEAgMJmJYLL198gP+i7HCtTDAtef1LMxG3CvHwrFnstzhJggFNf
# uvudWfPJBXWv+kUTeJYZl79lgRMM33nQR60OGcUbwQRsaohMsVSjHsLcZPb7kp2m
# +Gv4kXqPJWXogRUAZ7MT
# SIG # End signature block

