[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$global:url = "https://www.digicert.com/services/v2"
$global:array = New-Object System.Collections.ArrayList
$global:dir = [environment]::GetFolderPath("MyDocuments")

#Call this function to get a list of Expiring Orders in your DigiCert account
#Your API key needs to be passed into the function call in a hashtable, which looks like: DC-GetExpiringOrders @{key="API_KEY_HERE"}
#You can pass in the number of days you'd like to filter on to the function call, which would look like: DC-GetExpiringOrders @{key="API_KEY_HERE"} -days 30
function DC-GetExpiringOrders
{    
    Param(
        #pass in the API Key
        [Parameter(HelpMessage = "An API Key must be supplied", Mandatory = $true)]
        [System.Collections.Hashtable]$apikey,
        
        #pass in the number days within which to get expiring orders
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage ="Please specify the number of days to look at expiring orders")]
		[ValidateNotNullOrEmpty()]
		[string] $days        
    )

    #timestamp the report file
    $timestamp = Get-Date -Format o | foreach {$_ -replace ":", "."}

    try
    {
        # Build filter for expiring orders within the specified time
        $current_time = get-date -Format yyyy-MM-dd+hh:mm:ss
        $plus_30_days = (get-date).AddDays($days)
        $end_time = $plus_30_days.ToString("yyyy-MM-dd+hh:mm:ss")
        $filters = "?filters[status]=issued&filters[valid_till]=$current_time...$end_time"
        $list_url = $global:url + "/order/certificate" + $filters             

        # Set the request headers
        $headers = @{
            "X-DC-DEVKEY"=$apikey.key;            
            "Accept"="application/json"
        }        
        
        # Send the API query
        $resp = Invoke-RestMethod -Uri $list_url -Method Get -Headers $headers -ContentType "application/json"
        
        #Grab the order details of each Order expiring in the next [user-specified] days, including user assignments and additional emails; exclude those orders which have already been renewed
        foreach ( $order in $resp.orders | Where-Object {$_.is_renewed -ne "false"}) 
        {            
            $ord_url = $global:url + "/order/certificate/" + $order.id
            $ord_details = Invoke-RestMethod -Uri $ord_url -Method Get -Headers $headers -ContentType "application/json"
            $userarray = @(foreach ($user in $ord_details.user_assignments)
            {
                New-Object -TypeName PSCustomObject -Property @{              
                NameEmail = $user.first_name + " " + $user.last_name + ", " + $user.email                    
                }
            }
            )
            $users = $userarray.NameEmail -join "; "
            $emailarray = @(foreach ($email in $ord_details.additional_emails)
            {
                New-Object -TypeName PSCustomObject -Property @{              
                Emails = $email
                }
            }
            )
            $emails = $emailarray.Emails -join "; "       
        
            # Create a CSV report of the orders expiring in the next [user-specified] days                        
            New-Object -TypeName PSCustomObject -Property @{
                OrderID = $order.id
                CertificateID = $order.certificate.id
                CommonName = $order.certificate.common_name
                ExpirationDate = $order.certificate.valid_till
                Renewed = $order.is_renewed
                Renewal = $ord_details.is_renewal
                OrderDate = $order.date_created
                Organization = $order.organization.name
                Product = $order.product.name
                Duplicates = $order.has_duplicates
                MainThumbprint = $ord_details.certificate.thumbprint
                MainSerialNumber = $ord_details.certificate.serial_number
                Issuer = $ord_details.certificate.ca_cert.name
                RenewalNotifications = $ord_details.disable_renewal_notifications
                Emails = $emails
                Users = $users            
            } | Export-Csv -Path $global:dir\ExpiringOrders$timestamp.csv -NoTypeInformation -Append
        }                        
    }
    catch 
    {
        throw Error($_.Exception)
    }
}

function Error( [Exception] $exception )
{
    try
    {
        $result = $exception.Response.GetResponseStream()
        $read = New-Object System.IO.StreamReader($result)
        $read.BaseStream.Position = 0
        $read.DiscardBufferedData()
        $content_type = $read.ReadToEnd() | ConvertFrom-Json
        return $content_type.errors.message
    }
    catch
    {
        return $exception.Message
    }
}
# SIG # Begin signature block
# MIIOmwYJKoZIhvcNAQcCoIIOjDCCDogCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUYVwqLuo1wb845jOBOCvIAWNV
# ZFigggvgMIIFNTCCBB2gAwIBAgIQCtvOfO8MtmSb4KtGs1D3YjANBgkqhkiG9w0B
# AQUFADBvMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2VydCBBc3N1cmVk
# IElEIENvZGUgU2lnbmluZyBDQS0xMB4XDTE3MDgwNzAwMDAwMFoXDTE4MDgxNzEy
# MDAwMFowgYAxCzAJBgNVBAYTAlVTMQswCQYDVQQIEwJVVDENMAsGA1UEBxMET3Jl
# bTEVMBMGA1UEChMMQ2xpbnQgV2lsc29uMRUwEwYDVQQDEwxDbGludCBXaWxzb24x
# JzAlBgkqhkiG9w0BCQEWGGNsaW50LnQud2lsc29uQGdtYWlsLmNvbTCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBALhPEhMNFxyYuMbrMm6sDhIB1W5MWiVH
# Z7NCRjqZbJnDj+0j3MoHiEcCfEyLcZUlDJMj6bcwIVbMjbqgYyy2rzOxzkcI4VU8
# 1aL6hbqUw2uY0hF8U5b1E00NThN1LfCQ9rbEhhg/8OcLbTnRGO0yXaxbmld4QOm1
# BbdnH+ApXIKUwDdJA0/RxgN/K2+O2J3k0nvoF1Ob30ILDfmFr7BV/msJuRJ0QDDb
# NIY0TwzxK7kgkmn+cGDK9MhFYC+ozTKSi0DgcP7Rw/SljbZuKLuLk/51jdmZp2xl
# Ecgese8qqAFF/mNSTAacH4SO54p7EDreAq0XFbSIRU8cT/rgzzro6zkCAwEAAaOC
# AbkwggG1MB8GA1UdIwQYMBaAFHtozimqwBe+SXrh5T/Wp/dFjzUyMB0GA1UdDgQW
# BBSj0RrbnEA7AUN+/EKsGEifFRxM5DAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwbQYDVR0fBGYwZDAwoC6gLIYqaHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL2Fzc3VyZWQtY3MtZzEuY3JsMDCgLqAshipodHRwOi8vY3JsNC5kaWdp
# Y2VydC5jb20vYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1s
# AwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAI
# BgZngQwBBAEwgYIGCCsGAQUFBwEBBHYwdDAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEwGCCsGAQUFBzAChkBodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDb2RlU2lnbmluZ0NBLTEuY3J0MAwG
# A1UdEwEB/wQCMAAwDQYJKoZIhvcNAQEFBQADggEBAAQkBxyiTRr5kZGbDU9lIJPe
# vVRpX6tiZGVhng2//V++xwgKXZBSXl1UdGmuKSwHQkbiyMMvekE5UgKh4VKzJ5Bl
# xqbgEvZMeS581EakRgFstqOYEintk/ItIpw09iA7mAGsHf2Sqt0jZVmxuRxybYpN
# ITdqyC8EI3T7rFljWOZpUu5MUGNrtyIu+wSL03XCe04gKt0QKKN0lYXEuyExBVjE
# ccniqw2shCjHFODM9oOEYpt/IPW0yVztlQfuh94TwkdPeoeeqyaSThl+tewDYeyo
# k7ykEtAAQiIQTgP9A0OFdl1jDx7BDOBhPIsYyoW6+BpOwM1AUxGU7WmkaWeJdQkw
# ggajMIIFi6ADAgECAhAPqEkGFdcAoL4hdv3F7G29MA0GCSqGSIb3DQEBBQUAMGUx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9v
# dCBDQTAeFw0xMTAyMTExMjAwMDBaFw0yNjAyMTAxMjAwMDBaMG8xCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xLjAsBgNVBAMTJURpZ2lDZXJ0IEFzc3VyZWQgSUQgQ29kZSBTaWduaW5n
# IENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCcfPmgjwrKiUtT
# mjzsGSJ/DMv3SETQPyJumk/6zt/G0ySR/6hSk+dy+PFGhpTFqxf0eH/Ler6QJhx8
# Uy/lg+e7agUozKAXEUsYIPO3vfLcy7iGQEUfT/k5mNM7629ppFwBLrFm6aa43Abe
# ro1i/kQngqkDw/7mJguTSXHlOG1O/oBcZ3e11W9mZJRru4hJaNjR9H4hwebFHsng
# lrgJlflLnq7MMb1qWkKnxAVHfWAr2aFdvftWk+8b/HL53z4y/d0qLDJG2l5jvNC4
# y0wQNfxQX6xDRHz+hERQtIwqPXQM9HqLckvgVrUTtmPpP05JI+cGFvAlqwH4KEHm
# x9RkO12rAgMBAAGjggNDMIIDPzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMwggHDBgNVHSAEggG6MIIBtjCCAbIGCGCGSAGG/WwDMIIBpDA6Bggr
# BgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0
# b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8A
# ZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQA
# aQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUA
# IABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUA
# IABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQA
# IAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEA
# bgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUA
# aQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjASBgNVHRMBAf8ECDAGAQH/
# AgEAMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2
# hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290
# Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURSb290Q0EuY3JsMB0GA1UdDgQWBBR7aM4pqsAXvkl64eU/1qf3RY81
# MjAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUF
# AAOCAQEAe3IdZP+IyDrBt+nnqcSHu9uUkteQWTP6K4feqFuAJT8Tj5uDG3xDxOaM
# 3zk+wxXssNo7ISV7JMFyXbhHkYETRvqcP2pRON60Jcvwq9/FKAFUeRBGJNE4Dyah
# YZBNur0o5j/xxKqb9to1U0/J8j3TbNwj7aqgTWcJ8zqAPTz7NkyQ53ak3fI6v1Y1
# L6JMZejg1NrRx8iRai0jTzc7GZQY1NWcEDzVsRwZ/4/Ia5ue+K6cmZZ40c2cURVb
# QiZyWo0KSiOSQOiG3iLCkzrUm2im3yl/Brk8Dr2fxIacgkdCcTKGCZlyCXlLnXFp
# 9UH/fzl3ZPGEjb6LHrJ9aKOlkLEM/zGCAiUwggIhAgEBMIGDMG8xCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xLjAsBgNVBAMTJURpZ2lDZXJ0IEFzc3VyZWQgSUQgQ29kZSBTaWduaW5n
# IENBLTECEArbznzvDLZkm+CrRrNQ92IwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcC
# AQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYB
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFMo72S5ziUyU
# hWM1hA7zMdCflmqvMA0GCSqGSIb3DQEBAQUABIIBAB5RnYYicgXbl9WdYLyajzu0
# bd0auv2Czw80svtetnHqaeluvMa8tWohhuC4zGdhWCBnzYP1REjsP2utWSCmGaEH
# TKxU5KIvbhD4yDb7toCjmWstc6ycoCGAOy+PcZ5NA8GnKIJnuQGdGHE24ALfqbGz
# xgBn6OSZwSO5HH3OSeSIXeaU+xYMkdlmB9rwOClFFr2uD3M8CLA/nYfXpAxJIbMJ
# E3BRFBID3nc7R9S3jMLSmIGTWWyiOM3S1tuc+eUA5ayF9WkQFfLwKj1B5OUYa4zz
# 5JpHfuGEjLXj6ZSsjfywKzdvxPNu1VN2P5ZTkkM+NlbIFvTDaevksugB+d2xrv8=
# SIG # End signature block
