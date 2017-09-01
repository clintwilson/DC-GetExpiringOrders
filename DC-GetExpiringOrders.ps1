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