<#
    .SYNOPSIS

    This function tests each accepted domain on the contact to ensure it appears in Office 365.

    .DESCRIPTION

    This function tests each accepted domain on the contact to ensure it appears in Office 365.

    .EXAMPLE

    Test-AcceptedDomain -originalContactConfiguration $originalContactConfiguration

    #>
    Function Test-AcceptedDomain
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration
        )

        #Define variables that will be utilzed in the function.

        [array]$originalcontactAddresses=@()
        [array]$originalcontactDomainNames=@()

        #Initiate the test.
        
        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Test-AcceptedDomain"
        Out-LogFile -string "********************************************************************************"

        foreach ($address in $originalContactConfiguration.proxyAddresses)
        {
            Out-logfile -string "Testing proxy address for SMTP"
            out-logfile -string $address

            if ($address -like "smtp*")
            {
                out-logfile -string ("Address is smtp address: "+$address)

                $tempAddress=$address.split("@")

                $originalcontactDomainNames+=$tempAddress[1]
            }
            else 
            {
                out-logfile -string ("Address is not an SMTP Address - skip.")
            }
        }

        #It is possible that the contact does not have proxy address but just mail - this is now a supported scenario.
        #To get this far the object has to have mail.

        out-logfile -string ("The mail address is: "+$originalContactConfiguration.mail)
        $tempAddress=$originalContactConfiguration.mail.split("@")
        $originalcontactDomainNames+=$tempAddress[1]
        

        $originalcontactDomainNames=$originalcontactDomainNames | select-object -Unique

        out-logfile -string "Unique domain names on the contact."
        out-logfile -string $originalcontactDomainNames

        foreach ($domain in $originalcontactDomainNames)
        {
            out-logfile -string "Testing Office 365 for Domain Name."

            if (get-o365acceptedDomain -identity $domain)
            {
                out-logfile -string ("Domain exists in Office 365. "+$domain)
            }
            else 
            {
                out-logfile -string $domain
                out-logfile -string "contact cannot be migrated until the domain is an accepted domain in Office 365 or removed from the contact."    
                out-logfile -string "Email address exists on contact that is not in Office 365." -isError:$TRUE
            }
        }

        Out-LogFile -string "END Test-AcceptedDomain"
        Out-LogFile -string "********************************************************************************"
    }