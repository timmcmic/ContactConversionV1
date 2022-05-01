<#
    .SYNOPSIS

    This function enables the dynamic contact for hybird mail flow.
    
    .DESCRIPTION

    This function enables the dynamic contact for hybird mail flow.

    .PARAMETER GlobalCatalogServer

    The global catalog to make the query against.

    .PARAMETER routingContactConfig

    The original DN of the object.

    .PARAMETER originalContactConfiguration

    The original DN of the object.

    .OUTPUTS

    None

    .EXAMPLE

    enable-mailDynamiccontact -globalCatalogServer GC -routingContactConfig contactConfiguration -originalContactConfiguration contactConfiguration

    #>
    Function Enable-MailDyamiccontact
     {
        [cmcontactetbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$globalCatalogServer,
            [Parameter(Mandatory = $true)]
            $routingContactConfig,
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration,
            [Parameter(Mandatory = $false)]
            $isRetry=$FALSE
        )

        [string]$isTestError="No"

        #Declare function variables.

        $functionEmailAddress=$NULL

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Enable-MailDyamiccontact"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        #Create the dynamic distribution contact.
        #This is very import - the contact is scoped to the OU where it was created and uses the two custom attributes.
        #If the mail contact is ever moved from the OU that the contact originally existed in - hybrid mail flow breaks.

        try{
            out-logfile -string "Creating dynamic contact..."

            if ($isRetry -eq $false)
            {
                $tempOUSubstring = Get-OULocation -originalContactConfiguration $originalContactConfiguration

                new-dynamicDistributioncontact -name $originalContactConfiguration.name -alias $originalContactConfiguration.mailNickName -primarySMTPAddress $originalContactConfiguration.mail -organizationalUnit $tempOUSubstring -domainController $globalCatalogServer -includedRecipients AllRecipients -conditionalCustomAttribute1 $routingContactConfig.extensionAttribute1 -conditionalCustomAttribute2 $routingContactConfig.extensionAttribute2 -displayName $originalContactConfiguration.DisplayName 

            }
            else 
            {
                $tempOUSubstring = Get-OULocation -originalContactConfiguration $routingContactConfig

                new-dynamicDistributioncontact -name $originalContactConfiguration.name -alias $originalContactConfiguration.Alias -primarySMTPAddress $originalContactConfiguration.windowsEmailAddress -organizationalUnit $tempOUSubstring -domainController $globalCatalogServer -includedRecipients AllRecipients -conditionalCustomAttribute1 $routingContactConfig.extensionAttribute1 -conditionalCustomAttribute2 $routingContactConfig.extensionAttribute2 -displayName $originalContactConfiguration.DisplayName
            }

        }
        catch{
            out-logfile -string $_
            $isTestError="Yes"
            return $isTestError
        }

        #All of the email addresses that existed on the migrated contact need to be stamped on the new contact.

        if ($isRetry -eq $FALSE)
        {
            foreach ($address in $originalContactConfiguration.proxyAddresses)
            {
                out-logfile -string ("Adding proxy address = "+$address)

                #If the address is not a mail.onmicrosoft.com address - stamp it.
                #Otherwise skip it - this is because the address is stamped on the mail contact already.

                if (!$address.contains("mail.onmicrosoft.com"))
                {
                    out-logfile -string "Address is not a mail.onmicrosoft.com address."

                    try{
                        set-dynamicdistributioncontact -identity $originalContactConfiguration.mail -emailAddresses @{add=$address} -domainController $globalCatalogServer
                    }
                    catch{
                        out-logfile -string $_ 
                        $isTestError="Yes"
                        return $isTestError
                    }
                }
                else 
                {
                    out-logfile -string "Address is a mail.onmicrosoft.com address - skipping."    
                }
            }
        }
        else
        {
            foreach ($address in $originalContactConfiguration.emailAddresses)
            {
                out-logfile -string ("Adding proxy address = "+$address)

                #If the address is not a mail.onmicrosoft.com address - stamp it.
                #Otherwise skip it - this is because the address is stamped on the mail contact already.

                if (!$address.contains("mail.onmicrosoft.com"))
                {
                    out-logfile -string "Address is not a mail.onmicrosoft.com address."

                    try{
                        set-dynamicdistributioncontact -identity $originalContactConfiguration.windowsEmailAddress -emailAddresses @{add=$address} -domainController $globalCatalogServer
                    }
                    catch{
                        out-logfile -string $_ 
                        $isTestError="Yes"
                        return $isTestError
                    }
                }
                else 
                {
                    out-logfile -string "Address is a mail.onmicrosoft.com address - skipping."    
                }
            }
        }

        #The legacy Exchange DN must now be added to the contact.

        if ($isRetry -eq $FALSE)
        {
            $functionEmailAddress = "x500:"+$originalContactConfiguration.legacyExchangeDN

            out-logfile -string $originalContactConfiguration.legacyExchangeDN
            out-logfile -string ("Calculated x500 Address = "+$functionEmailAddress)

            try{
                set-dynamicDistributioncontact -identity $originalContactConfiguration.mail -emailAddresses @{add=$functionEmailAddress} -domainController $globalCatalogServer
            }
            catch{
                out-logfile -string $_
                $isTestError="Yes"
                return $isTestError        
            }
        }
        else 
        {
            out-logfile -string "X500 added in previous operation since it already existed on the contact."    
        }

        
        #The script intentionally does not set any other restrictions on the contact.
        #It allows all restriction to be evaluated once the mail reaches office 365.
        #The only restriction I set it require sender authentication - this ensures that anonymous email can still use the contact if the source is on prem.

        if ($isRetry -eq $FALSE)
        {
            if ($originalContactConfiguration.msExchRequireAuthToSendTo -eq $NULL)
            {
                out-logfile -string "The sender authentication setting was not set - maybe legacy version of Exchange."
                out-logfile -string "The sender authentication setting value FALSE in this instance."

                try {
                    set-dynamicdistributioncontact -identity $originalContactConfiguration.mail -RequireSenderAuthenticationEnabled $FALSE -domainController $globalCatalogServer
                }
                catch {
                    out-logfile -string $_
                    $isTestError="Yes"
                    return $isTestError
                }
            }
            else
            {
                out-logfile -string "Sender authentication setting is present - retaining setting as present."

                try {
                    set-dynamicdistributioncontact -identity $originalContactConfiguration.mail -RequireSenderAuthenticationEnabled $originalContactConfiguration.msExchRequireAuthToSendTo -domainController $globalCatalogServer
                }
                catch {
                    out-logfile -string $_
                    $isTestError="Yes"
                    return $isTestError
                }
            }
        }
        else 
        {
            try{
                set-dynamicDistributioncontact -identity $originalContactConfiguration.windowsEmailAddress -RequireSenderAuthenticationEnabled $originalContactConfiguration.RequireSenderAuthenticationEnabled -domainController $globalCatalogServer
            }
            catch{
                out-logfile -string "Unable to update require sender authentication on the contact."
                out-logfile -string $_ -isError:$TRUE
            }
        }

        #Evaluate hide from address book.

        if ($isRetry -eq $FALSE)
        {
            if (($originalContactConfiguration.msExchHideFromAddressLists -eq $TRUE) -or ($originalContactConfiguration.msExchHideFromAddressLists -eq $FALSE))
            {
                out-logfile -string "Evaluating hide from address list."

                try {
                    set-dynamicdistributioncontact -identity $originalContactConfiguration.mail -HiddenFromAddressListsEnabled $originalContactConfiguration.msExchHideFromAddressLists -domainController $globalCatalogServer
                }
                catch {
                    out-logfile -string $_
                    $isTestError="Yes"
                    return $isTestError
                }
            }
            else
            {
                out-logfile -string "Hide from address list settings retained at default value - not set."
            }
        }
        else 
        {
            try {
                set-dynamicdistributioncontact -identity $originalContactConfiguration.windowsEmailAddress -HiddenFromAddressListsEnabled $originalContactConfiguration.HiddenFromAddressListsEnabled -domainController $globalCatalogServer
            }
            catch {
                out-logfile -string $_
                $isTestError="Yes"
                return $isTestError
            }
        }

        Out-LogFile -string "END Enable-MailDyamiccontact"
        Out-LogFile -string "********************************************************************************"
    }