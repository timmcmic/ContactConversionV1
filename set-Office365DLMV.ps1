<#
    .SYNOPSIS

    This function sets the multi valued attributes of the contact

    .DESCRIPTION

    This function sets the multi valued attributes of the contact.
    For each of use - I've combined these into a single function instead of splitting them out.dddd

    .PARAMETER originalContactConfiguration

    The original configuration of the contact on premises.

    .PARAMETER exchangecontactMembership

    The array of members of the contact.

    .PARAMETER exchangeRejectMessages

    The array of objects with reject message permissions.

    .PARAMETER exchangeAcceptMessages

    The array of users with accept message permissions.

    .PARAMETER exchangeManagedBy

    The array of objects with managedBY permissions.

    .PARAMETER exchangeModeratedBy

    The array of moderators.

    .PARAMETER exchangeBypassModeration

    The list of users / contacts that have bypass moderation rights.

    .PARAMETER exchangeFrantSendOnBehalfTo

    The list of objecst that have grant send on behalf to rights.

    .OUTPUTS

    None

    .EXAMPLE

    set-Office365contactMV -originalContactConfiguration -exchangecontactMembership -exchangeRejectMessage -exchangeAcceptMessage -exchangeManagedBy -exchangeModeratedBy -exchangeBypassMOderation -exchangeGrantSendOnBehalfTo.

    [array$exchangecontactMembershipSMTP=$NULL
    [array]$exchangeRejectMessagesSMTP=$NULL
    [array]$exchangeAcceptMessageSMTP=$NULL
    [array]$exchangeManagedBySMTP=$NULL
    [array]$exchangeModeratedBySMTP=
    [array]$exchangeBypassModerationSMTP=$NULL 
    [array]$exchangeGrantSendOnBehalfToSMTP



    #>
    Function set-Office365contactMV
     {
        [cmcontactetbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration,
            [Parameter(Mandatory = $true)]
            $office365contactConfiguration,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$exchangecontactMembershipSMTP=$NULL,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$exchangeRejectMessagesSMTP=$NULL,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$exchangeAcceptMessageSMTP=$NULL,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$exchangeManagedBySMTP=$NULL,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$exchangeModeratedBySMTP=$NULL,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$exchangeBypassModerationSMTP=$NULL,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$exchangeGrantSendOnBehalfToSMTP=$NULL,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$exchangeSendAsSMTP=$NULL,
            [Parameter(Mandatory=$true)]
            [string]$contactTypeOverride,
            [Parameter(Mandatory=$true)]
            $office365contactConfigurationPostMigration,
            [Parameter(Mandatory=$TRUE)]
            $mailOnMicrosoftComDomain,
            [Parameter(Mandatory=$TRUE)]
            $allowNonSyncedcontact=$FALSE,
            [Parameter(Mandatory=$TRUE)]
            $allOffice365SendAsAccessOncontact=$NULL
        )

        #Declare function variables.

        [array]$functionDirectoryObjectID = $NULL
        $functionEmailAddress = $NULL
        [boolean]$routingAddressIsPresent=$FALSE
        [string]$hybridRemoteRoutingAddress=$NULL
        [int]$functionLoopCounter=0
        [boolean]$functionFirstRun=$TRUE
        [array]$functionRecipients=@()
        [array]$functionEmailAddresses=@()
        [string]$functionMail=""
        [string]$functionMailNickname=""
        [string]$functionExternalDirectoryObjectID = ""
        [string]$functionEmailAddressToRemove = ""

        [boolean]$isTestError=$false
        [array]$functionErrors=@()

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN SET-Office365contactMV"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("originalContactConfiguration = ")
        out-logfile -string $originalContactConfiguration

        out-logfile -string ("Office 365 contact Configuration:")
        out-logfile -string $office365contactConfiguration

        out-logfile -string ("Office 365 contact Configuration Post Migration")
        out-logfile -string $office365contactConfigurationPostMigration

        out-logfile -string "Resetting all SMTP addresses on the object to match on premises."

        foreach ($address in $originalContactConfiguration.proxyAddresses)
        {
            if ($address.contains("mail.onmicrosoft.com"))
            {
                out-logfile -string ("Hybrid remote routing address found.")
                out-logfile -string $address
                $routingAddressIsPresent=$TRUE
            }

            out-logfile -string $address
            $functionEmailAddresses+=$address.tostring()
        }

        foreach ($address in $office365contactConfiguration.emailAddresses)
        {
            if ($address.contains("mail.onmicrosoft.com"))
            {
                out-logfile -string ("Hybrid remote routing address found.")
                out-logfile -string $address
                $routingAddressIsPresent=$TRUE
            }

            out-logfile -string $address
            $functionEmailAddresses+=$address.tostring()
        }

        $functionEmailAddresses = $functionEmailAddresses | select-object -unique

        out-logfile -string $functionEmailAddresses

        $functionExternalDirectoryObjectID = $office365contactConfigurationPostMigration.externalDirectoryObjectID
        
        out-logfile -string "External directory object ID utilized for set commands:"
        out-logfile -string $functionExternalDirectoryObjectID

        $functionEmailAddressToRemove = $office365contactConfigurationPostMigration.primarySMTPAddress

        out-logfile -string "Email address to remove after resetting attributes."
        out-logfile -string $functionEmailAddressToRemove

        if ($originalContactConfiguration.mailNickName -ne $NULL)
        {
            out-logfile -string "Mail nickname present on premsies -> using this value."
            $functionMailNickName = $originalContactConfiguration.mailNickName
            out-logfile -string $functionMailNickName
        }
        else 
        {
            out-logfile -string "Mail nickname not present on premises -> using Office 365 value."
            $functionMailNickName = $office365contactConfiguration.alias
            out-logfile -string $functionMailNickName
        }

        #With the new temp contact logic - the fast deletion and then immediately moving into set operations sometimes caused cache collisions.
        #This caused the following bulk logic to fail - then the individual set logics would also fail.
        #This left us with the temp contact without any actual SMTP addresses.
        #New logic - try / sleep 10 times then try the individuals.

        $maxRetries = 0

        Do
        {
            try {
                $isTestError=$FALSE

                out-logfile -string ("Max retry attempt: "+$maxRetries.toString())

                Set-O365Distributioncontact -identity $functionExternalDirectoryObjectID -emailAddresses $functionEmailAddresses -errorAction STOP -BypassSecuritycontactManagerCheck

                $maxRetries = 10 #The previous set was successful - so immediately bail.
            }
            catch {
                out-logfile -string "Error bulk updating email addresses on distribution contact."
                out-logfile -string $_
                $isTestError=$TRUE
                out-logfile -string "Starting 10 second sleep before trying bulk update."
                start-sleep -s 10
                $maxRetries = $maxRetries+1
            }
        }
        while($maxRetries -lt 10)
        
        if ($isTestError -eq $TRUE)
        {
            out-logfile -string "Attempting SMTP address updates per address."

            out-logfile -string "Establishing contact primary SMTP Address."

            try {
                set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -primarySMTPAddress $originalContactConfiguration.mail -errorAction STOP
            }
            catch {
                out-logfile -string "Error establishing new contact primary SMTP Address."

                out-logfile -string $_
                
                $isErrorObject = new-Object psObject -property @{
                    PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                    ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                    Alias = $functionMailNickName
                    Name = $originalContactConfiguration.name
                    Attribute = "Cloud Proxy Addresses"
                    ErrorMessage = ("Unable to set cloud distribution contact primary SMTP address to match on-premsies mail address.")
                    ErrorMessageDetail = $_
                }

                out-logfile -string $isErrorObject

                $functionErrors+=$isErrorObject
            }

            foreach ($address in $functionEmailAddresses)
            {
                out-logfile -string ("Processing address: "+$address)

                try{
                    Set-O365Distributioncontact -identity $functionExternalDirectoryObjectID -emailAddresses @{add=$address} -errorAction STOP -BypassSecuritycontactManagerCheck
                }
                catch{
                    out-logfile -string ("Error processing address: "+$address)

                    out-logfile -string $_

                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                        ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                        Alias = $functionMailNickName
                        Name = $originalContactConfiguration.name
                        Attribute = "Cloud Proxy Addresses"
                        ErrorMessage = ("Address "+$address+" could not be added to new cloud distribution contact.  Manual addition required.")
                        ErrorMessageDetail = $_
                    }

                    out-logfile -string $isErrorObject

                    $functionErrors+=$isErrorObject
                }
            }
        }
        
        #Operation set complete - reset isError.

        $isTestError=$FALSE

        if ($originalContactConfiguration.legacyExchangeDN -ne $NULL)
        {
            out-logfile -string "Processing on premises legacy ExchangeDN to X500"
            out-logfile -string $originalContactConfiguration.legacyExchangeDN

            $functionEmailAddress = "X500:"+$originalContactConfiguration.legacyExchangeDN

            out-logfile -string ("The x500 address to process = "+$functionEmailAddress)

            try {
                Set-O365Distributioncontact -identity $functionExternalDirectoryObjectID -emailAddresses @{add=$functionEmailAddress} -errorAction STOP -BypassSecuritycontactManagerCheck
            }
            catch {
                out-logfile -string ("Error processing address: "+$functionEmailAddress)

                out-logfile -string $_

                $isErrorObject = new-Object psObject -property @{
                    PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                    ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                    Alias = $functionMailNickName
                    Name = $originalContactConfiguration.name
                    Attribute = "Cloud Proxy Addresses"
                    ErrorMessage = ("Address "+$functionEmailAddress+" could not be added to new cloud distribution contact.  Manual addition required.")
                    ErrorMessageDetail = $_
                }

                out-logfile -string $isErrorObject

                $functionErrors+=$isErrorObject
            }
        }

        if ($allowNonSyncedcontact -eq $FALSE)
        {
            out-logfile -string "Processing original cloud legacy ExchangeDN to X500"
            out-logfile -string $office365contactConfiguration.legacyExchangeDN

            $functionEmailAddress = "X500:"+$office365contactConfiguration.legacyExchangeDN

            out-logfile -string ("The x500 address to process = "+$functionEmailAddress)

            try {
                Set-O365Distributioncontact -identity $functionExternalDirectoryObjectID -emailAddresses @{add=$functionEmailAddress} -errorAction STOP -BypassSecuritycontactManagerCheck
            }
            catch {
                out-logfile -string ("Error processing address: "+$functionEmailAddress)

                out-logfile -string $_

                $isErrorObject = new-Object psObject -property @{
                    PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                    ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                    Alias = $functionMailNickName
                    Name = $originalContactConfiguration.name
                    Attribute = "Cloud Proxy Addresses"
                    ErrorMessage = ("Address "+$functionEmailAddress+" could not be added to new cloud distribution contact.  Manual addition required.")
                    ErrorMessageDetail = $_
                }

                out-logfile -string $isErrorObject

                $functionErrors+=$isErrorObject
            }
        }

        if ($routingAddressIsPresent -eq $FALSE)
        {
            out-logfile -string "A hybrid remote routing address was not present.  Adding hybrid remote routing address."
            $hybridRemoteRoutingAddress=$functionMailNickName+"@"+$mailOnMicrosoftComDomain

            out-logfile -string ("Hybrid remote routing address = "+$hybridRemoteRoutingAddress)

            try {
                Set-O365Distributioncontact -identity $functionExternalDirectoryObjectID -emailAddresses @{add=$hybridRemoteRoutingAddress} -errorAction STOP -BypassSecuritycontactManagerCheck
            }
            catch {
                out-logfile -string ("Error processing address: "+$hybridRemoteRoutingAddress)

                out-logfile -string $_

                $isErrorObject = new-Object psObject -property @{
                    PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                    ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                    Alias = $functionMailNickName
                    Name = $originalContactConfiguration.name
                    Attribute = "Cloud Proxy Addresses"
                    ErrorMessage = ("Address "+$hybridRemoteRoutingAddress+" could not be added to new cloud distribution contact.  Manual addition required.")
                    ErrorMessageDetail = $_
                }

                out-logfile -string $isErrorObject

                $functionErrors+=$isErrorObject
            }
        }

        $isTestError=$FALSE

        out-logFile -string "Evaluating exchangecontactMembershipSMTP"

        if ($exchangecontactMembershipSMTP -ne $NULL)
        {
            #All of the members were previously verified as present - so no member should be gone by now unless removed.
            #This adds all members as a single operation.  Errors we silently continue.

            #Ensureing all addresses in the array are unique.
            foreach ($member in $exchangecontactMembershipSMTP)
            {
                if ($member.externalDirectoryObjectID -ne $NULL)
                {
                    out-logfile -string ("Processing directory ID: "+$member.ExternalDirectoryObjectID)
                    $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")
                    $functionRecipients+=$functionDirectoryObjectID[1]
                }
                else 
                {
                    out-logfile -string ("Processing SMTPAddress: "+$member.primarySMTPAddressOrUPN)  
                    $functionRecipients+=$member.primarySMTPAddressOrUPN    
                }
            }

            #Becuase contacts could have been mirgated and retained - this ensures that all SMTP addresses and GUIDs in the array are unique.

            $functionRecipients = $functionRecipients | select-object -Unique

            out-logfile -string "Updating membership with unique values."
            out-logfile -string $functionRecipients

            #Using update to reset the entire membership of the contact to the unique array.
            #Alberto Larrinaga for the suggestion.

            try {
                update-o365DistributioncontactMember -identity $functionExternalDirectoryObjectID -members $functionRecipients -BypassSecuritycontactManagerCheck -confirm:$FALSE -errorAction Stop
            }
            catch {
                out-logfile -string "Unable to bulk update distribution contact membership."

                out-logfile -string $_

                $isTestError=$TRUE
            }
            
            if ($isTestError -eq $TRUE)
            {
                out-logfile -string "Attempting to update membership individually..."

                foreach ($recipient in $functionRecipients)
                {
                    out-logfile -string ("Attempting to add recipient: "+$recipient)


                    try {
                        add-O365DistributioncontactMember -identity $functionExternalDirectoryObjectID -member $recipient -BypassSecuritycontactManagerCheck -errorAction STOP
                    }
                    catch {
                        out-logfile -string ("Error procesing recipient: "+$recipient)

                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                            ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                            Alias = $originalContactConfiguration.mailNickName
                            Name = $functionMailNickName
                            Attribute = "Cloud Distribution contact Member"
                            ErrorMessage = ("Member "+$recipient+" unable to add to cloud distribution contact.  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
            }
        }
        else 
        {
            Out-LogFile -string "There were no members to process."    
        }

        $isTestError=$FALSE #Resetting error trigger.

        $functionRecipients=@() #Reset the test array.

        out-logFile -string "Evaluating exchangeRejectMessagesSMTP"

        if ($exchangeRejectMessagesSMTP -ne $NULL)
        {
            foreach ($member in $exchangeRejectMessagesSMTP)
            {
                #Implement some protections for larger operations to ensure we do not exhaust our powershell budget.

                if ($member.externalDirectoryObjectID -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.externalDirectoryObjectID)

                    $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

                    out-LogFile -string ("Processing updated member = "+$functionDirectoryObjectID[1])

                    $functionRecipients+=$functionDirectoryObjectID[1]
                }
                elseif ($member.primarySMTPAddressOrUPN -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.PrimarySMTPAddressOrUPN)

                    $functionRecipients+=$member.primarySMTPAddressOrUPN    
                }
                else 
                {
                    out-logfile -string "Invalid function object for recipient." -isError:$TRUE
                } 
            }

            #Becuase contacts could have been mirgated and retained - this ensures that all SMTP addresses and GUIDs in the array are unique.

            $functionRecipients = $functionRecipients | select-object -Unique

            out-logfile -string "Updating reject messages SMTP with unique values."
            out-logfile -string $functionRecipients

            try {
                set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -RejectMessagesFromSendersOrMembers $functionRecipients -errorAction STOP -BypassSecuritycontactManagerCheck
            }
            catch {
                out-logfile -string "Error bulk updating RejectMessagesFromSendersOrMembers"

                out-logfile -string $_

                $isTestError=$TRUE
            }

            if ($isTestError -eq $TRUE)
            {
                out-logfile -string "Attempting individual update of RejectMessagesFromSendersOrMembers"

                foreach ($recipient in $functionRecipients)
                {
                    out-logfile -string ("Attempting to add recipient: "+$recipient)

                    try {
                        set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -RejectMessagesFromSendersOrMembers @{Add=$recipient} -errorAction STOP -BypassSecuritycontactManagerCheck                    }
                    catch {
                        out-logfile -string ("Error procesing recipient: "+$recipient)

                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                            ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                            Alias = $functionMailNickName
                            Name = $originalContactConfiguration.name
                            Attribute = "Cloud Distribution contact RejectMessagesFromSendersOrMembers"
                            ErrorMessage = ("Member of RejectMessagesFromSendersOrMembers "+$recipient+" unable to add to cloud distribution contact.  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
            }

        }
        else 
        {
            Out-LogFile -string "There were no members to process."    
        }

        $isTestError = $FALSE #Reset error tracker.

        $functionRecipients=@() #Reset the test array.

        out-logFile -string "Evaluating exchangeAcceptMessagesSMTP"

        if ($exchangeAcceptMessageSMTP -ne $NULL)
        {
            foreach ($member in $exchangeAcceptMessageSMTP)
            {
                #Implement some protections for larger operations to ensure we do not exhaust our powershell budget.

                if ($member.externalDirectoryObjectID -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.externalDirectoryObjectID)

                    $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

                    out-LogFile -string ("Processing updated member = "+$functionDirectoryObjectID[1])

                    $functionRecipients+=$functionDirectoryObjectID[1]
                }
                elseif ($member.primarySMTPAddressOrUPN -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.PrimarySMTPAddressOrUPN)

                    $functionRecipients+=$member.primarySMTPAddressOrUPN    
                }
                else 
                {
                    out-logfile -string "Invalid function object for recipient." -isError:$TRUE
                } 
            }

            #Becuase contacts could have been mirgated and retained - this ensures that all SMTP addresses and GUIDs in the array are unique.

            $functionRecipients = $functionRecipients | select-object -Unique

            out-logfile -string "Updating accept messages SMTP with unique values."
            out-logfile -string $functionRecipients

            try {
                set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -AcceptMessagesOnlyFromSendersOrMembers $functionRecipients -errorAction STOP -BypassSecuritycontactManagerCheck

            }
            catch {
                out-logfile -string "Error bulk updating AcceptMessagesOnlyFromSendersOrMembers."

                out-logfile -string $_

                $isTestError = $TRUE
            }

            if ($isTestError -eq $TRUE)
            {
                out-logfile -string "Attempting individual update of AcceptMessagesOnlyFromSendersOrMembers"

                foreach ($recipient in $functionRecipients)
                {
                    out-logfile -string ("Attempting to add recipient: "+$recipient)

                    try {
                        set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -AcceptMessagesOnlyFromSendersOrMembers @{Add=$recipient} -errorAction STOP -BypassSecuritycontactManagerCheck                    }
                    catch {
                        out-logfile -string ("Error procesing recipient: "+$recipient)

                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                            ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                            Alias = $functionMailNickName
                            Name = $originalContactConfiguration.name
                            Attribute = "Cloud Distribution contact AcceptMessagesOnlyFromSendersOrMembers"
                            ErrorMessage = ("Member of AcceptMessagesOnlyFromSendersOrMembers "+$recipient+" unable to add to cloud distribution contact.  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
            }
        }
        else 
        {
            Out-LogFile -string "There were no members to process."    
        }

        $isTestError = $FALSE #Reset error tracker.

        $functionRecipients=@() #Reset the test array.

        out-logFile -string "Evaluating exchangeManagedBySMTP"

        if ($exchangeManagedBySMTP -ne $NULL)
        {
            foreach ($member in $exchangeManagedBySMTP)
            {
                #Implement some protections for larger operations to ensure we do not exhaust our powershell budget.

                if ($member.externalDirectoryObjectID -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.externalDirectoryObjectID)

                    $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

                    out-LogFile -string ("Processing updated member = "+$functionDirectoryObjectID[1])

                    $functionRecipients+=$functionDirectoryObjectID[1]
                }
                elseif ($member.primarySMTPAddressOrUPN -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.PrimarySMTPAddressOrUPN)

                    $functionRecipients+=$member.primarySMTPAddressOrUPN    
                }
                else 
                {
                    out-logfile -string "Invalid function object for recipient." -isError:$TRUE
                } 
            }

            #Becuase contacts could have been mirgated and retained - this ensures that all SMTP addresses and GUIDs in the array are unique.

            $functionRecipients = $functionRecipients | select-object -Unique

            out-logfile -string "Updating managed by SMTP with unique values."
            out-logfile -string $functionRecipients

            try {
                set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -managedBy $functionRecipients -errorAction STOP -BypassSecuritycontactManagerCheck
            }
            catch {
                out-logfile -string "Unable to bulk update managedBy"

                out-logfile $_

                $isTestError=$TRUE
            }

            if ($isTestError -eq $TRUE)
            {
                out-logfile -string "Attempting individual update of ManagedBy"

                foreach ($recipient in $functionRecipients)
                {
                    out-logfile -string ("Attempting to add recipient: "+$recipient)

                    try {
                        set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -managedBy @{Add=$recipient} -errorAction STOP -BypassSecuritycontactManagerCheck                    }
                    catch {
                        out-logfile -string ("Error procesing recipient: "+$recipient)

                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                            ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                            Alias = $functionMailNickName
                            Name = $originalContactConfiguration.name
                            Attribute = "Cloud Distribution contact ManagedBy"
                            ErrorMessage = ("Member of ManagedBy "+$recipient+" unable to add to cloud distribution contact.  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
            }
        }
        else 
        {
            Out-LogFile -string "There were no members to process."    
        }

        $isTestError = $FALSE #Reset error tracker.

        $functionRecipients=@() #Reset the test array.

        out-logFile -string "Evaluating exchangeModeratedBy"

        if ($exchangeModeratedBySMTP -ne $NULL)
        {
            foreach ($member in $exchangeModeratedBySMTP)
            {
                #Implement some protections for larger operations to ensure we do not exhaust our powershell budget.

                if ($member.externalDirectoryObjectID -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.externalDirectoryObjectID)

                    $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

                    out-LogFile -string ("Processing updated member = "+$functionDirectoryObjectID[1])

                    $functionRecipients+=$functionDirectoryObjectID[1]
                }
                elseif ($member.primarySMTPAddressOrUPN -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.PrimarySMTPAddressOrUPN)

                    $functionRecipients+=$member.primarySMTPAddressOrUPN    
                }
                else 
                {
                    out-logfile -string "Invalid function object for recipient." -isError:$TRUE
                } 
            }

            #Becuase contacts could have been mirgated and retained - this ensures that all SMTP addresses and GUIDs in the array are unique.

            $functionRecipients = $functionRecipients | select-object -Unique

            out-logfile -string "Updating moderated by SMTP with unique values."
            out-logfile -string $functionRecipients

            try {
                set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -moderatedBy $functionRecipients -errorAction STOP -BypassSecuritycontactManagerCheck
            }
            catch {
                out-logfile -string "Unable to bulk update moderatedBy."

                out-logfile -string $_

                $isTestError=$TRUE
            }

            if ($isTestError -eq $TRUE)
            {
                out-logfile -string "Attempting individual update of ModeratedBy"

                foreach ($recipient in $functionRecipients)
                {
                    out-logfile -string ("Attempting to add recipient: "+$recipient)

                    try {
                        set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -moderatedBy @{Add=$recipient} -errorAction STOP -BypassSecuritycontactManagerCheck                    }
                    catch {
                        out-logfile -string ("Error procesing recipient: "+$recipient)

                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                            ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                            Alias = $functionMailNickName
                            Name = $originalContactConfiguration.name
                            Attribute = "Cloud Distribution contact ModeratedBy"
                            ErrorMessage = ("Member of ModeratedBy "+$recipient+" unable to add to cloud distribution contact.  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
            }
        }
        else 
        {
            Out-LogFile -string "There were no members to process."    
        }

        $isTestError=$FALSE

        $functionRecipients=@() #Reset the test array.

        out-logFile -string "Evaluating exchangeBypassModerationSMTP"

        if ($exchangeBypassModerationSMTP -ne $NULL)
        {
            foreach ($member in $exchangeBypassModerationSMTP)
            {
                #Implement some protections for larger operations to ensure we do not exhaust our powershell budget.

                if ($member.externalDirectoryObjectID -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.externalDirectoryObjectID)

                    $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

                    out-LogFile -string ("Processing updated member = "+$functionDirectoryObjectID[1])

                    $functionRecipients+=$functionDirectoryObjectID[1]
                }
                elseif ($member.primarySMTPAddressOrUPN -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.PrimarySMTPAddressOrUPN)

                    $functionRecipients+=$member.primarySMTPAddressOrUPN    
                }
                else 
                {
                    out-logfile -string "Invalid function object for recipient." -isError:$TRUE
                } 
            }

            #Becuase contacts could have been mirgated and retained - this ensures that all SMTP addresses and GUIDs in the array are unique.

            $functionRecipients = $functionRecipients | select-object -Unique

            out-logfile -string "Updating bypass moderation from senders or members SMTP with unique values."
            out-logfile -string $functionRecipients

            try {
                set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -BypassModerationFromSendersOrMembers $functionRecipients -errorAction STOP -BypassSecuritycontactManagerCheck
            }
            catch {
                out-logfile -string "Unable to bulk modify bypassModerationFromSendersOrMembers"

                out-logfile -string $_

                $isTestError=$TRUE
            }

            if ($isTestError -eq $TRUE)
            {
                out-logfile -string "Attempting individual update of BypassModerationFromSendersOrMembers"

                foreach ($recipient in $functionRecipients)
                {
                    out-logfile -string ("Attempting to add recipient: "+$recipient)

                    try {
                        set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -BypassModerationFromSendersOrMembers @{Add=$recipient} -errorAction STOP -BypassSecuritycontactManagerCheck                    }
                    catch {
                        out-logfile -string ("Error procesing recipient: "+$recipient)

                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                            ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                            Alias = $functionMailNickName
                            Name = $originalContactConfiguration.name
                            Attribute = "Cloud Distribution contact BypassModerationFromSendersOrMembers"
                            ErrorMessage = ("Member of BypassModerationFromSendersOrMembers "+$recipient+" unable to add to cloud distribution contact.  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
            }
        }
        else 
        {
            Out-LogFile -string "There were no members to process."    
        }

        $isTestError=$FALSE

        $functionRecipients=@() #Reset the test array.

        out-logFile -string "Evaluating exchangeGrantSendOnBehalfToSMTP"

        if ($exchangeGrantSendOnBehalfToSMTP -ne $NULL)
        {
            foreach ($member in $exchangeGrantSendOnBehalfToSMTP)
            {
                #Implement some protections for larger operations to ensure we do not exhaust our powershell budget.

                if ($member.externalDirectoryObjectID -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.externalDirectoryObjectID)

                    $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

                    out-LogFile -string ("Processing updated member = "+$functionDirectoryObjectID[1])

                    $functionRecipients+=$functionDirectoryObjectID[1]
                }
                elseif ($member.primarySMTPAddressOrUPN -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.PrimarySMTPAddressOrUPN)

                    $functionRecipients+=$member.primarySMTPAddressOrUPN    
                }
                else 
                {
                    out-logfile -string "Invalid function object for recipient." -isError:$TRUE
                } 
            }

            #Becuase contacts could have been mirgated and retained - this ensures that all SMTP addresses and GUIDs in the array are unique.

            $functionRecipients = $functionRecipients | select-object -Unique

            out-logfile -string "Updating grant send on behalf to SMTP with unique values."
            out-logfile -string $functionRecipients

            try {
                set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -GrantSendOnBehalfTo $functionRecipients -errorAction STOP -BypassSecuritycontactManagerCheck
            }
            catch {
                out-logfile -string "Unable to bulk updated GrantSendOnBehalfTo."

                out-logfile -string $_

                $isTestError=$TRUE
            }

            if ($isTestError -eq $TRUE)
            {
                out-logfile -string "Attempting individual update of GrantSendOnBehalfTo"

                foreach ($recipient in $functionRecipients)
                {
                    out-logfile -string ("Attempting to add recipient: "+$recipient)

                    try {
                        set-o365Distributioncontact -identity $functionExternalDirectoryObjectID -GrantSendOnBehalfTo @{Add=$recipient} -errorAction STOP -BypassSecuritycontactManagerCheck                    }
                    catch {
                        out-logfile -string ("Error procesing recipient: "+$recipient)

                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                            ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                            Alias = $functionMailNickName
                            Name = $originalContactConfiguration.name
                            Attribute = "Cloud Distribution contact GrantSendOnBehalfTo"
                            ErrorMessage = ("Member of GrantSendOnBehalfTo "+$recipient+" unable to add to cloud distribution contact.  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
            }
        }
        else 
        {
            Out-LogFile -string "There were no members to process."    
        }

        $isTestError=$FALSE

        out-logFile -string "Evaluating exchangeSendAsSMTP"

        if ($exchangeSendAsSMTP -ne $NULL)
        {
            foreach ($member in $exchangeSendAsSMTP)
            {
                if ($member.externalDirectoryObjectID -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.externalDirectoryObjectID)

                    $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

                    out-LogFile -string ("Processing updated member = "+$functionDirectoryObjectID[1])

                    try {
                        add-o365RecipientPermission -Identity $functionExternalDirectoryObjectID -Trustee $functionDirectoryObjectID[1] -AccessRights "SendAs" -confirm:$FALSE
                    }
                    catch {
                        out-logfile -string "Unable to add member. "

                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $member.externalDirectoryObjectID
                            ExternalDirectoryObjectID = $NULL
                            Alias = $NULL
                            Name = $NULL
                            Attribute = "Cloud Distribution contact SendAs"
                            ErrorMessage = ("Unable to add migrated distribution contact with send as to "+$member.externalDirectoryObjectID+".  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
                elseif ($member.primarySMTPAddressOrUPN -ne $NULL)
                {
                    out-LogFile -string ("Processing member = "+$member.PrimarySMTPAddressOrUPN)

                    try {
                        add-o365RecipientPermission -Identity $functionExternalDirectoryObjectID -Trustee $member.primarySMTPAddressOrUPN -AccessRights "SendAs" -confirm:$FALSE
                    }
                    catch {
                        out-logfile -string "Unable to add member. "
                        out-logfile -string $_

                        $isErrorObject = new-Object psObject -property @{
                            PrimarySMTPAddressorUPN = $member.primarySMTPAddressorUPN
                            ExternalDirectoryObjectID = $NULL
                            Alias = $NULL
                            Name = $NULL
                            Attribute = "Cloud Distribution contact SendAs"
                            ErrorMessage = ("Unable to add migrated distribution contact with send as to "+$member.primarySMTPAddressOrUPN+".  Manual addition required.")
                            ErrorMessageDetail = $_
                        }

                        out-logfile -string $isErrorObject

                        $functionErrors+=$isErrorObject
                    }
                }
                else 
                {
                    out-logfile -string "Invalid function object for recipient." -isError:$TRUE
                } 
            }
        }
        else 
        {
            Out-LogFile -string "There were no members to process."    
        }

        out-logfile -string "Resetting send as directly set on the contact to be migrated."

        if ($allOffice365SendAsAccessOncontact -ne $NULL)
        {
            foreach ($member in $allOffice365SendAsAccessOncontact)
            {
                out-logfile -string ("Processing trustee: "+$member.trustee)

                try
                {
                    add-o365RecipientPermission -identity $functionExternalDirectoryObjectID -trustee $member.trustee -accessRights $member.accessRights -confirm:$FALSE -errorAction STOP
                }
                catch
                {
                    out-logfile -string "Unable to add member. "

                    out-logfile -string $_

                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $member.trustee
                        ExternalDirectoryObjectID = $null
                        Alias = $functionMailNickName
                        Name = $originalContactConfiguration.name
                        Attribute = "Send As On Migrated contact"
                        ErrorMessage = ("Unable to add "+$member.trustee+" to migrated distribution contact with send as rights.  Manual addition required.")
                        ErrorMessageDetail = $_
                    }

                    out-logfile -string $isErrorObject

                    $functionErrors+=$isErrorObject
                }
            }
        }

        out-logfile -string "Remove the SMTP Address added by creating the temporary contact."

        try {
            out-logfile -string ("Removing: "+$functionEmailAddressToRemove)
            Set-O365Distributioncontact -identity $functionExternalDirectoryObjectID -emailAddresses @{remove=$functionEmailAddressToRemove} -errorAction STOP -BypassSecuritycontactManagerCheck
        }
        catch {
            out-logfile -string "Unable to remove SMTP address assigned by default during contact creation."
            out-logfile -string $_

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                Alias = $functionMailNickName
                Name = $originalContactConfiguration.name
                Attribute = "Unable to remove temporary SMTP address of contact."
                ErrorMessage = ("Unable to remove" +$functionEmailAddressToRemove+" - manaual removal required.")
                ErrorMessageDetail = $_
            }

            out-logfile -string $isErrorObject

            $functionErrors+=$isErrorObject
        }

        Out-LogFile -string "END SET-Office365contactMV"
        Out-LogFile -string "********************************************************************************"

        out-logfile -string ("The number of function Errors = "+$functionErrors.count)
        $global:postCreateErrors += $functionErrors
    }