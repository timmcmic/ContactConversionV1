<#
    .SYNOPSIS

    This function sets the single value attributes of the contact created in Office 365.

    .DESCRIPTION

    This function sets the single value attributes of the contact created in Office 365.

    .PARAMETER originalContactConfiguration

    The original configuration of the contact on premises.

    .PARAMETER contactTypeOverride

    Submits the contact type override of specified by the administrator at run time.

    .OUTPUTS

    None

    .EXAMPLE

    set-Office365contact -originalContactConfiguration contactConfiguration -contactTypeOverride TYPEOVERRIDE.

    #>
    Function set-Office365contact
     {
        [cmcontactetbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration,
            [Parameter(Mandatory = $true)]
            $office365contactConfiguration,
            [Parameter(Mandatory = $true)]
            [string]$contactTypeOverride,
            [Parameter(Mandatory = $true)]
            $office365contactConfigurationPostMigration
        )

        #Declare function variables.

        $functionModerationFlags=$NULL
        $functionSendModerationNotifications=$NULL
        $functionModerationEnabled=$NULL
        $functionoofReplyToOriginator=$NULL
        $functionreportToOwner=$NULL
        $functionHiddenFromAddressList=$NULL
        $functionMemberJoinRestriction=$NULL
        $functionMemberDepartRestriction=$NULL
        $functionRequireAuthToSendTo=$NULL

        [string]$functionMailNickName=""
        [string]$functionDisplayName=""
        [string]$functionSimpleDisplayName=""
        [string]$functionWindowsEmailAddress=""
        [boolean]$functionReportToOriginator=$NULL
        [string]$functionExternalDirectoryObjectID = $office365contactConfigurationPostMigration.externalDirectoryObjectID

        [boolean]$isTestError=$FALSE
        [array]$functionErrors=@()

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN SET-Office365contact"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("originalContactConfiguration = ")
        out-logfile -string $originalContactConfiguration
        out-logfile -string ("contact Type Override = "+$contactTypeOverride)

        #There are several flags of a contact that are either calculated hashes <or> booleans not set by default.
        #The exchange commancontactets abstract this by performing a conversion or filling the values in.
        #Since we use ldap to get these values now - we must reverse engineer these and / or set them.

        #If the contact type was overridden from the default - the member join restriction has to be adjusted.
        #If the contact tyoe was not overriden - check to see if depart is NULL and set to closed which is default.
        #Otherwise take the value from the string.

        if ( $contactTypeOverride -eq "Security" )
		{
            out-logfile -string "contact type overriden to Security by administrator.  This requires depart restriction closed."

			$functionMemberDepartRestriction = "Closed"

            out-logfile -string ("Function member depart restrictions = "+$functionMemberDepartRestriction)
		}
        elseif ($originalContactConfiguration.msExchcontactDepartRestriction -eq $NULL)
        {
            out-logFile -string ("Member depart restriction is NULL.")

            $functionMemberDepartRestriction="Closed"

            out-LogFile -string ("The member depart restriction is now = "+$functionMemberDepartRestriction)
        }
        elseif (($originalContactConfiguration.contactType -eq "-2147483640") -or ($originalContactConfiguration.contactType -eq "-2147483646") -or ($originalContactConfiguration.contactType -eq "-2147483644"))
        {
            Out-logfile -string ("contact type is security - ensuring member depart restriction CLOSED")

            $functionMemberDepartRestriction="Closed"
        }
		else 
		{
			$functionMemberDepartRestriction = $originalContactConfiguration.msExchcontactDepartRestriction

            out-logfile -string ("Function member depart restrictions = "+$functionMemberDepartRestriction)
		}

        #The moderation settings a are a hash valued flag.
        #This test looks to see if bypass nested moderation is enabled from the hash.

        if (($originalContactConfiguration.msExchModerationFlags -eq "1") -or ($originalContactConfiguration.msExchModerationFlags -eq "3") -or ($originalContactConfiguration.msExchModerationFlags -eq "7") )
        {
            out-logfile -string ("The moderation flags are 1 / 3 / 7 - setting bypass nested moderation to TRUE - "+$originalContactConfiguration.msExchModerationFlags)

            $functionModerationFlags=$TRUE

            out-logfile ("The function moderation flags are = "+$functionModerationFlags)
        }
        else 
        {
            out-logfile -string ("The moderation flags are NOT 1 / 3 / 7 - setting bypass nested moderation to FALSE - "+$originalContactConfiguration.msExchModerationFlags)

            $functionModerationFlags=$FALSE

            out-logfile ("The function moderation flags is = "+$functionModerationFlags)
        }

        #Test now to see if the moderation settings are always, internal, or none.  This uses the same hash.

        if (($originalContactConfiguration.msExchModerationFlags -eq "0") -or ($originalContactConfiguration.msExchModerationFlags -eq "1")  )
        {
            out-logfile -string ("The moderation flags are 0 / 2 / 6 - send notifications to never."+$originalContactConfiguration.msExchModerationFlags)

            $functionSendModerationNotifications="Never"

            out-logfile -string ("The function send moderations notifications is = "+$functionSendModerationNotifications)
        }
        elseif (($originalContactConfiguration.msExchModerationFlags -eq "2") -or ($originalContactConfiguration.msExchModerationFlags -eq "3")  )
        {
            out-logfile -string ("The moderation flags are 0 / 2 / 6 - setting send notifications to internal."+$originalContactConfiguration.msExchModerationFlags)

            $functionSendModerationNotifications="Internal"

            out-logfile -string ("The function send moderations notifications is = "+$functionSendModerationNotifications)

        }
        elseif (($originalContactConfiguration.msExchModerationFlags -eq "6") -or ($originalContactConfiguration.msExchModerationFlags -eq "7")  )
        {
            out-logfile -string ("The moderation flags are 0 / 2 / 6 - setting send notifications to always."+$originalContactConfiguration.msExchModerationFlags)

            $functionSendModerationNotifications="Always"

            out-logfile -string ("The function send moderations notifications is = "+$functionSendModerationNotifications)
        }
        else 
        {
            out-logFile -string ("The moderation flags are not set.  Setting to default of always.")
            
            $functionSendModerationNotifications="Always"

            out-logFile -string ("The function send moderation notification is = "+$functionSendModerationNotifications)
        }

        #Evaluate moderation enabled.

        if ($originalContactConfiguration.msExchEnableModeration -eq $NULL)
        {
            out-logfile -string "The moderation enabled setting is null."

            $functionModerationEnabled=$FALSE

            out-logfile -string ("The updated moderation enabled flag is = "+$functionModerationEnabled)
        }
        else 
        {
            out-logfile -string "The moderation setting was set on premises."
            
            $functionModerationEnabled=$originalContactConfiguration.msExchEnableModeration

            out-Logfile -string ("The function moderation setting is "+$functionModerationEnabled)
        }

        #Evaluate oofReplyToOriginator

        if ($originalContactConfiguration.oofReplyToOriginator -eq $NULL)
        {
            out-logfile -string "The oofReplyToOriginator is null."

            $functionoofReplyToOriginator = $FALSE

            out-logfile -string ("The oofReplyToOriginator is now = "+$functionoofReplyToOriginator)
        }
        else 
        {
            out-logFile -string "The oofReplyToOriginator was set on premises."
            
            $functionoofReplyToOriginator=$originalContactConfiguration.oofReplyToOriginator

            out-logfile -string ("The function oofReplyToOriginator = "+$functionoofReplyToOriginator)
        }

        #Evaluate reportToOwner

        if ($originalContactConfiguration.reportToOwner -eq $NULL)
        {
            out-logfile -string "The reportToOwner is null."

            $functionreportToOwner = $FALSE

            out-logfile -string ("The reportToOwner is now = "+$functionreportToOwner)
        }
        else 
        {
            out-logfile -string "The reportToOwner was set on premises." 
            
            $functionReportToOwner = $originalContactConfiguration.reportToOwner

            out-logfile -string ("The function reportToOwner = "+$functionreportToOwner)
        }

        if ($originalContactConfiguration.reportToOriginator -eq $NULL)
        {
            out-logfile -string "The report to originator is NULL."

            $functionReportToOriginator = $FALSE
        }
        else 
        {
            $functionReportToOriginator = $originalContactConfiguration.reportToOriginator    
        }

        #Evaluate hidden from address list.

        if ($originalContactConfiguration.msExchHideFromAddressLists -eq $NULL)
        {
            out-logfile -string ("Hidden from adddress list is null.")

            $functionHiddenFromAddressList=$FALSE

            out-logfile -string ("The hidden from address list is now = "+$functionHiddenFromAddressList)
        }
        else 
        {
            out-logFile -string ("Hidden from address list is not null.")
            
            $functionHiddenFromAddressList=$originalContactConfiguration.msExchHideFromAddressLists
        }

        #Evaluate member join restrictions.

        if ($originalContactConfiguration.msExchcontactJoinRestriction -eq $NULL)
        {
            out-Logfile -string ("Member join restriction is NULL.")

            $functionMemberJoinRestriction="Closed"

            out-logfile -string ("The member join restriction is now = "+$functionMemberJoinRestriction)
        }
        else 
        {
            $functionMemberJoinRestriction = $originalContactConfiguration.msExchcontactJoinRestriction

            out-logfile -string ("The function member join restriction is: "+$functionMemberJoinRestriction)
        }

        #Evaluate require auth to send to contact.  If the contact is open to everyone - the value may not be present.

        if ($originalContactConfiguration.msExchRequireAuthToSendTo -eq $NULL)
        {
            out-logfile -string ("Require auth to send to is not set.")

            $functionRequireAuthToSendTo = $FALSE

            out-logfile -string ("The new require auth to sent to is: "+$functionRequireAuthToSendTo)
        }
        else 
        {
            out-logfile -string ("Require auth to send to is set - retaining value. "+ $originalContactConfiguration.msExchRequireAuthToSendTo)
            
            $functionRequireAuthToSendTo = $originalContactConfiguration.msExchRequireAuthToSendTo
        }

        #It is possible that the contact is not fully mail enabled.
        #contacts may now be represented as mail enabled if only MAIL is populated.
        #If on premsies attributes are not specified - use the attributes that were obtained from office 365.

        if ($originalContactConfiguration.mailNickName -eq $NULL)
        {
            out-logfile -string "On premsies contact does not have alias / mail nick name -> using Office 365 value."

            $functionMailNickName = $office365contactConfiguration.alias

            out-logfile -string ("Office 365 alias used for contact creation: "+$functionMailNickName)
        }
        else 
        {
            out-logfile -string "On premises contact has a mail nickname specified - using on premsies value."
            $functionMailNickName = $originalContactConfiguration.mailNickName
            out-logfile -string $functionMailNickName    
        }

        if ($originalContactConfiguration.displayName -ne $NULL)
        {
            $functionDisplayName = $originalContactConfiguration.displayName
        }
        else 
        {
            $functionDisplayName = $office365contactConfiguration.displayName    
        }

        if ($originalContactConfiguration.simpleDisplayNamePrintable -ne $NULL)
        {
            $functionSimpleDisplayName = $originalContactConfiguration.simpleDisplayNamePrintable
        }
        else 
        {
            $functionSimpleDisplayName = $office365contactConfiguration.simpleDisplayName    
        }

        try 
        {
            out-logfile -string "Setting core single values for the distribution contact."

            Set-O365Distributioncontact -Identity $functionExternalDirectoryObjectID -name $originalContactConfiguration.cn -Alias $functionMailNickName -DisplayName $functionDisplayName -HiddenFromAddressListsEnabled $functionHiddenFromAddressList -RequireSenderAuthenticationEnabled $functionRequireAuthToSendTo -SimpleDisplayName $functionSimpleDisplayName -WindowsEmailAddress $originalContactConfiguration.mail -MailTipTranslations $originalContactConfiguration.msExchSenderHintTranslations -BypassSecuritycontactManagerCheck -errorAction STOP
        }
        catch 
        {
            out-logfile "Error encountered setting core single valued attributes."

            out-logfile -string $_

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                Alias = $functionMailNickName
                Name = $originalContactConfiguration.name
                Attribute = "Cloud distribution list:  Alias / DisplayName / HiddenFromAddressList / RequireSenderAuthenticaiton / SimpleDisplayName / WindowsEmailAddress / MailTipTranslations / Name"
                ErrorMessage = "Error setting single valued attribute of the migrated distribution list."
                ErrorMessageDetail = $_
            }

            $functionErrors+=$isErrorObject
        }

        try 
        {
            out-logfile -string "Setting single valued moderation propeties for the contact.."

            Set-O365Distributioncontact -Identity $functionExternalDirectoryObjectID -BypassNestedModerationEnabled $functionModerationFlags -ModerationEnabled $functionModerationEnabled -SendModerationNotifications $functionSendModerationNotifications -BypassSecuritycontactManagerCheck -errorAction STOP
        }
        catch 
        {
            out-logfile "Error encountered setting moderation properties of the contact...."

            out-logfile -string $_

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                Alias = $functionMailNickName
                Name = $originalContactConfiguration.name
                Attribute = "Cloud distribution list:  BypassNedstedModerationEnabled / ModerationEnabled / SendModerationNotifications"
                ErrorMessage = "Error setting additional single valued attribute of the migrated distribution contact."
                ErrorMessageDetail = $_
            }

            $functionErrors+=$isErrorObject
        }

        try 
        {
            out-logfile -string "Setting member join depart restritions on the contact.."

            Set-O365Distributioncontact -Identity $functionExternalDirectoryObjectID -MemberJoinRestriction $functionMemberJoinRestriction -MemberDepartRestriction $functionMemberDepartRestriction -BypassSecuritycontactManagerCheck -errorAction STOP
        }
        catch 
        {
            out-logfile "Error encountered setting member join depart restritions on the contact...."

            out-logfile -string $_

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                Alias = $functionMailNickName
                Name = $originalContactConfiguration.name
                Attribute = "Cloud distribution list:  MemberJoinRestriction / MemberDepartRestriction"
                ErrorMessage = "Error setting join or depart restrictions."
                ErrorMessageDetail = $_
            }

            $functionErrors+=$isErrorObject
        }

        try 
        {
            out-logfile -string "Setting the single valued report to settings.."

            Set-O365Distributioncontact -Identity $functionExternalDirectoryObjectID -ReportToManagerEnabled $functionreportToOwner -ReportToOriginatorEnabled $functionReportToOriginator -SendOofMessageToOriginatorEnabled $functionoofReplyToOriginator -BypassSecuritycontactManagerCheck -errorAction STOP       
        }
        catch 
        {
            out-logfile "Error encountered setting single valued report to settings on the contact...."

            out-logfile -string $_

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                Alias = $functionMailNickName
                Name = $originalContactConfiguration.name
                Attribute = "Cloud distribution list:  ReportToManagerEnabled / ReportToOriginatorEnabled / SendOOFMessageToOriginatorEnabled"
                ErrorMessage = "Error setting report to attributes."
                ErrorMessageDetail = $_
            }

            $functionErrors+=$isErrorObject
        }

        try 
        {
            out-logfile -string "Setting the custom and extension attributes of the contact."

            Set-O365Distributioncontact -Identity $functionExternalDirectoryObjectID -CustomAttribute1 $originalContactConfiguration.extensionAttribute1 -CustomAttribute10 $originalContactConfiguration.extensionAttribute10 -CustomAttribute11 $originalContactConfiguration.extensionAttribute11 -CustomAttribute12 $originalContactConfiguration.extensionAttribute12 -CustomAttribute13 $originalContactConfiguration.extensionAttribute13 -CustomAttribute14 $originalContactConfiguration.extensionAttribute14 -CustomAttribute15 $originalContactConfiguration.extensionAttribute15 -CustomAttribute2 $originalContactConfiguration.extensionAttribute2 -CustomAttribute3 $originalContactConfiguration.extensionAttribute3 -CustomAttribute4 $originalContactConfiguration.extensionAttribute4 -CustomAttribute5 $originalContactConfiguration.extensionAttribute5 -CustomAttribute6 $originalContactConfiguration.extensionAttribute6 -CustomAttribute7 $originalContactConfiguration.extensionAttribute7 -CustomAttribute8 $originalContactConfiguration.extensionAttribute8 -CustomAttribute9 $originalContactConfiguration.extensionAttribute9 -ExtensionCustomAttribute1 $originalContactConfiguration.msExtensionCustomAttribute1 -ExtensionCustomAttribute2 $originalContactConfiguration.msExtensionCustomAttribute2 -ExtensionCustomAttribute3 $originalContactConfiguration.msExtensionCustomAttribute3 -ExtensionCustomAttribute4 $originalContactConfiguration.msExtensionCustomAttribute4 -ExtensionCustomAttribute5 $originalContactConfiguration.msExtensionCustomAttribute5 -BypassSecuritycontactManagerCheck -errorAction STOP        
        }
        catch 
        {
            out-logfile "Error encountered setting custom and extension attributes of the contact...."

            out-logfile -string $_

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $originalContactConfiguration.mail
                ExternalDirectoryObjectID = $originalContactConfiguration.'msDS-ExternalDirectoryObjectId'
                Alias = $functionMailNickName
                Name = $originalContactConfiguration.name
                Attribute = "Cloud distribution list:  CustomAttributeX / ExtensionAttributeX"
                ErrorMessage = "Error setting custom or extension attributes."
                ErrorMessageDetail = $_
            }

            $functionErrors+=$isErrorObject
        }

        Out-LogFile -string "END SET-Office365contact"
        Out-LogFile -string "********************************************************************************"

        out-logfile -string ("The number of function errors is: "+$functionerrors.count )
        $global:postCreateErrors += $functionErrors
    }