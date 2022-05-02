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
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration,
            [Parameter(Mandatory = $true)]
            $office365contactConfiguration,
            [Parameter(Mandatory = $true)]
            [Parameter(Mandatory = $true)]
            $office365contactConfigurationPostMigration
        )

        #Declare function variables.

        $functionSendModerationNotifications=$NULL
        $functionModerationEnabled=$NULL
        $functionHiddenFromAddressList=$NULL
        $functionRequireAuthToSendTo=$NULL

        $functionMacFormat = ""
        $functionMessageFormat = ""
        $functionMessageBodyFormat = ""

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

        #There are several flags of a contact that are either calculated hashes <or> booleans not set by default.
        #The exchange commancontactets abstract this by performing a conversion or filling the values in.
        #Since we use ldap to get these values now - we must reverse engineer these and / or set them.

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

        out-logfile -string "Determine mail contact internet coding."

        if ($originalDLConfiguration.internetEncoding -eq "0")
        {
            $functionMacFormat = "BinHex"
            $functionMessageFormat = "Text"
            $functionMessageBodyFormat = "Text"

            out-logfile -string "Internet encoding value 0"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "2097152")
        {
            $functionMacFormat = "uUEncode"
            $functionMessageFormat = "Text"
            $functionMessageBodyFormat = "Text"

            out-logfile -string "Internet encoding value 2097152"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "262144")
        {
            $functionMacFormat = "BinHex"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "Text"

            out-logfile -string "Internet encoding value 262144"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "4456448")
        {
            $functionMacFormat = "AppleSingle"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "Text"

            out-logfile -string "Internet encoding value 4456448"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "6553600")
        {
            $functionMacFormat = "AppleDouble"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "Text"

            out-logfile -string "Internet encoding value 6553600"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "786432")
        {
            $functionMacFormat = "BinHex"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "HTML"

            out-logfile -string "Internet encoding value 786432"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "7077888")
        {
            $functionMacFormat = "AppleDouble"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "HTML"

            out-logfile -string "Internet encoding value 7077888"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "4980736")
        {
            $functionMacFormat = "AppleSingle"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "HTML"

            out-logfile -string "Internet encoding value 4980736"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "5505024")
        {
            $functionMacFormat = "AppleSingle"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "TextandHTML"

            out-logfile -string "Internet encoding value 5505024"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "7602176")
        {
            $functionMacFormat = "AppleDouble"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "HTML"

            out-logfile -string "Internet encoding value 7602176"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        elseif ($originalDLConfiguration.internetEncoding -eq "1310720")
        {
            $functionMacFormat = "BinHex"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "TextandHTML"

            out-logfile -string "Defaults intentionally set as value is present."
            out-logfile -string "Internet encoding value 1310720"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
        }
        else 
        {
            $functionMacFormat = "BinHex"
            $functionMessageFormat = "Mime"
            $functionMessageBodyFormat = "TextandHTML"

            out-logfile -string "Using default values as no explicit value was established on premises."
            out-logfile -string "Internet encoding value 1310720"
            out-logfile -string $functionMacFormat
            out-logfile -string $functionMessageFormat
            out-logfile -string $functionMessageBodyFormat
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
            out-logfile -string "Setting core single values for the contact contact."

            set-o365MailContact -Identity $functionExternalDirectoryObjectID -name $originalContactConfiguration.cn -Alias $functionMailNickName -DisplayName $functionDisplayName -HiddenFromAddressListsEnabled $functionHiddenFromAddressList -RequireSenderAuthenticationEnabled $functionRequireAuthToSendTo -SimpleDisplayName $functionSimpleDisplayName -WindowsEmailAddress $originalContactConfiguration.mail -MailTipTranslations $originalContactConfiguration.msExchSenderHintTranslations -errorAction STOP
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
                Attribute = "Cloud contact list:  Alias / DisplayName / HiddenFromAddressList / RequireSenderAuthenticaiton / SimpleDisplayName / WindowsEmailAddress / MailTipTranslations / Name"
                ErrorMessage = "Error setting single valued attribute of the migrated contact list."
                ErrorMessageDetail = $_
            }

            $functionErrors+=$isErrorObject
        }

        try 
        {
            out-logfile -string "Setting single valued moderation propeties for the contact.."

            set-o365MailContact -Identity $functionExternalDirectoryObjectID -ModerationEnabled $functionModerationEnabled -SendModerationNotifications $functionSendModerationNotifications  -errorAction STOP
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
                Attribute = "Cloud contact list:  BypassNedstedModerationEnabled / ModerationEnabled / SendModerationNotifications"
                ErrorMessage = "Error setting additional single valued attribute of the migrated contact contact."
                ErrorMessageDetail = $_
            }

            $functionErrors+=$isErrorObject
        }

        try 
        {
            out-logfile -string "Setting the custom and extension attributes of the contact."

            set-o365MailContact -Identity $functionExternalDirectoryObjectID -CustomAttribute1 $originalContactConfiguration.extensionAttribute1 -CustomAttribute10 $originalContactConfiguration.extensionAttribute10 -CustomAttribute11 $originalContactConfiguration.extensionAttribute11 -CustomAttribute12 $originalContactConfiguration.extensionAttribute12 -CustomAttribute13 $originalContactConfiguration.extensionAttribute13 -CustomAttribute14 $originalContactConfiguration.extensionAttribute14 -CustomAttribute15 $originalContactConfiguration.extensionAttribute15 -CustomAttribute2 $originalContactConfiguration.extensionAttribute2 -CustomAttribute3 $originalContactConfiguration.extensionAttribute3 -CustomAttribute4 $originalContactConfiguration.extensionAttribute4 -CustomAttribute5 $originalContactConfiguration.extensionAttribute5 -CustomAttribute6 $originalContactConfiguration.extensionAttribute6 -CustomAttribute7 $originalContactConfiguration.extensionAttribute7 -CustomAttribute8 $originalContactConfiguration.extensionAttribute8 -CustomAttribute9 $originalContactConfiguration.extensionAttribute9 -ExtensionCustomAttribute1 $originalContactConfiguration.msExtensionCustomAttribute1 -ExtensionCustomAttribute2 $originalContactConfiguration.msExtensionCustomAttribute2 -ExtensionCustomAttribute3 $originalContactConfiguration.msExtensionCustomAttribute3 -ExtensionCustomAttribute4 $originalContactConfiguration.msExtensionCustomAttribute4 -ExtensionCustomAttribute5 $originalContactConfiguration.msExtensionCustomAttribute5  -errorAction STOP        
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
                Attribute = "Cloud contact list:  CustomAttributeX / ExtensionAttributeX"
                ErrorMessage = "Error setting custom or extension attributes."
                ErrorMessageDetail = $_
            }

            $functionErrors+=$isErrorObject
        }

        try 
        {
            out-logfile -string "Setting internet encoding information."

            set-o365MailContact -Identity $functionExternalDirectoryObjectID -macFormat $functionMacFormat -messageFormat $functionMessageFormat -messageBodyFormat $functionMessageBodyFormat -errorAction STOP        
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
                Attribute = "Cloud contact list:  Internet encoding information / macFormat / messageFormat / MessageBodyFormat"
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