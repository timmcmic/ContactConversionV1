
#############################################################################################
# DISCLAIMER:																				#
#																							#
# THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT					#
# PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY				#
# OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT		#
# LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR		#
# PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS		#
# AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR			#
# ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE	#
# FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS	#
# PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)	#
# ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,		#
# EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES						#
#############################################################################################


Function Start-ContactMigration
{
    <#
    .SYNOPSIS

    Thiis is the trigger function that allows administrators to "migrate" the source of authority of an on premsies mail enabled contact to Office 365.

    .DESCRIPTION

    Trigger function.

    .PARAMETER contactSMTPAddress

    *REQUIRED*
    The SMTP address of the mail contact to be migrated.

    .PARAMETER globalCatalogServer

    *REQUIRED*
    A global catalog server in the domain where the Contact to be migrated resides.


    .PARAMETER activeDirectoryCredential

    *REQUIRED*
    This is the credential that will be utilized to perform operations against the global catalog server.
    If the Contact and all it's dependencies reside in a single domain - a domain administrator is acceptable.
    If the Contact and it's dependencies span multiple domains in a forest - enterprise administrator is required.
      
    .PARAMETER logFolder

    *REQUIRED*
    The location where logging for the migration should occur including all XML outputs for backups.

    .PARAMETER aadConnectServer

    *OPTIONAL*
    This is the AADConnect server that automated sycn attempts will be attempted.
    If specified with an AADConnect credential - delta syncs will be triggered automatically in attempts to service the move.
    This requires WINRM be enabled on the ADConnect server and may have additional WINRM dependencies / configuration.
    Name should be specified in fully qualified domain format.

    .PARAMETER aadConnectCredential

    *OPTIONAL*
    The credential specified to perform remote powershell / winrm sessions to the AADConnect server.


    .PARAMETER exchangeOnlineCredential

    *REQUIRED IF NO OTHER CREDENTIALS SPECIFIED*
    This is the credential utilized for Exchange Online connections.  
    The credential must be specified if certificate based authentication is not configured.
    The account requires global administration rights / exchange organization management rights.
    An exchange online credential cannot be combined with an exchangeOnlineCertificateThumbprint.

    .PARAMETER exchangeOnlineCertificateThumbprint

    *REQUIRED IF NO OTHER CREDENTIALS SPECIFIED*
    This is the certificate thumbprint that will be utilzied for certificate authentication to Exchange Online.
    This requires all the pre-requists be established and configured prior to access.
    A certificate thumbprint cannot be specified with exchange online credentials.

    .PARAMETER exchangeAuthenticationMethod

    *OPTIONAL*
    This allows the administrator to specify either Kerberos or Basic authentication for on premises Exchange Powershell.
    Basic is the assumed default and requires basic authentication be enabled on the powershell virtual directory of the specified exchange server.

    .PARAMETER retainOffice365Settings

    *OPTIONAL*
    It is possible over the course of migrations that cloud only resources could have dependencies on objects that still remain on premises.
    The administrator can choose to scan office 365 to capture any cloud only dependencies that may exist.
    The default is true.

    .PARAMETER doNoSyncOU

    *REQUIRED IF RETAIN Contact FALSE*
    This is the administrator specified organizational unit that is NOT configured to sync in AD Connect.
    When the administrator specifies to NOT retain the Contact the Contact is moved to this OU to allow for deletion from Office 365.
    A doNOSyncOU must be specified if the administrator specifies to NOT retain the Contact.

    .PARAMETER retainOriginalContact

    *OPTIONAL*
    Allows the administrator to retain the Contact - for example if the Contact also has on premises security dependencies.
    This triggers a mail disable of the Contact resulting in Contact deletion from Office 365.
    The name of the Contact is randomized with a character ! to ensure no conflict with hybird mail flow - if hybrid mail flow enabled.

	.OUTPUTS

    Logs all activities and backs up all original data to the log folder directory.
    Moves the Contact from on premieses source of authority to office 365 source of authority.

    .EXAMPLE

    Start-ContactMigration

    #>
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$contactSMTPAddress,
        [Parameter(Mandatory = $true)]
        [string]$globalCatalogServer,
        [Parameter(Mandatory = $true)]
        [pscredential]$activeDirectoryCredential,
        [Parameter(Mandatory = $true)]
        [string]$logFolderPath,
        [Parameter(Mandatory = $false)]
        [string]$aadConnectServer=$NULL,
        [Parameter(Mandatory = $false)]
        [pscredential]$aadConnectCredential=$NULL,
        [Parameter(Mandatory = $false)]
        [pscredential]$exchangeOnlineCredential=$NULL,
        [Parameter(Mandatory = $false)]
        [string]$exchangeOnlineCertificateThumbPrint="",
        [Parameter(Mandatory = $false)]
        [string]$exchangeOnlineOrganizationName="",
        [Parameter(Mandatory = $false)]
        [ValidateSet("O365Default","O365GermanyCloud","O365China","O365USGovGCCHigh","O365USGovDoD")]
        [string]$exchangeOnlineEnvironmentName="O365Default",
        [Parameter(Mandatory = $false)]
        [string]$exchangeOnlineAppID="",
        [Parameter(Mandatory = $false)]
        [ValidateSet("Basic","Kerberos")]
        [string]$exchangeAuthenticationMethod="Basic",
        [Parameter(Mandatory = $true)]
        [string]$dnNoSyncOU = "NotSet",
        [Parameter(Mandatory = $false)]
        [boolean]$retainOriginalContact = $TRUE,
        [Parameter(Mandatory = $false)]
        [int]$threadNumberAssigned=0,
        [Parameter(Mandatory = $false)]
        [int]$totalThreadCount=0,
        [Parameter(Mandatory = $FALSE)]
        [boolean]$isMultiMachine=$FALSE,
        [Parameter(Mandatory = $FALSE)]
        [string]$remoteDriveLetter=$NULL,
    )

    $windowTitle = ("Start-ContactMigration"+$contactSMTPAddress)
    $host.ui.RawUI.WindowTitle = $windowTitle

    if ($isMultiMachine -eq $TRUE)
    {
        try{
            #At this point we know that multiple machines was in use.
            #For multiple machines - the local controller instance mapped the drive Z for us in windows.
            #Therefore we override the original log folder path passed in and just use Z.

            [string]$networkName=$remoteDriveLetter
            $logFolderPath = $networkName+":"
        }
        catch{
            exit
        }
    }

    #Define global variables.

    $global:threadNumber=$threadNumberAssigned
    $global:logFile=$NULL #This is the global variable for the calculated log file name
    [string]$global:staticFolderName="\DLMigration\"

    #Define variables utilized in the core function that are not defined by parameters.

    [boolean]$useAADConnect=$FALSE #Determines if function will utilize aadConnect during migration.
    [string]$aadConnectPowershellSessionName="AADConnect" #Defines universal name for aadConnect powershell session.
    [string]$ADGlobalCatalogPowershellSessionName="ADGlobalCatalog" #Defines universal name for ADGlobalCatalog powershell session.
    [string]$exchangeOnlinePowershellModuleName="ExchangeOnlineManagement" #Defines the exchage management shell name to test for.
    [string]$activeDirectoryPowershellModuleName="ActiveDirectory" #Defines the active directory shell name to test for.
    [string]$contactConversionPowershellModuleName="ContactConversionV1"
    [string]$globalCatalogPort=":3268"
    [string]$globalCatalogWithPort=$globalCatalogServer+$globalCatalogPort

    #The variables below are utilized to define working parameter sets.
    #Some variables are assigned to single values - since these will be utilized with functions that query or set information.
    
    [string]$acceptMessagesFromDLMembers="dlMemSubmitPerms" #Attribute for the allow email members.
    [string]$rejectMessagesFromDLMembers="dlMemRejectPerms"
    [string]$bypassModerationFromDL="msExchBypassModerationFromDLMembersLink"
    [string]$forwardingAddressForDL="altRecipient"
    [string]$grantSendOnBehalfToDL="publicDelegates"
    [array]$ContactPropertySet = '*'
    [array]$ContactPropertySetToClear = @()
    [array]$contactPropertiesToClearModern='authOrig','DisplayName','DisplayNamePrintable',$rejectMessagesFromDLMembers,$acceptMessagesFromDLMembers,'extensionAttribute1','extensionAttribute10','extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5','extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','legacyExchangeDN','mail','mailNickName','msExchRecipientDisplayType','msExchRecipientTypeDetails','msExchRemoteRecipientType',$bypassModerationFromDL,'msExchBypassModerationLink','msExchCoManagedByLink','msExchEnableModeration','msExchExtensionCustomAttribute1','msExchExtensionCustomAttribute2','msExchExtensionCustomAttribute3','msExchExtensionCustomAttribute4','msExchExtensionCustomAttribute5','msExchContactDepartRestriction','msExchContactJoinRestriction','msExchHideFromAddressLists','msExchModeratedByLink','msExchModerationFlags','msExchRequireAuthToSendTo','msExchSenderHintTranslations','oofReplyToOriginator','proxyAddresses',$grantSendOnBehalfToDL,'reportToOriginator','reportToOwner','unAuthOrig','msExchArbitrationMailbox','msExchPoliciesIncluded','msExchUMDtmfMap','msExchVersion','showInAddressBook','msExchAddressBookFlags','msExchBypassAudit','msExchContactExternalMemberCount','msExchContactMemberCount','msExchContactSecurityFlags','msExchLocalizationFlags','msExchMailboxAuditEnable','msExchMailboxAuditLogAgeLimit','msExchMailboxFolderSet','msExchMDBRulesQuota','msExchPoliciesIncluded','msExchProvisioningFlags','msExchRecipientSoftDeletedStatus','msExchRoleContactType','msExchTransportRecipientSettingsFlags','msExchUMDtmfMap','msExchUserAccountControl','msExchVersion'
    [array]$contactPropertiesToClearLegacy='authOrig','DisplayName','DisplayNamePrintable',$rejectMessagesFromDLMembers,$acceptMessagesFromDLMembers,'extensionAttribute1','extensionAttribute10','extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5','extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','legacyExchangeDN','mail','mailNickName','msExchRecipientDisplayType','msExchRecipientTypeDetails','msExchRemoteRecipientType',$bypassModerationFromDL,'msExchBypassModerationLink','msExchCoManagedByLink','msExchEnableModeration','msExchExtensionCustomAttribute1','msExchExtensionCustomAttribute2','msExchExtensionCustomAttribute3','msExchExtensionCustomAttribute4','msExchExtensionCustomAttribute5','msExchContactDepartRestriction','msExchContactJoinRestriction','msExchHideFromAddressLists','msExchModeratedByLink','msExchModerationFlags','msExchRequireAuthToSendTo','msExchSenderHintTranslations','oofReplyToOriginator','proxyAddresses',$grantSendOnBehalfToDL,'reportToOriginator','reportToOwner','unAuthOrig','msExchArbitrationMailbox','msExchPoliciesIncluded','msExchUMDtmfMap','msExchVersion','showInAddressBook','msExchAddressBookFlags','msExchBypassAudit','msExchContactExternalMemberCount','msExchContactMemberCount','msExchLocalizationFlags','msExchMailboxAuditEnable','msExchMailboxAuditLogAgeLimit','msExchMailboxFolderSet','msExchMDBRulesQuota','msExchPoliciesIncluded','msExchProvisioningFlags','msExchRecipientSoftDeletedStatus','msExchRoleContactType','msExchTransportRecipientSettingsFlags','msExchUMDtmfMap','msExchUserAccountControl','msExchVersion'

    #On premises variables for the list to be migrated.

    $originalContactConfiguration=$NULL #This holds the on premises DL configuration for the Contact to be migrated.
    $originalContactConfigurationUpdated=$NULL #This holds the on premises DL configuration post the rename operations.
    $routingContactConfig=$NULL
    $routingDynamicContactConfig=$NULL
    [array]$exchangeDLMembershipSMTP=@() #Array of DL membership from AD.
    [array]$exchangeRejectMessagesSMTP=@() #Array of members with reject permissions from AD.
    [array]$exchangeAcceptMessagesSMTP=@() #Array of members with accept permissions from AD.
    [array]$exchangeManagedBySMTP=@() #Array of members with manage by rights from AD.
    [array]$exchangeModeratedBySMTP=@() #Array of members  with moderation rights.
    [array]$exchangeBypassModerationSMTP=@() #Array of objects with bypass moderation rights from AD.
    [array]$exchangeGrantSendOnBehalfToSMTP=@()
    [array]$exchangeSendAsSMTP=@()

    #Define XML files to contain backups.

    [string]$originalContactConfigurationADXML = "originalContactConfigurationADXML" #Export XML file of the Contact attibutes direct from AD.
    [string]$originalContactConfigurationUpdatedXML = "originalContactConfigurationUpdatedXML"
    [string]$originalContactConfigurationObjectXML = "originalContactConfigurationObjectXML" #Export of the ad attributes after selecting objects (allows for NULL objects to be presented as NULL)
    [string]$office365ContactConfigurationXML = "office365ContactConfigurationXML"
    [string]$office365ContactConfigurationPostMigrationXML = "office365ContactConfigurationPostMigrationXML"
    [string]$office365DLMembershipPostMigrationXML = "office365DLMembershipPostMigrationXML"
    [string]$exchangeDLMembershipSMTPXML = "exchangeDLMemberShipSMTPXML"
    [string]$exchangeRejectMessagesSMTPXML = "exchangeRejectMessagesSMTPXML"
    [string]$exchangeAcceptMessagesSMTPXML = "exchangeAcceptMessagesSMTPXML"
    [string]$exchangeManagedBySMTPXML = "exchangeManagedBySMTPXML"
    [string]$exchangeModeratedBySMTPXML = "exchangeModeratedBYSMTPXML"
    [string]$exchangeBypassModerationSMTPXML = "exchangeBypassModerationSMTPXML"
    [string]$exchangeGrantSendOnBehalfToSMTPXML = "exchangeGrantSendOnBehalfToXML"
    [string]$exchangeSendAsSMTPXML = "exchangeSendASSMTPXML"
    [string]$allContactsMemberOfXML = "allContactsMemberOfXML"
    [string]$allContactsRejectXML = "allContactsRejectXML"
    [string]$allContactsAcceptXML = "allContactsAcceptXML"
    [string]$allContactsBypassModerationXML = "allContactsBypassModerationXML"
    [string]$allUsersForwardingAddressXML = "allUsersForwardingAddressXML"
    [string]$allContactsGrantSendOnBehalfToXML = "allContactsGrantSendOnBehalfToXML"
    [string]$allContactsManagedByXML = "allContactsManagedByXML"
    [string]$allContactsSendAsXML = "allContactSendAsXML"
    [string]$allContactsSendAsNormalizedXML="allContactsSendAsNormalizedXML"
    [string]$allOffice365MemberOfXML="allOffice365MemberOfXML"
    [string]$allOffice365AcceptXML="allOffice365AcceptXML"
    [string]$allOffice365RejectXML="allOffice365RejectXML"
    [string]$allOffice365BypassModerationXML="allOffice365BypassModerationXML"
    [string]$allOffice365GrantSendOnBehalfToXML="allOffice365GrantSentOnBehalfToXML"
    [string]$allOffice365ManagedByXML="allOffice365ManagedByXML"
    [string]$allOffice365ForwardingAddressXML="allOffice365ForwardingAddressXML"
    [string]$allOffic365SendAsAccessXML = "allOffice365SendAsAccessXML"
    [string]$allOffice365SendAsAccessOnContactXML = 'allOffice365SendAsAccessOnContactXML'
    [string]$allContactsCoManagedByXML="allContactsCoManagedByXML"

    #Define the retention files.

    [string]$retainOnPremRecipientSendAsXML="onPremRecipientSendAs.xml"

    #The following variables hold information regarding other Contacts in the environment that have dependnecies on the Contact to be migrated.

    [array]$allContactsMemberOf=$NULL #Complete AD information for all Contacts the migrated Contact is a member of.
    [array]$allContactsReject=$NULL #Complete AD inforomation for all Contacts that the migrated Contact has reject mesages from.
    [array]$allContactsAccept=$NULL #Complete AD information for all Contacts that the migrated Contact has accept messages from.
    [array]$allContactsBypassModeration=$NULL #Complete AD information for all Contacts that the migrated Contact has bypass moderations.
    [array]$allUsersForwardingAddress=$NULL #All users on premsies that have this Contact as a forwarding DN.
    [array]$allContactsGrantSendOnBehalfTo=$NULL #All dependencies on premsies that have grant send on behalf to.
    [array]$allContactsManagedBy=$NULL
    [array]$allObjectsFullMailboxAccess=$NULL
    [array]$allObjectSendAsAccess=$NULL
    [array]$allObjectsSendAsAccessNormalized=@()
    [array]$allContactsCoManagedByBL=$NULL

    #The following variables hold information regarding Office 365 objects that have dependencies on the migrated DL.

    #The following are for standard Contacts.

    [array]$allOffice365MemberOf=$NULL
    [array]$allOffice365Accept=$NULL
    [array]$allOffice365Reject=$NULL
    [array]$allOffice365BypassModeration=$NULL
    [array]$allOffice365ManagedBy=$NULL
    [array]$allOffice365GrantSendOnBehalfTo=$NULL

    #These are for other mail enabled objects.

    [array]$allOffice365ForwardingAddress=$NULL
    [array]$allOffice365SendAsAccess=$NULL
    [array]$allOffice365SendAsAccessOnContact = $NULL 

    #The following are the cloud parameters we query for to look for dependencies.

    [string]$office365AcceptMessagesFrom="AcceptMessagesOnlyFromDLMembers"
    [string]$office365BypassModerationFrom="BypassModerationFromDLMembers"
    [string]$office365CoManagers="CoManagedBy"
    [string]$office365GrantSendOnBehalfTo="GrantSendOnBehalfTo"
    [string]$office365ManagedBy="ManagedBy"
    [string]$office365Members="Members"
    [string]$office365RejectMessagesFrom="RejectMessagesFromDLMembers"
    [string]$office365ForwardingAddress="ForwardingAddress"

    [string]$office365BypassModerationusers="BypassModerationFromSendersOrMembers"

    [string]$office365UnifiedAccept="AcceptMessagesOnlyFromSendersOrMembers"
    [string]$office365UnifiedReject="RejectMessagesFromSendersOrMembers"


    #The following are the on premises parameters utilized for restoring depdencies.

    [string]$onPremUnAuthOrig="unauthorig"
    [string]$onPremAuthOrig="authOrig"
    [string]$onPremManagedBy="managedBy"
    [string]$onPremMSExchCoManagedByLink="msExchCoManagedByLink"
    [string]$onPremPublicDelegate="publicDelegates"
    [string]$onPremMsExchModeratedByLink="msExchModeratedByLink"
    [string]$onPremmsExchBypassModerationLink="msExchBypassModerationLink"
    [string]$onPremMemberOf="member"
    [string]$onPremAltRecipient="altRecipient"

    #Cloud variables for the list to be migrated.

    $office365ContactConfiguration = $NULL #This holds the office 365 contact configuration for the Contact to be migrated.
    $office365ContactConfigurationPostMigration = $NULL

    #Declare some variables for string processing as items move around.

    [string]$tempOU=$NULL
    [array]$tempNameArrayArray=@()
    [string]$tempName=$NULL
    [string]$tempDN=$NULL

    #For loop counter.

    [int]$forLoopCounter=0

    #Exchange Schema Version

    [int]$exchangeRangeUpper=$NULL
    [int]$exchangeLegacySchemaVersion=15317 #Exchange 2016 Preview Schema - anything less is legacy.

    #Define new arrays to check for errors instead of failing.

    [array]$preCreateErrors=@()
    [array]$global:postCreateErrors=@()
    [array]$onPremReplaceErrors=@()
    [array]$office365ReplaceErrors=@()
    [array]$global:office365ReplacePermissionsErrors=@()
    [array]$global:onPremReplacePermissionsErrors=@()
    [array]$generalErrors=@()
    [string]$isTestError="No"


    [int]$forLoopTrigger=1000

    #Define the sub folders for multi-threading.

    [array]$threadFolder="\Thread0","\Thread1","\Thread2","\Thread3","\Thread4","\Thread5","\Thread6","\Thread7","\Thread8","\Thread9","\Thread10"

    #Define the status directory.

    [string]$global:statusPath="\Status\"
    [string]$global:fullStatusPath=$NULL
    [int]$statusFileCount=0

    #To support the new feature for multiple onmicrosoft.com domains -> use this variable to hold the cross premsies routing domain.
    #This value can no longer be calculated off the address@domain.onmicrosoft.com value.

    [string]$mailOnMicrosoftComDomain = ""


    #If multi threaded - the log directory needs to be created for each thread.
    #Create the log folder path for status before changing the log folder path.

    if ($totalThreadCount -gt 0)
    {
        new-statusFile -logFolderPath $logFolderPath

        $logFolderPath=$logFolderPath+$threadFolder[$global:threadNumber]
    }

    #Ensure that no status files exist at the start of the run.

    if ($totalThreadCount -gt 0)
    {
        if ($global:threadNumber -eq 1)
        {
            remove-statusFiles -fullCleanup:$TRUE
        }
    }

    #Log start of DL migration to the log file.

    new-LogFile -contactSMTPAddress $contactSMTPAddress.trim() -logFolderPath $logFolderPath

    Out-LogFile -string "================================================================================"
    Out-LogFile -string "BEGIN START-CONTACTMIGRATION"
    Out-LogFile -string "================================================================================"

    out-logfile -string "Set error action preference to continue to allow write-error in out-logfile to service exception retrys"

    if ($errorActionPreference -ne "Continue")
    {
        out-logfile -string ("Current Error Action Preference: "+$errorActionPreference)
        $errorActionPreference = "Continue"
        out-logfile -string ("New Error Action Preference: "+$errorActionPreference)
    }
    else
    {
        out-logfile -string ("Current Error Action Preference is CONTINUE: "+$errorActionPreference)
    }

    out-logfile -string "Ensure that all strings specified have no leading or trailing spaces."

    #Perform cleanup of any strings so that no spaces existin trailing or leading.

    $contactSMTPAddress = remove-stringSpace -stringToFix $contactSMTPAddress
    $globalCatalogServer = remove-stringSpace -stringToFix $globalCatalogServer
    $logFolderPath = remove-stringSpace -stringToFix $logFolderPath 

    if ($aadConnectServer -ne $NULL)
    {
        $aadConnectServer = remove-stringSpace -stringToFix $aadConnectServer
    }

    if ($exchangeOnlineCertificateThumbPrint -ne "")
    {
        $exchangeOnlineCertificateThumbPrint=remove-stringSpace -stringToFix $exchangeOnlineCertificateThumbPrint
    }

    if ($exchangeOnlineEnvironmentName -ne "")
    {
        $exchangeOnlineEnvironmentName=remove-stringSpace -stringToFix $exchangeOnlineEnvironmentName
    }

    if ($exchangeOnlineOrganizationName -ne "")
    {
        $exchangeOnlineOrganizationName=remove-stringSpace -stringToFix $exchangeOnlineOrganizationName
    }

    if ($exchangeOnlineAppID -ne "")
    {
        $exchangeOnlineAppID=remove-stringSpace -stringToFix $exchangeOnlineAppID
    }
    
    $dnNoSyncOU = remove-StringSpace -stringToFix $dnNoSyncOU
    
    #Output parameters to the log file for recording.
    #For parameters that are optional if statements determine if they are populated for recording.

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "PARAMETERS"
    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string ("ContactSMTPAddress = "+$contactSMTPAddress)
    out-logfile -string ("Contact SMTP Address Length = "+$contactSMTPAddress.length.tostring())
    out-logfile -string ("Spaces Removed Contact SMTP Address: "+$contactSMTPAddress)
    out-logfile -string ("Contact SMTP Address Length = "+$contactSMTPAddress.length.toString())
    Out-LogFile -string ("GlobalCatalogServer = "+$globalCatalogServer)
    Out-LogFile -string ("ActiveDirectoryUserName = "+$activeDirectoryCredential.UserName.tostring())
    Out-LogFile -string ("LogFolderPath = "+$logFolderPath)

    if ($aadConnectServer -ne "")
    {
        $aadConnectServer = $aadConnectServer -replace '\s',''
        Out-LogFile -string ("AADConnectServer = "+$aadConnectServer)
    }

    if ($aadConnectCredential -ne $null)
    {
        Out-LogFile -string ("AADConnectUserName = "+$aadConnectCredential.UserName.tostring())
    }

    if ($exchangeOnlineCredential -ne $null)
    {
        Out-LogFile -string ("ExchangeOnlineUserName = "+ $exchangeOnlineCredential.UserName.toString())
    }

    if ($exchangeOnlineCertificateThumbPrint -ne "")
    {
        Out-LogFile -string ("ExchangeOnlineCertificateThumbprint = "+$exchangeOnlineCertificateThumbPrint)
    }

    out-logfile -string ("Retain Office 365 Settings = "+$retainOffice365Settings)
    out-logfile -string ("OU that does not sync to Office 365 = "+$dnNoSyncOU)
    out-logfile -string ("Will the original Contact be retained as part of migration = "+$retainOriginalContact)
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string " RECORD VARIABLES"
    Out-LogFile -string "********************************************************************************"

    out-logfile -string ("Global Catalog Port = "+$globalCatalogPort)
    out-logfile -string ("Global catalog string used for function queries ="+$globalCatalogWithPort)
    Out-LogFile -string ("Initial user of ADConnect = "+$useAADConnect)
    Out-LogFile -string ("AADConnect powershell session name = "+$aadConnectPowershellSessionName)
    Out-LogFile -string ("AD Global catalog powershell session name = "+$ADGlobalCatalogPowershellSessionName)
    Out-LogFile -string ("Exchange powershell module name = "+$exchangeOnlinePowershellModuleName)
    Out-LogFile -string ("Active directory powershell modulename = "+$activeDirectoryPowershellModuleName)
    out-logFile -string ("Static property for accept messages from members = "+$acceptMessagesFromDLMembers)
    out-logFile -string ("Static property for accept messages from members = "+$rejectMessagesFromDLMembers)
    Out-LogFile -string ("DL Properties to collect = ")

    foreach ($dlProperty in $ContactPropertySet)
    {
        Out-LogFile -string $dlProperty
    }

    Out-LogFile -string ("DL property set to be cleared legacy = ")

    foreach ($dlProperty in $contactPropertiesToClearLegacy)
    {
        Out-LogFile -string $dlProperty
    }

    Out-LogFile -string ("DL property set to be cleared modern = ")

    foreach ($dlProperty in $contactPropertiesToClearModern)
    {
        Out-LogFile -string $dlProperty
    }

    Out-LogFile -string ("Exchange on prem DL active directory configuration XML = "+$originalContactConfigurationADXML)
    Out-LogFile -string ("Exchange on prem DL object configuration XML = "+$originalContactConfigurationObjectXML)
    Out-LogFile -string ("office 365 contact configuration XML = "+$office365ContactConfigurationXML)
    Out-LogFile -string ("Exchange DL members XML Name - "+$exchangeDLMembershipSMTPXML)
    Out-LogFile -string ("Exchange Reject members XML Name - "+$exchangeRejectMessagesSMTPXML)
    Out-LogFile -string ("Exchange Accept members XML Name - "+$exchangeAcceptMessagesSMTPXML)
    Out-LogFile -string ("Exchange ManagedBY members XML Name - "+$exchangeManagedBySMTPXML)
    Out-LogFile -string ("Exchange ModeratedBY members XML Name - "+$exchangeModeratedBySMTPXML)
    Out-LogFile -string ("Exchange BypassModeration members XML Name - "+$exchangeBypassModerationSMTPXML)
    out-logfile -string ("Exchange GrantSendOnBehalfTo members XML name - "+$exchangeGrantSendOnBehalfToSMTPXML)
    Out-LogFile -string ("All Contact members XML Name - "+$allContactsMemberOfXML)
    Out-LogFile -string ("All Reject members XML Name - "+$allContactsRejectXML)
    Out-LogFile -string ("All Accept members XML Name - "+$allContactsAcceptXML)
    Out-Logfile -string ("All Co Managed By BL XML - "+$allContactsCoManagedByXML)
    Out-LogFile -string ("All BypassModeration members XML Name - "+$allContactsBypassModerationXML)
    out-logfile -string ("All Users Forwarding Address members XML Name - "+$allUsersForwardingAddressXML)
    out-logfile -string ("All Contacts Grand Send On Behalf To XML Name - "+$allContactsGrantSendOnBehalfToXML)
    out-logfile -string ("Property in office 365 for accept members = "+$office365AcceptMessagesFrom)
    out-logfile -string ("Property in office 365 for bypassmoderation members = "+$office365BypassModerationFrom)
    out-logfile -string ("Property in office 365 for coManagers members = "+$office365CoManagers)
    out-logfile -string ("Property in office 365 for coManagers members = "+$office365GrantSendOnBehalfTo)
    out-logfile -string ("Property in office 365 for grant send on behalf to members = "+$office365GrantSendOnBehalfTo)
    out-logfile -string ("Property in office 365 for managed by members = "+$office365ManagedBy)
    out-logfile -string ("Property in office 365 for members = "+$office365Members)
    out-logfile -string ("Property in office 365 for reject messages from members = "+$office365RejectMessagesFrom)
    Out-LogFile -string "********************************************************************************"

    #Perform paramter validation manually.

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "ENTERING PARAMTER VALIDATION"
    Out-LogFile -string "********************************************************************************"

    #Test to ensure that if any of the aadConnect parameters are passed - they are passed together.

    Out-LogFile -string "Validating that both AADConnectServer and AADConnectCredential are specified"
   
    if (($aadConnectServer -eq "") -and ($aadConnectCredential -ne $null))
    {
        #The credential was specified but the server name was not.

        Out-LogFile -string "ERROR:  AAD Connect Server is required when specfying AAD Connect Credential" -isError:$TRUE
    }
    elseif (($aadConnectCredential -eq $NULL) -and ($aadConnectServer -ne ""))
    {
        #The server name was specified but the credential was not.

        Out-LogFile -string "ERROR:  AAD Connect Credential is required when specfying AAD Connect Server" -isError:$TRUE
    }
    elseif (($aadConnectCredential -ne $NULL) -and ($aadConnectServer -ne ""))
    {
        #The server name and credential were specified for AADConnect.

        Out-LogFile -string "AADConnectServer and AADConnectCredential were both specified." 
    
        #Set useAADConnect to TRUE since the parameters necessary for use were passed.
        
        $useAADConnect=$TRUE

        Out-LogFile -string ("Set useAADConnect to TRUE since the parameters necessary for use were passed. - "+$useAADConnect)
    }
    else
    {
        Out-LogFile -string ("Neither AADConnect Server or AADConnect Credentials specified - retain useAADConnect FALSE - "+$useAADConnect)
    }

    #Validate that only one method of engaging exchange online was specified.

    Out-LogFile -string "Validating Exchange Online Credentials."

    if (($exchangeOnlineCredential -ne $NULL) -and ($exchangeOnlineCertificateThumbPrint -ne ""))
    {
        Out-LogFile -string "ERROR:  Only one method of cloud authentication can be specified.  Use either cloud credentials or cloud certificate thumbprint." -isError:$TRUE
    }
    elseif (($exchangeOnlineCredential -eq $NULL) -and ($exchangeOnlineCertificateThumbPrint -eq ""))
    {
        out-logfile -string "ERROR:  One permissions method to connect to Exchange Online must be specified." -isError:$TRUE
    }
    else
    {
        Out-LogFile -string "Only one method of Exchange Online authentication specified."
    }

    #Validate that all information for the certificate connection has been provieed.

    if (($exchangeOnlineCertificateThumbPrint -ne "") -and ($exchangeOnlineOrganizationName -eq "") -and ($exchangeOnlineAppID -eq ""))
    {
        out-logfile -string "The exchange organiztion name and application ID are required when using certificate thumbprint authentication to Exchange Online." -isError:$TRUE
    }
    elseif (($exchangeOnlineCertificateThumbPrint -ne "") -and ($exchangeOnlineOrganizationName -ne "") -and ($exchangeOnlineAppID -eq ""))
    {
        out-logfile -string "The exchange application ID is required when using certificate thumbprint authentication." -isError:$TRUE
    }
    elseif (($exchangeOnlineCertificateThumbPrint -ne "") -and ($exchangeOnlineOrganizationName -eq "") -and ($exchangeOnlineAppID -ne ""))
    {
        out-logfile -string "The exchange organization name is required when using certificate thumbprint authentication." -isError:$TRUE
    }
    else 
    {
        out-logfile -string "All components necessary for Exchange certificate thumbprint authentication were specified."    
    }

    #exit #Debug exit.

    #Validate that an OU was specified <if> retain Contact is not set to true.

    Out-LogFile -string "Validating that if retain original Contact is false a non-sync OU is specified."

    if (($retainOriginalContact -eq $FALSE) -and ($dnNoSyncOU -eq "NotSet"))
    {
        out-LogFile -string "A no SYNC OU is required if retain original Contact is false." -isError:$TRUE
    }

    Out-LogFile -string "END PARAMETER VALIDATION"
    Out-LogFile -string "********************************************************************************"

    Out-Logfile -string "Determine Exchange Schema Version"

    try{
        $exchangeRangeUpper = get-ExchangeSchemaVersion -globalCatalogServer $globalCatalogServer -adCredential $activeDirectoryCredential -errorAction STOP
        out-logfile -string ("The range upper for Exchange Schema is: "+ $exchangeRangeUpper)
    }
    catch{
        out-logfile -string "Error occured obtaining the Exchange Schema Version."
        out-logfile -string $_ -isError:$TRUE
    }
    
    if ($exchangeRangeUpper -ge $exchangeLegacySchemaVersion)
    {
        out-logfile -string "Modern exchange version detected - using modern parameters"
        $ContactPropertySetToClear=$contactPropertiesToClearModern
    }
    else 
    {
        out-logfile -string "Legacy exchange versions detected - using legacy parameters"
        $ContactPropertySetToClear = $contactPropertiesToClearLegacy   
    }

    Out-LogFile -string ("DL property set to be cleared after schema evaluation = ")

    foreach ($dlProperty in $ContactPropertySetToClear)
    {
        Out-LogFile -string $dlProperty
    }

    # EXIT #Debug Exit

    #If exchange server information specified - create the on premises powershell session.

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "ESTABLISH POWERSHELL SESSIONS"
    Out-LogFile -string "********************************************************************************"

   #Test to determine if the exchange online powershell module is installed.
   #The exchange online session has to be established first or the commandlet set from on premises fails.

   Out-LogFile -string "Calling Test-PowerShellModule to validate the Exchange Module is installed."

   Test-PowershellModule -powershellModuleName $exchangeOnlinePowershellModuleName -powershellVersionTest:$TRUE

   Out-LogFile -string "Calling Test-PowerShellModule to validate the Active Directory is installed."

   Test-PowershellModule -powershellModuleName $activeDirectoryPowershellModuleName

   out-logfile -string "Calling Test-PowershellModule to validate the DL Conversion Module version installed."

   Test-PowershellModule -powershellModuleName $contactConversionPowershellModuleName -powershellVersionTest:$TRUE

   #Create the connection to exchange online.

   Out-LogFile -string "Calling New-ExchangeOnlinePowershellSession to create session to office 365."

   if ($exchangeOnlineCredential -ne $NULL)
   {
      #User specified non-certifate authentication credentials.

        try {
            New-ExchangeOnlinePowershellSession -exchangeOnlineCredentials $exchangeOnlineCredential -exchangeOnlineEnvironmentName $exchangeOnlineEnvironmentName -debugLogPath $logFolderPath
        }
        catch {
            out-logfile -string "Unable to create the exchange online connection using credentials."
            out-logfile -string $_ -isError:$TRUE
        }
   }
   elseif ($exchangeOnlineCertificateThumbPrint -ne "")
   {
      #User specified thumbprint authentication.

        try {
            new-ExchangeOnlinePowershellSession -exchangeOnlineCertificateThumbPrint $exchangeOnlineCertificateThumbPrint -exchangeOnlineAppId $exchangeOnlineAppID -exchangeOnlineOrganizationName $exchangeOnlineOrganizationName -exchangeOnlineEnvironmentName $exchangeOnlineEnvironmentName -debugLogPath $logFolderPath
        }
        catch {
            out-logfile -string "Unable to create the exchange online connection using certificate."
            out-logfile -string $_ -isError:$TRUE
        }
   }

    #If the administrator has specified aad connect information - establish the powershell session.

    Out-LogFile -string "Determine if AAD Connect information specified and establish session if necessary."

    if ($useAADConnect -eq $TRUE)
    {
        try 
        {
            out-logfile -string "Creating powershell session to the AD Connect server."

            New-PowershellSession -Server $aadConnectServer -Credentials $aadConnectCredential -PowershellSessionName $aadConnectPowershellSessionName
        }
        catch 
        {
            out-logfile -string "Unable to create remote powershell session to the AD Connect server."
            out-logfile -string $_ -isError:$TRUE
        }
    }

    #Establish powershell session to the global catalog server.

    try 
    {
        Out-LogFile -string "Establish powershell session to the global catalog server specified."

        new-powershellsession -server $globalCatalogServer -credentials $activeDirectoryCredential -powershellsessionname $ADGlobalCatalogPowershellSessionName
    }
    catch 
    {
        out-logfile -string "Unable to create remote powershell session to the AD Global Catalog server."
        out-logfile -string $_ -isError:$TRUE
    }
    

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END ESTABLISH POWERSHELL SESSIONS"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN GET original contact CONFIGURATION LOCAL AND CLOUD"
    Out-LogFile -string "********************************************************************************"

    #At this point we are ready to capture the original contact configuration.  We'll use the ad provider to gather this information.

    Out-LogFile -string "Getting the original contact Configuration"

    try
    {
        $originalContactConfiguration = Get-ADObjectConfiguration -contactSMTPAddress $contactSMTPAddress -globalCatalogServer $globalCatalogWithPort -parameterSet $ContactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential
    }
    catch
    {
        out-logfile -string $_ -isError:$TRUE
    }
    
    Out-LogFile -string "Log original contact configuration."
    out-logFile -string $originalContactConfiguration

    Out-LogFile -string "Create an XML file backup of the on premises DL Configuration"

    Out-XMLFile -itemToExport $originalContactConfiguration -itemNameToExport $originalContactConfigurationADXML

    #exit #Debug Exit

    Out-LogFile -string "Capture the original office 365 list information."

    try 
    {
        $office365ContactConfiguration=Get-O365DLConfiguration -contactSMTPAddress $contactSMTPAddress -errorAction STOP
    }
    catch 
    {
        out-logFile -string $_ -isError:$TRUE
    }
    
    Out-LogFile -string $office365ContactConfiguration

    Out-LogFile -string "Create an XML file backup of the office 365 contact configuration."

    Out-XMLFile -itemToExport $office365ContactConfiguration -itemNameToExport $office365ContactConfigurationXML

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END GET original contact CONFIGURATION LOCAL AND CLOUD"
    Out-LogFile -string "********************************************************************************"

    try 
    {
        Invoke-Office365SafetyCheck -o365dlconfiguration $office365ContactConfiguration -errorAction STOP
    }
    catch 
    {
        out-logFile -string $_ -isError:$TRUE
    }
  
    #At this time we have the DL configuration on both sides and have checked to ensure it is dir synced.
    #Membership of attributes is via DN - these need to be normalized to SMTP addresses in order to find users in Office 365.

    #Start with DL membership and normallize.

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN NORMALIZE DNS FOR ALL ATTRIBUTES"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the members DN to Office 365 identifier."

    if ($originalContactConfiguration.member -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.member)
        {
            #Resetting error variable.

            $isTestError="No"

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -isMember:$TRUE -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "List Membership (ADAttribute: Members)"
                        errorMessage = $normalizedTest.isErrorMessage
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeDLMembershipSMTP+=$normalizedTest
                }
                
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeDLMembershipSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the Contact:"
        
        out-logfile -string $exchangeDLMembershipSMTP
    }
    else 
    {
        out-logFile -string "The Contact has no members."    
    }

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the reject members DN to Office 365 identifier."

    Out-LogFile -string "REJECT USERS"

    if ($originalContactConfiguration.unAuthOrig -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.unAuthOrig)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "RejectMessagesFrom (ADAttribute: UnAuthOrig)"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeRejectMessagesSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    Out-LogFile -string "REJECT ContactS"

    if ($originalContactConfiguration.dlMemRejectPerms -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.dlMemRejectPerms)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "RejectMessagesFromDLMembers (ADAttribute DLMemRejectPerms)"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else {
                    $exchangeRejectMessagesSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeRejectMessagesSMTP -ne $NULL)
    {
        out-logfile -string "The Contact has reject messages members."
        Out-logFile -string $exchangeRejectMessagesSMTP
    }
    else 
    {
        out-logfile "The Contact to be migrated has no reject messages from members."    
    }
    
    Out-LogFile -string "Invoke get-NormalizedDN to normalize the accept members DN to Office 365 identifier."

    Out-LogFile -string "ACCEPT USERS"

    if ($originalContactConfiguration.AuthOrig -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.AuthOrig)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "AcceptMessagesOnlyFrom (ADAttribute: AuthOrig)"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else {
                    $exchangeAcceptMessagesSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logFile -string $_ -isError:$TRUE
            }
        }
    }

    Out-LogFile -string "ACCEPT ContactS"

    if ($originalContactConfiguration.dlMemSubmitPerms -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.dlMemSubmitPerms)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "AcceptMessagesOnlyFromDLMembers (ADAttribute: DLMemSubmitPerms)"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeAcceptMessagesSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeAcceptMessagesSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the accept messages from senders:"
        
        out-logfile -string $exchangeAcceptMessagesSMTP
    }
    else
    {
        out-logFile -string "This Contact has no accept message from restrictions."    
    }
    
    Out-LogFile -string "Invoke get-NormalizedDN to normalize the managedBy members DN to Office 365 identifier."

    Out-LogFile -string "Process MANAGEDBY"

    if ($originalContactConfiguration.managedBy -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.managedBy)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "Owners (ADAttribute: ManagedBy)"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeManagedBySMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    Out-LogFile -string "Process CoMANAGERS"

    if ($originalContactConfiguration.msExchCoManagedByLink -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.msExchCoManagedByLink)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "Owners (ADAttribute: msExchCoManagedByLink"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeManagedBySMTP+=$normalizedTest
                }
                
            }
            catch 
            {
                out-logFile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeManagedBySMTP -ne $NULL)
    {
        #First scan is to ensure that any of the Contacts listed on the managed by objects are still security.
        #It is possible someone added it to managed by and changed the Contact type after.

        foreach ($object in $exchangeManagedBySMTP)
        {
            #If the objec thas a non-null Contact type (is a Contact) and the value of the Contact type matches none of the secuity Contact types.
            #The object is a list - no good.

            if (($object.ContactType -ne $NULL) -and ($object.ContactType -ne "-2147483640") -and ($object.ContactType -ne "-2147483646") -and ($object.ContactType -ne "-2147483644"))
            {
                $isErrorObject = new-Object psObject -property @{
                    primarySMTPAddressOrUPN = $object.primarySMTPAddressOrUPN
                    externalDirectoryObjectID = $object.externalDirectoryObjectID
                    alias=$normalizedTest.alias
                    name=$normalizedTest.name
                    attribute = "Test ManagedBy For Security Flag"
                    errorMessage = "A Contact was found on the owners attribute that is no longer a security Contact.  Security Contact is required.  Remove Contact or change Contact type to security."
                    errorMessageDetail = ""
                }

                out-logfile -string $isErrorObject

                $preCreateErrors+=$isErrorObject

                out-logfile -string "A list (not security enabled) was found on managed by."
                out-logfile -string "The Contact must be converted to security or removed from managed by."
                out-logfile -string $object.primarySMTPAddressOrUPN
            }

            #The Contact is not a list.
            #If the SMTP object of the managedBy object equals the original Contact - check to see if an override is found.
            #If an override of is found - this is not OK since security is required.

            elseif (($object.primarySMTPAddressOrUPN -eq $originalContactConfiguration.mail) -and ($ContactTypeOverride -eq "Distribution")) 
            {
                out-logfile -string "Contact type override detected - Contact has managed by permissions."

                #Contact type is not NULL / Contact type is security value.

                if (($object.ContactType -ne $NULL) -and (($object.ContactType -eq "-2147483640") -or ($object.ContactType -eq "-2147483646" -or ($object.ContactType -eq "-2147483644"))))
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $object.primarySMTPAddressOrUPN
                        externalDirectoryObjectID = $object.externalDirectoryObjectID
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "Test ManagedBy For Contact Override"
                        errorMessage = "The Contact being migrated was found on the Owners attribute.  The administrator has requested migration as not Security.  To remain an owner the Contact must be migrated as Security - remove override or remove owner."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject
    
                    $preCreateErrors+=$isErrorObject
        
                    out-logfile -string "A security Contact has managed by rights on the list."
                    out-logfile -string "The administrator has specified to override the Contact type."
                    out-logfile -string "The Contact override must be removed or the object removed from managedBY."
                    out-logfile -string $object.primarySMTPAddressOrUPN
                }
            }
        }

        Out-LogFile -string "The following objects are members of the managedBY:"
        
        out-logfile -string $exchangeManagedBySMTP
    }
    else 
    {
        out-logfile -string "The Contact has no managers."    
    }

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the moderatedBy members DN to Office 365 identifier."

    Out-LogFile -string "Process MODERATEDBY"

    if ($originalContactConfiguration.msExchModeratedByLink -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.msExchModeratedByLink)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "ModeratedBy (ADAttribute: msExchModeratedByLink"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeModeratedBySMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeModeratedBySMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the moderatedBY:"
        
        out-logfile -string $exchangeModeratedBySMTP    
    }
    else 
    {
        out-logfile "The Contact has no moderators."    
    }

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the bypass moderation users members DN to Office 365 identifier."

    Out-LogFile -string "Process BYPASS USERS"

    if ($originalContactConfiguration.msExchBypassModerationLink -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.msExchBypassModerationLink)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "BypassModerationFromSendersOrMembers (ADAttribute: msExchBypassModerationLink)"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeBypassModerationSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logFile -string $_ -isError:$TRUE
            }
        }
    }

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the bypass moderation Contacts members DN to Office 365 identifier."

    Out-LogFile -string "Process BYPASS ContactS"

    if ($originalContactConfiguration.msExchBypassModerationFromDLMembersLink -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.msExchBypassModerationFromDLMembersLink)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "BypassModerationFromSendersOrMembers (ADAttribute: msExchBypassModerationFromDLMembersLink"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeBypassModerationSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeBypassModerationSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the bypass moderation:"
        
        out-logfile -string $exchangeBypassModerationSMTP 
    }
    else 
    {
        out-logfile "The Contact has no bypass moderation."    
    }

    if ($originalContactConfiguration.publicDelegates -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.publicDelegates)
        {
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalContactDN $originalContactConfiguration.distinguishedName  -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "GrantSendOnBehalfTo (ADAttribute: publicDelegates"
                        errorMessage = $normalizedTest.isErrorMessage
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangeGrantSendOnBehalfToSMTP+=$normalizedTest
                }
                
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeGrantSendOnBehalfToSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the grant send on behalf to:"
        
        out-logfile -string $exchangeGrantSendOnBehalfToSMTP
    }
    else 
    {
        out-logfile "The Contact has no grant send on behalf to."    
    }
    
    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END NORMALIZE DNS FOR ALL ATTRIBUTES"
    Out-LogFile -string "********************************************************************************"

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logFile -string "Summary of Contact information:"
    out-logfile -string ("The number of objects included in the member migration: "+$exchangeDLMembershipSMTP.count)
    out-logfile -string ("The number of objects included in the reject memebers: "+$exchangeRejectMessagesSMTP.count)
    out-logfile -string ("The number of objects included in the accept memebers: "+$exchangeAcceptMessagesSMTP.count)
    out-logfile -string ("The number of objects included in the managedBY memebers: "+$exchangeManagedBySMTP.count)
    out-logfile -string ("The number of objects included in the moderatedBY memebers: "+$exchangeModeratedBySMTP.count)
    out-logfile -string ("The number of objects included in the bypassModeration memebers: "+$exchangeBypassModerationSMTP.count)
    out-logfile -string ("The number of objects included in the grantSendOnBehalfTo memebers: "+$exchangeGrantSendOnBehalfToSMTP.count)
    out-logfile -string ("The number of objects included in the send as rights: "+$exchangeSendAsSMTP.count)
    out-logfile -string ("The number of Contacts on premsies that this Contact has send as rights on: "+$allObjectsSendAsAccessNormalized.Count)
    out-logfile -string ("The number of Contacts on premises that this Contact has full mailbox access on: "+$allObjectsFullMailboxAccess.count)
    out-logfile -string ("The number of mailbox folders on premises that this Contact has access to: "+$allMailboxesFolderPermissions.count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

    #Exit #Debug Exit.

    #At this point we have obtained all the information relevant to the individual Contact.
    #Validate that the discovered dependencies are valid in Office 365.

    $forLoopCounter=0 #Resetting counter at next set of queries.

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN VALIDATE RECIPIENTS IN CLOUD"
    Out-LogFile -string "********************************************************************************"

    out-logfile -string "Begin accepted domain validation."

    try {
        test-AcceptedDomain -originalContactConfiguration $originalContactConfiguration -errorAction STOP
    }
    catch {
        out-logfile $_
        out-logfile -string "Unable to capture accepted domains for validation." -isError:$TRUE
    }

    try {
        $mailOnMicrosoftComDomain = Get-MailOnMicrosoftComDomain -errorAction STOP
    }
    catch {
        out-logfile -string $_
        out-logfile -string "Unable to obtain the onmicrosoft.com domain." -errorAction STOP    
    }

    out-logfile -string "Being validating all list members."
    
    if ($exchangeDLMembershipSMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each DL member is in Office 365 / Exchange Online"

        foreach ($member in $exchangeDLMembershipSMTP)
        {
            #Reset the failure.

            $isTestError="No"

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

            try{
                $isTestError=test-O365Recipient -member $member

                if ($isTestError -eq "Yes")
                {
                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $member.PrimarySMTPAddressorUPN
                        ExternalDirectoryObjectID = $member.ExternalDirectoryObjectID
                        Alias = $member.Alias
                        Name = $member.name
                        Attribute = "Member (ADAttribute: Members)"
                        ErrorMessage = "A member of the list is not found in Office 365."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
            }
            catch{
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }
    else 
    {
        out-logfile -string "There are no DL members to test."    
    }

    out-logfile -string "Begin evaluating all members with reject rights."

    if ($exchangeRejectMessagesSMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each DL reject messages is in Office 365."

        foreach ($member in $exchangeRejectMessagesSMTP)
        {
            #Reset error variable.

            $isTestError="No"

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

            try{
                $isTestError=test-O365Recipient -member $member

                if ($isTestError -eq "Yes")
                {
                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $member.PrimarySMTPAddressorUPN
                        ExternalDirectoryObjectID = $member.ExternalDirectoryObjectID
                        Alias = $member.Alias
                        Name = $member.name
                        Attribute = "RejectMessagesFromSendersorMembers / RejectMessagesFrom / RejectMessagesFromDLMembers (ADAttributes: UnAuthOrig / DLMemRejectPerms)"
                        ErrorMessage = "A member of RejectMessagesFromSendersOrMembers was not found in Office 365."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
            }
            catch{
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }
    else 
    {
        out-logfile -string "There are no reject members to test."    
    }

    out-logfile -string "Begin evaluating all members with accept rights."

    if ($exchangeAcceptMessagesSMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each DL accept messages is in Office 365 / Exchange Online"

        foreach ($member in $exchangeAcceptMessagesSMTP)
        {
            #Reset error variable.

            $isTestError="No"
            
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

            try{
                $isTestError=test-O365Recipient -member $member

                if ($isTestError -eq "Yes")
                {
                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $member.PrimarySMTPAddressorUPN
                        ExternalDirectoryObjectID = $member.ExternalDirectoryObjectID
                        Alias = $member.Alias
                        Name = $member.name
                        Attribute = "AcceptMessagesOnlyFromSendersorMembers / AcceptMessagesOnlyFrom / AcceptMessagesOnlyFromDLMembers (ADAttributes: authOrig / DLMemSubmitPerms)"
                        ErrorMessage = "A member of AcceptMessagesOnlyFromSendersorMembers was not found in Office 365."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
            }
            catch{
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }
    else 
    {
        out-logfile -string "There are no accept members to test."    
    }

    out-logfile -string "Begin evaluating all managed by members."

    if ($exchangeManagedBySMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each DL managed by is in Office 365 / Exchange Online"

        foreach ($member in $exchangeManagedBySMTP)
        {
            #Reset Error Variable.

            $isTestError="No"
            
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

            try{
                $isTestError=test-O365Recipient -member $member

                if ($isTestError -eq "Yes")
                {
                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $member.PrimarySMTPAddressorUPN
                        ExternalDirectoryObjectID = $member.ExternalDirectoryObjectID
                        Alias = $member.Alias
                        Name = $member.name
                        Attribute = "Owners (ADAttributes: ManagedBy,msExchCoManagedByLink)"
                        ErrorMessage = "A member of owners was not found in Office 365."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
            }
            catch{
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }
    else 
    {
        out-logfile -string "There were no managed by members to evaluate."    
    }

    out-logfile -string "Being evaluating all moderated by members."

    if ($exchangeModeratedBySMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each DL moderated by is in Office 365 / Exchange Online"

        foreach ($member in $exchangeModeratedBySMTP)
        {
            #Reset error variable.

            $isTestError="No"

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

            try{
                $isTestError=test-O365Recipient -member $member

                if ($isTestError -eq "Yes")
                {
                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $member.PrimarySMTPAddressorUPN
                        ExternalDirectoryObjectID = $member.ExternalDirectoryObjectID
                        Alias = $member.Alias
                        Name = $member.name
                        Attribute = "ModeratedBy (ADAttributes: msExchModeratedByLink)"
                        ErrorMessage = "A member of moderatedBy was not found in Office 365."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
            }
            catch{
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }
    else 
    {
        out-logfile -string "There were no moderated by members to evaluate."    
    }

    out-logfile -string "Being evaluating all bypass moderation members."

    if ($exchangeBypassModerationSMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each DL bypass moderation is in Office 365 / Exchange Online"

        foreach ($member in $exchangeBypassModerationSMTP)
        {
            #Reset error variable.

            $isTestError="No"

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

            try{
                $isTestError=test-O365Recipient -member $member

                if ($isTestError -eq "Yes")
                {
                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $member.PrimarySMTPAddressorUPN
                        ExternalDirectoryObjectID = $member.ExternalDirectoryObjectID
                        Alias = $member.Alias
                        Name = $member.name
                        Attribute = "BypassModerationFromSendersorMembers (ADAttributes: msExchBypassModerationLink,msExchBypassModerationFromDLMembersLink)"
                        ErrorMessage = "A member of BypassModerationFromSendersorMembers was not found in Office 365."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
            }
            catch{
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }
    else 
    {
        out-logfile -string "There were no bypass moderation members to evaluate."    
    }

    out-logfile -string "Begin evaluation of all grant send on behalf to members."

    if ($exchangeGrantSendOnBehalfToSMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each DL grant send on behalf to is in Office 365 / Exchange Online"

        foreach ($member in $exchangeGrantSendOnBehalfToSMTP)
        {
            $isTestError = "No"

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds..." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

            try{
                $isTestError=test-O365Recipient -member $member

                if ($isTestError -eq "Yes")
                {
                    $isErrorObject = new-Object psObject -property @{
                        PrimarySMTPAddressorUPN = $member.PrimarySMTPAddressorUPN
                        ExternalDirectoryObjectID = $member.ExternalDirectoryObjectID
                        Alias = $member.Alias
                        Name = $member.name
                        Attribute = "GrantSendOnBehalfTo (ADAttributes: publicDelegates)"
                        ErrorMessage = "A member of GrantSendOnBehalfTo was not found in Office 365."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
            }
            catch{
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }
    else 
    {
        out-logfile -string "There were no grant send on behalf to members to evaluate."    
    }

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END VALIDATE RECIPIENTS IN CLOUD"
    Out-LogFile -string "********************************************************************************"

    #It is possible that this Contact was a member of - or other Contacts have a dependency on this Contact.
    #We will implement a function to track those dependen$ocies.

    #At this time we have validated the on premises pre-requisits for Contact migration.
    #If anything is not in order - this code will provide the summary list to the customer and then trigger end.

    if ($preCreateErrors.count -gt 0)
    {
        out-logfile -string "+++++"
        out-logfile -string "Pre-requist checks failed.  Please refer to the following list of items that require addressing for migration to proceed."
        out-logfile -string "+++++"
        out-logfile -string ""

        foreach ($preReq in $preCreateErrors)
        {
            out-logfile -string "====="
            out-logfile -string ("Primary Email Address or UPN: " +$preReq.primarySMTPAddressOrUPN)
            out-logfile -string ("External Directory Object ID: " +$preReq.externalDirectoryObjectID)
            out-logfile -string ("Name: "+$preReq.name)
            out-logfile -string ("Alias: "+$preReq.Alias)
            out-logfile -string ("Attribute in Error: "+$preReq.attribute)
            out-logfile -string ("Error Message Details: "+$preReq.errorMessage)
            out-logfile -string "====="
        }

        out-logfile -string "Pre-requist checks failed.  Please refer to the previous list of items that require addressing for migration to proceed." -isError:$TRUE
    }

    #Exit #Debug Exit

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN RECORD DEPENDENCIES ON MIGRATED Contact"
    Out-LogFile -string "********************************************************************************"

    out-logfile -string "Get all the Contacts that this user is a member of - normalize to canonicalname."

    #Start with Contacts this DL is a member of remaining on premises.

    if ($originalContactConfiguration.memberOf -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.memberof)
        {
            try 
            {
                $allContactsMemberOf += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($allContactsMemberOf -ne $NULL)
    {
        out-logFile -string "The Contact to be migrated is a member of the following Contacts."
        out-logfile -string $allContactsMemberOf
    }
    else 
    {
        out-logfile -string "The Contact is not a member of any other Contacts on premises."
    }

    #Handle all recipients that have forwarding to this Contact based on forwarding address.

    if ($originalContactConfiguration.altRecipientBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.altRecipientBL)
        {
            try 
            {
                $allUsersForwardingAddress += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($allUsersForwardingAddress -ne $NULL)
    {
        out-logFile -string "The Contact has forwarding address set on the following users.."
        out-logfile -string $allUsersForwardingAddress
    }
    else 
    {
        out-logfile -string "The Contact does not have forwarding set on any other users."
    }

    #Handle all Contacts this object has reject permissions on.

    if ($originalContactConfiguration.dLMemRejectPermsBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.dLMemRejectPermsBL)
        {
            try 
            {
                $allContactsReject += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($allContactsReject -ne $NULL)
    {
        out-logFile -string "The Contact has reject permissions on the following Contacts:"
        out-logfile -string $allContactsReject
    }
    else 
    {
        out-logfile -string "The Contact does not have any reject permissions on other Contacts."
    }

    #Handle all Contacts this object has accept permissions on.

    if ($originalContactConfiguration.dLMemSubmitPermsBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.dLMemSubmitPermsBL)
        {
            try 
            {
                $allContactsAccept += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($allContactsAccept -ne $NULL)
    {
        out-logFile -string "The Contact has accept messages from on the following Contacts:"
        out-logfile -string $allContactsAccept
    }
    else 
    {
        out-logfile -string "The Contact does not have accept permissions on any Contacts."
    }

    if ($originalContactConfiguration.msExchCoManagedObjectsBL -ne $NULL)
    {
        out-logfile -string "Calling ge canonical name."

        foreach ($dn in $originalContactConfiguration.msExchCoManagedObjectsBL)
        {
            try 
            {
                $allContactsCoManagedByBL += get-canonicalName -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP

            }
            catch {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }
    else 
    {
        out-logfile -string "The Contact is not a co manager on any other Contacts."    
    }

    if ($allContactsCoManagedByBL -ne $NULL)
    {
        out-logFile -string "The Contact is a co-manager on the following objects:"
        out-logfile -string $allContactsCoManagedByBL
    }
    else 
    {
        out-logfile -string "The Contact is not a co manager on any other objects."
    }

    #Handle all Contacts this object has bypass moderation permissions on.

    if ($originalContactConfiguration.msExchBypassModerationFromDLMembersBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.msExchBypassModerationFromDLMembersBL)
        {
            try 
            {
                $allContactsBypassModeration += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($allContactsBypassModeration -ne $NULL)
    {
        out-logFile -string "This Contact has bypass moderation on the following Contacts:"
        out-logfile -string $allContactsBypassModeration
    }
    else 
    {
        out-logfile -string "This Contact does not have any bypass moderation on any Contacts."
    }

    #Handle all Contacts this object has accept permissions on.

    if ($originalContactConfiguration.publicDelegatesBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.publicDelegatesBL)
        {
            try 
            {
                $allContactsGrantSendOnBehalfTo += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($allContactsGrantSendOnBehalfTo -ne $NULL)
    {
        out-logFile -string "This Contact has grant send on behalf to to the following Contacts:"
        out-logfile -string $allContactsGrantSendOnBehalfTo
    }
    else 
    {
        out-logfile -string "The Contact does ont have any send on behalf of rights to other Contacts."
    }

    #Handle all Contacts this object has manager permissions on.

    if ($originalContactConfiguration.managedObjects -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.managedObjects)
        {
            try 
            {
                $allContactsManagedBy += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($allContactsManagedBy -ne $NULL)
    {
        out-logFile -string "This Contact has managedBY rights on the following Contacts."
        out-logfile -string $allContactsManagedBy
    }
    else 
    {
        out-logfile -string "The Contact is not a manager on any other Contacts."
    }

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logfile -string ("Summary of dependencies found:")
    out-logfile -string ("The number of Contacts that the migrated DL is a member of = "+$allContactsMemberOf.count)
    out-logfile -string ("The number of Contacts that this Contact is a manager of: = "+$allContactsManagedBy.count)
    out-logfile -string ("The number of Contacts that this Contact has grant send on behalf to = "+$allContactsGrantSendOnBehalfTo.count)
    out-logfile -string ("The number of Contacts that have this Contact as bypass moderation = "+$allContactsBypassModeration.count)
    out-logfile -string ("The number of Contacts with accept permissions = "+$allContactsAccept.count)
    out-logfile -string ("The number of Contacts with reject permissions = "+$allContactsReject.count)
    out-logfile -string ("The number of mailboxes forwarding to this Contact is = "+$allUsersForwardingAddress.count)
    out-logfile -string ("The number of Contacts this Contact is a co-manager on = "+$allContactsCoManagedByBL.Count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"


    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END RECORD DEPENDENCIES ON MIGRATED Contact"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "Recording all gathered information to XML to preserve original values."
    
    if ($exchangeDLMembershipSMTP -ne $NULL)
    {
        Out-XMLFile -itemtoexport $exchangeDLMembershipSMTP -itemNameToExport $exchangeDLMembershipSMTPXML
    }
    else 
    {
        $exchangeDLMembershipSMTP=@()
    }

    if ($exchangeRejectMessagesSMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeRejectMessagesSMTP -itemNameToExport $exchangeRejectMessagesSMTPXML
    }
    else 
    {
        $exchangeRejectMessagesSMTP=@()
    }

    if ($exchangeAcceptMessagesSMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeAcceptMessagesSMTP -itemNameToExport $exchangeAcceptMessagesSMTPXML
    }
    else 
    {
        $exchangeAcceptMessagesSMTP=@()
    }

    if ($exchangeManagedBySMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeManagedBySMTP -itemNameToExport $exchangeManagedBySMTPXML
    }
    else 
    {
        $exchangeManagedBySMTP=@()
    }

    if ($exchangeModeratedBySMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeModeratedBySMTP -itemNameToExport $exchangeModeratedBySMTPXML
    }
    else 
    {
        $exchangeModeratedBySMTP=@()
    }

    if ($exchangeBypassModerationSMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeBypassModerationSMTP -itemNameToExport $exchangeBypassModerationSMTPXML
    }
    else 
    {
        $exchangeBypassModerationSMTP=@()
    }

    if ($exchangeGrantSendOnBehalfToSMTP -ne $NULL)
    {
        out-xmlfile -itemToExport $exchangeGrantSendOnBehalfToSMTP -itemNameToExport $exchangeGrantSendOnBehalfToSMTPXML
    }
    else 
    {
        $exchangeGrantSendOnBehalfToSMTP=@()
    }

    if ($allContactsMemberOf -ne $NULL)
    {
        out-xmlfile -itemtoexport $allContactsMemberOf -itemNameToExport $allContactsMemberOfXML
    }
    else 
    {
        $allContactsMemberOf=@()
    }
    
    if ($allContactsReject -ne $NULL)
    {
        out-xmlfile -itemtoexport $allContactsReject -itemNameToExport $allContactsRejectXML
    }
    else 
    {
        $allContactsReject=@()
    }
    
    if ($allContactsAccept -ne $NULL)
    {
        out-xmlfile -itemtoexport $allContactsAccept -itemNameToExport $allContactsAcceptXML
    }
    else 
    {
        $allContactsAccept=@()
    }

    if ($allContactsCoManagedByBL -ne $NULL)
    {
        out-xmlfile -itemToExport $allContactsCoManagedByBL -itemNameToExport $allContactsCoManagedByXML
    }
    else 
    {
        $allContactsCoManagedByBL=@()    
    }

    if ($allContactsBypassModeration -ne $NULL)
    {
        out-xmlfile -itemtoexport $allContactsBypassModeration -itemNameToExport $allContactsBypassModerationXML
    }
    else 
    {
        $allContactsBypassModeration=@()
    }

    if ($allUsersForwardingAddress -ne $NULL)
    {
        out-xmlFile -itemToExport $allUsersForwardingAddress -itemNameToExport $allUsersForwardingAddressXML
    }
    else 
    {
        $allUsersForwardingAddress=@()
    }

    if ($allContactsManagedBy -ne $NULL)
    {
        out-xmlFile -itemToExport $allContactsManagedBy -itemNameToExport $allContactsManagedByXML
    }
    else 
    {
        $allContactsManagedBy=@()
    }

    if ($allContactsGrantSendOnBehalfTo -ne $NULL)
    {
        out-xmlFile -itemToExport $allContactsGrantSendOnBehalfTo -itemNameToExport $allContactsGrantSendOnBehalfToXML
    }
    else 
    {
        $allContactsGrantSendOnBehalfTo =@()
    }

    #EXIT #Debug Exit

    #Ok so at this point we have preserved all of the information regarding the on premises DL.
    #It is possible that there could be cloud only objects that this Contact was made dependent on.
    #For example - the dirSync Contact could have been added as a member of a cloud only Contact - or another Contact that was migrated.
    #The issue here is that this gets VERY expensive to track - since some of the word to do do is not filterable.
    #With the LDAP improvements we no longer offert the option to track on premises - but the administrator can choose to track the cloud

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "START RETAIN OFFICE 365 Contact DEPENDENCIES"
    Out-LogFile -string "********************************************************************************"

    #Process normal mail enabled Contacts.

    if ($retainOffice365Settings -eq $TRUE)
    {
        out-logFile -string "Office 365 settings are to be retained."

        try {
            $allOffice365MemberOf = Get-O365ContactDependency -dn $office365ContactConfiguration.distinguishedName -attributeType $office365Members -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of Contacts in Office 365 cloud only that the DL is a member of = "+$allOffice365MemberOf.count)

        try {
            $allOffice365Accept = Get-O365ContactDependency -dn $office365ContactConfiguration.distinguishedName -attributeType $office365AcceptMessagesFrom -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of Contacts in Office 365 cloud only that the DL has accept rights = "+$allOffice365Accept.count)

        try {
            $allOffice365Reject = Get-O365ContactDependency -dn $office365ContactConfiguration.distinguishedName -attributeType $office365RejectMessagesFrom -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of Contacts in Office 365 cloud only that the DL has reject rights = "+$allOffice365Reject.count)

        try {
            $allOffice365BypassModeration = Get-O365ContactDependency -dn $office365ContactConfiguration.distinguishedName -attributeType $office365BypassModerationFrom -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of Contacts in Office 365 cloud only that the DL has grant send on behalf to righbypassModeration rights = "+$allOffice365BypassModeration.count)

        try {
            $allOffice365GrantSendOnBehalfTo = Get-O365ContactDependency -dn $office365ContactConfiguration.distinguishedName -attributeType $office365GrantSendOnBehalfTo -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of Contacts in Office 365 cloud only that the DL has grantSendOnBehalFto = "+$allOffice365GrantSendOnBehalfTo.count)

        try {
            $allOffice365ManagedBy = Get-O365ContactDependency -dn $office365ContactConfiguration.distinguishedName -attributeType $office365ManagedBy -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of Contacts in Office 365 cloud only that the DL has managedBY = "+$allOffice365ManagedBy.count)

        #Process other mail enabled object dependencies.

        try {
            $allOffice365ForwardingAddress = Get-O365ContactDependency -dn $office365ContactConfiguration.distinguishedName -attributeType $office365ForwardingAddress -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of Contacts in Office 365 cloud only that the DL has forwarding on mailboxes = "+$allOffice365ForwardingAddress.count)

        if ($allOffice365MemberOf -ne $NULL)
        {
            out-logfile -string $allOffice365MemberOf
            out-xmlfile -itemtoexport $allOffice365MemberOf -itemNameToExport $allOffice365MemberofXML
        }
        else 
        {
            $allOffice365MemberOf=@()
        }

        if ($allOffice365Accept -ne $NULL)
        {
            out-logfile -string $allOffice365Accept
            out-xmlFile -itemToExport $allOffice365Accept -itemNameToExport $allOffice365AcceptXML
        }
        else 
        {
            $allOffice365Accept=@()    
        }

        if ($allOffice365Reject -ne $NULL)
        {
            out-logfile -string $allOffice365Reject
            out-xmlFile -itemToExport $allOffice365Reject -itemNameToExport $allOffice365RejectXML
        }
        else 
        {
            $allOffice365Reject=@()    
        }
        
        if ($allOffice365BypassModeration -ne $NULL)
        {
            out-logfile -string $allOffice365BypassModeration
            out-xmlFile -itemToExport $allOffice365BypassModeration -itemNameToExport $allOffice365BypassModerationXML
        }
        else 
        {
            $allOffice365BypassModeration=@()    
        }

        if ($allOffice365GrantSendOnBehalfTo -ne $NULL)
        {
            out-logfile -string $allOffice365GrantSendOnBehalfTo
            out-xmlfile -itemToExport $allOffice365GrantSendOnBehalfTo -itemNameToExport $allOffice365GrantSendOnBehalfToXML
        }
        else 
        {
            $allOffice365GrantSendOnBehalfTo=@()    
        }

        if ($allOffice365ManagedBy -ne $NULL)
        {
            out-logfile -string $allOffice365ManagedBy
            out-xmlFile -itemToExport $allOffice365ManagedBy -itemNameToExport $allOffice365ManagedByXML

            out-logfile -string "Setting Contact type override to security - the Contact type may have changed on premises after the permission was added."

            $ContactTypeOverride="Security"
        }
        else 
        {
            $allOffice365ManagedBy=@()    
        }

        if ($allOffice365ForwardingAddress -ne $NULL)
        {
            out-logfile -string $allOffice365ForwardingAddress
            out-xmlfile -itemToExport $allOffice365ForwardingAddress -itemNameToExport $allOffice365ForwardingAddressXML
        }
        else 
        {
            $allOffice365ForwardingAddress=@()    
        }

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logfile -string ("Summary of dependencies found:")
    out-logfile -string ("The number of office 365 objects that the migrated DL is a member of = "+$allOffice365MemberOf.count)
    out-logfile -string ("The number of office 365 objects that this Contact is a manager of: = "+$allOffice365ManagedBy.count)
    out-logfile -string ("The number of office 365 objects that this Contact has grant send on behalf to = "+$allOffice365GrantSendOnBehalfTo.count)
    out-logfile -string ("The number of office 365 objects that have this Contact as bypass moderation = "+$allOffice365BypassModeration.count)
    out-logfile -string ("The number of office 365 objects with accept permissions = "+$allOffice365Accept.count)
    out-logfile -string ("The number of office 365 objects with reject permissions = "+$allOffice365Reject.count)
    out-logfile -string ("The number of office 365 mailboxes forwarding to this Contact is = "+$allOffice365ForwardingAddress.count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

    EXIT #Debug Exit

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END RETAIN OFFICE 365 Contact DEPENDENCIES"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "START Remove on premises Contact from office 365.."
    Out-LogFile -string "********************************************************************************"

    #At this stage we will move the Contact to the non-Sync OU and then re-record the attributes.
    #The move here will allow us to preserve the original Contacts with attributes until we know that the migration was successful.
    #We will use the move to the non-SYNC OU to trigger deletion.

    #EXIT #Debug exit

    #In this case we should write a status file for all threads.
    #When all threads have reached this point it is safe to have them all move their DLs to the non-Sync OU.

    #If there are multiple threads in use hold all of them for thread 

    out-logfile -string "Determine if multiple migration threads are in use..."

    if ($totalThreadCount -eq 0)
    {
        out-logfile -string "Multiple threads are not in use.  Continue functions..."
    }
    else 
    {
        out-logfile -string "Multiple threads are in use.  Hold at this point for all threads to reach the point of moving to non-Sync OU."

        try{
            out-statusFile -threadNumber $global:threadNumber -errorAction STOP
        }
        catch{
            out-logfile -string "Unable to write status file." -isError:$TRUE
        }

        do 
        {
            out-logfile -string "All threads are not ready - sleeping."
        } until ((get-statusFileCount) -eq  $totalThreadCount)
    }

    try {
        move-toNonSyncOU -dn $originalContactConfiguration.distinguishedName -OU $dnNoSyncOU -globalCatalogServer $globalCatalogServer -adCredential $activeDirectoryCredential -errorAction STOP
    }
    catch {
        out-logfile -string $_ -isError:$TRUE
    }

    #If there are multiple threads have all threads > 1 sleep for 15 seconds while thread one deletes all status files.
    #This should cover the 5 seconds that any other threads may be sleeping looking to read the status directory.

    if ($totalThreadCount -gt 0)
    {
        start-sleepProgress -sleepString "Starting sleep before removing individual status files.." -sleepSeconds 5

        out-logfile -string "Trigger cleanup of individual status files."

        try{
            remove-statusFiles -functionThreadNumber $global:threadNumber
        }
        catch{
            out-logfile -string "Unable to remove status files" -isError:$TRUE
        }

        start-sleepProgress -sleepString "Starting sleep after removing individual status files.." -sleepSeconds 5
    }

    #$Capture the moved DL configuration (since attibutes change upon move.)

    try {
        $originalContactConfigurationUpdated = Get-ADObjectConfiguration -contactSMTPAddress $contactSMTPAddress -globalCatalogServer $globalCatalogWithPort -parameterSet $ContactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 
    }
    catch {
        out-logFile -string $_ -isError:$TRUE
    }

    out-LogFile -string $originalContactConfigurationUpdated
    out-xmlFile -itemToExport $originalContactConfigurationUpdated -itemNameTOExport $originalContactConfigurationUpdatedXML

    

    

    #If there are multiple threads and we've reached this point - we're ready to write a status file.

    out-logfile -string "If thread number > 1 - write the status file here."

    if ($global:threadNumber -gt 1)
    {
        out-logfile -string "Thread number is greater than 1."

        try{
            out-statusFile -threadNumber $global:threadNumber -errorAction STOP
        }
        catch{
            out-logfile -string $_
            out-logfile -string "Unable to write status file." -isError:$TRUE
        }
    }
    else 
    {
        out-logfile -string "Thread number is 1 - do not write status at this time."    
    }

    #If there are multiple threads - only replicate the domain controllers and trigger AD connect if all threads have completed their work.

    out-logfile -string "Determine if multiple migration threads are in use..."

    if ($totalThreadCount -eq 0)
    {
        out-logfile -string "Multiple threads are not in use.  Continue functions..."
    }
    else 
    {
        out-logfile -string "Multiple threads are in use - depending on thread number take different actions."
        
        if ($global:threadNumber -eq 1)
        {
            out-logfile -string "This is the master thread responsible for triggering operations."
            out-logfile -string "Search status directory and count files - if file count = number of threads - 1 thread 1 can proceed."

            #Do the following until the count of the files in the directory = number of threads - 1.

            do 
            {
                out-logfile -string "Other threads are pending.  Sleep 5 seconds."
            } until ((get-statusFileCount) -eq ($totalThreadCount - 1))
        }
        elseif ($global:threadNumber -gt 1)
        {
            out-logfile -string "This is not the master thread responsible for triggering operations."
            out-logfile -string "Search directory and count files.  If the file count = number of threads proceed."

            do 
            {
                out-logfile -string "Thread 1 is not ready to trigger.  Sleep 5 seconds."
            } until ((get-statusFileCount) -eq  $totalThreadCount)
        }
    }

    #Replicate domain controllers so that the change is received as soon as possible.()
    
    if ($global:threadNumber -eq 0 -or ($global:threadNumber -eq 1))
    {
        start-sleepProgress -sleepString "Starting sleep before invoking AD replication - 15 seconds." -sleepSeconds 15

        out-logfile -string "Invoking AD replication."

        try {
            invoke-ADReplication -globalCatalogServer $globalCatalogServer -powershellSessionName $ADGlobalCatalogPowershellSessionName -errorAction STOP
        }
        catch {
            out-logfile -string $_
        }
    }

    #Start the process of syncing the deletion to the cloud if the administrator has provided credentials.
    #Note:  If this is not done we are subject to sitting and waiting for it to complete.

    if ($global:threadNumber -eq 0 -or ($global:threadNumber -eq 1))
    {
        if ($useAADConnect -eq $TRUE)
        {
            start-sleepProgress -sleepString "Starting sleep before invoking AD Connect - one minute." -sleepSeconds 60

            out-logfile -string "Invoking AD Connect."

            invoke-ADConnect -powerShellSessionName $aadConnectPowershellSessionName

            start-sleepProgress -sleepString "Starting sleep after invoking AD Connect - one minute." -sleepSeconds 60
        }   
        else 
        {
            out-logfile -string "AD Connect information not specified - allowing ad connect to run on normal cycle and process deletion."    
        }
    }   

    #The single functions have triggered operations.  Other threads may continue.

    if ($global:threadNumber -eq 1)
    {
        out-statusFile -threadNumber $global:threadNumber
        start-sleepProgress -sleepString "Starting sleep after writing file..." -sleepSeconds 3
    }

    #If this is the main thread - introduce a sleep for 10 seconds - allows the other threads to detect 5 files.
    #Reset the status directory for furture thread dependencies.

    if ($totalThreadCount -gt 0)
    {
        start-sleepProgress -sleepString "Starting sleep before removing individual status files.." -sleepSeconds 5

        out-logfile -string "Trigger cleanup of individual status files."

        try{
            remove-statusFiles -functionThreadNumber $global:threadNumber
        }
        catch{
            out-logfile -string "Unable to remove status files" -isError:$TRUE
        }

        start-sleepProgress -sleepString "Starting sleep after removing individual status files.." -sleepSeconds 5
    }
    
    #At this time we have processed the deletion to azure.
    #We need to wait for that deletion to occur in Exchange Online.

    out-logfile -string "Monitoring Exchange Online for list deletion."

    try {
        test-CloudDLPresent -contactSMTPAddress $contactSMTPAddress -errorAction SilentlyContinue
    }
    catch {
        out-logfile -string $_ -isError:$TRUE
    }

    #At this point we have validated that the Contact is gone from office 365.
    #We can begin the process of recreating the Contact in Exchange Online.

    out-logfile "Attempting to create the DL in Office 365."

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try {
            $office365ContactConfigurationPostMigration=new-office365dl -originalContactConfiguration $originalContactConfiguration -office365ContactConfiguration $office365ContactConfiguration -Contacttypeoverride $ContactTypeOverride -errorAction STOP

            #If we made it this far then the Contact was created.

            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 10)
            {
                out-logFile -string $_ -isError:$TRUE 
            }
            else 
            {
                out-logfile -string "Unable to create the list on attempt.  Retry"

                if ($loopCounter -gt 0)
                {
                    start-sleepProgress -sleepSeconds ($loopCounter * 5) -sleepstring "Invoke sleep - error creating Contact."
                }
                $loopCounter=$loopCounter+1
            }
        }
    } while ($stopLoop -eq $FALSE)

    #Sometimes the configuration is not immediately available due to ad sync time in Office 365.
    #Implement a loop that protects us here - trying 10 times and sleeping the bare minimum in between to eliminate longer static sleeps.

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do 
    {
        try {
                        
            #If we hit here we did not get a terminating error.  Write the configuration.

            out-LogFile -string "Write new DL configuration to XML."

            out-Logfile -string $office365ContactConfigurationPostMigration
            out-xmlFile -itemToExport $office365ContactConfigurationPostMigration -itemNameToExport $office365ContactConfigurationPostMigrationXML
            
            #If we made it this far we can end the loop - we were succssful.

            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 10)
            {
                out-logfile -string "Unable to get Office 365 list configuration after 10 tries."
                $stopLoop -eq $TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to capture the office 365 contact configuration.  Sleeping 15 seconds." -sleepSeconds 15

                $loopCounter = $loopCounter+1 
            }
        }   
    } while ($stopLoop -eq $false)

    #EXIT #Debug Exit.

    #Now it is time to set the multi valued attributes on the DL in Office 365.
    #Setting these first must occur since moderators have to be established before moderation can be enabled.

    out-logFile -string "Setting the multivalued attributes of the migrated Contact."

    out-logfile -string $office365ContactConfigurationPostMigration.primarySMTPAddress

    [int]$loopCounter=0
    [boolean]$stopLoop = $FALSE
    
    do {
        try {
            set-Office365DLMV -originalContactConfiguration $originalContactConfiguration -office365ContactConfiguration $office365ContactConfiguration -office365ContactConfigurationPostMigration $office365ContactConfigurationPostMigration -exchangeDLMembership $exchangeDLMembershipSMTP -exchangeRejectMessage $exchangeRejectMessagesSMTP -exchangeAcceptMessage $exchangeAcceptMessagesSMTP -exchangeModeratedBy $exchangeModeratedBySMTP -exchangeManagedBy $exchangeManagedBySMTP -exchangeBypassMOderation $exchangeBypassModerationSMTP -exchangeGrantSendOnBehalfTo $exchangeGrantSendOnBehalfToSMTP -errorAction STOP -ContactTypeOverride $ContactTypeOverride -exchangeSendAsSMTP $exchangeSendAsSMTP -mailOnMicrosoftComDomain $mailOnMicrosoftComDomain -allowNonSyncedContact $allowNonSyncedContact -allOffice365SendAsAccessOnContact $allOffice365SendAsAccessOnContact 

            $stopLoop = $TRUE
        }
        catch {
            if ($loopCounter -gt 4)
            {
                out-logFile -string $_ -isError:$TRUE
            }
            else {
                start-sleepProgress -sleepString "Uanble to set office 365 contact Multi Value attributes - try again." -sleepSeconds 5

                $loopCounter = $loopCounter +1
            } 
        }
    } while ($stopLoop -eq $FALSE)

    out-logfile -string ("The number of post create errors is: "+$global:postCreateErrors.count)

    #Sometimes the configuration is not immediately available due to ad sync time in Office 365.
    #Implement a loop that protects us here - trying 10 times and sleeping the bare minimum in between to eliminate longer static sleeps.

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try {
            $office365ContactConfigurationPostMigration = Get-O365DLConfiguration -contactSMTPAddress $office365ContactConfigurationPostMigration.externalDirectoryObjectID -errorAction STOP

            #If we made it this far we were successful - output the information to XML.

            out-LogFile -string "Write new DL configuration to XML."

            out-Logfile -string $office365ContactConfigurationPostMigration
            out-xmlFile -itemToExport $office365ContactConfigurationPostMigration -itemNameToExport $office365ContactConfigurationPostMigrationXML

            #Now that we are this far - we can exit the loop.

            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 10)
            {
                out-logfile -string "Unable to get Office 365 list configuration after 10 tries."
                $stopLoop -eq $TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to capture the office 365 contact configuration.  Sleeping 15 seconds." -sleepSeconds 15

                $loopCounter = $loopCounter+1 
            }
        }
        
    } while ($stopLoop -eq $FALSE)

    

    

    #The list has now been created.  There are single value attributes that we're now ready to update.

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try {
            set-Office365DL -originalContactConfiguration $originalContactConfiguration -office365ContactConfiguration $office365ContactConfiguration -ContactTypeOverride $ContactTypeOverride -office365ContactConfigurationPostMigration $office365ContactConfigurationPostMigration
            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 4)
            {
                out-logfile -string $_ -isError:$TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Transient error updating Contact - retrying." -sleepSeconds 5

                $loopCounter=$loopCounter+1
            }
        }
    } while ($stopLoop -eq $FALSE)

    
    out-logfile -string ("The number of post create errors is: "+$global:postCreateErrors.count)
    

    out-logFile -string ("Capture the DL status post migration.")

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try {
            $office365ContactConfigurationPostMigration = Get-O365DLConfiguration -contactSMTPAddress $office365ContactConfigurationPostMigration.externalDirectoryObjectID -errorAction STOP

            #If we made it this far we successfully got the DL.  Write it.

            out-LogFile -string "Write new DL configuration to XML."

            out-Logfile -string $office365ContactConfigurationPostMigration
            out-xmlFile -itemToExport $office365ContactConfigurationPostMigration -itemNameToExport $office365ContactConfigurationPostMigrationXML

            #Now that we wrote it - stop the loop.

            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 10)
            {
                out-logfile -string "Unable to get Office 365 list configuration after 10 tries."
                $stopLoop -eq $TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to capture the office 365 contact configuration.  Sleeping 15 seconds." -sleepSeconds 15

                $loopCounter = $loopCounter+1 
            }
        }   
    } while ($stopLoop -eq $false)

    out-logfile -string "Obtain the migrated DL membership and record it for validation."

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try{
            $office365DLMembershipPostMigration = get-O365DLMembership -contactSMTPAddress $office365ContactConfigurationPostMigration.externalDirectoryObjectID -errorAction STOP

            #Membership obtained - export.

            out-logFile -string "Write the new DL membership to XML."
            out-logfile -string office365DLMembershipPostMigration

            out-xmlFile -itemToExport office365DLMembershipPostMigration -itemNametoExport $office365DLMembershipPostMigrationXML

            #Exports complete - stop loop

            $stopLoop=$TRUE
        }
        catch{
            if ($loopCounter -gt 10)
            {
                out-logfile -string "Unable to get Office 365 list configuration after 10 tries."
                $stopLoop -eq $TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to capture the office 365 contact configuration.  Sleeping 15 seconds." -sleepSeconds 15
 
                $loopCounter = $loopCounter+1 
            }
        }
    } while ($stopLoop -eq $FALSE)

    #The Contact has been created and both single and multi valued attributes have been updated.
    #The Contact is fully availablle in exchange online.
    #The Contact as this point sits in the non-sync OU.  This was to service the deletion.
    #The administrator may have reasons for keeping the Contact.
    #If they do the plan is to do two things.
    ###Rename the Contact by adding a ! to the name - this ensures that if the Contact is every accidentally mail enabled it will not soft match the migrated Contact.
    ###We'll stamp custom attribute flags on it to ensure that we know the Contact has been mirgated - in case it's a member of another Contact to be migrated.

    if ($retainOriginalContact -eq $TRUE)
    {
        Out-LogFile -string "Administrator has choosen to retain the original Contact."
        out-logfile -string "Rename the Contact by adding the fixed character !"

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE   

        do {
            try {
                set-newDLName -globalCatalogServer $globalCatalogServer -dlName $originalContactConfigurationUpdated.Name -dlSAMAccountName $originalContactConfigurationUpdated.SAMAccountName -dn $originalContactConfigurationUpdated.distinguishedName -adCredential $activeDirectoryCredential -errorAction STOP

                $stopLoop=$TRUE
            }
            catch {
                if($loopCounter -gt 4)
                {
                    out-logfile -string $_ -isError:$TRUE
                }
                else 
                {
                    start-sleepProgress -sleepString "Uanble to change DL name - try again." -sleepSeconds 5
                    $loopCounter = $loopCounter+1    
                }
            }
        } while ($stopLoop=$FALSE)

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE

        do {
            try {
                $originalContactConfigurationUpdated = Get-ADObjectConfiguration -contactSMTPAddress $contactSMTPAddress -globalCatalogServer $globalCatalogWithPort -parameterSet $ContactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

                $stopLoop=$TRUE
            }
            catch {
                if ($loopCounter -gt 4)
                {
                    out-logFile -string $_ -isError:$TRUE
                }
                else 
                {
                    start-sleepProgress -sleepString "Unable to obtain updated original contact Configuration - try again." -sleepSeconds 5

                    $loopCounter = $loopCounter+1
                }
            }
        } while ($stopLoop -eq $FALSE)

        out-logfile -string $originalContactConfigurationUpdated
        out-xmlFile -itemToExport $originalContactConfigurationUpdated -itemNameTOExport $originalContactConfigurationUpdatedXML+$global:unDoStatus

        Out-LogFile -string "Administrator has choosen to regain the original Contact."
        out-logfile -string "Disabling the mail attributes on the Contact."

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE
        
        do {
            try{
                Disable-OriginalDL -originalContactConfiguration $originalContactConfigurationUpdated -globalCatalogServer $globalCatalogServer -parameterSet $ContactPropertySetToClear -adCredential $activeDirectoryCredential -useOnPremisesExchange $useOnPremisesExchange -errorAction STOP

                $stopLoop = $TRUE
            }
            catch{
                if ($loopCounter -gt 4)
                {
                    out-LogFile -string $_ -isError:$TRUE
                }
                else 
                {
                    start-sleepProgress -sleepString "Unable to disable Contact - try again." -sleepSeconds 5

                    $loopCounter = $loopCounter + 1
                }
            }
        } while ($stopLoop -eq $false)

        

        

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE
        
        do {
            try {
                $originalContactConfigurationUpdated = Get-ADObjectConfiguration -dn $originalContactConfigurationUpdated.distinguishedName -globalCatalogServer $globalCatalogWithPort -parameterSet $ContactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

                $stopLoop = $TRUE
            }
            catch {
                if ($loopCounter -gt 4)
                {
                    out-logFile -string $_ -isError:$TRUE
                }
                else {
                    start-sleeProgress -sleepString "Attempt to gather updated DL configuration failed - try again." -sleepSeconds 5

                    $loopCounter = $loopCounter + 1
                } 
            }
        } while ($stopLoop -eq $FALSE)


        out-logfile -string $originalContactConfigurationUpdated
        out-xmlFile -itemToExport $originalContactConfigurationUpdated -itemNameTOExport $originalContactConfigurationUpdatedXML+$global:unDoStatus

        Out-LogFile -string "Move the original Contact back to the OU it came from.  The Contact will no longer be soft matched."

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE

        do {
            try {
                #Discovered that it's possible someone used the name "Test Contact".  This breaks the following DN search as OU appears in the name - WHOOPS
                #So we need to try to make the substring call more unique - as to avoid detecting OU in a name.
                #To do so - we know that the DN has ,OU= so the first substring we'll search is ,OU=. 
                #Then we'll do it again - this time for just OU.  And that should give us what we need for the OU.

                $tempOUSubstring = Get-OULocation -originalContactConfiguration $originalContactConfiguration -errorAction STOP

                move-toNonSyncOU -DN $originalContactConfigurationUpdated.distinguishedName -ou $tempOUSubstring -globalCatalogServer $globalCatalogServer -adCredential $activeDirectoryCredential -errorAction STOP

                $stopLoop = $TRUE
            }
            catch {
                if ($loopCounter -gt 4)
                {
                    out-logfile -string $_ -isError:$TRUE
                }
                else {

                    out-logfile -string $_
                    start-sleepProgress -sleepString "Unable to move the DL to a non-sync OU - try again." -sleepSeconds 5

                    $loopCounter = $loopCounter +1
                }
            }
        } while ($stopLoop -eq $FALSE)

        [int]$loopCounter = 0
        [boolean]$stopLoop = $FALSE

        do {
            try {
                $tempOU=get-OULocation -originalContactConfiguration $originalContactConfiguration
                $tempNameArray=$originalContactConfigurationUpdated.distinguishedName.split(",")
                $tempDN=$tempNameArray[0]+","+$tempOU
                $originalContactConfigurationUpdated = Get-ADObjectConfiguration -dn $tempDN -globalCatalogServer $globalCatalogWithPort -parameterSet $ContactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

                $stopLoop = $TRUE
            }
            catch {
                if ($loopCounter -gt 4)
                {
                    out-logFile -string $_ -isError:$TRUE
                }
                else {
                    start-sleepProgress -sleepString "Unable to obtain moved DL configuration - try again." -sleepSeconds 5

                    $loopCounter = $loopCounter +1
                }
            }
        } while ($stopLoop = $FALSE)

        out-logfile -string $originalContactConfigurationUpdated
        out-xmlFile -itemToExport $originalContactConfigurationUpdated -itemNameTOExport $originalContactConfigurationUpdatedXML+$global:unDoStatus

        

        
    }

    #Now it is time to create the routing contact.

    [int]$loopCounter = 0
    [boolean]$stopLoop = $FALSE
    
    do {
        try {
            new-routingContact -originalContactConfiguration $originalContactConfiguration -office365ContactConfiguration $office365ContactConfigurationPostMigration -globalCatalogServer $globalCatalogServer -adCredential $activeDirectoryCredential

            $stopLoop = $TRUE
        }
        catch {
            if ($loopCounter -gt 4)
            {
                out-logfile -string $_ -isError:$TRUE
            }
            else {
                start-sleepProgress -sleepString "Unable to create routing contact - try again." -sleepSeconds 5

                $loopCounter = $loopCounter +1
            }
        }
    } while ($stopLoop -eq $FALSE)

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try {
            <#
            $tempOU=get-OULocation -originalContactConfiguration $originalContactConfiguration
            out-logfile -string $tempOU
            $tempName=$originalContactConfiguration.cn
            out-logfile -string $tempName
            $tempName=$tempname.replace(' ','')
            out-logfile -string $tempname
            $tempName=$tempName+"-MigratedByScript"
            out-logfile -string $tempName
            $tempName="CN="+$tempName
            out-logfile -string $tempName
            $tempDN=$tempName+","+$tempOU
            out-logfile -string $tempDN
            #>

            $tempMailArray = $originalContactConfiguration.mail.split("@")

            foreach ($member in $tempMailArray)
            {
                out-logfile -string ("Temp Mail Address Member: "+$member)
            }

            $tempMailAddress = $tempMailArray[0]+"-MigratedByScript"

            out-logfile -string ("Temp routing contact address: "+$tempMailAddress)

            $tempMailAddress = $tempMailAddress+"@"+$tempMailArray[1]

            out-logfile -string ("Temp routing contact address: "+$tempMailAddress)

            $routingContactConfiguration = Get-ADObjectConfiguration -contactSMTPAddress $tempMailAddress -globalCatalogServer $globalCatalogWithPort -parameterSet $ContactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

            $stopLoop=$TRUE
        }
        catch 
        {
            if ($loopCounter -gt 5)
            {
                out-logfile -string "Unable to obtain routing contact information post creation."
                out-logfile -string $_ -isError:$TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to obtain routing contact after creation - sleep try again." -sleepSeconds 10
                $loopCounter = $loopCounter + 1                
            }
        }
    } while ($stopLoop -eq $FALSE)

    out-logfile -string $routingContactConfiguration
    out-xmlFile -itemToExport $routingContactConfiguration -itemNameTOExport $routingContactXML

    

    

    #At this time the contact is created - issuing a replication of domain controllers and sleeping one minute.
    #We've gotta get the contact pushed out so that cross domain operations function - otherwise reconciling memership fails becuase the contacts not available.

    start-sleepProgress -sleepString "Starting sleep before invoking AD replication.  Sleeping 15 seconds." -sleepSeconds 15

    out-logfile -string "Invoking AD replication."

    try {
        invoke-ADReplication -globalCatalogServer $globalCatalogServer -powershellSessionName $ADGlobalCatalogPowershellSessionName -errorAction STOP
    }
    catch {
        out-logfile -string $_
    }

    $forLoopCounter=0 #Restting loop counter for next series of operations.

    #At this time we are ready to begin resetting the on premises dependencies.

    $isTestError = "No" #Reset error tracking.

    out-logfile -string ("Starting on premies DL members.")

    if ($allContactsMemberOf.count -gt 0)
    {
        foreach ($member in $allContactsMemberOf)
        {  
            $isTestError = "No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5

                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-logfile -string ("Processing member = "+$member.canonicalName)
            out-logfile -string ("Routing contact DN = "+$routingContactConfiguration.distinguishedName)
            out-logfile -string ("Attribute Operation = "+$onPremMemberOf)

            if ($member.distinguishedName -ne $originalContactConfiguration.distinguishedName)
            {
                try{
                    $isTestError=start-replaceOnPrem -routingContact $routingContactConfiguration -attributeOperation $onPremMemberOf -canonicalObject $member -adCredential $activeDirectoryCredential -globalCatalogServer $globalCatalogServer -errorAction STOP
                }
                catch{
                    out-logfile -string $_
                    $isTestError="Yes"
                }

                if ($isTestError -eq "Yes")
                {
                    out-logfile -string "Error adding routing contact to on premises resource."

                    $isErrorObject = new-Object psObject -property @{
                        distinguishedName = $member.distinguishedName
                        canonicalDomainName = $member.canonicalDomainName
                        canonicalName=$member.canonicalName
                        attribute = "List Membership (ADAttribute: Members)"
                        errorMessage = "Unable to add mail routing contact to on premises Contact.  Manual add required."
                        erroMessageDetail = $isTestErrorDetail
                    }

                    out-logfile -string $isErrorObject

                    $onPremReplaceErrors+=$isErrorObject
                }
            }
            else 
            {
                out-logfile -string "The original Contact had permissions to itself - skipping as it no longer exists."
            }
        }
    }
    else 
    {
        out-logfile -string "No on premises Contact memberships to process."    
    }

    

    

    out-logfile -string ("Starting on premises reject messages from.")

    if ($allContactsReject.Count -gt 0)
    {
        foreach ($member in $allContactsReject)
        {  
            $isTestError="No" #Reset error test.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-logfile -string ("Processing member = "+$member.canonicalName)
            out-logfile -string ("Routing contact DN = "+$routingContactConfiguration.distinguishedName)
            out-logfile -string ("Attribute Operation = "+$onPremUnAuthOrig)

            if ($member.distinguishedname -ne $originalContactConfiguration.distinguishedname)
            {
                try{
                    $isTestError=start-replaceOnPrem -routingContact $routingContactConfiguration -attributeOperation $onPremUnAuthOrig -canonicalObject $member -adCredential $activeDirectoryCredential -globalCatalogServer $globalCatalogServer -errorAction STOP
                }
                catch{
                    out-logfile -string $_
                    $isTestError="Yes"
                }

                if ($isTestError -eq "Yes")
                {
                    out-logfile -string "Error adding routing contact to on premises resource."

                    $isErrorObject = new-Object psObject -property @{
                        distinguishedName = $member.distinguishedName
                        canonicalDomainName = $member.canonicalDomainName
                        canonicalName=$member.canonicalName
                        attribute = "List RejectMessagesFromSendersOrMembers (ADAttribute: DLMemRejectPerms)"
                        errorMessage = "Unable to add mail routing contact to on premises Contact.  Manual add required."
                        erroMessageDetail = $isTestErrorDetail
                    }

                    out-logfile -string $isErrorObject

                    $onPremReplaceErrors+=$isErrorObject
                }
            }
            else
            {
                out-logfile -string "The original Contact had permissions to itself - skipping as it no longer exists."
            }
        }
    }
    else 
    {
        out-logfile -string "No on premises reject permission to evaluate."    
    }

    

    

    out-logfile -string ("Starting on premises accept messages from.")

    if ($allContactsAccept.Count -gt 0)
    {
        foreach ($member in $allContactsAccept)
        {  
            $isTestError="No" #Reset test 

            out-logfile -string ("Processing member = "+$member.canonicalName)
            out-logfile -string ("Routing contact DN = "+$routingContactConfiguration.distinguishedName)
            out-logfile -string ("Attribute Operation = "+$onPremAuthOrig)

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            if ($member.distinguishedName -ne $originalContactConfiguration.distinguishedname)
            {
                try{
                    $isTestError=start-replaceOnPrem -routingContact $routingContactConfiguration -attributeOperation $onPremAuthOrig -canonicalObject $member -adCredential $activeDirectoryCredential -globalCatalogServer $globalCatalogServer -errorAction STOP
                }
                catch{
                    out-logfile -string $_
                    $isTestError="Yes"
                }

                if ($isTestError -eq "Yes")
                {
                    out-logfile -string "Error adding routing contact to on premises resource."

                    $isErrorObject = new-Object psObject -property @{
                        distinguishedName = $member.distinguishedName
                        canonicalDomainName = $member.canonicalDomainName
                        canonicalName=$member.canonicalName
                        attribute = "List AcceptMessagesOnlyFromSendersorMembers (ADAttribute: DLMemSubmitPerms)"
                        errorMessage = "Unable to add mail routing contact to on premises Contact.  Manual add required."
                        erroMessageDetail = $isTestErrorDetail
                    }

                    out-logfile -string $isErrorObject

                    $onPremReplaceErrors+=$isErrorObject
                }
            }
            else 
            {
                out-logfile -string "The original Contact had permissions to itself - skipping as it no longer exists."
            }
        }
    }
    else 
    {
        out-logfile -string "No on premsies accept permissions to evaluate."    
    }

    

    

    out-logfile -string ("Starting on premises co managed by BL.")

    if ($allContactsCoManagedByBL.Count -gt 0)
    {
        foreach ($member in $allContactsCoManagedByBL)
        {  
            $isTestError="No" #Reset error tracking.

            out-logfile -string ("Processing member = "+$member.canonicalName)
            out-logfile -string ("Routing contact DN = "+$routingContactConfiguration.distinguishedName)
            out-logfile -string ("Attribute Operation = "+$onPremMSExchCoManagedByLink)

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            if ($member.distinguishedName -ne $originalContactConfiguration.distinguishedname)
            {
                try{
                    $isTestError=start-replaceOnPrem -routingContact $routingContactConfiguration -attributeOperation $onPremMSExchCoManagedByLink -canonicalObject $member -adCredential $activeDirectoryCredential -globalCatalogServer $globalCatalogServer -errorAction STOP
                }
                catch{
                    out-logfile -string $_
                    $isTestError="Yes"
                }

                if ($isTestError -eq "Yes")
                {
                    out-logfile -string "Error adding routing contact to on premises resource."

                    $isErrorObject = new-Object psObject -property @{
                        distinguishedName = $member.distinguishedName
                        canonicalDomainName = $member.canonicalDomainName
                        canonicalName=$member.canonicalName
                        attribute = "List ManagedBy (ADAttribute: MSExchCoManagedBy)"
                        errorMessage = "Unable to add mail routing contact to on premises Contact.  Manual add required."
                        erroMessageDetail = $isTestErrorDetail
                    }

                    out-logfile -string $isErrorObject

                    $onPremReplaceErrors+=$isErrorObject
                }
            }
            else 
            {
                out-logfile -string "The original Contact was a co-manager of itself."
            }
        }
    }
    else 
    {
        out-logfile -string "No on premsies accept permissions to evaluate."    
    }

    

    


    out-logfile -string ("Starting on premises bypass moderation.")

    if ($allContactsBypassModeration.Count -gt 0)
    {
        foreach ($member in $allContactsBypassModeration)
        {  
            $isTestError="No" #Reset error tracking.

            out-logfile -string ("Processing member = "+$member.canonicalName)
            out-logfile -string ("Routing contact DN = "+$routingContactConfiguration.distinguishedName)
            out-logfile -string ("Attribute Operation = "+$onPremmsExchBypassModerationLink)

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            if ($member.distinguishedname -ne $originalContactConfiguration.distinguishedName)
            {
                try{
                    $isTestError=start-replaceOnPrem -routingContact $routingContactConfiguration -attributeOperation $onPremmsExchBypassModerationLink -canonicalObject $member -adCredential $activeDirectoryCredential -globalCatalogServer $globalCatalogServer -errorAction STOP
                }
                catch{
                    out-logfile -string $_
                    $isTestError="Yes"
                }

                if ($isTestError -eq "Yes")
                {
                    out-logfile -string "Error adding routing contact to on premises resource."

                    $isErrorObject = new-Object psObject -property @{
                        distinguishedName = $member.distinguishedName
                        canonicalDomainName = $member.canonicalDomainName
                        canonicalName=$member.canonicalName
                        attribute = "List BypassModerationFromSendersOrMembers (ADAttribute: msExchBypassModerationFromDLMembers)"
                        errorMessage = "Unable to add mail routing contact to on premises Contact.  Manual add required."
                        erroMessageDetail = $isTestErrorDetail
                    }

                    out-logfile -string $isErrorObject

                    $onPremReplaceErrors+=$isErrorObject
                }
            }
            else 
            {
                out-logfile -string "The original Contact had permissions to itself - skipping as it no longer exists."
            }
        }
    }
    else 
    {
        out-logfile -string "No on premsies accept permissions to evaluate."    
    }

    

    
    
    out-logfile -string ("Starting on premises grant send on behalf to.")

    if ($allContactsGrantSendOnBehalfTo.Count -gt 0)
    {
        foreach ($member in $allContactsGrantSendOnBehalfTo)
        {  
            $isTestError="No" #Reset error tracking

            out-logfile -string ("Processing member = "+$member.canonicalName)
            out-logfile -string ("Routing contact DN = "+$routingContactConfiguration.distinguishedName)
            out-logfile -string ("Attribute Operation = "+$onPremPublicDelegate)

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            if ($member.distinguishedname -ne $originalContactConfiguration.distinguishedname)
            {
                try{
                    $isTestError=start-replaceOnPrem -routingContact $routingContactConfiguration -attributeOperation $onPremPublicDelegate -canonicalObject $member -adCredential $activeDirectoryCredential -globalCatalogServer $globalCatalogServer -errorAction STOP
                }
                catch{
                    out-logfile -string $_
                    $isTestError="Yes"
                }

                if ($isTestError -eq "Yes")
                {
                    out-logfile -string "Error adding routing contact to on premises resource."

                    $isErrorObject = new-Object psObject -property @{
                        distinguishedName = $member.distinguishedName
                        canonicalDomainName = $member.canonicalDomainName
                        canonicalName=$member.canonicalName
                        attribute = "List GrantSendOnBehalfTo (ADAttribute: PublicDelegates)"
                        errorMessage = "Unable to add mail routing contact to on premises Contact.  Manual add required."
                        erroMessageDetail = $isTestErrorDetail
                    }

                    out-logfile -string $isErrorObject

                    $onPremReplaceErrors+=$isErrorObject
                }
            }
            else 
            {
                out-logfile -string "The original Contact had permissions to itself - skipping as it no longer exists."
            }
        }
    }
    else 
    {
        out-logfile -string "No on premsies grant send on behalf to evaluate."    
    }

    

    

    #Managed by is a unique animal.
    #Managed by is represented by the single valued AD attribute and the multi-evalued exchange attribute.
    #From an exchange standpoint - as long as the member is in one of them it works.
    #We will use the multi-valued attriute so we can recycle the same code.

    out-logfile -string ("Starting on premises managed by.")

    if ($allContactsManagedBy.Count -gt 0)
    {
        foreach ($member in $allContactsManagedBy)
        {  
            $isTestError="No" #Reset error tracking.

            out-logfile -string ("Processing member = "+$member.canonicalName)
            out-logfile -string ("Routing contact DN = "+$routingContactConfiguration.distinguishedName)
            out-logfile -string ("Attribute Operation = "+$onPremMSExchCoManagedByLink)

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            if ($member.distinguishedname -ne $originalContactConfiguration.distinguishedname)
            {
                #More than Contacts can have managed by set.
                #If the object is NOT a Contact - then we should skip it.

                if ($member.objectClass -eq "Contact")
                {
                    out-logfile -string "Object class is Contact - proceed."          

                    try{
                        $isTestError=start-replaceOnPrem -routingContact $routingContactConfiguration -attributeOperation $onPremMSExchCoManagedByLink -canonicalObject $member -adCredential $activeDirectoryCredential -globalCatalogServer $globalCatalogServer -errorAction STOP
                    }
                    catch{
                        out-logfile -string $_
                        $isTestError="Yes"
                    }

                    if ($isTestError -eq "Yes")
                    {
                        out-logfile -string "Error adding routing contact to on premises resource."

                        $isErrorObject = new-Object psObject -property @{
                            distinguishedName = $member.distinguishedName
                            canonicalDomainName = $member.canonicalDomainName
                            canonicalName=$member.canonicalName
                            attribute = "List ManagedBy (ADAttribute: managedBy)"
                            errorMessage = "Unable to add mail routing contact to on premises Contact.  Manual add required."
                            erroMessageDetail = $isTestErrorDetail
                        }

                        out-logfile -string $isErrorObject

                        $onPremReplaceErrors+=$isErrorObject
                    }
                }
                else 
                {
                    out-logfile -string "Other objects than Contacts have this Contact as a manager.  Not processing the routing contact change as manager."
                    out-logfile -string "Automatically setting preserve Contact as to not break permissions on objects."    

                    $retainOriginalContact = $TRUE

                    out-logfile -string ("Retain Original Contact: "+$retainOriginalContact)
                }
            }
            else 
            {
                out-logfile -string "The original Contact had permissions to itself - skipping as it no longer exists."
            }
        }
    }
    else 
    {
        out-logfile -string "No on premsies grant send on behalf to evaluate."    
    }

    

    

    #Forwarding address is a single value replacemet.
    #Created separate function for single values and have called that function here.

    out-logfile -string ("Starting on premises forwarding.")

    if ($allUsersForwardingAddress.Count -gt 0)
    {
        foreach ($member in $allUsersForwardingAddress)
        { 
            $isTestError="No" #Reset error tracking.

            out-logfile -string ("Processing member = "+$member.canonicalName)
            out-logfile -string ("Routing contact DN = "+$routingContactConfiguration.distinguishedName)
            out-logfile -string ("Attribute Operation = "+$onPremAltRecipient)

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-replaceOnPremSV -routingContact $routingContactConfiguration -attributeOperation $onPremAltRecipient -canonicalObject $member -adCredential $activeDirectoryCredential -globalCatalogServer $globalCatalogServer -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to on premises resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    canonicalDomainName = $member.canonicalDomainName
                    canonicalName=$member.canonicalName
                    attribute = "Mailbox Attribute Forwarding Address (ADAttribute: forwardingAddress)"
                    errorMessage = "Unable to add mail routing contact to on premises mailbox object.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $onPremReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-logfile -string "No on premsies grant send on behalf to evaluate."    
    }

    $forLoopCounter=0 #Resetting loop counter now that we're switching to cloud operations.

    out-logfile -string "Processing Office 365 Accept Messages From"

    if ($allOffice365Accept.count -gt 0)
    {
        foreach ($member in $allOffice365Accept)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365UnifiedAccept -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List AcceptMessagesOnlyFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Contacts with accept permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Reject Messages From"

    if ($allOffice365Reject.count -gt 0)
    {
        foreach ($member in $allOffice365Reject)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365UnifiedReject -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List RejectMessagesFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Contacts with reject permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Bypass Moderation From Users"

    if ($allOffice365BypassModeration.count -gt 0)
    {
        foreach ($member in $allOffice365BypassModeration)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365BypassModerationusers -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List BypassModerationFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Contacts with bypass moderation permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Grant Send On Behalf To Users"

    if ($allOffice365GrantSendOnBehalfTo.count -gt 0)
    {
        foreach ($member in $allOffice365GrantSendOnBehalfTo)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365GrantSendOnBehalfTo -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List GrantSendOnBehalfTo"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Contacts with grant send on behalf to permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Managed By"

    if ($allOffice365ManagedBy.count -gt 0)
    {
        foreach ($member in $allOffice365ManagedBy)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365ManagedBy -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List ManagedBy"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 managed by permissions."    
    }

    

    <#

    #Start the process of updating any dynamic Contacts.

    $forLoopCounter=0 #Resetting loop counter now that we're switching to cloud operations.

    out-logfile -string "Processing Office 365 Dynamic Accept Messages From"

    if ($allOffice365DynamicAccept.count -gt 0)
    {
        foreach ($member in $allOffice365DynamicAccept)
        {
            $isTestError="No"

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365Dynamic -office365Attribute $office365AcceptMessagesFrom -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to Office 365 Dynamic DL resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List AcceptMessagesFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Dynamic Contacts with accept permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Dynamic Reject Messages From"

    if ($allOffice365DynamicReject.count -gt 0)
    {
        foreach ($member in $allOffice365DynamicReject)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365Dynamic -office365Attribute $office365RejectMessagesFrom -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to Office 365 Dynamic DL resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List RejectMessagesFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Dynamic Contacts with reject permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Dynamic Bypass Moderation From Users"

    if ($allOffice365DynamicBypassModeration.count -gt 0)
    {
        foreach ($member in $allOffice365DynamicBypassModeration)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365Dynamic -office365Attribute $office365BypassModerationusers -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to Office 365 Dynamic DL resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List BypassModerationFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Dynamic Contacts with bypass moderation permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Dynamic Grant Send On Behalf To Users"

    if ($allOffice365DynamicGrantSendOnBehalfTo.count -gt 0)
    {
        foreach ($member in $allOffice365DynamicGrantSendOnBehalfTo)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365Dynamic -office365Attribute $office365GrantSendOnBehalfTo -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to Office 365 Dynamic DL resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List GrantSendOnBehalfTo"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Dynamic Contacts with grant send on behalf to permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Dynamic Managed By"

    if ($allOffice365DynamicManagedBy.count -gt 0)
    {
        foreach ($member in $allOffice365DynamicManagedBy)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365Dynamic -office365Attribute $office365ManagedBy -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to Office 365 Dynamic DL resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List ManagedBy"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 Dynamic managed by permissions."    
    }

    #>

    <#

    #Start the process of updating the unified Contact dependencies.

    out-logfile -string "Processing Office 365 Unified Accept From"

    if ($allOffice365UniversalAccept.count -gt 0)
    {
        foreach ($member in $allOffice365UniversalAccept)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365Unified -office365Attribute $office365UnifiedAccept -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to Office 365 Universal Modern DL resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List AcceptMessagesOnlyFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 accept from permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Unified Reject From"

    if ($allOffice365UniversalReject.count -gt 0)
    {
        foreach ($member in $allOffice365UniversalReject)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365Unified -office365Attribute $office365UnifiedReject -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to Office 365 Universal Modern DL resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List RejectMessagesFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 reject from permissions."    
    }

    

    

    out-logfile -string "Processing Office 365 Grant Send On Behalf To"

    if ($allOffice365UniversalGrantSendOnBehalfTo.count -gt 0)
    {
        foreach ($member in $allOffice365UniversalGrantSendOnBehalfTo)
        {
            $isTestError="No" #Reset error tracking.

            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            try{
                $isTestError=start-ReplaceOffice365Unified -office365Attribute $office365GrantSendOnBehalfTo -office365Member $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding routing contact to Office 365 Universal Modern DL resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List GrantSendOnBehalfTo"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 grant send on behalf to permissions."    
    }

    #>

    

    #Process any Contact memberships to the service.

    out-logfile -string ("Adding migrated Contact to any cloud only Contacts.")

    if ($allOffice365MemberOf.count -gt 0)
    {
        out-logfile -string "Adding cloud only Contact member."

        foreach ($member in $allOffice365MemberOf )
        {
            $isTestError="No" #Reset error tracking.
            if ($forLoopCounter -eq $forLoopTrigger)
            {
                start-sleepProgress -sleepString "Throttling for 5 seconds...." -sleepSeconds 5
                $forLoopCounter = 0
            }
            else 
            {
                $forLoopCounter++    
            }

            out-logfile -string ("Processing Contact = "+$member.primarySMTPAddress)
            try {
                $isTestError=start-replaceOffice365Members -office365Contact $member -contactSMTPAddress $contactSMTPAddress -errorAction STOP
            }
            catch {
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated list to Office 365 List."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "List Membership"
                    errorMessage = "Unable to add the migrated list to Office 365 Contact.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-logfile -string "No cloud only Contacts had the migrated Contact as a member."
    }   
    
    if ($allowNonSyncedContact -eq $FALSE)
    {
        out-logFile -string "Start replacing Office 365 permissions."

        try 
        {
            set-Office365DLPermissions -allSendAs $allOffice365SendAsAccess -allFullMailboxAccess $allOffice365FullMailboxAccess -allFolderPermissions $allOffice365MailboxFolderPermissions -allOnPremSendAs $allObjectsSendAsAccessNormalized -originalContactPrimarySMTPAddress $contactSMTPAddress -errorAction STOP
        }
        catch 
        {
            out-logfile -string "Unable to set office 365 send as or full mailbox access permissions."
            out-logfile -string $_
            $isTestErrorDetail=$_

            $isErrorObject = new-Object psObject -property @{
                permissionIdentity = "ALL"
                attribute = "Send As / Full Mailbox Access / Mailbox Folder Permissions"
                errorMessage = "Unable to call function to reset send as, full mailbox access, and mailbox folder permissions in Office 365."
                erroMessageDetail = $isTestErrorDetail
            }

            out-logfile -string $isErrorObject

            $global:office365ReplacePermissionsErrors+=$isErrorObject
        }
    }    

    if ($enableHybridMailflow -eq $TRUE)
    {
        #The first step is to upgrade the contact to a full mail contact and remove the target address from proxy addresses.

        $isTestError="No"

        out-logfile -string "The administrator has enabled hybrid mail flow."

        try{
            $isTestError=Enable-MailRoutingContact -globalCatalogServer $globalCatalogServer -routingContactConfig $routingContactConfiguration
        }
        catch{
            out-logfile -string $_
            $isTestError="Yes"
            $errorMessageDetail=$_
        }

        if ($isTestError -eq "Yes")
        {
            $isErrorObject = new-Object psObject -property @{
                errorMessage = "Unable to enable the mail routing contact as a full recipient.  Manually enable the mail routing contact."
                errorMessaegDetail = $errorMessageDetail
            }

            out-logfile -string $isErrorObject

            $generalErrors+=$isErrorObject
        }

        

        #The mail contact has been created and upgrade.  Now we need to capture the updated configuration.

        try{
            $routingContactConfiguration = Get-ADObjectConfiguration -dn $tempDN -globalCatalogServer $globalCatalogWithPort -parameterSet $ContactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 
        }
        catch{
            out-logfile -string $_ -isError:$TRUE
        }

        out-logfile -string $routingContactConfiguration
        out-xmlFile -itemToExport $routingContactConfiguration -itemNameTOExport $routingContactXML+$global:unDoStatus

        #The routing contact configuration has been updated and retained.
        #Now create the dynamic Contact.  This gives us our address book object and our proxy addressed object that cannot collide with the previous object migrated.

        out-logfile -string "Enabling the dynamic Contact to complete the mail routing scenario."

        try{
            $isTestError="No"

            #It is possible that we may need to support a list that is missing attributes.
            #The enable mail dynamic has a retry flag - which is designed to create the DL post migration if necessary.
            #We're going to overload this here - if any of the attributes necessary are set to NULL - then pass in the O365 config and the retry flag.
            #This is what the enable post migration does - bases this off the O365 object.

            if ( ($originalContactConfiguration.name -eq $NULL) -or ($originalContactConfiguration.mailNickName -eq $NULL) -or ($originalContactConfiguration.mail -eq $NULL) -or ($originalContactConfiguration.displayName -eq $NULL) )
            {
                out-logfile -string "Using Office 365 attributes for the mail dynamic Contact."
                $isTestError=Enable-MailDyamicContact -globalCatalogServer $globalCatalogServer -originalContactConfiguration $office365ContactConfiguration -routingContactConfig $routingContactConfiguration -isRetry:$TRUE
            }
            else
            {
                out-logfile -string "Using on premises attributes for the mail dynamic Contact."
                $isTestError=Enable-MailDyamicContact -globalCatalogServer $globalCatalogServer -originalContactConfiguration $originalContactConfiguration -routingContactConfig $routingContactConfiguration
            }
        }
        catch{
            out-logfile -string $_
            $isTestError="Yes"
        }

        if ($isTestError -eq "Yes")
        {
            $isErrorObject = new-Object psObject -property @{
                errorMessage = "Unable to create the mail dynamic Contact to service hybrid mail routing.  Manually create the dynamic Contact."
                erroMessageDetail = $isTestErrorDetail
            }

            out-logfile -string $isErrorObject

            $generalErrors+=$isErrorObject
        }

        [boolean]$stopLoop=$FALSE
        [int]$loopCounter=0

        do {
            try{
                $routingDynamicContactConfig = $originalContactConfiguration = Get-ADObjectConfiguration -contactSMTPAddress $contactSMTPAddress -globalCatalogServer $globalCatalogWithPort -parameterSet $ContactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential

                $stopLoop = $TRUE
            }
            catch{
                if($loopCounter -gt 10)
                {
                    out-logfile -string "Unable to obtain the routing Contact after multiple tries."

                    $isErrorObject = new-Object psObject -property @{
                        errorMessage = "Unable to obtain the routing Contact after multiple tries."
                        erroMessageDetail = $isTestErrorDetail
                    }
        
                    out-logfile -string $isErrorObject
        
                    $generalErrors+=$isErrorObject

                    $stopLoop=$TRUE
                }
                else 
                {
                    out-logfile -string "Unable to obtain the dynamic Contact - retrying..."
                    start-sleepProgress -sleepstring "Unable to obtain the dynamic Contact - retrying..." -sleepSeconds 10

                    $loopCounter = $loopCounter+1
                }
            }
        } while ($stopLoop -eq $FALSE)

        out-logfile -string $routingDynamicContactConfig
        out-xmlfile -itemToExport $routingDynamicContactConfig -itemNameToExport $routingDynamicContactXML
    }


    #At this time the Contact has been migrated.
    #All on premises settings have been reconciled.
    #All cloud settings have been reconciled.
    #If exchange hybrid mail flow was enabled - the routing components were completed.

    #If the administrator has choosen to migrate and request upgrade to Office 365 Contact - trigger the ugprade.

    if ($triggerUpgradeToOffice365Contact -eq $TRUE)
    {
        out-logfile -string "Administrator has choosen to trigger modern Contact upgrade."

        try{
            $isTestError="No"

            $isTestError=start-upgradeToOffice365Contact -contactSMTPAddress $contactSMTPAddress
        }
        catch{
            out-logfile -string $_
            $isTestError="Yes"
        }
    }
    else
    {
        $isTestError="No"
    }


    if ($isTestError -eq "Yes")
    {
        $isErrorObject = new-Object psObject -property @{
            errorMessage = "Unable to trigger upgrade to Office 365 Unified / Modern Contact.  Administrator may need to manually perform the operation."
            erroMessageDetail = $isTestErrorDetail
        }

        out-logfile -string $isErrorObject

        $generalErrors+=$isErrorObject
    }

    #If the administrator has selected to not retain the Contact - remove it.

    if ($retainOriginalContact -eq $FALSE)
    {
        $isTestError="No"

        out-logfile -string "Deleting the original Contact."

        $isTestError=remove-OnPremContact -globalCatalogServer $globalCatalogServer -originalContactConfiguration $originalContactConfigurationUpdated -adCredential $activeDirectoryCredential -errorAction STOP
    }
    else
    {
        $isTestError = "No"
    }


    if ($isTestError -eq "Yes")
    {
        $isErrorObject = new-Object psObject -property @{
            errorMessage = "Uanble to remove the on premises Contact at request of administrator.  Contact may need to be manually removed."
            erroMessageDetail = $isTestErrorDetail
        }

        out-logfile -string $isErrorObject

        $generalErrors+=$isErrorObject
    }

    

   

   #If there are multiple threads and we've reached this point - we're ready to write a status file.

   out-logfile -string "If thread number > 1 - write the status file here."

   if ($global:threadNumber -gt 1)
   {
       out-logfile -string "Thread number is greater than 1."

       try{
           out-statusFile -threadNumber $global:threadNumber -errorAction STOP
       }
       catch{
           out-logfile -string $_
           out-logfile -string "Unable to write status file." -isError:$TRUE
       }
   }
   else 
   {
       out-logfile -string "Thread number is 1 - do not write status at this time."    
   }

   #If there are multiple threads - only replicate the domain controllers and trigger AD connect if all threads have completed their work.

   out-logfile -string "Determine if multiple migration threads are in use..."

   if ($totalThreadCount -eq 0)
   {
       out-logfile -string "Multiple threads are not in use.  Continue functions..."
   }
   else 
   {
       out-logfile -string "Multiple threads are in use - depending on thread number take different actions."
       
       if ($global:threadNumber -eq 1)
       {
           out-logfile -string "This is the master thread responsible for triggering operations."
           out-logfile -string "Search status directory and count files - if file count = number of threads - 1 thread 1 can proceed."

           #Do the following until the count of the files in the directory = number of threads - 1.

           do 
           {
               start-sleepProgress -sleepString "Other threads are pending.  Sleep 5 seconds." -sleepSeconds 5
           } until ((get-statusFileCount) -eq ($totalThreadCount - 1))
       }
       elseif ($global:threadNumber -gt 1)
       {
           out-logfile -string "This is not the master thread responsible for triggering operations."
           out-logfile -string "Search directory and count files.  If the file count = number of threads proceed."

           do 
           {               
               start-sleepProgress -sleepString "Thread 1 is not ready to trigger.  Sleep 5 seconds." -sleepSeconds 5

           } until ((get-statusFileCount) -eq  $totalThreadCount)
       }
   }

   #Replicate domain controllers so that the change is received as soon as possible.()
   
   if ($global:threadNumber -eq 0 -or ($global:threadNumber -eq 1))
   {
       start-sleepProgress -sleepString "Starting sleep before invoking AD replication - 15 seconds." -sleepSeconds 15

       out-logfile -string "Invoking AD replication."

       try {
           invoke-ADReplication -globalCatalogServer $globalCatalogServer -powershellSessionName $ADGlobalCatalogPowershellSessionName -errorAction STOP
       }
       catch {
           out-logfile -string $_
       }
   }

   #Start the process of syncing the deletion to the cloud if the administrator has provided credentials.
   #Note:  If this is not done we are subject to sitting and waiting for it to complete.

   if ($global:threadNumber -eq 0 -or ($global:threadNumber -eq 1))
   {
       if ($useAADConnect -eq $TRUE)
       {
           start-sleepProgress -sleepString "Starting sleep before invoking AD Connect - one minute." -sleepSeconds 60

           out-logfile -string "Invoking AD Connect."

           invoke-ADConnect -powerShellSessionName $aadConnectPowershellSessionName

           start-sleepProgress -sleepString "Starting sleep after invoking AD Connect - one minute." -sleepSeconds 60

       }   
       else 
       {
           out-logfile -string "AD Connect information not specified - allowing ad connect to run on normal cycle and process deletion."    
       }
   }   

   #The single functions have triggered operations.  Other threads may continue.

   if ($global:threadNumber -eq 1)
   {
       out-statusFile -threadNumber $global:threadNumber
   }

    #If this is the main thread - introduce a sleep for 10 seconds - allows the other threads to detect 5 files.
    #Reset the status directory for furture thread dependencies.

   if ($totalThreadCount -gt 0)
   {
        start-sleepProgress -sleepString "Sleep..." -sleepSeconds 10

        try{
        remove-statusFiles -functionThreadNumber $global:threadNumber
        }
        catch{
            out-logfile -string "Unable to remove status files" -isError:$TRUE
        }
   }

    out-logfile -string "Calling function to disconnect all powershell sessions."

    disable-allPowerShellSessions

    Out-LogFile -string "================================================================================"
    Out-LogFile -string "END START-CONTACTMIGRATION"
    Out-LogFile -string "================================================================================"

    if (($global:office365ReplacePermissionsErrors.count -gt 0) -or ($global:postCreateErrors.count -gt 0) -or ($onPremReplaceErrors.count -gt 0) -or ($office365ReplaceErrors.count -gt 0) -or ($global:office365ReplacePermissionsErrors.count -gt 0) -or ($generalErrors.count -gt 0))
    {
        out-logfile -string ""
        out-logfile -string "+++++"
        out-logfile -string "++++++++++"
        out-logfile -string "MIGRATION ERRORS OCCURED - REFER TO LIST BELOW FOR ERRORS"
        out-logfile -string ("Post Create Errors: "+$global:postCreateErrors.count)
        out-logfile -string ("On-Premises Replace Errors :"+$onPremReplaceErrors.count)
        out-logfile -string ("Office 365 Replace Errors: "+$office365ReplaceErrors.count)
        out-logfile -string ("Office 365 Replace Permissions Errors: "+$global:office365ReplacePermissionsErrors.count)
        out-logfile -string ("On Prem Replace Permissions Errors: "+$global:onPremReplacePermissionsErrors.count)
        out-logfile -string ("General Errors: "+$generalErrors.count)
        out-logfile -string "++++++++++"
        out-logfile -string "+++++"
        out-logfile -string ""

        if ($global:postCreateErrors.count -gt 0)
        {
            foreach ($createError in $global:postCreateErrors)
            {
                out-logfile -string "====="
                out-logfile -string "Post Create Errors:"
                out-logfile -string ("Primary Email Address or UPN: " +$CreateError.primarySMTPAddressOrUPN)
                out-logfile -string ("External Directory Object ID: " +$CreateError.externalDirectoryObjectID)
                out-logfile -string ("Name: "+$CreateError.name)
                out-logfile -string ("Alias: "+$CreateError.Alias)
                out-logfile -string ("Attribute in Error: "+$CreateError.attribute)
                out-logfile -string ("Error Message: "+$CreateError.errorMessage)
                out-logfile -string ("Error Message Details: "+$CreateError.errorMessageDetail)
                out-logfile -string "====="
            }
        }

        if ($onPremReplaceErrors.count -gt 0)
        {
            foreach ($onPremReplaceError in $onPremReplaceErrors)
            {
                out-logfile -string "====="
                out-logfile -string "Replace On Premises Errors:"
                out-logfile -string ("Distinguished Name: "+$onPremReplaceError.distinguishedName)
                out-logfile -string ("Canonical Domain Name: "+$onPremReplaceError.canonicalDomainName)
                out-logfile -string ("Canonical Name: "+$onPremReplaceError.canonicalName)
                out-logfile -string ("Attribute in Error: "+$onPremReplaceError.attribute)
                out-logfile -string ("Error Message: "+$onPremReplaceError.errorMessage)
                out-logfile -string ("Error Message Details: "+$onPremReplaceError.errorMessageDetail)
                out-logfile -string "====="
            }
        }

       
        if ($office365ReplaceErrors.count -gt 0)
        {
            foreach ($office365ReplaceError in $office365ReplaceErrors)
            {
                out-logfile -string "====="
                out-logfile -string "Replace Office 365 Errors:"
                out-logfile -string ("Distinguished Name: "+$office365ReplaceError.distinguishedName)
                out-logfile -string ("Primary SMTP Address: "+$office365ReplaceError.primarySMTPAddress)
                out-logfile -string ("Alias: "+$office365ReplaceError.alias)
                out-logfile -string ("Display Name: "+$office365ReplaceError.displayName)
                out-logfile -string ("Attribute in Error: "+$office365ReplaceError.attribute)
                out-logfile -string ("Error Message: "+$office365ReplaceError.errorMessage)
                out-logfile -string ("Error Message Details: "+$office365Replace.errorMessageDetail)
                out-logfile -string "====="
            }
        }
        
        if ($global:office365ReplacePermissionsErrors.count -gt 0)
        {
            foreach ($office365ReplacePermissionsError in $global:office365ReplacePermissionsErrors)
            {
                out-logfile -string "====="
                out-logfile -string "Office 365 Permissions Error: "
                out-logfile -string ("Permission in Error: "+$office365ReplacePermissionsError.permissionidentity)
                out-logfile -string ("Attribute in Error: "+$office365ReplacePermissionsError.attribute)
                out-logfile -string ("Error Message: "+$office365ReplacePermissionsError.errorMessage)
                out-logfile -string ("Error Message Detail: "+$office365ReplacePermissionsError.errorMessageDetail)
                out-logfile -string "====="
            }
        }

        if ($global:onPremReplacePermissionsErrors.count -gt 0)
        {
            foreach ($onPremReplacePermissionsError in $global:office365ReplacePermissionsErrors)
            {
                out-logfile -string "====="
                out-logfile -string "On Prem Permissions Error: "
                out-logfile -string ("Permission in Error: "+$office365ReplacePermissionsError.permissionidentity)
                out-logfile -string ("Attribute in Error: "+$office365ReplacePermissionsError.attribute)
                out-logfile -string ("Error Message: "+$office365ReplacePermissionsError.errorMessage)
                out-logfile -string ("Error Message Detail: "+$office365ReplacePermissionsError.errorMessageDetail)
                out-logfile -string "====="
            }
        }
        
        if ($generalErrors.count -gt 0)
        {
            foreach ($generalError in $generalErrors)
            {
                out-logfile -string "====="
                out-logfile -string "General Errors:"
                out-logfile -string ("Error Message: "+$generalError.errorMessage)
                out-logfile -string ("Error Message Detail: "+$generalError.errorMessageDetail)
                out-logfile -string "====="
            }
        }

        out-logfile -string ""
        out-logfile -string "+++++"
        out-logfile -string "++++++++++"
        out-logfile -string "Errors were encountered in the list creation process requireing administrator review."
        out-logfile -string "Although the migration may have been successful - manual actions may need to be taken to full complete the migration."
        out-logfile -string "++++++++++"
        out-logfile -string "+++++"
        out-logfile -string "" -isError:$TRUE

    }

    #Archive the files into a date time success folder.

    Start-ArchiveFiles -isSuccess:$TRUE -logFolderPath $logFolderPath
}