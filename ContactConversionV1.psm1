
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

    This is the trigger function that begins the process of allowing an administrator to migrate a distribution list from
    on premises to Office 365.

    .DESCRIPTION

    Trigger function.

    .PARAMETER contactSMTPAddress

    *REQUIRED*
    The SMTP address of the distribution list to be migrated.

    .PARAMETER globalCatalogServer

    *REQUIRED*
    A global catalog server in the domain where the contact to be migrated resides.


    .PARAMETER activeDirectoryCredential

    *REQUIRED*
    This is the credential that will be utilized to perform operations against the global catalog server.
    If the contact and all it's dependencies reside in a single domain - a domain administrator is acceptable.
    If the contact and it's dependencies span multiple domains in a forest - enterprise administrator is required.
      
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

    .PARAMETER doNoSyncOU

    *REQUIRED IF RETAIN contact FALSE*
    This is the administrator specified organizational unit that is NOT configured to sync in AD Connect.
    When the administrator specifies to NOT retain the contact the contact is moved to this OU to allow for deletion from Office 365.
    A doNOSyncOU must be specified if the administrator specifies to NOT retain the contact.

    .PARAMETER retainOriginalContact

    *OPTIONAL*
    Allows the administrator to retain the contact - for example if the contact also has on premises security dependencies.
    This triggers a mail disable of the contact resulting in contact deletion from Office 365.
    The name of the contact is randomized with a character ! to ensure no conflict with hybird mail flow - if hybrid mail flow enabled.

    .PARAMETER enableHybridMailFlow

    *OPTIONAL*
    Allows the administrator to decide that they want mail flow from on premises to cloud to work for the migrated contact.
    This involves provisioning a mail contact and a dynamic distribution contact.
    The dynamic distribution contact is intentionally choosen to prevent soft matching of a contact and an undo of the migration.
    This option requires on premises Exchange be specified and configured.

    .PARAMETER contactTypeOverride

    *OPTIONAL*
    This allows the administrator to override the contact type created in the cloud from on premises.
    For example - if the contact was provisioned on premises as security but does not require security rights in Office 365 - the administrator can override to DISTRIBUTION.
    Mandatory types -> SECURITY or DISTRIBUTION

	.OUTPUTS

    Logs all activities and backs up all original data to the log folder directory.
    Moves the distribution contact from on premieses source of authority to office 365 source of authority.

    .EXAMPLE

    Start-DistributionListMigration

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
        [Parameter(Mandatory = $true)]
        [string]$dnNoSyncOU = "NotSet",
        [Parameter(Mandatory = $false)]
        [boolean]$retainOriginalcontact = $TRUE,
        [Parameter(Mandatory = $false)]
        [int]$threadNumberAssigned=0,
        [Parameter(Mandatory = $false)]
        [int]$totalThreadCount=0,
        [Parameter(Mandatory = $FALSE)]
        [boolean]$isMultiMachine=$FALSE,
        [Parameter(Mandatory = $FALSE)]
        [string]$remoteDriveLetter=$NULL,
        [Parameter(Mandatory=$false)]
        [boolean]$allowNonSyncedcontact=$FALSE
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
    [string]$global:staticFolderName="\contactMigration\"
    [string]$global:staticAuditFolderName="\AuditData\"
    [string]$global:importFile=$logFolderPath+$global:staticAuditFolderName
    [int]$global:unDoStatus=0
    [array]$importData=@()
    [string]$importFilePath=$NULL

    #Define variables utilized in the core function that are not defined by parameters.

    [boolean]$useAADConnect=$FALSE #Determines if function will utilize aadConnect during migration.
    [string]$aadConnectPowershellSessionName="AADConnect" #Defines universal name for aadConnect powershell session.
    [string]$ADGlobalCatalogPowershellSessionName="ADGlobalCatalog" #Defines universal name for ADGlobalCatalog powershell session.
    [string]$exchangeOnlinePowershellModuleName="ExchangeOnlineManagement" #Defines the exchage management shell name to test for.
    [string]$activeDirectoryPowershellModuleName="ActiveDirectory" #Defines the active directory shell name to test for.
    [string]$contactConversionPowershellModule="ContactConversionV1"
    [string]$globalCatalogPort=":3268"
    [string]$globalCatalogWithPort=$globalCatalogServer+$globalCatalogPort

    #The variables below are utilized to define working parameter sets.
    #Some variables are assigned to single values - since these will be utilized with functions that query or set information.
    
    [string]$acceptMessagesFromcontactMembers="dlMemSubmitPerms" #Attribute for the allow email members.
    [string]$rejectMessagesFromcontactMembers="dlMemRejectPerms"
    [string]$bypassModerationFromcontact="msExchBypassModerationFromcontactMembersLink"
    [string]$forwardingAddressForcontact="altRecipient"
    [string]$grantSendOnBehalfTocontact="publicDelegates"
    [array]$contactPropertySet = '*'
    [array]$contactPropertySetToClear = @()
    [array]$contactPropertiesToClearModern='authOrig','DisplayName','DisplayNamePrintable',$rejectMessagesFromcontactMembers,$acceptMessagesFromcontactMembers,'extensionAttribute1','extensionAttribute10','extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5','extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','legacyExchangeDN','mail','mailNickName','msExchRecipientDisplayType','msExchRecipientTypeDetails','msExchRemoteRecipientType',$bypassModerationFromcontact,'msExchBypassModerationLink','msExchCoManagedByLink','msExchEnableModeration','msExchExtensionCustomAttribute1','msExchExtensionCustomAttribute2','msExchExtensionCustomAttribute3','msExchExtensionCustomAttribute4','msExchExtensionCustomAttribute5','msExchcontactDepartRestriction','msExchcontactJoinRestriction','msExchHideFromAddressLists','msExchModeratedByLink','msExchModerationFlags','msExchRequireAuthToSendTo','msExchSenderHintTranslations','oofReplyToOriginator','proxyAddresses',$grantSendOnBehalfTocontact,'reportToOriginator','reportToOwner','unAuthOrig','msExchArbitrationMailbox','msExchPoliciesIncluded','msExchUMDtmfMap','msExchVersion','showInAddressBook','msExchAddressBookFlags','msExchBypassAudit','msExchcontactExternalMemberCount','msExchcontactMemberCount','msExchcontactSecurityFlags','msExchLocalizationFlags','msExchMailboxAuditEnable','msExchMailboxAuditLogAgeLimit','msExchMailboxFolderSet','msExchMDBRulesQuota','msExchPoliciesIncluded','msExchProvisioningFlags','msExchRecipientSoftDeletedStatus','msExchRolecontactType','msExchTransportRecipientSettingsFlags','msExchUMDtmfMap','msExchUserAccountControl','msExchVersion'
    [array]$contactPropertiesToClearLegacy='authOrig','DisplayName','DisplayNamePrintable',$rejectMessagesFromcontactMembers,$acceptMessagesFromcontactMembers,'extensionAttribute1','extensionAttribute10','extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5','extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','legacyExchangeDN','mail','mailNickName','msExchRecipientDisplayType','msExchRecipientTypeDetails','msExchRemoteRecipientType',$bypassModerationFromcontact,'msExchBypassModerationLink','msExchCoManagedByLink','msExchEnableModeration','msExchExtensionCustomAttribute1','msExchExtensionCustomAttribute2','msExchExtensionCustomAttribute3','msExchExtensionCustomAttribute4','msExchExtensionCustomAttribute5','msExchcontactDepartRestriction','msExchcontactJoinRestriction','msExchHideFromAddressLists','msExchModeratedByLink','msExchModerationFlags','msExchRequireAuthToSendTo','msExchSenderHintTranslations','oofReplyToOriginator','proxyAddresses',$grantSendOnBehalfTocontact,'reportToOriginator','reportToOwner','unAuthOrig','msExchArbitrationMailbox','msExchPoliciesIncluded','msExchUMDtmfMap','msExchVersion','showInAddressBook','msExchAddressBookFlags','msExchBypassAudit','msExchcontactExternalMemberCount','msExchcontactMemberCount','msExchLocalizationFlags','msExchMailboxAuditEnable','msExchMailboxAuditLogAgeLimit','msExchMailboxFolderSet','msExchMDBRulesQuota','msExchPoliciesIncluded','msExchProvisioningFlags','msExchRecipientSoftDeletedStatus','msExchRolecontactType','msExchTransportRecipientSettingsFlags','msExchUMDtmfMap','msExchUserAccountControl','msExchVersion'

    #On premises variables for the distribution list to be migrated.

    $originalContactConfiguration=$NULL #This holds the on premises contact configuration for the contact to be migrated.
    $originalContactConfigurationUpdated=$NULL #This holds the on premises contact configuration post the rename operations.
    [array]$exchangecontactMembershipSMTP=@() #Array of contact membership from AD.
    [array]$exchangeRejectMessagesSMTP=@() #Array of members with reject permissions from AD.
    [array]$exchangeAcceptMessagesSMTP=@() #Array of members with accept permissions from AD.
    [array]$exchangeManagedBySMTP=@() #Array of members with manage by rights from AD.
    [array]$exchangeModeratedBySMTP=@() #Array of members  with moderation rights.
    [array]$exchangeBypassModerationSMTP=@() #Array of objects with bypass moderation rights from AD.
    [array]$exchangeGrantSendOnBehalfToSMTP=@()

    #Define XML files to contain backups.

    [string]$originalContactConfigurationADXML = "originalContactConfigurationADXML" #Export XML file of the contact attibutes direct from AD.
    [string]$originalContactConfigurationUpdatedXML = "originalContactConfigurationUpdatedXML"
    [string]$originalContactConfigurationObjectXML = "originalContactConfigurationObjectXML" #Export of the ad attributes after selecting objects (allows for NULL objects to be presented as NULL)
    [string]$office365contactConfigurationXML = "office365contactConfigurationXML"
    [string]$office365contactConfigurationPostMigrationXML = "office365contactConfigurationPostMigrationXML"
    [string]$office365contactMembershipPostMigrationXML = "office365contactMembershipPostMigrationXML"
    [string]$exchangecontactMembershipSMTPXML = "exchangecontactMemberShipSMTPXML"
    [string]$exchangeRejectMessagesSMTPXML = "exchangeRejectMessagesSMTPXML"
    [string]$exchangeAcceptMessagesSMTPXML = "exchangeAcceptMessagesSMTPXML"
    [string]$exchangeManagedBySMTPXML = "exchangeManagedBySMTPXML"
    [string]$exchangeModeratedBySMTPXML = "exchangeModeratedBYSMTPXML"
    [string]$exchangeBypassModerationSMTPXML = "exchangeBypassModerationSMTPXML"
    [string]$exchangeGrantSendOnBehalfToSMTPXML = "exchangeGrantSendOnBehalfToXML"
    [string]$allcontactsMemberOfXML = "allcontactsMemberOfXML"
    [string]$allcontactsRejectXML = "allcontactsRejectXML"
    [string]$allcontactsAcceptXML = "allcontactsAcceptXML"
    [string]$allcontactsBypassModerationXML = "allcontactsBypassModerationXML"
    [string]$allUsersForwardingAddressXML = "allUsersForwardingAddressXML"
    [string]$allcontactsGrantSendOnBehalfToXML = "allcontactsGrantSendOnBehalfToXML"
    [string]$allcontactsManagedByXML = "allcontactsManagedByXML"
    [string]$allOffice365MemberOfXML="allOffice365MemberOfXML"
    [string]$allOffice365AcceptXML="allOffice365AcceptXML"
    [string]$allOffice365RejectXML="allOffice365RejectXML"
    [string]$allOffice365BypassModerationXML="allOffice365BypassModerationXML"
    [string]$allOffice365GrantSendOnBehalfToXML="allOffice365GrantSentOnBehalfToXML"
    [string]$allOffice365ManagedByXML="allOffice365ManagedByXML"
    [string]$allOffice365ForwardingAddressXML="allOffice365ForwardingAddressXML"
    [string]$allcontactsCoManagedByXML="allcontactsCoManagedByXML"

    #The following variables hold information regarding other contacts in the environment that have dependnecies on the contact to be migrated.

    [array]$allcontactsMemberOf=$NULL #Complete AD information for all contacts the migrated contact is a member of.
    [array]$allcontactsReject=$NULL #Complete AD inforomation for all contacts that the migrated contact has reject mesages from.
    [array]$allcontactsAccept=$NULL #Complete AD information for all contacts that the migrated contact has accept messages from.
    [array]$allcontactsBypassModeration=$NULL #Complete AD information for all contacts that the migrated contact has bypass moderations.
    [array]$allUsersForwardingAddress=$NULL #All users on premsies that have this contact as a forwarding DN.
    [array]$allcontactsGrantSendOnBehalfTo=$NULL #All dependencies on premsies that have grant send on behalf to.
    [array]$allcontactsManagedBy=$NULL
    [array]$allcontactsCoManagedByBL=$NULL

    #The following variables hold information regarding Office 365 objects that have dependencies on the migrated contact.

    #The following are for standard distribution contacts.

    [array]$allOffice365MemberOf=$NULL
    [array]$allOffice365Accept=$NULL
    [array]$allOffice365Reject=$NULL
    [array]$allOffice365BypassModeration=$NULL
    [array]$allOffice365ManagedBy=$NULL


    #These are for other mail enabled objects.

    [array]$allOffice365ForwardingAddress=$NULL
  
    #The following are the cloud parameters we query for to look for dependencies.

    [string]$office365AcceptMessagesFrom="AcceptMessagesOnlyFrom"
    [string]$office365BypassModerationFrom="BypassModerationFrom"
    [string]$office365ManagedBy="ManagedBy"
    [string]$office365GrantSendOnBehalfTo="GrantSendOnBehalfTo"
    [string]$office365Members="Members"
    [string]$office365RejectMessagesFrom="RejectMessagesFrom"
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

    #Cloud variables for the distribution list to be migrated.

    $office365contactConfiguration = $NULL #This holds the office 365 contact configuration for the contact to be migrated.
    $office365contactConfigurationPostMigration = $NULL

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

    #Log start of contact migration to the log file.

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

    $exchangeOnlineEnvironmentName=remove-stringSpace -stringToFix $exchangeOnlineEnvironmentName

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
    Out-LogFile -string ("contactSMTPAddress = "+$contactSMTPAddress)
    out-logfile -string ("contact SMTP Address Length = "+$contactSMTPAddress.length.tostring())
    out-logfile -string ("Spaces Removed contact SMTP Address: "+$contactSMTPAddress)
    out-logfile -string ("contact SMTP Address Length = "+$contactSMTPAddress.length.toString())
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

    out-logfile -string ("OU that does not sync to Office 365 = "+$dnNoSyncOU)
    out-logfile -string ("Will the original contact be retained as part of migration = "+$retainOriginalContact)
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
    out-logFile -string ("Static property for accept messages from members = "+$acceptMessagesFromcontactMembers)
    out-logFile -string ("Static property for accept messages from members = "+$rejectMessagesFromcontactMembers)
    Out-LogFile -string ("contact Properties to collect = ")

    foreach ($contactProperty in $contactPropertySet)
    {
        Out-LogFile -string $contactProperty
    }

    Out-LogFile -string ("contact property set to be cleared legacy = ")

    foreach ($contactProperty in $contactPropertiesToClearLegacy)
    {
        Out-LogFile -string $contactProperty
    }

    Out-LogFile -string ("contact property set to be cleared modern = ")

    foreach ($contactProperty in $contactPropertiesToClearModern)
    {
        Out-LogFile -string $contactProperty
    }

    Out-LogFile -string ("Exchange on prem contact active directory configuration XML = "+$originalContactConfigurationADXML)
    Out-LogFile -string ("Exchange on prem contact object configuration XML = "+$originalContactConfigurationObjectXML)
    Out-LogFile -string ("Office 365 contact configuration XML = "+$office365contactConfigurationXML)
    Out-LogFile -string ("Exchange contact members XML Name - "+$exchangecontactMembershipSMTPXML)
    Out-LogFile -string ("Exchange Reject members XML Name - "+$exchangeRejectMessagesSMTPXML)
    Out-LogFile -string ("Exchange Accept members XML Name - "+$exchangeAcceptMessagesSMTPXML)
    Out-LogFile -string ("Exchange ManagedBY members XML Name - "+$exchangeManagedBySMTPXML)
    Out-LogFile -string ("Exchange ModeratedBY members XML Name - "+$exchangeModeratedBySMTPXML)
    Out-LogFile -string ("Exchange BypassModeration members XML Name - "+$exchangeBypassModerationSMTPXML)
    out-logfile -string ("Exchange GrantSendOnBehalfTo members XML name - "+$exchangeGrantSendOnBehalfToSMTPXML)
    Out-LogFile -string ("All contact members XML Name - "+$allcontactsMemberOfXML)
    Out-LogFile -string ("All Reject members XML Name - "+$allcontactsRejectXML)
    Out-LogFile -string ("All Accept members XML Name - "+$allcontactsAcceptXML)
    Out-Logfile -string ("All Co Managed By BL XML - "+$allcontactsCoManagedByXML)
    Out-LogFile -string ("All BypassModeration members XML Name - "+$allcontactsBypassModerationXML)
    out-logfile -string ("All Users Forwarding Address members XML Name - "+$allUsersForwardingAddressXML)
    out-logfile -string ("All contacts Grand Send On Behalf To XML Name - "+$allcontactsGrantSendOnBehalfToXML)
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

    #Validate that an OU was specified <if> retain contact is not set to true.

    Out-LogFile -string "Validating that if retain original contact is false a non-sync OU is specified."

    if (($retainOriginalContact -eq $FALSE) -and ($dnNoSyncOU -eq "NotSet"))
    {
        out-LogFile -string "A no SYNC OU is required if retain original contact is false." -isError:$TRUE
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
        $contactPropertySetToClear=$contactPropertiesToClearModern
    }
    else 
    {
        out-logfile -string "Legacy exchange versions detected - using legacy parameters"
        $contactPropertySetToClear = $contactPropertiesToClearLegacy   
    }

    Out-LogFile -string ("contact property set to be cleared after schema evaluation = ")

    foreach ($contactProperty in $contactPropertySetToClear)
    {
        Out-LogFile -string $contactProperty
    }

    # EXIT #Debug Exit

    #If exchange server information specified - create the on premises powershell session.

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "ESTABLISH POWERSHELL SESSIONS"
    Out-LogFile -string "********************************************************************************"

   #Test to determine if the exchange online powershell module is installed.
   #The exchange online session has to be established first or the commancontactet set from on premises fails.

   Out-LogFile -string "Calling Test-PowerShellModule to validate the Exchange Module is installed."

   Test-PowershellModule -powershellModuleName $exchangeOnlinePowershellModuleName -powershellVersionTest:$TRUE

   Out-LogFile -string "Calling Test-PowerShellModule to validate the Active Directory is installed."

   Test-PowershellModule -powershellModuleName $activeDirectoryPowershellModuleName

   out-logfile -string "Calling Test-PowershellModule to validate the contact Conversion Module version installed."

   Test-PowershellModule -powershellModuleName $contactConversionPowershellModule -powershellVersionTest:$TRUE

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

   #exit #debug exit

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

    #EXIT #Debug Exit

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN GET ORIGINAL contact CONFIGURATION LOCAL AND CLOUD"
    Out-LogFile -string "********************************************************************************"

    #At this point we are ready to capture the original contact configuration.  We'll use the ad provider to gather this information.

    Out-LogFile -string "Getting the original contact Configuration"

    try
    {
        $originalContactConfiguration = Get-ADObjectConfiguration -contactSMTPAddress $contactSMTPAddress -globalCatalogServer $globalCatalogWithPort -parameterSet $contactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential
    }
    catch
    {
        out-logfile -string $_ -isError:$TRUE
    }
    
    Out-LogFile -string "Log original contact configuration."
    out-logFile -string $originalContactConfiguration

    Out-LogFile -string "Create an XML file backup of the on premises contact Configuration"

    Out-XMLFile -itemToExport $originalContactConfiguration -itemNameToExport $originalContactConfigurationADXML

    #exit #Debug Exit

    Out-LogFile -string "Capture the original office 365 distribution list information."

    if ($allowNonSyncedcontact -eq $FALSE)
    {
        try 
        {
            $office365contactConfiguration=Get-O365contactConfiguration -contactSMTPAddress $contactSMTPAddress -errorAction STOP
        }
        catch 
        {
            out-logFile -string $_ -isError:$TRUE
        }
    }
    else 
    {
        $office365contactConfiguration="DistributionListIsNonSynced"
    }

    Out-LogFile -string $office365contactConfiguration

    Out-LogFile -string "Create an XML file backup of the office 365 contact configuration."

    Out-XMLFile -itemToExport $office365contactConfiguration -itemNameToExport $office365contactConfigurationXML

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END GET ORIGINAL contact CONFIGURATION LOCAL AND CLOUD"
    Out-LogFile -string "********************************************************************************"

    if ($allowNonSyncedcontact -eq $FALSE)
    {
        Out-LogFile -string "Perform a safety check to ensure that the distribution list is directory sync."

        try 
        {
            Invoke-Office365SafetyCheck -o365contactconfiguration $office365contactConfiguration -errorAction STOP
        }
        catch 
        {
            out-logFile -string $_ -isError:$TRUE
        }
    }
    else 
    {
        out-logfile -string "The administrator is attempting to migrate a non-synced contact.  Office 365 check skipped."
        
        try 
        {
            test-nonSynccontact -originalContactConfiguration $originalContactConfiguration -errorAction STOP    
        }
        catch 
        {
            out-logfile -string $_ -isError:$TRUE   
        }
    }

    
    #At this time we have the contact configuration on both sides and have checked to ensure it is dir synced.
    #Membership of attributes is via DN - these need to be normalized to SMTP addresses in order to find users in Office 365.

    #Start with contact membership and normallize.

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
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -isMember:$TRUE -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "Distribution List Membership (ADAttribute: Members)"
                        errorMessage = $normalizedTest.isErrorMessage
                    }

                    out-logfile -string $isErrorObject

                    $preCreateErrors+=$isErrorObject
                }
                else 
                {
                    $exchangecontactMembershipSMTP+=$normalizedTest
                }
                
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangecontactMembershipSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the contact:"
        
        out-logfile -string $exchangecontactMembershipSMTP
    }
    else 
    {
        out-logFile -string "The distribution contact has no members."    
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
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

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

    Out-LogFile -string "REJECT contactS"

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
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "RejectMessagesFromcontactMembers (ADAttribute dlMemRejectPerms)"
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
        out-logfile -string "The contact has reject messages members."
        Out-logFile -string $exchangeRejectMessagesSMTP
    }
    else 
    {
        out-logfile "The contact to be migrated has no reject messages from members."    
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
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

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

    Out-LogFile -string "ACCEPT contactS"

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
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "AcceptMessagesOnlyFromcontactMembers (ADAttribute: dlMemSubmitPerms)"
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
        out-logFile -string "This contact has no accept message from restrictions."    
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
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

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
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

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
        #First scan is to ensure that any of the contacts listed on the managed by objects are still security.
        #It is possible someone added it to managed by and changed the contact type after.

        foreach ($object in $exchangeManagedBySMTP)
        {
            #If the objec thas a non-null contact type (is a contact) and the value of the contact type matches none of the secuity contact types.
            #The object is a distribution list - no good.

            if (($object.contactType -ne $NULL) -and ($object.contactType -ne "-2147483640") -and ($object.contactType -ne "-2147483646") -and ($object.contactType -ne "-2147483644"))
            {
                $isErrorObject = new-Object psObject -property @{
                    primarySMTPAddressOrUPN = $object.primarySMTPAddressOrUPN
                    externalDirectoryObjectID = $object.externalDirectoryObjectID
                    alias=$normalizedTest.alias
                    name=$normalizedTest.name
                    attribute = "Test ManagedBy For Security Flag"
                    errorMessage = "A contact was found on the owners attribute that is no longer a security contact.  Security contact is required.  Remove contact or change contact type to security."
                    errorMessageDetail = ""
                }

                out-logfile -string $isErrorObject

                $preCreateErrors+=$isErrorObject

                out-logfile -string "A distribution list (not security enabled) was found on managed by."
                out-logfile -string "The contact must be converted to security or removed from managed by."
                out-logfile -string $object.primarySMTPAddressOrUPN
            }

            #The contact is not a distribution list.
            #If the SMTP object of the managedBy object equals the original contact - check to see if an override is found.
            #If an override of distribution is found - this is not OK since security is required.

            elseif (($object.primarySMTPAddressOrUPN -eq $originalContactConfiguration.mail) -and ($contactTypeOverride -eq "Distribution")) 
            {
                out-logfile -string "contact type override detected - contact has managed by permissions."

                #contact type is not NULL / contact type is security value.

                if (($object.contactType -ne $NULL) -and (($object.contactType -eq "-2147483640") -or ($object.contactType -eq "-2147483646" -or ($object.contactType -eq "-2147483644"))))
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $object.primarySMTPAddressOrUPN
                        externalDirectoryObjectID = $object.externalDirectoryObjectID
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "Test ManagedBy For contact Override"
                        errorMessage = "The contact being migrated was found on the Owners attribute.  The administrator has requested migration as Distribution not Security.  To remain an owner the contact must be migrated as Security - remove override or remove owner."
                        errorMessageDetail = ""
                    }

                    out-logfile -string $isErrorObject
    
                    $preCreateErrors+=$isErrorObject
        
                    out-logfile -string "A security contact has managed by rights on the distribution list."
                    out-logfile -string "The administrator has specified to override the contact type."
                    out-logfile -string "The contact override must be removed or the object removed from managedBY."
                    out-logfile -string $object.primarySMTPAddressOrUPN
                }
            }
        }

        Out-LogFile -string "The following objects are members of the managedBY:"
        
        out-logfile -string $exchangeManagedBySMTP
    }
    else 
    {
        out-logfile -string "The contact has no managers."    
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
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

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
        out-logfile "The contact has no moderators."    
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
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

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

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the bypass moderation contacts members DN to Office 365 identifier."

    Out-LogFile -string "Process BYPASS contactS"

    if ($originalContactConfiguration.msExchBypassModerationFromcontactMembersLink -ne $NULL)
    {
        foreach ($DN in $originalContactConfiguration.msExchBypassModerationFromcontactMembersLink)
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
                $normalizedTest = get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName -errorAction STOP -cn "None"

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $isErrorObject = new-Object psObject -property @{
                        primarySMTPAddressOrUPN = $normalizedTest.name
                        externalDirectoryObjectID = $NULL
                        alias=$normalizedTest.alias
                        name=$normalizedTest.name
                        attribute = "BypassModerationFromSendersOrMembers (ADAttribute: msExchBypassModerationFromcontactMembersLink"
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
        out-logfile "The contact has no bypass moderation."    
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
                $normalizedTest=get-normalizedDN -globalCatalogServer $globalCatalogWithPort -DN $DN -adCredential $activeDirectoryCredential -originalcontactDN $originalContactConfiguration.distinguishedName  -errorAction STOP -cn "None"

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
        out-logfile "The contact has no grant send on behalf to."    
    }

    Out-LogFile -string "Invoke get-normalizedDN for any on premises object that the migrated contact has send as permissions."

    #exit #Debug Exit

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END NORMALIZE DNS FOR ALL ATTRIBUTES"
    Out-LogFile -string "********************************************************************************"

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logFile -string "Summary of contact information:"
    out-logfile -string ("The number of objects included in the member migration: "+$exchangecontactMembershipSMTP.count)
    out-logfile -string ("The number of objects included in the reject memebers: "+$exchangeRejectMessagesSMTP.count)
    out-logfile -string ("The number of objects included in the accept memebers: "+$exchangeAcceptMessagesSMTP.count)
    out-logfile -string ("The number of objects included in the managedBY memebers: "+$exchangeManagedBySMTP.count)
    out-logfile -string ("The number of objects included in the moderatedBY memebers: "+$exchangeModeratedBySMTP.count)
    out-logfile -string ("The number of objects included in the bypassModeration memebers: "+$exchangeBypassModerationSMTP.count)
    out-logfile -string ("The number of objects included in the grantSendOnBehalfTo memebers: "+$exchangeGrantSendOnBehalfToSMTP.count)
    out-logfile -string ("The number of objects included in the send as rights: "+$exchangeSendAsSMTP.count)
    out-logfile -string ("The number of contacts on premsies that this contact has send as rights on: "+$allObjectsSendAsAccessNormalized.Count)
    out-logfile -string ("The number of contacts on premises that this contact has full mailbox access on: "+$allObjectsFullMailboxAccess.count)
    out-logfile -string ("The number of mailbox folders on premises that this contact has access to: "+$allMailboxesFolderPermissions.count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

    #Exit #Debug Exit.

    #At this point we have obtained all the information relevant to the individual contact.
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

    out-logfile -string "Being validating all distribution list members."
    
    if ($exchangecontactMembershipSMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each contact member is in Office 365 / Exchange Online"

        foreach ($member in $exchangecontactMembershipSMTP)
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
                        ErrorMessage = "A member of the distribution list is not found in Office 365."
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
        out-logfile -string "There are no contact members to test."    
    }

    out-logfile -string "Begin evaluating all members with reject rights."

    if ($exchangeRejectMessagesSMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each contact reject messages is in Office 365."

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
                        Attribute = "RejectMessagesFromSendersorMembers / RejectMessagesFrom / RejectMessagesFromcontactMembers (ADAttributes: UnAuthOrig / dlMemRejectPerms)"
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
        out-logfile -string "Ensuring each contact accept messages is in Office 365 / Exchange Online"

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
                        Attribute = "AcceptMessagesOnlyFromSendersorMembers / AcceptMessagesOnlyFrom / AcceptMessagesOnlyFromcontactMembers (ADAttributes: authOrig / dlMemSubmitPerms)"
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
        out-logfile -string "Ensuring each contact managed by is in Office 365 / Exchange Online"

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
        out-logfile -string "Ensuring each contact moderated by is in Office 365 / Exchange Online"

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
        out-logfile -string "Ensuring each contact bypass moderation is in Office 365 / Exchange Online"

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
                        Attribute = "BypassModerationFromSendersorMembers (ADAttributes: msExchBypassModerationLink,msExchBypassModerationFromcontactMembersLink)"
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
        out-logfile -string "Ensuring each contact grant send on behalf to is in Office 365 / Exchange Online"

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

    out-logfile -string "Begin evaluation all members with send as rights."

    if ($exchangeSendAsSMTP.count -gt 0)
    {
        out-logfile -string "Ensuring each contact send as is in Office 365."

        foreach ($member in $exchangeSendAsSMTP)
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
                        Attribute = "SendAs"
                        ErrorMessage = "A member with SendAs permissions was not found in Office 365."
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
        out-logfile -string "There were no members with send as rights."    
    }

    out-logfile -string "Begin evaluation of contacts on premises that the contact to be migrated has send as rights on."

    if ($allObjectsSendAsAccessNormalized.count -gt 0)
    {
        out-logfile -string "Ensuring that each contact on premises that the migrated contact has send as rights on is in Office 365."

        foreach ($member in $allObjectsSendAsAccessNormalized)
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
                        Attribute = "contact with SendAs"
                        ErrorMessage = "The contact to be migrated has send as rights on an on premises object.  The object is not present in Office 365."
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
        out-logfile -string "There were no members with send as rights."    
    }

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END VALIDATE RECIPIENTS IN CLOUD"
    Out-LogFile -string "********************************************************************************"

    #It is possible that this contact was a member of - or other contacts have a dependency on this contact.
    #We will implement a function to track those dependen$ocies.

    #Exit #Debug Exit

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN RECORD DEPENDENCIES ON MIGRATED contact"
    Out-LogFile -string "********************************************************************************"

    out-logfile -string "Get all the contacts that this user is a member of - normalize to canonicalname."

    #Start with contacts this contact is a member of remaining on premises.

    if ($originalContactConfiguration.memberOf -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.memberof)
        {
            try 
            {
                $allcontactsMemberOf += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $NULL
                ExternalDirectoryObjectID = $NULL
                Alias = $NULL
                Name = $NULL
                Attribute = "MemberOf"
                ErrorMessage = "The contact to be migrated is a member of an on premises group.  The contact must be removed <or> is ineligable for migration."
                errorMessageDetail = $DN
            }

            out-logfile -string $isErrorObject

            $preCreateErrors+=$isErrorObject
        }
    }

    if ($allcontactsMemberOf -ne $NULL)
    {
        out-logFile -string "The contact to be migrated is a member of the following contacts."
        out-logfile -string $allcontactsMemberOf
    }
    else 
    {
        out-logfile -string "The contact is not a member of any other contacts on premises."
    }

    #Hancontacte all recipients that have forwarding to this contact based on forwarding address.

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

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $NULL
                ExternalDirectoryObjectID = $NULL
                Alias = $NULL
                Name = $NULL
                Attribute = "AlternateRecipient / ForwardingAddress"
                ErrorMessage = "The contact to be migrated is set as a forwarding address on a mailbox.  The contact must be removed from the mailbox <or> is ineligable for migration."
                errorMessageDetail = $DN
            }

            out-logfile -string $isErrorObject

            $preCreateErrors+=$isErrorObject
        }
    }

    if ($allUsersForwardingAddress -ne $NULL)
    {
        out-logFile -string "The contact has forwarding address set on the following users.."
        out-logfile -string $allUsersForwardingAddress
    }
    else 
    {
        out-logfile -string "The contact does not have forwarding set on any other users."
    }

    #Hancontacte all contacts this object has reject permissions on.

    if ($originalContactConfiguration.unAuthOrigBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.unAuthOrigBL)
        {
            try 
            {
                $allcontactsReject += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $NULL
                ExternalDirectoryObjectID = $NULL
                Alias = $NULL
                Name = $NULL
                Attribute = "RejectMessagesFromSendersOrmembers"
                ErrorMessage = "The contact to be migrated has reject permissions on the following object.  The permissions need to be removed <or> the contact is ineligable for migration."
                errorMessageDetail = $DN
            }

            out-logfile -string $isErrorObject

            $preCreateErrors+=$isErrorObject
        }
    }

    if ($allcontactsReject -ne $NULL)
    {
        out-logFile -string "The contact has reject permissions on the following contacts:"
        out-logfile -string $allcontactsReject
    }
    else 
    {
        out-logfile -string "The contact does not have any reject permissions on other contacts."
    }

    #Hancontacte all contacts this object has accept permissions on.

    if ($originalContactConfiguration.authOrigBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.authOrigBL)
        {
            try 
            {
                $allcontactsAccept += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $NULL
                ExternalDirectoryObjectID = $NULL
                Alias = $NULL
                Name = $NULL
                Attribute = "AcceptMessagesFromSendersOrMembers"
                ErrorMessage = "The contact has Accept permissions on a group or object.  The permissions must be removed <or> the contact is ineligable for migration."
                errorMessageDetail = $DN
            }

            out-logfile -string $isErrorObject

            $preCreateErrors+=$isErrorObject
        }
    }

    if ($allcontactsAccept -ne $NULL)
    {
        out-logFile -string "The contact has accept messages from on the following contacts:"
        out-logfile -string $allcontactsAccept
    }
    else 
    {
        out-logfile -string "The contact does not have accept permissions on any contacts."
    }

    if ($originalContactConfiguration.msExchCoManagedObjectsBL -ne $NULL)
    {
        out-logfile -string "Calling ge canonical name."

        foreach ($dn in $originalContactConfiguration.msExchCoManagedObjectsBL)
        {
            try 
            {
                $allcontactsCoManagedByBL += get-canonicalName -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP

            }
            catch {
                out-logfile -string $_ -isError:$TRUE
            }

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $NULL
                ExternalDirectoryObjectID = $NULL
                Alias = $NULL
                Name = $NULL
                Attribute = "ManagedBy"
                ErrorMessage = "The contact to be migrated has managed by rights on the following object.  The rights must be removed <or> the contact is ineligable for migration."
                errorMessageDetail = $DN
            }

            out-logfile -string $isErrorObject

            $preCreateErrors+=$isErrorObject
        }
    }
    else 
    {
        out-logfile -string "The contact is not a co manager on any other contacts."    
    }

    if ($allcontactsCoManagedByBL -ne $NULL)
    {
        out-logFile -string "The contact is a co-manager on the following objects:"
        out-logfile -string $allcontactsCoManagedByBL
    }
    else 
    {
        out-logfile -string "The contact is not a co manager on any other objects."
    }

    #Hancontacte all contacts this object has bypass moderation permissions on.

    if ($originalContactConfiguration.msExchBypassModerationBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.msExchBypassModerationBL)
        {
            try 
            {
                $allcontactsBypassModeration += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }

        $isErrorObject = new-Object psObject -property @{
            PrimarySMTPAddressorUPN = $NULL
            ExternalDirectoryObjectID = $NULL
            Alias = $NULL
            Name = $NULL
            Attribute = "BypassModerationFromSendersOrMembers"
            ErrorMessage = "The contact has bypass permissions on an object.  The permissions must be removed <or> the contact is ineligable for migration."
            errorMessageDetail = $DN
        }

        out-logfile -string $isErrorObject

        $preCreateErrors+=$isErrorObject
    }

    if ($allcontactsBypassModeration -ne $NULL)
    {
        out-logFile -string "This contact has bypass moderation on the following contacts:"
        out-logfile -string $allcontactsBypassModeration
    }
    else 
    {
        out-logfile -string "This contact does not have any bypass moderation on any contacts."
    }

    #Hancontacte all contacts this object has accept permissions on.

    if ($originalContactConfiguration.publicDelegatesBL -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.publicDelegatesBL)
        {
            try 
            {
                $allcontactsGrantSendOnBehalfTo += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $NULL
                ExternalDirectoryObjectID = $NULL
                Alias = $NULL
                Name = $NULL
                Attribute = "GrantSendOnBehalfTo"
                ErrorMessage = "The contact has grant send on behalf to rights on an object.  The rights must be removed <or> the contact is ineligable for migration."
                errorMessageDetail = $DN
            }

            out-logfile -string $isErrorObject

            $preCreateErrors+=$isErrorObject
        }
    }

    if ($allcontactsGrantSendOnBehalfTo -ne $NULL)
    {
        out-logFile -string "This contact has grant send on behalf to to the following contacts:"
        out-logfile -string $allcontactsGrantSendOnBehalfTo
    }
    else 
    {
        out-logfile -string "The contact does ont have any send on behalf of rights to other contacts."
    }

    #Hancontacte all contacts this object has manager permissions on.

    if ($originalContactConfiguration.managedObjects -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalContactConfiguration.managedObjects)
        {
            try 
            {
                $allcontactsManagedBy += get-canonicalname -globalCatalog $globalCatalogWithPort -dn $DN -adCredential $activeDirectoryCredential -errorAction STOP
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }

            $isErrorObject = new-Object psObject -property @{
                PrimarySMTPAddressorUPN = $NULL
                ExternalDirectoryObjectID = $NULL
                Alias = $NULL
                Name = $NULL
                Attribute = "ManagedBy"
                ErrorMessage = "The contat has managed by rights on an object.  The rights must be removed <or> the contact is ineligable for migration."
                errorMessageDetail = $DN
            }

            out-logfile -string $isErrorObject

            $preCreateErrors+=$isErrorObject
        }
    }

    if ($allcontactsManagedBy -ne $NULL)
    {
        out-logFile -string "This contact has managedBY rights on the following contacts."
        out-logfile -string $allcontactsManagedBy
    }
    else 
    {
        out-logfile -string "The contact is not a manager on any other contacts."
    }

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logfile -string ("Summary of dependencies found:")
    out-logfile -string ("The number of contacts that the migrated contact is a member of = "+$allcontactsMemberOf.count)
    out-logfile -string ("The number of contacts that this contact is a manager of: = "+$allcontactsManagedBy.count)
    out-logfile -string ("The number of contacts that this contact has grant send on behalf to = "+$allcontactsGrantSendOnBehalfTo.count)
    out-logfile -string ("The number of contacts that have this contact as bypass moderation = "+$allcontactsBypassModeration.count)
    out-logfile -string ("The number of contacts with accept permissions = "+$allcontactsAccept.count)
    out-logfile -string ("The number of contacts with reject permissions = "+$allcontactsReject.count)
    out-logfile -string ("The number of mailboxes forwarding to this contact is = "+$allUsersForwardingAddress.count)
    out-logfile -string ("The number of contacts this contact is a co-manager on = "+$allcontactsCoManagedByBL.Count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

    #Moved the pre-create errors to this section.
    #If any of the above attributes are set - this means that there are dependencies that we cannot rebuild in Office 365.

    #At this time we have validated the on premises pre-requisits for contact migration.
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
            out-logfile -string ("Error Message: "+$preReq.errorMessage)
            out-logfile -string ("Error Message Details: "+$preReq.errorMessageDetail)
            out-logfile -string "====="
        }

        out-logfile -string "Pre-requist checks failed.  Please refer to the previous list of items that require addressing for migration to proceed." -isError:$TRUE
    }



    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END RECORD DEPENDENCIES ON MIGRATED contact"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "Recording all gathered information to XML to preserve original values."

    if ($allObjectsSendAsAccessNormalized.count -ne 0)
    {
        out-logfile -string $allObjectsSendAsAccessNormalized

        out-xmlFile -itemToExport $allObjectsSendAsAccessNormalized -itemNameToExport $allcontactsSendAsNormalizedXML
    }
    
    if ($exchangecontactMembershipSMTP -ne $NULL)
    {
        Out-XMLFile -itemtoexport $exchangecontactMembershipSMTP -itemNameToExport $exchangecontactMembershipSMTPXML
    }
    else 
    {
        $exchangecontactMembershipSMTP=@()
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

    if ($exchangeSendAsSMTP -ne $NULL)
    {
        out-xmlfile -itemToExport $exchangeSendAsSMTP -itemNameToExport $exchangeSendAsSMTPXML
    }
    else 
    {
        $exchangeSendAsSMTP=@()
    }

    if ($allcontactsMemberOf -ne $NULL)
    {
        out-xmlfile -itemtoexport $allcontactsMemberOf -itemNameToExport $allcontactsMemberOfXML
    }
    else 
    {
        $allcontactsMemberOf=@()
    }
    
    if ($allcontactsReject -ne $NULL)
    {
        out-xmlfile -itemtoexport $allcontactsReject -itemNameToExport $allcontactsRejectXML
    }
    else 
    {
        $allcontactsReject=@()
    }
    
    if ($allcontactsAccept -ne $NULL)
    {
        out-xmlfile -itemtoexport $allcontactsAccept -itemNameToExport $allcontactsAcceptXML
    }
    else 
    {
        $allcontactsAccept=@()
    }

    if ($allcontactsCoManagedByBL -ne $NULL)
    {
        out-xmlfile -itemToExport $allcontactsCoManagedByBL -itemNameToExport $allcontactsCoManagedByXML
    }
    else 
    {
        $allcontactsCoManagedByBL=@()    
    }

    if ($allcontactsBypassModeration -ne $NULL)
    {
        out-xmlfile -itemtoexport $allcontactsBypassModeration -itemNameToExport $allcontactsBypassModerationXML
    }
    else 
    {
        $allcontactsBypassModeration=@()
    }

    if ($allUsersForwardingAddress -ne $NULL)
    {
        out-xmlFile -itemToExport $allUsersForwardingAddress -itemNameToExport $allUsersForwardingAddressXML
    }
    else 
    {
        $allUsersForwardingAddress=@()
    }

    if ($allcontactsManagedBy -ne $NULL)
    {
        out-xmlFile -itemToExport $allcontactsManagedBy -itemNameToExport $allcontactsManagedByXML
    }
    else 
    {
        $allcontactsManagedBy=@()
    }

    if ($allcontactsGrantSendOnBehalfTo -ne $NULL)
    {
        out-xmlFile -itemToExport $allcontactsGrantSendOnBehalfTo -itemNameToExport $allcontactsGrantSendOnBehalfToXML
    }
    else 
    {
        $allcontactsGrantSendOnBehalfTo =@()
    }

    #EXIT #Debug Exit

    #Ok so at this point we have preserved all of the information regarding the on premises contact.
    #It is possible that there could be cloud only objects that this contact was made dependent on.
    #For example - the dirSync contact could have been added as a member of a cloud only contact - or another contact that was migrated.
    #The issue here is that this gets VERY expensive to track - since some of the word to do do is not filterable.
    #With the LDAP improvements we no longer offert the option to track on premises - but the administrator can choose to track the cloud

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "START RETAIN OFFICE 365 contact DEPENDENCIES"
    Out-LogFile -string "********************************************************************************"

    #Process normal mail enabled contacts.

    if ($allowNonSyncedcontact -eq $FALSE)
    {
        out-logFile -string "Office 365 settings are to be retained."

        try {
            $allOffice365MemberOf = Get-O365contactDependency -dn $office365contactConfiguration.distinguishedName -attributeType $office365Members -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of contacts in Office 365 cloud only that the contact is a member of = "+$allOffice365MemberOf.count)

        try {
            $allOffice365Accept = Get-O365contactDependency -dn $office365contactConfiguration.distinguishedName -attributeType $office365AcceptMessagesFrom -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of contacts in Office 365 cloud only that the contact has accept rights = "+$allOffice365Accept.count)

        try {
            $allOffice365Reject = Get-O365contactDependency -dn $office365contactConfiguration.distinguishedName -attributeType $office365RejectMessagesFrom -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of contacts in Office 365 cloud only that the contact has reject rights = "+$allOffice365Reject.count)

        try {
            $allOffice365BypassModeration = Get-O365contactDependency -dn $office365contactConfiguration.distinguishedName -attributeType $office365BypassModerationFrom -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of contacts in Office 365 cloud only that the contact has grant send on behalf to righbypassModeration rights = "+$allOffice365BypassModeration.count)

        try {
            $allOffice365ManagedBy = Get-O365ContactDependency -dn $office365ContactConfiguration.distinguishedName -attributeType $office365ManagedBy -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of groups in Office 365 cloud only that the DL has managedBY = "+$allOffice365ManagedBy.count)

        

        try {
            $allOffice365ForwardingAddress = Get-O365contactDependency -dn $office365contactConfiguration.distinguishedName -attributeType $office365ForwardingAddress -errorAction STOP
        }
        catch {
            out-logFile -string $_ -isError:$TRUE
        }

        out-logfile -string ("The number of contacts in Office 365 cloud only that the contact has forwarding on mailboxes = "+$allOffice365ForwardingAddress.count)

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

        if ($allOffice365ManagedBy -ne $NULL)
        {
            out-logfile -string $allOffice365ManagedBy
            out-xmlFile -itemToExport $allOffice365ManagedBy -itemNameToExport $allOffice365ManagedByXML

            out-logfile -string "Setting group type override to security - the group type may have changed on premises after the permission was added."

            $groupTypeOverride="Security"
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
    }
    else 
    {
        out-logfile -string "Administrator opted out of recording Office 365 dependencies."
    }

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logfile -string ("Summary of dependencies found:")
    out-logfile -string ("The number of office 365 objects that the migrated contact is a member of = "+$allOffice365MemberOf.count)
    out-logfile -string ("The number of office 365 objects that have this contact as bypass moderation = "+$allOffice365BypassModeration.count)
    out-logfile -string ("The number of office 365 objects that this group is a manager of: = "+$allOffice365ManagedBy.count)
    out-logfile -string ("The number of office 365 objects with accept permissions = "+$allOffice365Accept.count)
    out-logfile -string ("The number of office 365 objects with reject permissions = "+$allOffice365Reject.count)
    out-logfile -string ("The number of office 365 mailboxes forwarding to this contact is = "+$allOffice365ForwardingAddress.count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

    #EXIT #Debug Exit

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END RETAIN OFFICE 365 contact DEPENDENCIES"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "START Remove on premises distribution contact from office 365.."
    Out-LogFile -string "********************************************************************************"

    #At this stage we will move the contact to the non-Sync OU and then re-record the attributes.
    #The move here will allow us to preserve the original contacts with attributes until we know that the migration was successful.
    #We will use the move to the non-SYNC OU to trigger deletion.

    #EXIT #Debug exit

    #In this case we should write a status file for all threads.
    #When all threads have reached this point it is safe to have them all move their contacts to the non-Sync OU.

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

    #$Capture the moved contact configuration (since attibutes change upon move.)

    try {
        $originalContactConfigurationUpdated = Get-ADObjectConfiguration -contactSMTPAddress $contactSMTPAddress -globalCatalogServer $globalCatalogWithPort -parameterSet $contactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 
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

    out-logfile -string "Monitoring Exchange Online for distribution list deletion."

    try {
        test-CloudcontactPresent -contactSMTPAddress $contactSMTPAddress -errorAction SilentlyContinue
    }
    catch {
        out-logfile -string $_ -isError:$TRUE
    }

    #At this point we have validated that the contact is gone from office 365.
    #We can begin the process of recreating the distribution contact in Exchange Online.

    out-logfile "Attempting to create the contact in Office 365."

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try {
            $office365contactConfigurationPostMigration=new-office365contact -originalContactConfiguration $originalContactConfiguration -office365contactConfiguration $office365contactConfiguration -errorAction STOP

            #If we made it this far then the contact was created.

            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 10)
            {
                out-logFile -string $_ -isError:$TRUE 
            }
            else 
            {
                out-logfile -string "Unable to create the distribution list on attempt.  Retry"

                if ($loopCounter -gt 0)
                {
                    start-sleepProgress -sleepSeconds ($loopCounter * 5) -sleepstring "Invoke sleep - error creating distribution contact."
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

            out-LogFile -string "Write new contact configuration to XML."

            out-Logfile -string $office365contactConfigurationPostMigration
            out-xmlFile -itemToExport $office365contactConfigurationPostMigration -itemNameToExport $office365contactConfigurationPostMigrationXML
            
            #If we made it this far we can end the loop - we were succssful.

            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 10)
            {
                out-logfile -string "Unable to get Office 365 distribution list configuration after 10 tries."
                $stopLoop -eq $TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to capture the Office 365 contact configuration.  Sleeping 15 seconds." -sleepSeconds 15

                $loopCounter = $loopCounter+1 
            }
        }   
    } while ($stopLoop -eq $false)

    #EXIT #Debug Exit.

    #Now it is time to set the multi valued attributes on the contact in Office 365.
    #Setting these first must occur since moderators have to be established before moderation can be enabled.

    out-logFile -string "Setting the multivalued attributes of the migrated contact."

    out-logfile -string $office365contactConfigurationPostMigration.primarySMTPAddress

    [int]$loopCounter=0
    [boolean]$stopLoop = $FALSE
    
    do {
        try {
            set-Office365contactMV -originalContactConfiguration $originalContactConfiguration -office365contactConfiguration $office365contactConfiguration -office365contactConfigurationPostMigration $office365contactConfigurationPostMigration -exchangeRejectMessage $exchangeRejectMessagesSMTP -exchangeAcceptMessage $exchangeAcceptMessagesSMTP -exchangeModeratedBy $exchangeModeratedBySMTP -exchangeBypassMOderation $exchangeBypassModerationSMTP -exchangeGrantSendOnBehalfTo $exchangeGrantSendOnBehalfToSMTP -errorAction STOP -mailOnMicrosoftComDomain $mailOnMicrosoftComDomain -allowNonSyncedcontact $allowNonSyncedcontact  

            $stopLoop = $TRUE
        }
        catch {
            if ($loopCounter -gt 4)
            {
                out-logFile -string $_ -isError:$TRUE
            }
            else {
                start-sleepProgress -sleepString "Uanble to set Office 365 contact Multi Value attributes - try again." -sleepSeconds 5

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
            $office365contactConfigurationPostMigration = Get-O365contactConfiguration -contactSMTPAddress $office365contactConfigurationPostMigration.externalDirectoryObjectID -errorAction STOP

            #If we made it this far we were successful - output the information to XML.

            out-LogFile -string "Write new contact configuration to XML."

            out-Logfile -string $office365contactConfigurationPostMigration
            out-xmlFile -itemToExport $office365contactConfigurationPostMigration -itemNameToExport $office365contactConfigurationPostMigrationXML

            #Now that we are this far - we can exit the loop.

            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 10)
            {
                out-logfile -string "Unable to get Office 365 distribution list configuration after 10 tries."
                $stopLoop -eq $TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to capture the Office 365 contact configuration.  Sleeping 15 seconds." -sleepSeconds 15

                $loopCounter = $loopCounter+1 
            }
        }
        
    } while ($stopLoop -eq $FALSE)

    
    #The distribution list has now been created.  There are single value attributes that we're now ready to update.

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try {
            set-Office365contact -originalContactConfiguration $originalContactConfiguration -office365contactConfiguration $office365contactConfiguration -office365contactConfigurationPostMigration $office365contactConfigurationPostMigration
            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 4)
            {
                out-logfile -string $_ -isError:$TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Transient error updating distribution contact - retrying." -sleepSeconds 5

                $loopCounter=$loopCounter+1
            }
        }
    } while ($stopLoop -eq $FALSE)

    
    out-logfile -string ("The number of post create errors is: "+$global:postCreateErrors.count)
    

    out-logFile -string ("Capture the contact status post migration.")

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        try {
            $office365contactConfigurationPostMigration = Get-O365contactConfiguration -contactSMTPAddress $office365contactConfigurationPostMigration.externalDirectoryObjectID -errorAction STOP

            #If we made it this far we successfully got the contact.  Write it.

            out-LogFile -string "Write new contact configuration to XML."

            out-Logfile -string $office365contactConfigurationPostMigration
            out-xmlFile -itemToExport $office365contactConfigurationPostMigration -itemNameToExport $office365contactConfigurationPostMigrationXML

            #Now that we wrote it - stop the loop.

            $stopLoop=$TRUE
        }
        catch {
            if ($loopCounter -gt 10)
            {
                out-logfile -string "Unable to get Office 365 distribution list configuration after 10 tries."
                $stopLoop -eq $TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to capture the Office 365 contact configuration.  Sleeping 15 seconds." -sleepSeconds 15

                $loopCounter = $loopCounter+1 
            }
        }   
    } while ($stopLoop -eq $false)

    #The distribution contact has been created and both single and multi valued attributes have been updated.
    #The contact is fully availablle in exchange online.
    #The contact as this point sits in the non-sync OU.  This was to service the deletion.
    #The administrator may have reasons for keeping the contact.
    #If they do the plan is to do two things.
    ###Rename the contact by adding a ! to the name - this ensures that if the contact is every accidentally mail enabled it will not soft match the migrated contact.
    ###We'll stamp custom attribute flags on it to ensure that we know the contact has been mirgated - in case it's a member of another contact to be migrated.

    if ($retainOriginalContact -eq $TRUE)
    {
        Out-LogFile -string "Administrator has choosen to retain the original group."
        out-logfile -string "Rename the group by adding the fixed character !"

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE   

        do {
            try {
                set-newContactName -globalCatalogServer $globalCatalogServer -ContactName $originalContactConfigurationUpdated.Name -dn $originalContactConfigurationUpdated.distinguishedName -adCredential $activeDirectoryCredential -errorAction STOP

                $stopLoop=$TRUE
            }
            catch {
                if($loopCounter -gt 4)
                {
                    out-logfile -string $_ -isError:$TRUE
                }
                else 
                {
                    out-logfile -string $_ 
                    start-sleepProgress -sleepString "Uanble to change DL name - try again." -sleepSeconds 5
                    $loopCounter = $loopCounter+1    
                }
            }
        } while ($stopLoop=$FALSE)

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE

        do {
            try {
                $originalContactConfigurationUpdated = Get-ADObjectConfiguration -contactSMTPAddress $contactSMTPAddress -globalCatalogServer $globalCatalogWithPort -parameterSet $contactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

                $stopLoop=$TRUE
            }
            catch {
                if ($loopCounter -gt 4)
                {
                    out-logFile -string $_ -isError:$TRUE
                }
                else 
                {
                    start-sleepProgress -sleepString "Unable to obtain updated original DL Configuration - try again." -sleepSeconds 5

                    $loopCounter = $loopCounter+1
                }
            }
        } while ($stopLoop -eq $FALSE)

        out-logfile -string $originalContactConfigurationUpdated
        out-xmlFile -itemToExport $originalContactConfigurationUpdated -itemNameTOExport $originalContactConfigurationUpdatedXML+$global:unDoStatus

        Out-LogFile -string "Administrator has choosen to regain the original group."
        out-logfile -string "Disabling the mail attributes on the group."

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE
        
        do {
            try{
                Disable-originalContact -originalContactConfiguration $originalContactConfigurationUpdated -globalCatalogServer $globalCatalogServer -parameterSet $contactPropertySetToClear -adCredential $activeDirectoryCredential -errorAction STOP

                $stopLoop = $TRUE
            }
            catch{
                if ($loopCounter -gt 4)
                {
                    out-LogFile -string $_ -isError:$TRUE
                }
                else 
                {
                    start-sleepProgress -sleepString "Unable to disable distribution group - try again." -sleepSeconds 5

                    $loopCounter = $loopCounter + 1
                }
            }
        } while ($stopLoop -eq $false)

        

        

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE
        
        do {
            try {
                $originalContactConfigurationUpdated = Get-ADObjectConfiguration -dn $originalContactConfigurationUpdated.distinguishedName -globalCatalogServer $globalCatalogWithPort -parameterSet $contactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

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

        Out-LogFile -string "Move the original group back to the OU it came from.  The group will no longer be soft matched."

        [int]$loopCounter=0
        [boolean]$stopLoop=$FALSE

        do {
            try {
                #Discovered that it's possible someone used the name "Test Group".  This breaks the following DN search as OU appears in the name - WHOOPS
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
                $originalContactConfigurationUpdated = Get-ADObjectConfiguration -dn $tempDN -globalCatalogServer $globalCatalogWithPort -parameterSet $contactPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

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
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365UnifiedAccept -office365Member $member -groupSMTPAddress $groupSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated distribution list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "Distribution List AcceptMessagesOnlyFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated distribution list to Office 365 distribution group.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 groups with accept permissions."    
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
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365UnifiedReject -office365Member $member -groupSMTPAddress $groupSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated distribution list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "Distribution List RejectMessagesFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated distribution list to Office 365 distribution group.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 groups with reject permissions."    
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
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365BypassModerationusers -office365Member $member -groupSMTPAddress $groupSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated distribution list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "Distribution List BypassModerationFromSendersOrMembers"
                    errorMessage = "Unable to add the migrated distribution list to Office 365 distribution group.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-LogFile -string "There were no Office 365 groups with bypass moderation permissions."    
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
                $isTestError=start-ReplaceOffice365 -office365Attribute $office365ManagedBy -office365Member $member -groupSMTPAddress $groupSMTPAddress -errorAction STOP
            }
            catch{
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated distribution list to Office 365 Resource."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "Distribution List ManagedBy"
                    errorMessage = "Unable to add the migrated distribution list to Office 365 distribution group.  Manual add required."
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

    out-logfile -string ("Adding migrated group to any cloud only groups.")

    if ($allOffice365MemberOf.count -gt 0)
    {
        out-logfile -string "Adding cloud only group member."

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

            out-logfile -string ("Processing group = "+$member.primarySMTPAddress)
            try {
                $isTestError=start-replaceOffice365Members -office365Group $member -groupSMTPAddress $groupSMTPAddress -errorAction STOP
            }
            catch {
                out-logfile -string $_
                $isTestErrorDetail = $_
                $isTestError="Yes"
            }

            if ($isTestError -eq "Yes")
            {
                out-logfile -string "Error adding migrated distribution list to Office 365 Distribution List."

                $isErrorObject = new-Object psObject -property @{
                    distinguishedName = $member.distinguishedName
                    primarySMTPAddress = $member.primarySMTPAddress
                    alias = $member.Alias
                    displayName = $member.displayName
                    attribute = "Distribution List Membership"
                    errorMessage = "Unable to add the migrated distribution list to Office 365 distribution group.  Manual add required."
                    erroMessageDetail = $isTestErrorDetail
                }

                out-logfile -string $isErrorObject

                $office365ReplaceErrors+=$isErrorObject
            }
        }
    }
    else 
    {
        out-logfile -string "No cloud only groups had the migrated group as a member."
    }       

    #If the administrator has selected to not retain the group - remove it.

    if ($retainOriginalContact -eq $FALSE)
    {
        $isTestError="No"

        out-logfile -string "Deleting the original group."

        $isTestError=remove-onPremContact -globalCatalogServer $globalCatalogServer -originalContactConfiguration $originalContactConfigurationUpdated -adCredential $activeDirectoryCredential -errorAction STOP
    }
    else
    {
        $isTestError = "No"
    }


    if ($isTestError -eq "Yes")
    {
        $isErrorObject = new-Object psObject -property @{
            errorMessage = "Uanble to remove the on premises group at request of administrator.  Group may need to be manually removed."
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
    Out-LogFile -string "END START-DISTRIBUTIONLISTMIGRATION"
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
        out-logfile -string "Errors were encountered in the distribution list creation process requireing administrator review."
        out-logfile -string "Although the migration may have been successful - manual actions may need to be taken to full complete the migration."
        out-logfile -string "++++++++++"
        out-logfile -string "+++++"
        out-logfile -string "" -isError:$TRUE

    }

    #Archive the files into a date time success folder.

    Start-ArchiveFiles -isSuccess:$TRUE -logFolderPath $logFolderPath
}