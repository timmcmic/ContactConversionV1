<#
    .SYNOPSIS

    This function enables the administrator to create the hybrid mail flow objects post migration.

    .DESCRIPTION

    This function enables the administrator to create the hybrid mail flow objects post migration.

    .PARAMETER contactSMTPAddress

    The mail attribute of the contact to search.

    .OUTPUTS

    None

    .EXAMPLE

    enable-HybridMailFlow -contactSMTPAddress SMTPAddress -globalCatalogServer GC.domain.com -activeDirectoryCredential $cred -exchangeServer server.domain.com -exchangeServerCredential $cred -exchangeOnlineCredential $cred

    #>
    Function enable-hybridMailFlowPostMigration
     {
        [cmcontactetbinding()]

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
            [string]$exchangeServer=$NULL,
            [Parameter(Mandatory = $false)]
            [pscredential]$exchangeCredential=$NULL,
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
            [string]$OU=$NULL
        )

        #Declare function variables.

        $global:logFile=$NULL #This is the global variable for the calculated log file name
        [string]$global:staticFolderName="\contactMigration\"

        [boolean]$useOnPremisesExchange=$FALSE #Determines if function will utilize onpremises exchange during migration.
        [string]$exchangeOnPremisesPowershellSessionName="ExchangeOnPremises" #Defines universal name for on premises Exchange Powershell session.
        [string]$exchangeOnlinePowershellModuleName="ExchangeOnlineManagement" #Defines the exchage management shell name to test for.
        [string]$activeDirectoryPowershellModuleName="ActiveDirectory" #Defines the active directory shell name to test for.
        [string]$contactConversionPowershellModule="contactConversionV2"
        [string]$globalCatalogPort=":3268"
        [string]$globalCatalogWithPort=$globalCatalogServer+$globalCatalogPort

        #Static variables utilized for the Exchange On-Premsies Powershell.
   
        [string]$exchangeServerConfiguration = "Microsoft.Exchange" #Powershell configuration.
        [boolean]$exchangeServerAllowRedirection = $TRUE #Allow redirection of URI call.
        [string]$exchangeServerURI = "https://"+$exchangeServer+"/powershell" #Full URL to the on premises powershell instance based off name specified parameter.

        #Declare logging variables.

        [string]$office365contactConfigurationXML = "office365contactConfigurationXML"
        [string]$routingContactXML="routingContactXML"
        [string]$routingDynamiccontactXML="routingDynamiccontactXML"

        $routingContactConfig=$NULL
        $office365contactConfiguration = $NULL

        #Create the log file.

        new-LogFile -contactSMTPAddress $contactSMTPAddress.trim() -logFolderPath $logFolderPath

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "enable-hybridMailFlowPostMigration"
        Out-LogFile -string "********************************************************************************"

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

        $contactSMTPAddress = remove-stringSpace -stringToFix $contactSMTPAddress
        $globalCatalogServer = remove-stringSpace -stringToFix $globalCatalogServer
        $logFolderPath = remove-stringSpace -stringToFix $logFolderPath 

        if ($exchangeServer -ne $NULL)
        {
            $exchangeServer=remove-stringSpace -stringToFix $exchangeServer
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
    
        $exchangeAuthenticationMethod=remove-StringSpace -stringToFix $exchangeAuthenticationMethod

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

        if ($exchangeServer -ne "")
        {
            Out-LogFile -string ("ExchangeServer = "+$exchangeServer)
        }

        if ($exchangecredential -ne $null)
        {
            Out-LogFile -string ("ExchangeUserName = "+$exchangeCredential.UserName.toString())
        }

        if ($exchangeOnlineCredential -ne $null)
        {
            Out-LogFile -string ("ExchangeOnlineUserName = "+ $exchangeOnlineCredential.UserName.toString())
        }

        if ($exchangeOnlineCertificateThumbPrint -ne "")
        {
            Out-LogFile -string ("ExchangeOnlineCertificateThumbprint = "+$exchangeOnlineCertificateThumbPrint)
        }

        Out-LogFile -string ("ExchangeAuthenticationMethod = "+$exchangeAuthenticationMethod)

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string " RECORD VARIABLES"
        Out-LogFile -string "********************************************************************************"

        out-logfile -string ("Global Catalog Port = "+$globalCatalogPort)
        out-logfile -string ("Global catalog string used for function queries ="+$globalCatalogWithPort)
        out-logFile -string ("Initial use of Exchange On Prem = "+$useOnPremisesExchange)
        Out-LogFile -string ("Exchange on prem powershell session name = "+$exchangeOnPremisesPowershellSessionName)
        Out-LogFile -string ("AD Global catalog powershell session name = "+$ADGlobalCatalogPowershellSessionName)
        Out-LogFile -string ("Exchange powershell module name = "+$exchangeOnlinePowershellModuleName)
        Out-LogFile -string ("Active directory powershell modulename = "+$activeDirectoryPowershellModuleName)

        #Validate that both the exchange credential and exchange server are presented together.

        Out-LogFile -string "Validating that both ExchangeServer and ExchangeCredential are specified."

        if (($exchangeServer -eq "") -and ($exchangeCredential -ne $null))
        {
            #The exchange credential was specified but the exchange server was not specified.

            Out-LogFile -string "ERROR:  Exchange Server is required when specfying Exchange Credential." -isError:$TRUE
        }
        elseif (($exchangeCredential -eq $NULL) -and ($exchangeServer -ne ""))
        {
            #The exchange server was specified but the exchange credential was not.

            Out-LogFile -string "ERROR:  Exchange Credential is required when specfying Exchange Server." -isError:$TRUE
        }
        elseif (($exchangeCredential -ne $NULL) -and ($exchangetServer -ne ""))
        {
            #The server name and credential were specified for Exchange.

            Out-LogFile -string "The server name and credential were specified for Exchange."

            #Set useOnPremisesExchange to TRUE since the parameters necessary for use were passed.

            $useOnPremisesExchange=$TRUE

            Out-LogFile -string ("Set useOnPremsiesExchanget to TRUE since the parameters necessary for use were passed - "+$useOnPremisesExchange)
        }
        else
        {
            Out-LogFile -string ("Neither Exchange Server or Exchange Credentials specified - retain useOnPremisesExchange FALSE - "+$useOnPremisesExchange)
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

        if ($useOnPremisesExchange -eq $False)
        {
            out-logfile -string "Exchange on premsies information must be provided in order to enable hybrid mail flow." -isError:$TRUE
        }

        Out-LogFile -string "END PARAMETER VALIDATION"
        Out-LogFile -string "********************************************************************************"
        
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

        #Now we can determine if exchange on premises is utilized and if so establish the connection.
   
        Out-LogFile -string "Determine if Exchange On Premises specified and create session if necessary."

        if ($useOnPremisesExchange -eq $TRUE)
        {
            try 
            {
                Out-LogFile -string "Calling New-PowerShellSession"

                $sessiontoImport=new-PowershellSession -credentials $exchangecredential -powershellSessionName $exchangeOnPremisesPowershellSessionName -connectionURI $exchangeServerURI -authenticationType $exchangeAuthenticationMethod -configurationName $exchangeServerConfiguration -allowredirection $exchangeServerAllowRedirection -requiresImport:$TRUE
            }
            catch 
            {
                Out-LogFile -string "ERROR:  Unable to create powershell session." -isError:$TRUE
            }
            try 
            {
                Out-LogFile -string "Calling import-PowerShellSession"

                import-powershellsession -powershellsession $sessionToImport
            }
            catch 
            {
                Out-LogFile -string "ERROR:  Unable to create powershell session." -isError:$TRUE
            }
            try 
            {
                out-logfile -string "Calling set entire forest."

                enable-ExchangeOnPremEntireForest
            }
            catch 
            {
                Out-LogFile -string "ERROR:  Unable to view entire forest." -isError:$TRUE
            }
        }
        else
        {
            Out-LogFile -string "No on premises Exchange specified - skipping setup of powershell session."
        }

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END ESTABLISH POWERSHELL SESSIONS"
        Out-LogFile -string "********************************************************************************"

        #First step - gather the Office 365 contact Information.
        #The contact should be present in the service and previously migrated.

        try {
            out-logfile -string "Obtaining Office 365 Distribution List Configuration"

            $office365contactConfiguration = get-o365contactconfiguration -contactSMTPAddress $contactSMTPAddress -errorAction STOP
        }
        catch {
            out-logfile -string "Unable to obtain the distribution list information from Office 365."
            out-logfile -string $_ -isError:$TRUE
        }

        out-xmlFile -itemToExport $office365contactConfiguration -itemNameToExport $office365contactConfigurationXML

        #Now that we have the configuration - we need to ensure dir sync is set to false.

        out-logfile -string "Testing to ensure that the distribution list is directory synchornized."

        out-logfile -string ("IsDirSynced: "+$office365contactConfiguration.isDirSynced)

        if ($office365contactConfiguration.isDirSynced -eq $FALSE)
        {
            out-logfile -string "The distribution list is cloud only - proceed."
        }
        else 
        {
            out-logfile -string "The distribution list is directory synchronized - this function may only run on cloud only contacts." -isError:$TRUE    
        }

        #At this time test to ensure the routing contact is present.

        $tempMailArray = $office365contactConfiguration.windowsEmailAddress.split("@")

        foreach ($member in $tempMailArray)
        {
            out-logfile -string ("Temp Mail Address Member: "+$member)
        }

        $tempMailAddress = $tempMailArray[0]+"-MigratedByScript"

        out-logfile -string ("Temp routing contact address: "+$tempMailAddress)

        $tempMailAddress = $tempMailAddress+"@"+$tempMailArray[1]

        out-logfile -string ("Temp routing contact address: "+$tempMailAddress)

        try {
            $routingContactConfiguration = Get-ADObjectConfiguration -contactSMTPAddress $tempMailAddress -globalCatalogServer $globalCatalogWithPort -parameterSet "*" -errorAction STOP -adCredential $activeDirectoryCredential 

            out-logfile -string "Overriding OU selection by adminsitrator - contact already exists.  Must be the same as contact."

            $OU = get-OULocation -originalContactConfiguration $routingContactConfiguration

            out-logfile -string "The routing contact was found and recorded."

            out-xmlFile -itemToExport $routingContactConfiguration -itemNameToExport $routingContactXML+0
        }
        catch {
            out-logfile -string "The routing contact is not present - create the routing contact."
            out-logfile -string $_

            try{
                out-logfile -string "Creating the routing contact that is missing."

                new-routingContact -originalContactConfiguration $office365contactConfiguration -office365contactConfiguration $office365contactConfiguration -globalCatalogServer $globalCatalogServer -adCredential $activeDirectoryCredential -isRetry:$TRUE -isRetryOU $OU -errorAction STOP

                out-logfile -string "The routing contact was created successfully."
            }
            catch{
                out-logfile -string "The routing contact could not be created."
                out-logfile -string $_ -isError:$TRUE
            }
        }

        $loopCounter=0
        $stopLoop=$FALSE

        do {
            try {
                out-logfile -string "Re-obtaining the routing contact configuration."
    
                $routingContactConfiguration = Get-ADObjectConfiguration -contactSMTPAddress $tempMailAddress -globalCatalogServer $globalCatalogWithPort -parameterSet "*" -errorAction STOP -adCredential $activeDirectoryCredential 

                $stopLoop = $TRUE
            }
            catch {

                if ($loopCounter -lt 5)
                {
                    start-sleeProgress -sleepSeconds 5 -sleepString "Sleeping failed obtain contact..."
                    $loopCounter=$loopCounter+1
                }
                else
                {
                    out-logfile -string $_
                    out-logfile -string "Unable to obtain the routing contact information." -isError:$TRUE
                }
            }
        } until ($stopLoop -eq $TRUE)

       

        out-xmlFile -itemToExport $routingContactConfiguration -itemNameToExport $routingContactXML+1

        #At this time the mail contact needs to be mail enabled.

        try {
            enable-mailRoutingContact -globalCatalogServer $globalCatalogServer -routingContactConfig $routingContactConfiguration -errorAction STOP
        }
        catch {
            out-logfile -string $_
            out-logfile -string "Unable to mail enable the routing contact." -isError:$TRUE
        }

        #Obtain the updated routing contact.

        try{
            out-logfile -string "Re-obtaining the routing contact configuration."

            $routingContactConfiguration = Get-ADObjectConfiguration -contactSMTPAddress $tempMailAddress -globalCatalogServer $globalCatalogWithPort -parameterSet "*" -errorAction STOP -adCredential $activeDirectoryCredential 
        }
        catch{
            out-logfile -string $_
            out-logfile -string "Unable to obtain the routing contact." -isError:$TRUE
        }

        out-xmlFile -itemToExport $routingContactConfiguration -itemNameToExport $routingContactXML+2

        #The routing contact is now mail enabled.  Create the dynamic distribution contact.

        try {
            out-logfile -string "Creating the dynamic distribution contact for mail routing."

            Enable-MailDyamiccontact -globalCatalogServer $globalCatalogServer -originalContactConfiguration $office365contactConfiguration -routingContactConfig $routingContactConfiguration -isRetry:$TRUE -errorAction STOP
        }
        catch {
            out-logfile -string "Unable to create the dynamic distribution contact."
            out-logfile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END enable-hybridMailFlowPostMigration"
        Out-LogFile -string "********************************************************************************"
    }