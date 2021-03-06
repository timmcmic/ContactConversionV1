<#
    .SYNOPSIS

    This function gets the original contact configuration for the on premises contact using AD providers.

    .DESCRIPTION

    This function gets the original contact configuration for the on premises contact using AD providers.

    .PARAMETER parameterSet

    These are the parameters that the GET will gather from AD for the contact.  This should be the full map.

    .PARAMETER contactSMTPAddress

    The mail attribute of the contact to search.

    .PARAMETER GlobalCatalog

    The global catalog to utilize for the query.

    .OUTPUTS

    Returns the contact configuration from the LDAP / AD call to the calling function.

    .EXAMPLE

    Get-ADObjectConfiguration -powershellsessionname NAME -contactSMTPAddress Address

    #>
    Function Get-ADObjectConfiguration
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true,ParameterSetName = "BySMTPAddress")]
            [string]$contactSMTPAddress="None",
            [Parameter(Mandatory = $true,ParameterSetName = "ByDN")]
            [string]$dn="None",
            [Parameter(Mandatory = $true,ParameterSetName = "BySMTPAddress")]
            [Parameter(Mandatory = $true,ParameterSetName = "ByDN")]
            [string]$globalCatalogServer,
            [Parameter(Mandatory = $true,ParameterSetName = "BySMTPAddress")]
            [Parameter(Mandatory = $true,ParameterSetName = "ByDN")]
            [array]$parameterSet,
            [Parameter(Mandatory = $TRUE,ParameterSetName = "BySMTPAddress")]
            [Parameter(Mandatory = $true,ParameterSetName = "ByDN")]
            $adCredential
        )

        #Declare function variables.

        $functioncontactConfiguration=$NULL #Holds the return information for the contact query.

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Get-ADObjectConfiguration"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("contactSMTPAddress = "+$contactSMTPAddress)
        Out-LogFile -string ("GlobalCatalogServer = "+$globalCatalogServer)
        OUt-LogFile -string ("Parameter Set:")
        
        foreach ($parameterIncluded in $parameterSet)
        {
            Out-Logfile -string $parameterIncluded
        }

        out-logfile -string ("Credential user name = "+$adCredential.UserName)

        #Get the contact using LDAP / AD providers.
        
        try 
        {
            Out-LogFile -string "Using AD / LDAP provider to get original contact configuration"

            if ($contactSMTPAddress -ne "None")
            {
                out-logfile -string ("Searching by mail address "+$contactSMTPAddress)
                out-logfile -string ("Imported Address Length: "+$contactSMTPAddress.length.toString())

                #Ensure that there are no spaces contained in the string (account for import errors.)

                out-logfile -string ("Spaces Removed Address Length: "+$contactSMTPAddress.length.toString())

                $functioncontactConfiguration=Get-ADObject -filter "mail -eq '$contactSMTPAddress'" -properties $parameterSet -server $globalCatalogServer -credential $adCredential -errorAction STOP
            }
            elseif ($DN -ne "None")
            {
                out-logfile -string ("Searching by distinguished name "+$dn)

                $functioncontactConfiguration=get-adObject -identity $DN -properties $parameterSet -server $globalCatalogServer -credential $adCredential
            }
            else 
            {
                out-logfile -string "No value query found for local object." -isError:$TRUE    
            }
            

            #If the ad provider command cannot find the contact - the variable is NULL.  An error is not thrown.

            if ($functioncontactConfiguration -eq $NULL)
            {
                throw "The contact cannot be found in Active Directory by email address."
            }

            Out-LogFile -string "Original contact configuration found and recorded."
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END Get-ADObjectConfiguration"
        Out-LogFile -string "********************************************************************************"
        
        #This function is designed to open local and remote powershell sessions.
        #If the session requires import - for example exchange - return the session for later work.
        #If not no return is required.
        
        return $functioncontactConfiguration
    }