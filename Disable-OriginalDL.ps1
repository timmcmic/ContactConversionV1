<#
    .SYNOPSIS

    This function disabled the on premies distribution list - removing it from azure ad and exchange online.

    .DESCRIPTION

    This function disabled the on premies distribution list - removing it from azure ad and exchange online.

    .PARAMETER parameterSet

    These are the parameters that will be manually cleared from the object in AD mode.

    .PARAMETER DN

    The DN of the contact to remove.

    .PARAMETER GlobalCatalog

    The global catalog server the operation should be performed on.

    .PARAMETER UseExchange

    If set to true disablement will occur using the exchange on premises powershell commands.

    .OUTPUTS

    No return.

    .EXAMPLE

    Get-ADObjectConfiguration -powershellsessionname NAME -contactSMTPAddress Address

    #>
    Function Disable-Originalcontact
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration,
            [Parameter(Mandatory = $true)]
            [string]$globalCatalogServer,
            [Parameter(Mandatory = $false)]
            [array]$parameterSet="None",
            [Parameter(Mandatory = $true)]
            $adCredential
        )

        #Declare function variables.

        $functioncontactConfiguration=$NULL #Holds the return information for the contact query.


        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Disable-originalContactConfiguration"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("originalContactConfiguration = "+$originalContactConfiguration)
        Out-LogFile -string ("GlobalCatalogServer = "+$globalCatalogServer)
        out-logfile -string ("Use Exchange On Premises ="+$useOnPremisesExchange)
        out-logfile -string ("DN of object to modify / disable "+$originalContactConfiguration.distinguishedName)

        OUt-LogFile -string ("Parameter Set:")
        
        foreach ($parameterIncluded in $parameterSet)
        {
            Out-Logfile -string $parameterIncluded
        }

        #Get the contact using LDAP / AD providers.
        
        try 
        {
            set-adcObject -identity $originalContactConfiguration.distinguishedName -server $globalCatalogServer -clear $parameterSet -credential $adCredential

        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END Disable-originalContactConfiguration"
        Out-LogFile -string "********************************************************************************"
    }