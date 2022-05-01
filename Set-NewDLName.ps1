<#
    .SYNOPSIS

    This function add a character to the contact name if exchange hybrid is enabled (allows for the dynamic contact creation.)
    
    .DESCRIPTION

    This function add a character to the contact name if exchange hybrid is enabled (allows for the dynamic contact creation.)

    .PARAMETER GlobalCatalogServer

    The global catalog to make the query against.

    .PARAMETER DN

    The original DN of the object.

    .PARAMETER contactName

    The name of the contact from the original configuration.

    .PARAMETER contactSamAccountName

    The original DN of the object.

    .OUTPUTS

    None

    .EXAMPLE

    set-newcontactName -contactConfiguration contactConfiguration -globalCatalogServer globalCatalogServer

    #>
    Function set-newcontactName
     {
        [cmcontactetbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$globalCatalogServer,
            [Parameter(Mandatory = $true)]
            $contactName,
            [Parameter(Mandatory = $true)]
            $contactSAMAccountName,
            [Parameter(Mandatory = $true)]
            $DN,
            [Parameter(Mandatory = $true)]
            $adCredential
        )

        #Declare function variables.

        [string]$functioncontactName=$NULL #Holds the calculated name.
        [string]$functioncontactSAMAccountName=$NULL #Holds the calculated sam account name.

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN SET-NEWcontactNAME"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("GlobalCatalogServer = "+$globalCatalogServer)
        OUt-LogFile -string ("contactName = "+$contactName)
        out-logfile -string ("contactSamAccontName = "+$contactSAMAccountName)
        out-logfile -string ("DN = "+$dn)

        #Establish new names

        [string]$functioncontactName = $contactName+"!"
        [string]$functioncontactSAMAccountName = $contactSAMAccountName+"!"

        out-logfile -string ("New contact name = "+$functioncontactName)
        out-logfile -string ("New contact sam account name = "+$functioncontactSAMAccountName)
        
        #Get the specific user using ad providers.
        
        try 
        {
            Out-LogFile -string "Set the AD contact name."

            set-adcontact -identity $dn -samAccountName $functioncontactSAMAccountName -server $globalCatalogServer -Credential $adCredential
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        try
        {
            out-logfile -string "Setting the new contact name.."

            rename-adobject -identity $dn -newName $functioncontactName -server $globalCatalogServer -credential $adCredential
        }
        catch
        {
            Out-LogFile -string $_ -isError:$true  
        }

        Out-LogFile -string "END Set-NewcontactName"
        Out-LogFile -string "********************************************************************************"
    }