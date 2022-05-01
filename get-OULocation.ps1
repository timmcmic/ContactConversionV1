<#
    .SYNOPSIS

    This function calculates the correct OU to place an object.

    .DESCRIPTION

    This function calculates the correct OU to place an object.

    .PARAMETER originalContactConfiguration

    The mail attribute of the contact to search.

    .OUTPUTS

    Returns the organizational unit where the object should be stored.

    .EXAMPLE

    get-OULocation -originalContactConfiguration $originalContactConfiguration

    #>

    Function Get-OULocation
     {
        [cmcontactetbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration
        )

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Get-OULocation"
        Out-LogFile -string "********************************************************************************"

        #Declare function variables.

        [string]$returnOU=$NULL

        #Test to see if the DN contains an OU.

        out-logfile -string $originalContactConfiguration.distinguishedname

        $testOUSubstringLocation = $originalContactConfiguration.distinguishedName.indexof(",OU=")
        out-logfile -string ("The location of ,OU= is:"+$testOUSubstringLocation)

        if ($testOUSubStringLocation -ge 0)
        {
            out-logfile -string "The contact is in an organizational unit."
            out-logfile -string $testOUSubstringLocation.tostring()
            $tempOUSubstring = $originalContactConfiguration.distinguishedname.substring($testOUSubstringLocation)
            out-logfile -string "Temp OU Substring = "
            out-logfile -string $tempOUSubstring
            $testOUSubstringLocation = $tempOUSubstring.indexof("OU=")
            out-logfile -string $testOUSubstringLocation.tostring()
            $tempOUSubstring = $tempOUSubstring.substring($testOUSubstringLocation)
            out-logfile -string "Temp OU Substring Substring ="
            out-logfile -string $tempOUSubstring
        }
        else 
        {
            out-logfile -string "The contact is in a container and not an OU."
            $testOUSubstringLocation = $originalContactConfiguration.distinguishedName.indexof(",CN=")    
            out-logfile -string $testOUSubstringLocation.tostring()
            $tempOUSubstring = $originalContactConfiguration.distinguishedname.substring($testOUSubstringLocation)
            out-logfile -string "Temp OU Substring = "
            out-logfile -string $tempOUSubstring
            $testOUSubstringLocation = $tempOUSubstring.indexof("CN=")
            out-logfile -string $testOUSubstringLocation.tostring()
            $tempOUSubstring = $tempOUSubstring.substring($testOUSubstringLocation)
            out-logfile -string "Temp OU Substring Substring ="
            out-logfile -string $tempOUSubstring
        }

        

        $returnOU = $tempOUSubstring

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END Get-OULocation"
        Out-LogFile -string "********************************************************************************"

        return $returnOU
     }