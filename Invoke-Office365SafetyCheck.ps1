<#
    .SYNOPSIS

    This function confirms that the distribution list specified and found in Office 365 is DirSynced=TRUE
    
    .DESCRIPTION

    This function confirms that the distribution list specified and found in Office 365 is DirSynced=TRUE

    .PARAMETER O365contactConfiguration

    The contact configuration obtained by the service.

    .OUTPUTS

    No returns.

    .EXAMPLE

    invoke-office365safetycheck -o365contactconfiguration o365contactconfiguration

    #>
    Function Invoke-Office365SafetyCheck
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $o365contactconfiguration
        )

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN INVOKE-OFFICE365SAFETYCHECK"
        Out-LogFile -string "********************************************************************************"

        #Comapre the isDirSync attribute.
        
        try 
        {
            Out-LogFile -string ("Distribution list isDirSynced = "+$o365contactconfiguration.isDirSynced)

            if ($o365contactconfiguration.isDirSynced -eq $FALSE)
            {
                Out-LogFile -string "The distribution list requested is not directory synced and cannot be migrated." -isError:$TRUE
            }
            else 
            {
                Out-LogFile -string "The distribution list requested is directory synced."
            }
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END INVOKE-OFFICE365SAFETYCHECK"
        Out-LogFile -string "********************************************************************************"
    }