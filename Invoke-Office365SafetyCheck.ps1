<#
    .SYNOPSIS

    This function confirms that the list specified and found in Office 365 is DirSynced=TRUE
    
    .DESCRIPTION

    This function confirms that the list specified and found in Office 365 is DirSynced=TRUE

    .PARAMETER O365DLConfiguration

    The DL configuration obtained by the service.

    .OUTPUTS

    No returns.

    .EXAMPLE

    invoke-office365safetycheck -o365dlconfiguration o365dlconfiguration

    #>
    Function Invoke-Office365SafetyCheck
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $o365dlconfiguration
        )

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN INVOKE-OFFICE365SAFETYCHECK"
        Out-LogFile -string "********************************************************************************"

        #Comapre the isDirSync attribute.
        
        try 
        {
            Out-LogFile -string ("list isDirSynced = "+$o365dlconfiguration.isDirSynced)

            if ($o365dlconfiguration.isDirSynced -eq $FALSE)
            {
                Out-LogFile -string "The list requested is not directory synced and cannot be migrated." -isError:$TRUE
            }
            else 
            {
                Out-LogFile -string "The list requested is directory synced."
            }
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END INVOKE-OFFICE365SAFETYCHECK"
        Out-LogFile -string "********************************************************************************"
    }