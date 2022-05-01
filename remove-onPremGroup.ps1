<#
    .SYNOPSIS

    This function disables all open powershell sessions.

    .DESCRIPTION

    This function disables all open powershell sessions.

    .OUTPUTS

    No return.

    .EXAMPLE

    disable-allPowerShellSessions

    #>
    Function remove-onPremcontact
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$globalCatalogServer,
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration,
            [Parameter(Mandatory = $true)]
            $adCredential
        )

        [string]$isTestError="No"

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN remove-onPremcontact"
        Out-LogFile -string "********************************************************************************"

        out-logFile -string "Remove on premises distribution contact."

        try
        {
            remove-adobject -identity $originalContactConfiguration.distinguishedName -server $globalCatalogServer -credential $adCredential -confirm:$FALSE -errorAction STOP
        }
        catch
        {
            out-logfile -string $_
            $isTestError="Yes"
        }

        Out-LogFile -string "END remove-onPremcontact"
        Out-LogFile -string "********************************************************************************"

        return $isTestError
    }