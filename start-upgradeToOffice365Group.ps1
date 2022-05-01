<#
    .SYNOPSIS

    This function triggers the upgrade of the contact to an Office 365 Modern / Unified contact
    
    .DESCRIPTION

    This function triggers the upgrade of the contact to an Office 365 Modern / Unified contact

    .PARAMETER contactSMTPAddress

    .OUTPUTS

    None

    .EXAMPLE

    start-upgradeToOffice365contact -contactSMTPAddress address

    #>
    Function start-upgradeToOffice365contact
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress
        )

        [string]$isTestError="No"

        #Declare function variables.

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN start-upgradeToOffice365contact"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        out-logfile -string ("contact SMTP Address = "+$contactSMTPAddress)

        #Call the command to begin the upgrade process.

        out-logFile -string "Calling command to being the upgrade process."
        out-logfile -string "NOTE:  This command runs in the background and no status is provided."
        out-logfile -string "Administrators MUST validate the upgrade as successful manually."

        try{
            $attempt=upgrade-o365Distributioncontact -contactIdentities $contactSMTPAddress
        }
        catch{
            out-logFile -string $_
            $isTestError="Yes"
        }

        out-logfile -string $attempt
        out-logfile -string ("Upgrade attempt successfully submitted = "+$attempt.SuccessfullySubmittedForUpgrade)

        if ($attempt.reason -ne $NULL)
        {
            out-logfile -string ("Error Reason = "+$attempt.errorReason)
            $isTestError="Yes"
        }
        
        Out-LogFile -string "END start-upgradeToOffice365contact"
        Out-LogFile -string "********************************************************************************"

        return $isTestError
    }