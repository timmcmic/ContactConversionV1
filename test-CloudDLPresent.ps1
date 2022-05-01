<#
    .SYNOPSIS

    This function loops until we detect that the cloud contact is no longer present.
    
    .DESCRIPTION

    This function loops until we detect that the cloud contact is no longer present.

    .PARAMETER contactSMTPAddress

    The SMTP Address of the contact.

    .OUTPUTS

    None

    .EXAMPLE

    test-CloudcontactPresent -contactSMTPAddress SMTPAddress

    #>
    Function test-CloudcontactPresent
     {
        [cmcontactetbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress
        )

        #Declare function variables.

        [boolean]$firstLoopProcessing=$TRUE

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN TEST-CLOUDcontactPRESENT"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        out-Logfile -string ("contact SMTP Address = "+$contactSMTPAddress)

        do 
        {
            if ($firstLoopProcessing -eq $TRUE)
            {
                Out-LogFile -string "First time checking for contact - do not sleep."
                $firstLoopProcessing = $FALSE
            }
            else 
            {
                start-sleepProgress -sleepString "contact found in Office 365 - sleep for 30 seconds - try again." -sleepSeconds 30
            }

        } while (get-exoRecipient -identity $contactSMTPAddress)

        Out-LogFile -string "END TEST-CLOUDcontactPRESENT"
        Out-LogFile -string "********************************************************************************"
    }