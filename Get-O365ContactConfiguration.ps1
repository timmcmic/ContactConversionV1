<#
    .SYNOPSIS

    This function uses the exchange online powershell session to gather the office 365 distribution list configuration.

    .DESCRIPTION

    This function uses the exchange online powershell session to gather the office 365 distribution list configuration.

    .PARAMETER contactSMTPAddress

    The mail attribute of the contact to search.

    .OUTPUTS

    Returns the PS object associated with the recipient from get-o365recipient

    .EXAMPLE

    get-o365contactconfiguration -contactSMTPAddress Address

    #>
    Function Get-o365contactConfiguration
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress,
            [Parameter(Mandatory = $false)]
            [string]$contactTypeOverride=""
        )

        #Declare function variables.

        $functioncontactConfiguration=$NULL #Holds the return information for the contact query.
        $functionMailSecurity="MailUniversalSecuritycontact"
        $functionMailDistribution="MailUniversalDistributioncontact"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN GET-O365contactCONFIGURATION"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("contactSMTPAddress = "+$contactSMTPAddress)

        #Get the recipient using the exchange online powershell session.
        
        try 
        {
            if ($contactTypeOverride -eq "")
            {
                Out-LogFile -string "Using Exchange Online to capture the distribution contact."

                $functioncontactConfiguration=get-o365MailContact -identity $contactSMTPAddress -errorAction STOP
            
                Out-LogFile -string "Original contact configuration found and recorded."
            }
            elseif ($contactTypeOverride -eq "Security")
            {
                Out-logfile -string "Using Exchange Online to capture distribution contact with filter security"

                $functioncontactConfiguration=get-o365MailContact -identity $contactSMTPAddress -RecipientTypeDetails $functionMailSecurity -errorAction STOP

                out-logfile -string "Original contact configuration found and recorded by filter security."
            }
            elseif ($contactTypeOverride -eq "Distribution")
            {
                out-logfile -string "Using Exchange Online to capture distribution contact with filter distribution."

                $functioncontactConfiguration=get-o365MailContact -identity $contactSMTPAddress -RecipientTypeDetails $functionMailDistribution

                out-logfile -string "Original contact configuration found and recorded by filter distribution."
            }
            
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END GET-O365contactCONFIGURATION"
        Out-LogFile -string "********************************************************************************"
        
        #This function is designed to open local and remote powershell sessions.
        #If the session requires import - for example exchange - return the session for later work.
        #If not no return is required.
        
        return $functioncontactConfiguration
    }