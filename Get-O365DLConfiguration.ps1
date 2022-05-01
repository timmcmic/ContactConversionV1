<#
    .SYNOPSIS

    This function uses the exchange online powershell session to gather the office 365 distribution list configuration.

    .DESCRIPTION

    This function uses the exchange online powershell session to gather the office 365 distribution list configuration.

    .PARAMETER contactSMTPAddress

    The mail attribute of the group to search.

    .OUTPUTS

    Returns the PS object associated with the recipient from get-o365recipient

    .EXAMPLE

    get-o365dlconfiguration -contactSMTPAddress Address

    #>
    Function Get-o365DLConfiguration
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress,
            [Parameter(Mandatory = $false)]
            [string]$groupTypeOverride=""
        )

        #Declare function variables.

        $functionDLConfiguration=$NULL #Holds the return information for the group query.
        $functionMailSecurity="MailUniversalSecurityGroup"
        $functionMailDistribution="MailUniversalDistributionGroup"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN GET-O365DLCONFIGURATION"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("contactSMTPAddress = "+$contactSMTPAddress)

        #Get the recipient using the exchange online powershell session.
        
        try 
        {
            if ($groupTypeOverride -eq "")
            {
                Out-LogFile -string "Using Exchange Online to capture the distribution group."

                $functionDLConfiguration=get-O365DistributionGroup -identity $contactSMTPAddress -errorAction STOP
            
                Out-LogFile -string "Original DL configuration found and recorded."
            }
            elseif ($groupTypeOverride -eq "Security")
            {
                Out-logfile -string "Using Exchange Online to capture distribution group with filter security"

                $functionDLConfiguration=get-o365DistributionGroup -identity $contactSMTPAddress -RecipientTypeDetails $functionMailSecurity -errorAction STOP

                out-logfile -string "Original DL configuration found and recorded by filter security."
            }
            elseif ($groupTypeOverride -eq "Distribution")
            {
                out-logfile -string "Using Exchange Online to capture distribution group with filter distribution."

                $functionDLConfiguration=get-o365DistributionGroup -identity $contactSMTPAddress -RecipientTypeDetails $functionMailDistribution

                out-logfile -string "Original DL configuration found and recorded by filter distribution."
            }
            
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END GET-O365DLCONFIGURATION"
        Out-LogFile -string "********************************************************************************"
        
        #This function is designed to open local and remote powershell sessions.
        #If the session requires import - for example exchange - return the session for later work.
        #If not no return is required.
        
        return $functionDLConfiguration
    }