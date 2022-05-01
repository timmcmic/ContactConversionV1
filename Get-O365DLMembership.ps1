<#
    .SYNOPSIS

    This function obtains the contact membership of the Office 365 distribution contact.

    .DESCRIPTION

    This function obtains the contact membership of the Office 365 distribution contact.

    .PARAMETER contactSMTPAddress

    The mail attribute of the contact to search.

    .OUTPUTS

    Returns the membership array of the contact in Office 365.

    .EXAMPLE

    get-o365contactMembership -contactSMTPAddress Address

    #>
    Function Get-o365contactMembership
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress
        )

        #Declare function variables.

        $functioncontactMembership=$NULL #Holds the return information for the contact query.

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN GET-O365contactMEMBERSHIP"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("contactSMTPAddress = "+$contactSMTPAddress)

        #Get the recipient using the exchange online powershell session.
        
        try 
        {
            Out-LogFile -string "Using Exchange Online to obtain the contact membership."

            $functioncontactMembership=get-O365DistributioncontactMember -identity $contactSMTPAddress -errorAction STOP
            
            Out-LogFile -string "Distribution contact membership recorded."
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END GET-O365contactMEMBERSHIP"
        Out-LogFile -string "********************************************************************************"
        
        #Return the membership to the caller.
        
        return $functioncontactMembership
    }