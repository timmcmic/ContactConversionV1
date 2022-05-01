<#
    .SYNOPSIS

    This function obtains the DL membership of the Office 365 distribution group.

    .DESCRIPTION

    This function obtains the DL membership of the Office 365 distribution group.

    .PARAMETER contactSMTPAddress

    The mail attribute of the group to search.

    .OUTPUTS

    Returns the membership array of the DL in Office 365.

    .EXAMPLE

    get-o365dlMembership -contactSMTPAddress Address

    #>
    Function Get-o365DLMembership
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress
        )

        #Declare function variables.

        $functionDLMembership=$NULL #Holds the return information for the group query.

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN GET-O365DLMEMBERSHIP"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("contactSMTPAddress = "+$contactSMTPAddress)

        #Get the recipient using the exchange online powershell session.
        
        try 
        {
            Out-LogFile -string "Using Exchange Online to obtain the group membership."

            $functionDLMembership=get-O365DistributionGroupMember -identity $contactSMTPAddress -errorAction STOP
            
            Out-LogFile -string "Distribution group membership recorded."
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END GET-O365DLMEMBERSHIP"
        Out-LogFile -string "********************************************************************************"
        
        #Return the membership to the caller.
        
        return $functionDLMembership
    }