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

    Get-O365contactSendAs -contactSMTPAddress Address

    #>
    Function Get-O365contactSendAs
     {
        [cmcontactetbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress,
            [Parameter(Mandatory = $false)]
            [string]$isTrustee=$FALSE
        )

        #Declare function variables.

        [array]$functionSendAs=@()

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Get-O365contactSendAs"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("contactSMTPAddress = "+$contactSMTPAddress)

        #Get the recipient using the exchange online powershell session.

        if ($isTrustee -eq $TRUE)
        {
            try 
            {
                Out-LogFile -string "Obtaining all Office 365 contacts the migrated contact has send as permissions on."

                $functionSendAs = get-o365RecipientPermission -Trustee $contactSMTPAddress -resultsize unlimited -errorAction STOP
            }
            catch 
            {
                Out-LogFile -string $_ -isError:$TRUE
            }
        }
        else
        {
            try
            {
                out-logfile -string "Obtaining all send as permissions set directly in Office 365 on the contact to be migrated."

                $functionSendAs = get-O365RecipientPermission -identity $contactSMTPAddress -resultsize unlimited -errorAction STOP
            }
            catch
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
        
        

        Out-LogFile -string "END Get-O365contactSendAs"
        Out-LogFile -string "********************************************************************************"
        
        return $functionSendAs
    }