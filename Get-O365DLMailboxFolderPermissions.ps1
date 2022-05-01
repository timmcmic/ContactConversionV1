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

    Get-O365contactFullMaiboxAccess -contactSMTPAddress Address

    #>
    Function Get-O365contactMailboxFolderPermissions
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress,
            [Parameter(Mandatory = $false)]
            $collectedData=$NULL
        )

        #Declare function variables.

        [array]$functionFolderAccess=@()

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Get-O365contactMailboxFolderPermissions"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("contactSMTPAddress = "+$contactSMTPAddress)

        #Get the recipient using the exchange online powershell session.

        if ($collectedData -eq $NULL)
        {
            out-logfile -string "No folder permissions were provided for evaluation."
        }
        elseif ($collectedData -ne $NULL)
        {
            <#
            $ProgressDelta = 100/($collectedData.count); $PercentComplete = 0; $MbxNumber = 0

            out-logfile -string "Processing folder permissions for imported data."

            foreach ($folder in $collectedData)
            {
                $MbxNumber++
    
                write-progress -activity "Processing Recipient" -status $folder.identity -PercentComplete $PercentComplete

                $PercentComplete += $ProgressDelta

                if ($folder.user.tostring() -eq $contactSMTPAddress )
                {
                    $functionFolderAccess+=$folder
                }
            }

            #>

            out-logfile -string "Filter all entries for objects that have been removed."
            out-logfile -string ("Pre count: "+$collectedData.count)

            $collectedData = $collectedData | where {$_.user.userPrincipalName -ne $NULL}

            out-logfile -string ("Post count: "+$collectedData.count)

            $functionFolderAccess = $collectedData | where {$_.user.tostring() -eq $contactSMTPAddress}
        }

        write-progress -activity "Processing Recipient" -completed

        Out-LogFile -string "END Get-O365contactMailboxFolderPermissions"
        Out-LogFile -string "********************************************************************************"
        
        if ($functionFolderAccess.count -gt 0)
        {
            return $functionFolderAccess
        }
    }