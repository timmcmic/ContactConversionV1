<#
    .SYNOPSIS

    This function updates the membership of any cloud only distribution lists for the migrated distribution contact.

    .DESCRIPTION

    This function updates the membership of any cloud only distribution lists for the migrated distribution contact.

    .PARAMETER office365contact

    The member that is being added.

    .PARAMETER contactSMTPAddress

    The member that is being added.

    .OUTPUTS

    None

    .EXAMPLE

    sstart-replaceOffice365 -office365Attribute Attribute -office365Member contactMember -contactSMTPAddress smtpAddess

    #>
    Function start-replaceOffice365Members
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $office365contact,
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress
        )

        [string]$isTestError="No"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN start-ReplaceOffice365Members"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        $functionCommand=$NULL

        Out-LogFile -string ("Office 365 Attribute = "+$office365contact)
        out-logfile -string ("Office 365 Member = "+$contactSMTPAddress)

        #Declare function variables.

        out-Logfile -string "Processing operation..."

        try{
            add-o365DistributionGroupMember -identity $office365contact.primarySMTPAddress -member $contactSMTPAddress -errorAction STOP 
        }
        catch{
            out-logfile -string $_
            $isTestError="Yes"
        }


        Out-LogFile -string "END start-replaceOffice365Members"
        Out-LogFile -string "********************************************************************************"

        return $isTestError
    }