<#
    .SYNOPSIS

    This function begins the process of replacing the Office 365 settings for contacts that have been migrated that had cloud only dependencies.

    .DESCRIPTION

    This function begins the process of replacing the Office 365 settings for contacts that have been migrated that had cloud only dependencies.

    .PARAMETER office365Attribute

    The office 365 attribute.

    .PARAMETER office365Member

    The member that is being added.

    .PARAMETER contactSMTPAddress

    The member that is being added.

    .OUTPUTS

    None

    .EXAMPLE

    sstart-replaceOffice365 -office365Attribute Attribute -office365Member contactMember -contactSMTPAddress smtpAddess

    #>
    Function start-ReplaceOffice365
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $office365Attribute,
            [Parameter(Mandatory = $true)]
            $office365Member,
            [Parameter(Mandatory = $true)]
            [string]$contactSMTPAddress
        )

        [string]$isTestError="No"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN start-ReplaceOffice365"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        $functionCommand=$NULL
        $functionMailboxRecipientType = "UserMailbox"
        $functionDistributioncontactRecipientType = "MailUniversalDistributionGroup"
        $functionSecuritycontactRecipientType = "MailUniversalSecurityGroup"
        $functionMailUserRecipientType = "MailUser"
        $functionMailContactRecipientType = "MailContact"
        $functionUniveralRecipientDisplayType = "GroupMailbox"
        $functionDynamicDistributioncontactRecipientType = "DynamicDistributionGroup"

        $functionForwarding = "ForwardingAddress"

        Out-LogFile -string ("Office 365 Attribute = "+$office365Attribute)
        out-logfile -string ("Office 365 Member = "+$office365Member.primarySMTPAddress)
        out-logfile -string ("Office 365 Recipient Type = "+$office365Member.recipientType)
        out-logfile -string ("Office 365 Recipient Display Type Details ="+$office365Member.recipientTypeDetails)

        #Declare function variables.

        out-Logfile -string "Processing operation..."

        if ($office365Attribute -eq $functionForwarding)
        {
            out-logfile -string "Recipient is a mailbox with forwarding rights."

            $functionCommand="set-o365Mailbox -identity $office365Member -$office365Attribute '$contactSMTPAddress' -errorAction STOP"
            out-logfile -string ("The command to execute:  "+$functionCommand)
        }
        elseif (($office365Member.recipientType -eq $functionDistributioncontactRecipientType) -and ($office365Member.recipientTypeDetails -eq $functionUniveralRecipientDisplayType))
        {
            out-logfile -string "Recipient is a unified contact."

            $functionCommand="set-o365UnifiedGroup -identity $office365Member -$office365Attribute @{add='$contactSMTPAddress'} -errorAction STOP"
            out-logfile -string ("The command to execute:  "+$functionCommand)
        }
        elseif (($office365Member.recipientType -eq $functionDistributioncontactRecipientType) -or ($office365Member.recipientType -eq $functionSecuritycontactRecipientType))
        {
            out-logfile -string "Recipient is a mail enabled distribution contact or mail enabled security contact."

            $functionCommand="set-o365DistributionGroup -identity $office365Member -$office365Attribute @{add='$contactSMTPAddress'} -BypassSecurityGroupManagerCheck -errorAction STOP "
            out-logfile -string ("The command to execute:  "+$functionCommand)
        }
        elseif ($office365Member.recipientType -eq $functionDynamicDistributioncontactRecipientType)
        {
            out-logfile -string "Recipient is a dynamic distribution contact."

            $functionCommand="set-o365DynamicDistributionGroup -identity $office365Member -$office365Attribute '$contactSMTPAddress' -errorAction STOP"
            out-logfile -string ("The command to execute:  "+$functionCommand)
        }
        elseif ($office365member.recipientType -eq $functionMailboxRecipientType)
        {
            out-logfile -string "Recipient is a mailbox."

            $functionCommand="set-o365Mailbox -identity $office365Member -$office365Attribute @{add='$contactSMTPAddress'} -errorAction STOP"
            out-logfile -string ("The command to execute:  "+$functionCommand)
        }
        elseif ($office365Member.recipientType -eq $functionMailUserRecipientType)
        {
            out-logfile -string "Recipient is a mail user."

            $functionCommand="set-o365MailUser -identity $office365Member -$office365Attribute @{add='$contactSMTPAddress'} -errorAction STOP"
            out-logfile -string ("The command to execute:  "+$functionCommand)
        }
        elseif ($office365Member.recipientType -eq $functionMailContactRecipientType)
        {
            out-logfile -string "Recipient is a mail user."

            $functionCommand="set-o365MailContact -identity $office365Member -$office365Attribute @{add='$contactSMTPAddress'} -errorAction STOP"
            out-logfile -string ("The command to execute:  "+$functionCommand)
        }
        else 
        {
            out-logfile "There is no acceptable recipient type specified.  Manual intervention required."
            $isTestError="Yes"    
        }

        out-logfile -string "Recipient type is validated and correct command built."

        if ($isTestError -ne "Yes")
        {
            $scriptBlock = [scriptBlock]::create($functionCommand)

            out-logfile -string ("The script block to execute: "+$scriptBlock)
    
            try {
                & $scriptBlock
            }
            catch {
                out-logfile -string $_
                $isTestError="Yes"
            }
        }
        else
        {
            out-logfile -string "Previous error encountered - no command executed."
        }

        Out-LogFile -string "END start-replaceOffice365"
        Out-LogFile -string "********************************************************************************"

        return $isTestError
    }