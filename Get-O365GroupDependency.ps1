<#
    .SYNOPSIS

    This function queries Office 365 for any cloud only dependencies on the migrated contacts.
    
    .DESCRIPTION

    This function queries Office 365 for any cloud only dependencies on the migrated contacts.

    .PARAMETER DN

    The DN of the object to search attributes for.

    .PARAMETER ATTRIBUTETYPE

    The attribute type of the object we're looking for.

    .OUTPUTS

    An array of PS objects that are the canonicalNames of the dependencies.

    .EXAMPLE

    get-o36contactDependency -dn DN -attributeType multiValuedExchangeAttribute

    #>
    Function Get-O365contactDependency
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$DN,
            [Parameter(Mandatory = $TRUE)]
            [string]$attributeType
        )

        #Declare function variables.

        $functionTest=$NULL #Holds the return information for the contact query.
        $functionCommand=$NULL #Holds the expression that will be utilized to query office 365.
        [array]$functionObjectArray=$NULL #This is used to hold the object that will be returned.

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN GET-O365contactDependency"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        OUt-LogFile -string ("DN Set = "+$DN)
        out-logfile -string ("Attribute Type = "+$attributeType)
        out-logfile -string ("contact Type = "+$contactType)
        
        #Get the specific user using ad providers.
        
        try 
        {
            Out-LogFile -string "Attempting to search Office365 for any contacts or users that have the requested dependency."

            if ($attributeType -eq "Members")
            {
                #The attribute type is member - so we need to query recipients.

                Out-LogFile -string "Entering query office 365 for contact membership."

                $functionCommand = "Get-o365Recipient -Filter { ($attributeType -eq '$dn') -and (isDirSynced -eq '$FALSE') } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest = invoke-command -scriptBlock $scriptBlock

                out-logfile -string ("The function command executed = "+$functionCommand)
            }
            elseif ($attributeType -eq "ForwardingAddress")
            {
                 #The attribute type is forwarding address - search only mailboxes.

                 Out-LogFile -string "Entering query office 365 mailboxes."

                 $functionCommand = "Get-o365Mailbox -Filter { $attributeType -eq '$dn' } -errorAction 'STOP'"

                 $scriptBlock=[scriptBlock]::create($functionCommand)

                 $functionTest = invoke-command -scriptBlock $scriptBlock
                 
                 out-logfile -string ("The function command executed = "+$functionCommand)
            }
            elseif ($attributeType -eq "ManagedBy")
            {
                #The attribute type is managed by.  This is only relevant to contacts.

                out-logfile "Managed by is only relevant to contacts - performing query on only contacts."

                out-logfile -string "Starting collection of distribution contacts."

                $functionCommand = "Get-o365Distributioncontact -Filter { ($attributeType -eq '$dn') -and (isDirSynced -eq '$FALSE') } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest = invoke-command -scriptBlock $scriptBlock
                
                out-logfile -string ("The function command executed = "+$functionCommand)

                out-logfile -string "Starting collection of dynamic distribution contacts."

                $functionCommand = "Get-o365DynamicDistributioncontact -Filter { $attributeType -eq '$dn' } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest += invoke-command -scriptBlock $scriptBlock
                
                out-logfile -string ("The function command executed = "+$functionCommand)
            }
            else
            {
                #The attribute type is a property of the contact - attempt to obtain.

                <#

                Out-LogFile -string "Entering query office 365 for contact to be set on property."

                if ($contactType -eq "Standard")
                {
                    out-logfile -string "The contact type is standard - querying distribution contacts."
                    
                    $functionCommand = "Get-o365Distributioncontact -Filter { ($attributeType -eq '$dn') -and (isDirSynced -eq '$FALSE') } -errorAction 'STOP'"

                    $scriptBlock=[scriptBlock]::create($functionCommand)

                    $functionTest = invoke-command -scriptBlock $scriptBlock
                    
                    out-logfile -string ("The function command executed = "+$functionCommand)
                }
                elseif ($contactType -eq "Unified")
                {
                    out-logfile -string "The contact type is unified - querying distribution contacts."
                    
                    $functionCommand = "Get-o365Unifiedcontact -Filter { $attributeType -eq '$dn' } -errorAction 'STOP'"

                    $scriptBlock=[scriptBlock]::create($functionCommand)

                    $functionTest = invoke-command -scriptBlock $scriptBlock
                    
                    out-logfile -string ("The function command executed = "+$functionCommand)
                }
                elseif ($contactType -eq "Dynamic")
                {
                    out-logfile -string "The contact type is dynamic - querying distribution contacts."
                    
                    $functionCommand = "Get-o365DynamicDistributioncontact -Filter { $attributeType -eq '$dn' } -errorAction 'STOP'"

                    $scriptBlock=[scriptBlock]::create($functionCommand)

                    $functionTest = invoke-command -scriptBlock $scriptBlock
                    
                    out-logfile -string ("The function command executed = "+$functionCommand)
                }
                else 
                {
                    throw "Invalid contact type specified in function call.  Acceptable Standard or Universal"    
                } 

                #>

                out-logfile -string "Starting to gather attribute for all recipient types."
                out-logfile -string "Starting collection of distribution contacts."

                $functionCommand = "Get-o365Distributioncontact -Filter { ($attributeType -eq '$dn') -and (isDirSynced -eq '$FALSE') } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest = invoke-command -scriptBlock $scriptBlock
                
                out-logfile -string ("The function command executed = "+$functionCommand)

                out-logfile -string "Starting collection of dynamic distribution contacts."

                $functionCommand = "Get-o365DynamicDistributioncontact -Filter { $attributeType -eq '$dn' } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest += invoke-command -scriptBlock $scriptBlock
                
                out-logfile -string ("The function command executed = "+$functionCommand)

                out-logfile -string "Starting collection of universal distribution contacts."

                $functionCommand = "Get-o365Unifiedcontact -Filter { $attributeType -eq '$dn' } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest += invoke-command -scriptBlock $scriptBlock
                
                out-logfile -string ("The function command executed = "+$functionCommand)

                out-logfile -string "Starting collection of mailbox recipients."

                $functionCommand = "Get-o365Mailbox -Filter { ($attributeType -eq '$dn') -and (isDirSynced -eq '$FALSE') } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest += invoke-command -scriptBlock $scriptBlock

                out-logfile -string ("The function command executed = "+$functionCommand)

                out-logfile -string "Starting collection of mail user recipients."

                $functionCommand = "Get-o365Mailuser -Filter { ($attributeType -eq '$dn') -and (isDirSynced -eq '$FALSE') } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest += invoke-command -scriptBlock $scriptBlock

                out-logfile -string ("The function command executed = "+$functionCommand)

                out-logfile -string "Starting collection of mail contact recipients."

                $functionCommand = "Get-o365MailContact -Filter { ($attributeType -eq '$dn') -and (isDirSynced -eq '$FALSE') } -errorAction 'STOP'"

                $scriptBlock=[scriptBlock]::create($functionCommand)

                $functionTest += invoke-command -scriptBlock $scriptBlock

                out-logfile -string ("The function command executed = "+$functionCommand)

            }

            if ($functionTest -eq $NULL)
            {
                out-logfile -string "There were no contacts or users with the request dependency."
            }
            else 
            {
                $functionObjectArray = $functionTest
            }
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        return $functionObjectArray
    }