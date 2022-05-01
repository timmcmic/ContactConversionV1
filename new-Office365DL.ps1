<#
    .SYNOPSIS

    This function creates the new distribution contact in office 365.

    .DESCRIPTION

    This function creates the new distribution contact in office 365.

    .PARAMETER originalContactConfiguration

    The original configuration of the contact on premises.

    .PARAMETER contactTypeOverride

    Submits the contact type override of specified by the administrator at run time.

    .OUTPUTS

    None

    .EXAMPLE

    new-Office365contact -contactTypeOverride "Security" -originalContactConfiguration adConfigVariable.

    #>
    Function new-office365contact
     {
        [cmcontactetbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration,
            [Parameter(Mandatory = $true)]
            $office365contactConfiguration,
            [Parameter(Mandatory = $true)]
            [string]$contactTypeOverride
        )

        #Declare function variables.

        [string]$functioncontactType=$NULL #Holds the return information for the contact query.
        [string]$functionMailNickName = ""
        [string]$functionName = ((Get-Date -Format FileDateTime)+(Get-Random)).tostring()
        $functioncontact = $NULL

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN New-Office365contact"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        Out-LogFile -string ("originalContactConfiguration = ")
        out-logfile -string $originalContactConfiguration
        out-logfile -string ("Office365contactConfiguration = ")
        out-logfile -string $office365contactConfiguration
        out-logfile -string ("contact Type Override = "+$contactTypeOverride)

        #Calculate the contact type to be utilized.
        #Three values - either NULL,Security,or Distribution.

        out-Logfile -string ("The contact type for evaluation is = "+$originalContactConfiguration.contactType)

        if ($contactTypeOverride -Eq "Security")
        {
            out-logfile -string "The administrator overrode the contact type to security."

            $functioncontactType = "Security"
        }
        elseif ($contactTypeOverride -eq "Distribution")
        {
            out-logfile -string "The administrator overrode the contact type to distribution."

            $functioncontactType = "Distribution"
        }
        elseif ($contactTypeOverride -eq "None") 
        {
            out-logfile -string "A contact type override was not specified.  Using contact type from on premises."

            if (($originalContactConfiguration.contactType -eq "-2147483640") -or ($originalContactConfiguration.contactType -eq "-2147483646") -or ($originalContactConfiguration.contactType -eq "-2147483644"))
            {
                out-logfile -string "The contact type from on premises is security."

                $functioncontactType = "Security"
            }
            elseif (($originalContactConfiguration.contacttype -eq "8") -or ($originalContactConfiguration.contacttype -eq "4") -or ($originalContactConfiguration.contacttype -eq "2"))
            {
                out-logfile -string "The contact type from on premises is distribution."

                $functioncontactType = "Distribution"
            }
            else 
            {
                out-logfile -string "A contact type override was not provided and the input did not include a valid on premises contact type."    
            }
        }
        else 
        {
            out-logfile -string "An invalid contact type was utilized in function new-Office365contact" -isError:$TRUE    
        }

        out-logfile -string ("Random contact name: "+$functionName)

        #Create the distribution contact in office 365.
        
        try 
        {
            out-logfile -string "Creating the distribution contact in Office 365."

            $functioncontact = new-o365distributioncontact -name $functionName -type $functioncontactType -ignoreNamingPolicy:$TRUE -errorAction STOP 

            out-logfile -string $functioncontact
        }
        catch 
        {
            Out-LogFile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END New-Office365contact"
        Out-LogFile -string "********************************************************************************"

        return $functioncontact
    }