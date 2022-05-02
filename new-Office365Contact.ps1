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
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $originalContactConfiguration,
            [Parameter(Mandatory = $true)]
            $office365contactConfiguration
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


        #Calculate the contact type to be utilized.
        #Three values - either NULL,Security,or Distribution.

        out-logfile -string ("Random contact name: "+$functionName)

        #Create the distribution contact in office 365.
        
        try 
        {
            out-logfile -string "Creating the distribution contact in Office 365."

            $functioncontact = new-o365mailcontact -externalEmailAddress $originalDLConfiguration.mail -name $functionName -errorAction STOP 

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