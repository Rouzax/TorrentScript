<#
.SYNOPSIS
    Gets user input with optional validation.
.DESCRIPTION
    This function prompts the user for input and validates it based on the optional 'Required' switch.
.PARAMETER Message
    Specifies the message displayed to the user as a prompt for input.
.PARAMETER Required
    Indicates whether the input is required. If used, the function continues to prompt until valid input is provided.
.INPUTS
    None. This function does not accept piped input.
.OUTPUTS 
    System.String. The user-provided input.
.EXAMPLE
    Get-Input -Message "Enter your name" -Required
    Prompts the user to enter their name, and the input is required.
.EXAMPLE
    Get-Input -Message "Enter your age"
    Prompts the user to enter their age, and the input is optional.
#>
function Get-Input {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [switch]$Required
    )

    # Variable to store user input
    $UserInput = $null

    # Loop until a valid input is provided (if Required switch is used)
    while ($Required -and ($null -eq $UserInput -or $UserInput -eq '')) {
        $UserInput = Read-Host -Prompt "Required | $Message"
        $UserInput = $UserInput.Trim()
    }

    # If Required switch is not used, get input without validation
    if (-not $Required) {
        $UserInput = Read-Host -Prompt "$Message"
        $UserInput = $UserInput.Trim()
    }

    # Return the user input
    return $UserInput
}