Function Get-Input
{
    <#
    .SYNOPSIS
    Get input from user
    
    .DESCRIPTION
    Stop the script and get input from user and answer will be returned
    
    .PARAMETER Message
    The question to show to the user
    
    .PARAMETER Required
    If provided will force an answer to be given
    
    .EXAMPLE
    $UserName = Get-Input -Message "What is your name" -Required
    $UserAge = Get-Input -Message "What is your age"
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true
        )]
        [string]    $Message,

        [Parameter(
            Mandatory = $false
        )]
        [switch]    $Required
        
    )
    if ($Required)
    {
        While ( ($Null -eq $Variable) -or ($Variable -eq '') )
        {
            $Variable = Read-Host -Prompt "$Message"
            $Variable = $Variable.Trim()
        }
    }
    else
    {
        $Variable = Read-Host -Prompt "$Message"
        $Variable = $Variable.Trim()
    }
    Return $Variable
}