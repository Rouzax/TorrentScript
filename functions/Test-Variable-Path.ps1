function Test-Variable-Path {
    <#
    .SYNOPSIS
        Test path to variables.
    .DESCRIPTION
        This function tests the existence of a specified file path.
    .PARAMETER Path
        Specifies the file path to be tested for existence.
    .INPUTS
        Accepts a string representing the file path to be tested.
    .EXAMPLE
        Test-Variable-Path -Path 'c:\Windows\notepad.exe'
        # Checks if the specified file path exists.
    #>
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $true
        )]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-Host "Path cannot be empty or null." -ForegroundColor Red
        Write-Host 'Will now exit!' -ForegroundColor Red
        Exit 1
    }

    if (!(Test-Path -LiteralPath $Path)) {
        Write-Host "Cannot find: $Path" -ForegroundColor Red
        Write-Host "As defined in config" -ForegroundColor Red
        Write-Host 'Will now exit!' -ForegroundColor Red
        Exit 1
    } 
}
