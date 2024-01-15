function Test-Variable-Path {
    <#
    .SYNOPSIS
    Test path to variables
    
    .DESCRIPTION
    Test path to
    
    .PARAMETER Path
    File path to test if exist
    
    .EXAMPLE
    Test-Variable-Path -Path 'c:\Windows\notepad.exe'
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    if (!(Test-Path -LiteralPath $Path)) {
        Write-Host "Cannot find: $Path" -ForegroundColor Red
        Write-Host "As defined in config" -ForegroundColor Red
        Write-Host 'Will now exit!' -ForegroundColor Red
        Exit 1
    } 
}