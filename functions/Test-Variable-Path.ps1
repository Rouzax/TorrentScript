function Test-Variable-Path {
    <#
    .SYNOPSIS
    Test path to variables
    
    .DESCRIPTION
    Test path to
    
    .PARAMETER Path
    File path to test if exist
    
    .PARAMETER Name
    Variable that uses this path
    
    .EXAMPLE
    Test-Variable-Path -Path 'c:\Windows\notepad.exe' -Name 'NotepadPath'
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true
        )]
        [string]$Path,
        
        [Parameter(
            Mandatory = $true
        )]
        [string]$Name
    )
    if (!(Test-Path -LiteralPath $Path)) {
        Write-Host "Cannot find: $Path" -ForegroundColor Red
        Write-Host "As defined in variable: $Name" -ForegroundColor Red
        Write-Host 'Will now exit!' -ForegroundColor Red
        Exit 1
    } 
}