function Remove-Mutex
{
    <#
	.SYNOPSIS
	Removes a previously created Mutex
	.DESCRIPTION
	This function attempts to release a lock on a mutex created by an earlier call
	to New-Mutex.
	.PARAMETER MutexObject
	The PSObject object as output by the New-Mutex function.
	.INPUTS
	None. You cannot pipe objects to this function.
	.OUTPUTS
	None.
	#Requires -Version 2.0
	#>
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory
        )]
        [ValidateNotNull()]
        [PSObject]$MutexObject
    )

    # $MutexObject | fl * | Out-String | Write-Host
    Write-Host "Releasing lock [$($MutexObject.Name)]..." -ForegroundColor DarkGray
    try
    {
        [void]$MutexObject.Mutex.ReleaseMutex() 
    }
    catch
    { 
    }
} # Remove-Mutex