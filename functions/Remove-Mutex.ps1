<#
.SYNOPSIS
    Releases the lock held by the specified mutex object.
.DESCRIPTION
    The Remove-Mutex function releases the lock held by the specified mutex object. 
    It is intended for use with mutex objects created by the New-Mutex function, 
    allowing the release of the lock and enabling other processes or scripts to 
    acquire the mutex.
.PARAMETER MutexObject
    Specifies the mutex object to release. This parameter is mandatory and must be 
    a valid PSObject containing mutex information, typically obtained from the 
    New-Mutex function.
.OUTPUTS 
    None. The function releases the lock on the specified mutex.
.EXAMPLE
    PS C:\> $myMutex = New-Mutex -MutexName "MyMutex"
    PS C:\> # Perform actions requiring exclusive access to the shared resource
    PS C:\> Remove-Mutex -MutexObject $myMutex

    This example creates a new mutex named "MyMutex," acquires the lock, performs 
    actions requiring exclusive access, and then releases the lock using the 
    Remove-Mutex function.
#>

function Remove-Mutex {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSObject]$MutexObject
    )

    # $MutexObject | fl * | Out-String | Write-Host
    Write-Host "Releasing lock [$($MutexObject.Name)]..." -ForegroundColor DarkGray
    try {
        # Release the Mutex
        [void]$MutexObject.Mutex.ReleaseMutex() 
        Write-Host "Lock released: $($MutexObject.Name)" -ForegroundColor DarkGray
    } catch { 
        Write-Warning "Failed to release lock: $($MutexObject.Name). $_"
    }
}