<#
.SYNOPSIS
    Creates or opens a named mutex to control access to a shared resource.
.DESCRIPTION
    The New-Mutex function creates or opens a named mutex, providing a simple mechanism 
    for interprocess synchronization. It allows a script or function to acquire and release 
    a lock on a specified mutex, ensuring exclusive access to a shared resource.
.PARAMETER MutexName
    Specifies the name of the mutex. This parameter is mandatory and must be a non-null, 
    non-empty string.
.OUTPUTS 
    Returns a PSObject containing information about the created or opened mutex. 
    The PSObject includes the MutexName and the Mutex object itself.
.EXAMPLE
    PS C:\> $myMutex = New-Mutex -MutexName "MyMutex"
    PS C:\> # Perform actions requiring exclusive access to the shared resource
    PS C:\> Remove-Mutex -MutexObject $myMutex

    This example creates a new mutex named "MyMutex," acquires the lock, performs 
    actions requiring exclusive access, releases the lock, and continues with 
    other script or function logic.
#>

function New-Mutex {
    [CmdletBinding()]
    [OutputType(
        [PSObject]
    )]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$MutexName
    )

    # Variables for tracking mutex status
    $MutexWasCreated = $false
    $Mutex = $null

    Write-Host "Waiting to acquire lock [$MutexName]..." -ForegroundColor DarkGray
    
    # Attempt to open an existing mutex or create a new one
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Threading')
    try {
        $Mutex = [System.Threading.Mutex]::OpenExisting($MutexName)
    } catch {
        $Mutex = New-Object System.Threading.Mutex($true, $MutexName, [ref]$MutexWasCreated)
    }

    # Acquire the lock if the mutex was created successfully
    try {
        if (!$MutexWasCreated) {
            $Mutex.WaitOne() | Out-Null 
        } 
    } catch { 
        # Handle any errors during mutex acquisition
    }

    Write-Host "Lock [$MutexName] acquired. Executing..." -ForegroundColor DarkGray

    # Output a PSObject with mutex information
    Write-Output ([PSCustomObject]@{ Name = $MutexName; Mutex = $Mutex })
} 