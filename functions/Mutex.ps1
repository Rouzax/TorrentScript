function New-Mutex
{
    <#
	.SYNOPSIS
	Create a Mutex
	.DESCRIPTION
	This function attempts to get a lock to a mutex with a given name. If a lock
	cannot be obtained this function waits until it can.

	Using mutexes, multiple scripts/processes can coordinate exclusive access
	to a particular work product. One script can create the mutex then go about
	doing whatever work is needed, then release the mutex at the end. All other
	scripts will wait until the mutex is released before they too perform work
	that only one at a time should be doing.

	This function outputs a PSObject with the following NoteProperties:

		Name
		Mutex

	Use this object in a followup call to Remove-Mutex once done.
	.PARAMETER MutexName
	The name of the mutex to create.
	.INPUTS
	None. You cannot pipe objects to this function.
	.OUTPUTS
	PSObject
	#Requires -Version 2.0
	#>
    [CmdletBinding()]
    [OutputType(
        [PSObject]
    )]
    Param(
        [Parameter(
            Mandatory = $true
        )]
        [ValidateNotNullOrEmpty()]
        [string]    $MutexName
    )

    $MutexWasCreated = $false
    $Mutex = $Null
    Write-Host "Waiting to acquire lock [$MutexName]..." -ForegroundColor DarkGray
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Threading')
    try
    {
        $Mutex = [System.Threading.Mutex]::OpenExisting($MutexName)
    }
    catch
    {
        $Mutex = New-Object System.Threading.Mutex($true, $MutexName, [ref]$MutexWasCreated)
    }
    try
    {
        if (!$MutexWasCreated)
        {
            $Mutex.WaitOne() | Out-Null 
        } 
    }
    catch
    { 
    }
    Write-Host "Lock [$MutexName] acquired. Executing..." -ForegroundColor DarkGray
    Write-Output ([PSCustomObject]@{ Name = $MutexName; Mutex = $Mutex })
} # New-Mutex

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
            Mandatory = $true
        )]
        [ValidateNotNull()]
        [PSObject]  $MutexObject
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