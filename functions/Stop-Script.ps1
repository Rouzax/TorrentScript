function Stop-Script {
    <#
    .SYNOPSIS
    Stops the script and removes Mutex
    
    .DESCRIPTION
    Stops the script and removes the Mutex
    
    .PARAMETER ExitReason
    Reason for exit
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true
        )]
        [string]$ExitReason
    )  
    
    # Make sure needed functions are available otherwise try to load them.
    $commands = 'Write-HTMLLog', 'Send-Mail', 'Remove-Mutex'
    foreach ($commandName in $commands) {
        if (!($command = Get-Command $commandName -ErrorAction SilentlyContinue)) {
            Try {
                . $PSScriptRoot\$commandName.ps1
                Write-Host "$commandName Function loaded." -ForegroundColor Green
            }
            Catch {
                Write-Error -Message "Failed to import $commandName function: $_"
                exit 1
            }
        }
    }
    # Start

    # Stop the Stopwatch
    $StopWatch.Stop()

    Write-HTMLLog -Column1 '***  Script Exection time  ***' -Header
    Write-HTMLLog -Column1 'Time Taken:' -Column2 $($StopWatch.Elapsed.ToString('mm\:ss'))
      
    Format-Table
    Write-Log -LogFile $LogFilePath
    Send-Mail -SMTPserver $SMTPserver -SMTPport $SMTPport -MailTo $MailTo -MailFrom $MailFrom -MailFromName $MailFromName -MailSubject $ExitReason -MailBody $LogFilePath -SMTPuser $SMTPuser -SMTPpass $SMTPpass
    
    # Clean up the Mutex
    Remove-Mutex -MutexObject $ScriptMutex
    Exit
}