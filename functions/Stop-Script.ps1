function Stop-Script {
    <#
    .SYNOPSIS
    Stops the script, removes Mutex and send log file.
    
    .DESCRIPTION
    Stops the script and removes the Mutex.
    
    .PARAMETER ExitReason
    Reason for exit.
    
    .EXAMPLE
    Stop-Script -ExitReason "Script completed successfully."
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$ExitReason
    )  
    
    # Make sure needed functions are available otherwise try to load them.
    $commands = 'Write-HTMLLog', 'Send-HtmlMail', 'Remove-Mutex'
    foreach ($commandName in $commands) {
        if (!($command = Get-Command $commandName -ErrorAction SilentlyContinue)) {
            Try {
                . $PSScriptRoot\$commandName.ps1
                Write-Host "$commandName Function loaded." -ForegroundColor Green
            } Catch {
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
    # Handle Empty Download Label
    if ($DownloadLabel -ne 'NoMail') {
        try {
            $HTMLBody = Get-Content -Path $LogFilePath -Raw -ErrorAction Stop
        } catch {
            # Handle errors
            $errorMessage = $_.Exception.Message
            Write-Host "Error reading HTML content from file: $errorMessage"
            # Exit the script
            return
        }
        # Send-Mail -SMTPserver $SMTPserver -SMTPport $SMTPport -MailTo $MailTo -MailFrom $MailFrom -MailFromName $MailFromName -MailSubject $ExitReason -MailBody $LogFilePath -SMTPuser $SMTPuser -SMTPpass $SMTPpass
        Send-HtmlMail -SMTPServer $SMTPServer -SMTPServerPort $SMTPport -SmtpUser $SMTPuser -SmtpPassword $SMTPpass -To $MailTo -From "$MailFromName <$MailFrom>" -Subject $ExitReason -HTMLBody $HTMLBody
    }
    
    # Clean up the Mutex
    Remove-Mutex -MutexObject $ScriptMutex
    Exit
}