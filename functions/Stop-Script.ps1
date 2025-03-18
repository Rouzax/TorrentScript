<#
.SYNOPSIS
    Stop-Script function stops the script execution, logs the execution time, and sends an email notification.
.DESCRIPTION
    This function is designed to stop script execution, record the execution time, and send an email notification.
.PARAMETER ExitReason
    Mandatory parameter specifying the reason for stopping the script.
.EXAMPLE
    Stop-Script -ExitReason "Script completed successfully"
    Stops the script, logs execution time, and sends an email with the specified exit reason.
.NOTES
    This functions relies heavily on variables that have been set in the main script, 
    they are not passed as parameters to this functions
#>
function Stop-Script {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$ExitReason
    )  
    
    # Make sure needed functions are available otherwise try to load them.
    $functionsToLoad = @('Write-HTMLLog', 'Send-HtmlMail', 'Remove-Mutex')
    foreach ($functionName in $functionsToLoad) {
        if (-not (Get-Command $functionName -ErrorAction SilentlyContinue)) {
            try {
                . "$PSScriptRoot\$functionName.ps1"
                Write-Host "$functionName function loaded." -ForegroundColor Green
            } catch {
                Write-Error "Failed to import $functionName function: $_"
                exit 1
            }
        }
    }
    # Start

    # Stop the Stopwatch
    $ScriptTimer.Stop()

    Write-HTMLLog -Column1 '***  Script Execution time  ***' -Header
    Write-HTMLLog -Column1 'Time Taken:' -Column2 $($ScriptTimer.Elapsed.ToString('mm\:ss'))
      
    Format-Table
    Write-Log -LogFile $LogFilePath
    # Handle Empty Download Label
    if ($DownloadLabel -ne 'NoMail') {
        try {
            $HTMLBody = Get-Content -LiteralPath $LogFilePath -Raw -ErrorAction Stop
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