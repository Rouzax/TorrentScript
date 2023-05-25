function Send-Mail {
    <#
    .SYNOPSIS
    Send mail
    
    .DESCRIPTION
    Send email with file contents as body
    
    .PARAMETER SMTPserver
    SMTP server
    
    .PARAMETER SMTPport
    SMTP Port
    
    .PARAMETER MailTo
    Mail to
    
    .PARAMETER MailFrom
    Mail From
    
    .PARAMETER MailFromName
    Mail From Name
    
    .PARAMETER MailSubject
    Mail Subject
    
    .PARAMETER MailBody
    Path to html file as mail body
    
    .PARAMETER SMTPuser
    SMTP User
    
    .PARAMETER SMTPpass
    SMTP Password
    
    .EXAMPLE
    Send-Mail -SMTPserver 'mail.domain.com' -SMTPPort '25' -MailTo 'recipient@mail.com' -MailFrom 'sender@mail.com' -MailFromName 'Sender Name' -MailSubject 'Mail Subject' -MailBody 'C:\Temp\log.html' -SMTPuser 'user' -SMTPpass 'p@ssw0rd'
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true
        )] 
        [string]$SMTPserver,

        [Parameter(
            Mandatory = $true
        )] 
        [string]$SMTPport,

        [Parameter(
            Mandatory = $true
        )] 
        [string]$MailTo,

        [Parameter(
            Mandatory = $true
        )] 
        [string]$MailFrom,

        [Parameter(
            Mandatory = $true
        )] 
        [string]$MailFromName,

        [Parameter(
            Mandatory = $true
        )] 
        [string]$MailSubject,

        [Parameter(
            Mandatory = $true
        )] 
        [string]$MailBody,

        [Parameter(
            Mandatory = $true
        )] 
        [string]$SMTPuser,

        [Parameter(
            Mandatory = $true
        )] 
        [string]$SMTPpass
    )

    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $MailSendPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @("-smtp $SMTPserver", "-port $SMTPport", "-domain $SMTPserver", "-t $MailTo", "-f $MailFrom", "-fname `"$MailFromName`"", "-sub `"$MailSubject`"", 'body', "-file `"$MailBody`"", "-mime-type `"text/html`"", '-ssl', "auth -user $SMTPuser -pass $SMTPpass")
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    # $stdout = $Process.StandardOutput.ReadToEnd()
    # $stderr = $Process.StandardError.ReadToEnd()
    # Write-Host "stdout: $stdout"
    # Write-Host "stderr: $stderr"
    $Process.WaitForExit()
    # Write-Host "exit code: " + $p.ExitCode
    # return $stdout
}