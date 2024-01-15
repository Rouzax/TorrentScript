<#
.SYNOPSIS
    Sends an HTML email using the specified SMTP server and credentials.

.DESCRIPTION
    This function sends an HTML email using the specified SMTP server, port, credentials,
    recipient details, and email content.

.PARAMETER SMTPServer
    The SMTP server address.

.PARAMETER SMTPServerPort
    The SMTP server port.

.PARAMETER SmtpUser
    The username for SMTP authentication.

.PARAMETER SmtpPassword
    The password for SMTP authentication.

.PARAMETER To
    The email address of the recipient(s). Multiple recipients should be separated by commas.

.PARAMETER From
    The sender's email address.

.PARAMETER Subject
    The subject of the email.

.PARAMETER HTMLBody
    The HTML content of the email body.

.EXAMPLE
    Send-HtmlMail -SMTPServer 'smtp.example.com' -SMTPServerPort 587 -SmtpUser 'user@example.com'
                  -SmtpPassword 'password' -To 'recipient@example.com' -From 'sender@example.com'
                  -Subject 'Test Email' -HTMLBody '<p>This is a test email.</p>'
#>
function Send-HtmlMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SMTPServer,

        [Parameter(Mandatory = $true)]
        [int]$SMTPServerPort,   

        [Parameter(Mandatory = $true)]
        [string]$SmtpUser,

        [Parameter(Mandatory = $true)]
        [string]$SmtpPassword,

        [Parameter(Mandatory = $true)]
        [string]$To,

        [Parameter(Mandatory = $true)]
        [string]$From,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$HTMLBody
    )

    # Create PSCredential object
    $Credential = [System.Management.Automation.PSCredential]::new($SmtpUser, (ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force))

    # Define parameters for Send-MailKitMessage
    $Parameters = @{
        "UseSecureConnectionIfAvailable" = $true    
        "Credential"                     = $Credential
        "SMTPServer"                     = $SMTPServer
        "Port"                           = $SMTPServerPort
        "From"                           = $From
        "RecipientList"                  = $To
        "Subject"                        = $Subject
        "HTMLBody"                       = $HTMLBody
    }



    if (-not (Get-Module -Name Send-MailKitMessage -ListAvailable)) {
        Install-Module -Name Send-MailKitMessage -AllowClobber -Force -Confirm:$false -Scope CurrentUser
    } else {
        Import-Module -Name Send-MailKitMessage    
    }

    # Send the email
    try {
        Send-MailKitMessage @Parameters
    } catch {
        # Handle errors
        $errorMessage = $_.Exception.Message
        Write-Host "Error sending email: $errorMessage" -ForegroundColor DarkRed
    }
}