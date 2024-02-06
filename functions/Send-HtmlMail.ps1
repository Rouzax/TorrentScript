<#
.SYNOPSIS
    Sends an HTML email using the specified SMTP server and credentials.
.DESCRIPTION
    This function sends an HTML email using the Send-MailKitMessage cmdlet. It requires the
    SMTP server details, sender's and recipient's email addresses, subject, and HTML body.
.PARAMETER SMTPServer
    Specifies the address of the SMTP server for sending the email.
.PARAMETER SMTPServerPort
    Specifies the port number to be used when connecting to the SMTP server.
.PARAMETER SmtpUser
    Specifies the username for authenticating with the SMTP server.
.PARAMETER SmtpPassword
    Specifies the password for authenticating with the SMTP server.
.PARAMETER To
    Specifies the email address of the recipient.
.PARAMETER From
    Specifies the email address of the sender.
.PARAMETER Subject
    Specifies the subject of the email.
.PARAMETER HTMLBody
    Specifies the HTML content of the email body.
.OUTPUTS 
    None. The function does not return any objects.
.EXAMPLE
    Send-HtmlMail -SMTPServer "smtp.example.com" -SMTPServerPort 587 -SmtpUser "user@example.com"
    -SmtpPassword "P@ssw0rd" -To "recipient@example.com" -From "sender@example.com"
    -Subject "Test Email" -HTMLBody "<p>This is a test email.</p>"
    Sends a test email with HTML content.
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

    # Send the email
    try {
        Send-MailKitMessage @Parameters
    } catch {
        # Handle errors
        $errorMessage = $_.Exception.Message
        Write-Host "Error sending email: $errorMessage" -ForegroundColor DarkRed
    }
}