function Send-HtmlMail {
    <#
    .SYNOPSIS
        Sends an HTML email via Mailozaurr; auto-detects inline HTML vs. HTML file and selects TLS by port.
    .DESCRIPTION
        Uses Mailozaurr's Send-EmailMessage. If -HTMLBody is a path (or file:/// URI) to an .html/.htm file, the file is read as UTF-8; 
        otherwise it's treated as inline HTML. TLS mode auto-selects: 465=SSL (implicit), 587=STARTTLS, others try STARTTLS.
    .PARAMETER SMTPServer
        SMTP host (e.g., smtp.example.com).
    .PARAMETER SMTPServerPort
        Submission port (465 implicit TLS, 587 STARTTLS).
    .PARAMETER SmtpUser
        SMTP username.
    .PARAMETER SmtpPassword
        SMTP password (converted to SecureString).
    .PARAMETER To
        Recipient(s); single or comma/semicolon-separated.
    .PARAMETER From
        Sender address.
    .PARAMETER Subject
        Email subject.
    .PARAMETER HTMLBody
        Inline HTML string or path/URI to an HTML file.
    .EXAMPLE
        Send-HtmlMail -SMTPServer smtp.example.com -SMTPServerPort 587 `
        -SmtpUser user@example.com -SmtpPassword "P@ssw0rd!" `
        -To "a@example.com;b@example.com" -From user@example.com `
        -Subject "Report" -HTMLBody "<h1>Hello</h1>"
    .EXAMPLE
        Send-HtmlMail -SMTPServer smtp.example.com -SMTPServerPort 465 `
        -SmtpUser user@example.com -SmtpPassword "P@ssw0rd!" `
        -To ops@example.com -From noreply@example.com `
        -Subject "Daily" -HTMLBody "C:\Reports\Daily.html"
    .OUTPUTS
        None.
    
    .NOTES
        Requires: Mailozaurr module. Recipients list supports comma/semicolon. HTML files read as UTF-8.
    .LINK
        https://github.com/EvotecIT/Mailozaurr
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SMTPServer,

        [Parameter(Mandatory)]
        [int]$SMTPServerPort,

        [Parameter(Mandatory)]
        [string]$SmtpUser,

        [Parameter(Mandatory)]
        [string]$SmtpPassword,

        [Parameter(Mandatory)]
        [string]$To,   # single or comma/semicolon-separated list

        [Parameter(Mandatory)]
        [string]$From,

        [Parameter(Mandatory)]
        [string]$Subject,

        [Parameter(Mandatory)]
        [string]$HTMLBody  # can be HTML string OR a path to an .html/.htm file
    )

    try {
        # Recipients → array
        $toList = $To -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

        # Credentials
        $secure = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
        $cred = [PSCredential]::new($SmtpUser, $secure)

        # Detect file vs inline HTML
        $isFile = $false
        if (Test-Path -LiteralPath $HTMLBody -PathType Leaf) {
            $isFile = $true
        } elseif ([Uri]::IsWellFormedUriString($HTMLBody, [UriKind]::Absolute)) {
            $uri = [Uri]$HTMLBody
            if ($uri.Scheme -eq 'file' -and (Test-Path -LiteralPath $uri.LocalPath -PathType Leaf)) {
                $HTMLBody = $uri.LocalPath
                $isFile = $true
            }
        }
        $bodyHtml = if ($isFile) {
            Get-Content -LiteralPath $HTMLBody -Raw -Encoding utf8 
        } else {
            $HTMLBody 
        }

        # Base params (note: HTML, not BodyHtml)
        $params = @{
            Server     = $SMTPServer
            Port       = $SMTPServerPort
            From       = $From
            To         = $toList
            Subject    = $Subject
            HTML       = $bodyHtml
            Credential = $cred
        }

        # TLS via MailKit SecureSocketOptions
        switch ($SMTPServerPort) {
            465 {
                $params.SecureSocketOptions = [MailKit.Security.SecureSocketOptions]::SslOnConnect 
            }
            587 {
                $params.SecureSocketOptions = [MailKit.Security.SecureSocketOptions]::StartTls 
            }
            default {
                $params.SecureSocketOptions = [MailKit.Security.SecureSocketOptions]::StartTlsWhenAvailable 
            }
        }

        $mode = $params.SecureSocketOptions
        $src = if ($isFile) {
            "HTML file '$HTMLBody'" 
        } else {
            'inline HTML' 
        }

        Write-Host "Sending mail via $($SMTPServer):$SMTPServerPort using $mode; body from $src..." -ForegroundColor DarkGray
        $null = Send-EmailMessage @params
        Write-Host "Mail sent." -ForegroundColor DarkGray
    } catch {
        Write-Host ("Error sending email: " + $_.Exception.Message) -ForegroundColor DarkRed
    }
}
