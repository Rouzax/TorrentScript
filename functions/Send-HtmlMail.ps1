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


function Send-HtmlMail {
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
        # Normalize recipients to an array
        $toList = $To -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

        # Credentials
        $secure = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
        $cred = [System.Management.Automation.PSCredential]::new($SmtpUser, $secure)

        # --- Detect whether HTMLBody is a file path or inline HTML ---
        $bodyHtml = $null
        $isFile = $false

        # If it's a valid local file path, use it
        if (Test-Path -LiteralPath $HTMLBody -PathType Leaf) {
            $isFile = $true
        } else {
            # Also allow file:/// URIs that point to local files
            if ([Uri]::IsWellFormedUriString($HTMLBody, [UriKind]::Absolute)) {
                $uri = [Uri]$HTMLBody
                if ($uri.Scheme -eq 'file' -and (Test-Path -LiteralPath $uri.LocalPath -PathType Leaf)) {
                    $HTMLBody = $uri.LocalPath
                    $isFile = $true
                }
            }
        }

        if ($isFile) {
            # Read whole file as UTF-8 (works well with UTF-8 & UTF-8-BOM)
            $bodyHtml = Get-Content -LiteralPath $HTMLBody -Raw -Encoding utf8
        } else {
            # Treat the provided value as inline HTML
            $bodyHtml = $HTMLBody
        }
        # -------------------------------------------------------------

        # Base params
        $params = @{
            Server     = $SMTPServer
            Port       = $SMTPServerPort
            From       = $From
            To         = $toList
            Subject    = $Subject
            BodyHtml   = $bodyHtml
            Credential = $cred
        }

        # Auto-select TLS mode by port
        switch ($SMTPServerPort) {
            465 {
                $params.SSL = $true 
            }                  # implicit TLS (SMTPS)
            587 {
                $params.UseSecureConnection = $true 
            }  # STARTTLS
            default {
                $params.UseSecureConnection = $true 
            } # try STARTTLS if supported
        }

        $mode = if ($params.SSL) {
            'SSL (implicit TLS)' 
        } elseif ($params.UseSecureConnection) {
            'STARTTLS' 
        } else {
            'Plain' 
        }
        $src = if ($isFile) {
            "HTML file '$HTMLBody'" 
        } else {
            'inline HTML' 
        }

        Write-Host "Sending mail via $($SMTPServer):$SMTPServerPort using $mode; body from $src..." -ForegroundColor DarkGray
        Send-EmailMessage @params
        Write-Host "Mail sent." -ForegroundColor DarkGray
    } catch {
        Write-Host ("Error sending email: " + $_.Exception.Message) -ForegroundColor DarkRed
    }
}
