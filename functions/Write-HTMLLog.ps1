# =========================
# HTML Logging (Inline Style Version - Gmail Safe)
# =========================

# Script-scope storage instead of globals
$script:HtmlLogEntries = @()
$script:HtmlLogStarted = $false

<#
.SYNOPSIS
Encodes a string for safe HTML rendering.

.DESCRIPTION
Replaces special characters (&, <, >, ", ') with HTML entities so text does not break HTML markup.
Empty or null strings are allowed and return as empty.

.PARAMETER Text
The input string to encode.

.EXAMPLE
ConvertTo-HtmlEncoded -Text "<Hello & Goodbye>"
# Returns: &lt;Hello &amp; Goodbye&gt;
#>
function ConvertTo-HtmlEncoded {
    param(
        [Parameter()] [AllowEmptyString()] [string]$Text
    )
    if ($null -eq $Text) {
        return '' 
    }
    $Text -replace '&', '&amp;' `
        -replace '<', '&lt;' `
        -replace '>', '&gt;' `
        -replace '"', '&quot;' `
        -replace "'", '&#39;'
}

<#
.SYNOPSIS
Initialize or close the in-memory HTML log (compat shim).

.DESCRIPTION
When called with -Start, clears and initializes the in-memory log buffer.
Calling without -Start closes the buffer (no longer accepts entries).
Rendering is performed by Write-Log.

.PARAMETER Start
If specified, initializes/clears the in-memory log buffer.

.EXAMPLE
Format-Table -Start
# Initializes the log buffer.

.EXAMPLE
Format-Table
# Closes the log buffer.
#>
function Format-Table {
    [CmdletBinding()]
    param ([Parameter()][switch]$Start)
    if ($Start) {
        $script:HtmlLogEntries = @()
        $script:HtmlLogStarted = $true
    } else {
        $script:HtmlLogStarted = $false
    }
}

<#
.SYNOPSIS
Add a row to the in-memory HTML log.

.DESCRIPTION
Stores a structured log entry with two columns and optional header.
Visual emphasis is controlled by -ColorBg (Default, Success, Warning, Error).
Rendering to HTML occurs later with Convert-HTMLLogToString or Write-Log.

.PARAMETER Column1
Text for the first column (usually the label).

.PARAMETER Column2
Text for the second column (usually the value or message).

.PARAMETER Header
Treat this row as a table header (spans two columns). Column2 is ignored.

.PARAMETER ColorBg
Row type: Default, Success, Warning, or Error. Determines the background color of Column2.

.EXAMPLE
Write-HTMLLog -Column1 "Result:" -Column2 "Successful" -ColorBg Success

.EXAMPLE
Write-HTMLLog -Column1 "*** Information ***" -Header
#>
function Write-HTMLLog {
    [CmdletBinding()]
    param(
        [Parameter()] [string]$Column1,
        [Parameter()] [string]$Column2,
        [Parameter()] [switch]$Header,
        [Parameter()] [ValidateSet('Default', 'Success', 'Warning', 'Error')] [string]$ColorBg = 'Default'
    )

    if (-not $script:HtmlLogStarted) {
        $script:HtmlLogEntries = @()
        $script:HtmlLogStarted = $true
    }

    $entry = [PSCustomObject]@{
        Header  = [bool]$Header
        Type    = $ColorBg
        Column1 = $Column1
        Column2 = $Column2
    }

    $script:HtmlLogEntries += $entry

    # Console echo
    $toShow = @($Column1, $Column2) | Where-Object { $_ -and $_.Trim() -ne '' }
    if ($toShow.Count) {
        Write-Host ($toShow -join ' ')
    }
}

<#
.SYNOPSIS
Render in-memory log entries to an HTML string.

.DESCRIPTION
Generates a complete HTML <table> element with inline styles suitable for Gmail or browsers.
Includes all log rows, headers, and formatting.

.EXAMPLE
$html = Convert-HTMLLogToString
# Renders the log to an HTML string.

.EXAMPLE
Convert-HTMLLogToString | Out-File "C:\Logs\log.html"
# Saves the rendered HTML to a file.
#>
function Convert-HTMLLogToString {
    [CmdletBinding()]
    param()

    $rows = New-Object System.Collections.Generic.List[string]

    # Table wrapper
    $rows.Add('<table border="0" align="center" cellspacing="0" style="border-collapse:collapse;background-color:#555555;color:#FFFFFF;font-family:Arial,Helvetica,sans-serif;font-size:10pt;">')
    $rows.Add('<col width="125"><col width="500">')
    $rows.Add('<tbody>')

    foreach ($e in $script:HtmlLogEntries) {
        $rows.Add('<tr>')
        if ($e.Header) {
            $c1 = ConvertTo-HtmlEncoded $e.Column1
            $rows.Add("<td colspan=""2"" style=""background-color:#398AA4;text-align:center;font-size:10pt;font-weight:bold;"">$c1</td>")
            $rows.Add('</tr>')
            continue
        }

        $c1 = ConvertTo-HtmlEncoded $e.Column1
        $c2 = ConvertTo-HtmlEncoded $e.Column2

        $rows.Add("<td style=""vertical-align:top;padding:0px 10px;""><b>$c1</b></td>")

        switch ($e.Type) {
            'Success' {
                $rows.Add("<td style=""vertical-align:top;padding:0px 10px;background-color:#2C5000;"">$c2</td>")
            }
            'Error' {
                $rows.Add("<td style=""vertical-align:top;padding:0px 10px;background-color:#8C161F;"">$c2</td>")
            }
            'Warning' {
                $rows.Add("<td style=""vertical-align:top;padding:0px 10px;background-color:#925400;"">$c2</td>")
            }
            Default {
                $rows.Add("<td style=""vertical-align:top;padding:0px 10px;"">$c2</td>")
            }
        }

        $rows.Add('</tr>')
    }

    $rows.Add('</tbody>')
    $rows.Add('</table>')

    return ($rows -join "`r`n")
}

<#
.SYNOPSIS
Write the HTML log to disk.

.DESCRIPTION
Renders the current in-memory log to an HTML string and writes it to the specified file.
Uses UTF-8 without BOM encoding for compatibility with most tools and email systems.

.PARAMETER LogFile
The target file path (including .html extension).

.EXAMPLE
Write-Log -LogFile "C:\Temp\logfile.html"
# Saves the rendered log to an HTML file.

.EXAMPLE
Write-Log -LogFile "/var/www/html/log.html"
# On Linux, saves the rendered log to the web root.
#>
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$LogFile
    )
    $html = Convert-HTMLLogToString

    if ($PSVersionTable.PSVersion.Major -ge 6) {
        Set-Content -LiteralPath $LogFile -Value $html -Encoding utf8NoBOM
    } else {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($LogFile, $html, $utf8NoBom)
    }
}
