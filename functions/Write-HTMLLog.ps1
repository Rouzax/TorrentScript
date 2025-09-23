# =========================
# HTML Logging (CSS Version)
# Compatible with Windows PowerShell 5.1+
# =========================

# Script-scope storage instead of globals
$script:HtmlLogEntries = @()
$script:HtmlLogStarted = $false

# Centralized CSS (adjust colors/fonts here)
$script:HtmlLogCss = @"
<style>
    table.html-log {
        border-collapse: collapse;
        background-color: #555555;
        color: #FFFFFF;
        font-family: Arial, Helvetica, sans-serif;
        font-size: 10pt;
        margin: auto;
    }
    table.html-log col:first-child { width: 125px; }
    table.html-log col:nth-child(2) { width: 500px; }

    table.html-log td {
        vertical-align: top;
        padding: 0px 10px;
    }
    table.html-log td.header {
        background-color: #398AA4;
        text-align: center;
        font-size: 10pt;
        font-weight: bold;
    }
    table.html-log td.success { background-color: #555000; }
    table.html-log td.error   { background-color: #550000; }
    table.html-log td.warning { background-color: #AA5500; }
</style>
"@

function ConvertTo-HtmlEncoded {
    param(
        [Parameter()] [AllowEmptyString()] [string]$Text
    )
    if ($null -eq $Text) { return '' }
    # Minimal HTML encoding (PS 5.1-safe)
    $Text -replace '&','&amp;' `
         -replace '<','&lt;' `
         -replace '>','&gt;' `
         -replace '"','&quot;' `
         -replace "'",'&#39;'
}

<#
.SYNOPSIS
Initialize or close the in-memory HTML log (compat shim).
#>
function Format-Table {
    [CmdletBinding()]
    param ([Parameter()][switch]$Start)
    if ($Start) {
        $script:HtmlLogEntries = @()
        $script:HtmlLogStarted = $true
    } else {
        # No-op now; rendering happens in Write-Log
        $script:HtmlLogStarted = $false
    }
}

<#
.SYNOPSIS
Add a row to the in-memory HTML log.
#>
function Write-HTMLLog {
    [CmdletBinding()]
    Param(
        [Parameter()] [string]$Column1,
        [Parameter()] [string]$Column2,
        [Parameter()] [switch]$Header,
        [Parameter()] [ValidateSet('Default','Success','Warning','Error')] [string]$ColorBg = 'Default'
    )

    if (-not $script:HtmlLogStarted) {
        # Auto-start if the caller forgot to
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

    # Console echo (kept from original behavior)
    $toShow = @($Column1, $Column2) | Where-Object { $_ -and $_.Trim() -ne '' }
    if ($toShow.Count) {
        Write-Host ($toShow -join ' ')
    }
}

<#
.SYNOPSIS
Render in-memory log entries to an HTML string.
#>
function Convert-HTMLLogToString {
    [CmdletBinding()]
    param()

    $rows = New-Object System.Collections.Generic.List[string]

    # Add CSS block once
    $rows.Add($script:HtmlLogCss)

    # Table wrapper
    $rows.Add('<table class="html-log" border="0" cellspacing="0">')
    $rows.Add('<col><col>')
    $rows.Add('<tbody>')

    foreach ($e in $script:HtmlLogEntries) {
        $rows.Add('<tr>')
        if ($e.Header) {
            $c1 = ConvertTo-HtmlEncoded $e.Column1
            $rows.Add("<td colspan=""2"" class=""header"">$c1</td>")
            $rows.Add('</tr>')
            continue
        }

        $c1 = ConvertTo-HtmlEncoded $e.Column1
        $c2 = ConvertTo-HtmlEncoded $e.Column2

        $rows.Add("<td><b>$c1</b></td>")

        $class = switch ($e.Type) {
            'Success' { 'success' }
            'Error'   { 'error' }
            'Warning' { 'warning' }
            default   { '' }
        }

        if ($class -ne '') {
            $rows.Add("<td class=""$class"">$c2</td>")
        } else {
            $rows.Add("<td>$c2</td>")
        }

        $rows.Add('</tr>')
    }

    $rows.Add('</tbody>')
    $rows.Add('</table>')

    return ($rows -join "`r`n")
}

<#
.SYNOPSIS
Write the HTML log to disk (UTF-8 without BOM).
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
