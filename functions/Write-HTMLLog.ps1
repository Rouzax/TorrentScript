function Format-Table
{
    <#
    .SYNOPSIS
    Starts and stops Log file
    
    .DESCRIPTION
    Will either initiate the Log Variable in memory and open the HTML Table or clos the table
    
    .PARAMETER Start
    If defined indicated to open the HTML Table, without it the HTML table will be closed
    
    .EXAMPLE
    Format-Table -Start
    Format-Table
    
    .NOTES
    General notes
    #>
    param (
        [Parameter(Mandatory = $false)]
        [switch] $Start
    )
    if ($Start)
    {
        $global:Log = @()
        $global:Log += "<table border=`"0`" align=`"center`" cellspacing=`"0`""
        $global:Log += "<table border=`"0`" align=`"center`" cellspacing=`"0`""
        $global:Log += "style=`"border-collapse:collapse;background-color:#555555;color:#FFFFFF;font-family:arial,helvetica,sans-serif;font-size:10pt;`">"
        $global:Log += "<col width=`"125`">"
        $global:Log += "<col width=`"500`">"
        $global:Log += '<tbody>'
    }
    else
    {
        $global:Log += '</tbody>'
        $global:Log += '</table>'
    }
}

Function Write-HTMLLog
{
    <#
    .SYNOPSIS
    Add line to in memory log
    
    .DESCRIPTION
    Adds a line to the in memory log file and based on the parameters will do formating
    
    .PARAMETER Column1
    Text to be put in first column, mandatory
    
    .PARAMETER Column2
    Text to be put in the second column, not mandatory
    
    .PARAMETER Header
    Define that the Text from the parameter Column1 should be treated as new Header in the log table.
    If switch is defined Column2 is ignored
    
    .PARAMETER ColorBg
    Background color of Table Cell, this is a switch indicating a Success or Error. If not defined the standard color will be used.
    
    Success will get Green Table Cell color
    Error will get Red Table Cell Color
    
    .EXAMPLE
    Write-HTMLLog -Column1 '***  Header of the table  ***' -Header
    Write-HTMLLog -Column1 'Column1 Text' -Column2 'Column2 Text'
    Write-HTMLLog -Column1 'Exit Code:' -Column2 'Failed to do X' -ColorBg 'Error'
    Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
    
    .NOTES
    General notes
    #>
    Param(
        [Parameter(
            Mandatory = $true
        )]
        [string]    $Column1,

        [Parameter(
            Mandatory = $false
        )]
        [string]    $Column2,

        [Parameter(
            Mandatory = $false
        )]
        [switch]    $Header,

        [Parameter(
            Mandatory = $false
        )]
        [ValidateSet(
            'Success', 'Error'
        )]
        [string]    $ColorBg
    )

    $global:Log += '<tr>'
    if ($Header)
    {
        $global:Log += "<td colspan=`"2`" style=`"background-color:#398AA4;text-align:center;font-size:10pt`"><b>$Column1</b></td>"
    }
    else
    {
        if ($ColorBg -eq '')
        {
            $global:Log += "<td style=`"vertical-align:top;padding: 0px 10px;`"><b>$Column1</b></td>"
            $global:Log += "<td style=`"vertical-align:top;padding: 0px 10px;`">$Column2</td>"
            $global:Log += '</tr>'
        }
        elseif ($ColorBg -eq 'Success')
        {
            $global:Log += "<td style=`"vertical-align:top;padding: 0px 10px;`"><b>$Column1</b></td>"
            $global:Log += "<td style=`"vertical-align:top;padding: 0px 10px;background-color:#555000`">$Column2</td>"
            $global:Log += '</tr>'  
        }
        elseif ($ColorBg -eq 'Error')
        {
            $global:Log += "<td style=`"vertical-align:top;padding: 0px 10px;`"><b>$Column1</b></td>"
            $global:Log += "<td style=`"vertical-align:top;padding: 0px 10px;background-color:#550000`">$Column2</td>"
            $global:Log += '</tr>'  
        }
    }
    Write-Output "$Column1 $Column2"

}


function Write-Log
{
    <#
    .SYNOPSIS
    Write log to disk
    
    .DESCRIPTION
    Takes the Global Variable that hold the log in memory and writes it to disk
    
    .PARAMETER LogFile
    Log File including the Path to write
    
    .EXAMPLE
    Write-Log -LogFile 'C:\Temp\logfile.html'
    
    .NOTES
    General notes
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $LogFile
    )
    Set-Content -Path $LogFile -Value $global:Log
    
}