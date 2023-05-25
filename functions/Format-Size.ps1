function Format-Size() {
    <#
    .SYNOPSIS
    Takes bytes and converts it to KB.MB,GB,TB,PB
    
    .DESCRIPTION
    Takes bytes and converts it to KB.MB,GB,TB,PB
    
    .PARAMETER SizeInBytes
    Input bytes
    
    .EXAMPLE
    Format-Size -SizeInBytes 864132
    843,88 KB
    	
    Format-Size -SizeInBytes 8641320
    8,24 MB
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true, 
            ValueFromPipeline = $true
        )]
        [double]$SizeInBytes
    )
    switch ([math]::Max($SizeInBytes, 0)) {
        { $_ -ge 1PB } {
            '{0:N2} PB' -f ($SizeInBytes / 1PB); break
        }
        { $_ -ge 1TB } {
            '{0:N2} TB' -f ($SizeInBytes / 1TB); break
        }
        { $_ -ge 1GB } {
            '{0:N2} GB' -f ($SizeInBytes / 1GB); break
        }
        { $_ -ge 1MB } {
            '{0:N2} MB' -f ($SizeInBytes / 1MB); break
        }
        { $_ -ge 1KB } {
            '{0:N2} KB' -f ($SizeInBytes / 1KB); break
        }
        default {
            "$SizeInBytes Bytes"
        }
    }
}