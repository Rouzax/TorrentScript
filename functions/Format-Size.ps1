function Format-Size {
    <#
    .SYNOPSIS
        Formats a size in bytes into a human-readable format.
    .DESCRIPTION
        This function takes a size in bytes as input and converts it into a human-readable format, 
        displaying the size in terabytes (TB), gigabytes (GB), megabytes (MB), kilobytes (KB), 
        or bytes based on the magnitude of the input.
    .PARAMETER SizeInBytes
        Specifies the size in bytes that needs to be formatted.
    .INPUTS
        Accepts a double-precision floating-point number representing the size in bytes.
    .OUTPUTS 
        Returns a formatted string representing the size in TB, GB, MB, KB, or bytes.
    .EXAMPLE
        Format-Size -SizeInBytes 150000000000
        # Output: "139.81 GB"
        # Description: Formats 150,000,000,000 bytes into gigabytes.
    .EXAMPLE
        5000000 | Format-Size
        # Output: "4.77 MB"
        # Description: Pipes 5,000,000 bytes to the function and formats the size into megabytes.
    #>
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true, 
            ValueFromPipeline = $true
        )]
        [double]$SizeInBytes
    )

    switch ($SizeInBytes) {
        { $_ -ge 1PB } {
            # Convert to PB
            '{0:N2} PB' -f ($SizeInBytes / 1PB)
            break
        }
        { $_ -ge 1TB } {
            # Convert to TB
            '{0:N2} TB' -f ($SizeInBytes / 1TB)
            break
        }
        { $_ -ge 1GB } {
            # Convert to GB
            '{0:N2} GB' -f ($SizeInBytes / 1GB)
            break
        }
        { $_ -ge 1MB } {
            # Convert to MB
            '{0:N2} MB' -f ($SizeInBytes / 1MB)
            break
        }
        { $_ -ge 1KB } {
            # Convert to KB
            '{0:N2} KB' -f ($SizeInBytes / 1KB)
            break
        }
        default {
            # Display in bytes if less than 1KB
            "$SizeInBytes Bytes"
        }
    }
}