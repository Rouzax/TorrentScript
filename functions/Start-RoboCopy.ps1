function Start-RoboCopy
{
    <#
    .SYNOPSIS
    RoboCopy wrapper
    .DESCRIPTION
    Wrapper for RoboCopy since it it way faster to copy that way
    .PARAMETER Source
    Source path
    .PARAMETER Destination
    Destination path
    .PARAMETER File
    File patern to copy *.* or name.ext
    This allows for folder or single filecopy 
    .EXAMPLE
    Start-RoboCopy -Source 'C:\Temp\Source' -Destination 'C:\Temp\Destination -File '*.*'
    Start-RoboCopy -Source 'C:\Temp\Source' -Destination 'C:\Temp\Destination -File 'file.ext'
    #>
    param (
        [Parameter(
            Mandatory = $true
        )]
        [string]    $Source,
        
        [Parameter(
            Mandatory = $true
        )]
        [string]    $Destination,
        
        [Parameter(
            Mandatory = $true
        )]
        [string]    $File
    )

    if ($File -ne '*.*')
    {
        $options = @('/R:1', '/W:1', '/J', '/NP', '/NP', '/NJH', '/NFL', '/NDL', '/MT8')
    }
    elseif ($File -eq '*.*')
    {
        $options = @('/R:1', '/W:1', '/E', '/J', '/NP', '/NJH', '/NFL', '/NDL', '/MT8')
    }
    
    $cmdArgs = @("`"$Source`"", "`"$Destination`"", "`"$File`"", $options)
 
    #executing unrar command
    Write-HTMLLog -Column1 'Starting:' -Column2 'Copy files'
    #executing Robocopy command
    $Output = robocopy @cmdArgs
  
    foreach ($line in $Output)
    {
        switch -Regex ($line)
        {
            #Dir metrics
            '^\s+Dirs\s:\s*'
            {
                #Example:  Dirs :        35         0         0         0         0         0
                $dirs = $_.Replace('Dirs :', '').Trim()
                #Now remove the white space between the values.'
                $dirs = $dirs -split '\s+'
    
                #Assign the appropriate column to values.
                $TotalDirs = $dirs[0]
                $CopiedDirs = $dirs[1]
                $FailedDirs = $dirs[4]
            }
            #File metrics
            '^\s+Files\s:\s[^*]'
            {
                #Example:  Files :      8318         0      8318         0         0         0
                $files = $_.Replace('Files :', '').Trim()
                #Now remove the white space between the values.'
                $files = $files -split '\s+'
    
                #Assign the appropriate column to values.
                $TotalFiles = $files[0]
                $CopiedFiles = $files[1]
                $FailedFiles = $files[4]
            }
            #Byte metrics
            '^\s+Bytes\s:\s*'
            {
                #Example:   Bytes :   1.607 g         0   1.607 g         0         0         0
                $bytes = $_.Replace('Bytes :', '').Trim()
                #Now remove the white space between the values.'
                $bytes = $bytes -split '\s+'
    
                #The raw text from the log file contains a k,m,or g after the non zero numers.
                #This will be used as a multiplier to determin the size in MB.
                $counter = 0
                $tempByteArray = 0, 0, 0, 0, 0, 0
                $tempByteArrayCounter = 0
                foreach ($column in $bytes)
                {
                    if ($column -eq 'k')
                    {
                        $tempByteArray[$tempByteArrayCounter - 1] = '{0:N2}' -f ([single]($bytes[$counter - 1]) / 1024)
                        $counter += 1
                    }
                    elseif ($column -eq 'm')
                    {
                        $tempByteArray[$tempByteArrayCounter - 1] = '{0:N2}' -f $bytes[$counter - 1]
                        $counter += 1
                    }
                    elseif ($column -eq 'g')
                    {
                        $tempByteArray[$tempByteArrayCounter - 1] = '{0:N2}' -f ([single]($bytes[$counter - 1]) * 1024)
                        $counter += 1
                    }
                    else
                    {
                        $tempByteArray[$tempByteArrayCounter] = $column
                        $counter += 1
                        $tempByteArrayCounter += 1
                    }
                }
                #Assign the appropriate column to values.
                $TotalMBytes = $tempByteArray[0]
                $CopiedMBytes = $tempByteArray[1]
                $FailedMBytes = $tempByteArray[4]
                #array columns 2,3, and 5 are available, but not being used currently.
            }
            #Speed metrics
            '^\s+Speed\s:.*sec.$'
            {
                #Example:   Speed :             120.816 MegaBytes/min.
                $speed = $_.Replace('Speed :', '').Trim()
                $speed = $speed.Replace('Bytes/sec.', '').Trim()
                #Assign the appropriate column to values.
                $speed = $speed / 1048576
                $SpeedMBSec = [math]::Round($speed, 2)
            }
        }
    }

    if ($FailedDirs -gt 0 -or $FailedFiles -gt 0)
    {
        Write-HTMLLog -Column1 'Dirs' -Column2 "$TotalDirs Total" -ColorBg 'Error'
        Write-HTMLLog -Column1 'Dirs' -Column2 "$FailedDirs Failed" -ColorBg 'Error'
        Write-HTMLLog -Column1 'Files:' -Column2 "$TotalFiles Total" -ColorBg 'Error'
        Write-HTMLLog -Column1 'Files:' -Column2 "$FailedFiles Failed" -ColorBg 'Error'
        Write-HTMLLog -Column1 'Size:' -Column2 "$TotalMBytes MB Total" -ColorBg 'Error'
        Write-HTMLLog -Column1 'Size:' -Column2 "$FailedMBytes MB Failed" -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Copy Error: $DownloadLabel - $DownloadName" 
    }
    else
    {
        Write-HTMLLog -Column1 'Dirs:' -Column2 "$CopiedDirs Copied"
        Write-HTMLLog -Column1 'Files:' -Column2 "$CopiedFiles Copied"
        Write-HTMLLog -Column1 'Size:' -Column2 "$CopiedMBytes MB"
        Write-HTMLLog -Column1 'Throughput:' -Column2 "$SpeedMBSec MB/s"
        Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
    }
}