<#
.SYNOPSIS
    Copies files using RoboCopy or falls back to PowerShell Copy if RoboCopy is not available.
.DESCRIPTION
    This function copies files from a source to a destination using RoboCopy or PowerShell Copy.
.PARAMETER Source
    Specifies the source path of the files to be copied.
.PARAMETER Destination
    Specifies the destination path where the files will be copied.
.PARAMETER File
    Specifies the filter for files to be copied. Use '*.*' for all files.
.PARAMETER DownloadLabel
    Specifies the label for the download operation (for logging purposes).
.PARAMETER DownloadName
    Specifies the name of the download operation (for logging purposes).
.OUTPUTS 
    None
.EXAMPLE
    Start-FileCopy -Source "C:\Source" -Destination "D:\Destination" -File '*.*' -DownloadLabel "Label" -DownloadName "Download"
    Copies all files from C:\Source to D:\Destination and logs the results.
.EXAMPLE
    Start-FileCopy -Source "C:\Source" -Destination "D:\Destination" -File 'example.txt' -DownloadLabel "Label" -DownloadName "Download"
    Copies a single file named 'example.txt' from C:\Source to D:\Destination and logs the results.
#>


function Start-FileCopy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$Destination,

        [Parameter(Mandatory = $true)]
        [string]$File,

        [Parameter(Mandatory = $true)] 
        [string]$DownloadLabel,

        [Parameter(Mandatory = $true)] 
        [string]$DownloadName
    )

    # Make sure needed functions are available otherwise try to load them.
    $functionsToLoad = @('Write-HTMLLog', 'Stop-Script', 'Format-Size')
    foreach ($functionName in $functionsToLoad) {
        if (-not (Get-Command $functionName -ErrorAction SilentlyContinue)) {
            try {
                . "$PSScriptRoot\$functionName.ps1"
                Write-Host "$functionName function loaded." -ForegroundColor Green
            } catch {
                Write-Error "Failed to import $functionName function: $_"
                exit 1
            }
        }
    }

    if (Test-Path "$env:SystemRoot\system32\Robocopy.exe") {

        # Set RoboCopy options based on the file parameter.
        $options = @('/R:1', '/W:1', '/J', '/NP', '/NJH', '/NFL', '/NDL', '/MT8')
        if ($File -eq '*.*') {
            $options += '/E'
        }
    
        $cmdArgs = @("`"$Source`"", "`"$Destination`"", "`"$File`"", $options)
 
        # Execute RoboCopy command
        Write-HTMLLog -Column1 'Starting:' -Column2 'RoboCopy files'
        try {
            $Output = robocopy @cmdArgs
        } catch {
            Write-Host 'Exception:' $_.Exception.Message -ForegroundColor Red
            Write-Host 'RoboCopy not found' -ForegroundColor Red
            exit 1
        }

        # Parse RoboCopy output
        foreach ($line in $Output) {
            switch -Regex ($line) {
                # Dir metrics
                '^\s+Dirs\s:\s*' {
                    # Example:  Dirs :        35         0         0         0         0         0
                    $dirs = $_.Replace('Dirs :', '').Trim()
                    # Now remove the white space between the values.'
                    $dirs = $dirs -split '\s+'
    
                    # Assign the appropriate column to values.
                    $TotalDirs = $dirs[0]
                    $CopiedDirs = $dirs[1]
                    $FailedDirs = $dirs[4]
                }
                # File metrics
                '^\s+Files\s:\s[^*]' {
                    # Example:  Files :      8318         0      8318         0         0         0
                    $files = $_.Replace('Files :', '').Trim()
                    # Now remove the white space between the values.'
                    $files = $files -split '\s+'
    
                    # Assign the appropriate column to values.
                    $TotalFiles = $files[0]
                    $CopiedFiles = $files[1]
                    $FailedFiles = $files[4]
                }
                # Byte metrics
                '^\s+Bytes\s:\s*' {
                    # Example:   Bytes :   1.607 g         0   1.607 g         0         0         0
                    $bytes = $_.Replace('Bytes :', '').Trim()
                    # Now remove the white space between the values.'
                    $bytes = $bytes -split '\s+'
    
                    # The raw text from the log file contains a k,m,or g after the non zero numbers.
                    # This will be used as a multiplier to determine the size in kb.
                    $counter = 0
                    $tempByteArray = 0, 0, 0, 0, 0, 0
                    $tempByteArrayCounter = 0
                    foreach ($column in $bytes) {
                        if ($column -eq 'k') {
                            $tempByteArray[$tempByteArrayCounter - 1] = '{0:N2}' -f ([single]($bytes[$counter - 1]) * 1024)
                            $counter += 1
                        } elseif ($column -eq 'm') {
                            $tempByteArray[$tempByteArrayCounter - 1] = '{0:N2}' -f ([single]($bytes[$counter - 1]) * 1048576)
                            $counter += 1
                        } elseif ($column -eq 'g') {
                            $tempByteArray[$tempByteArrayCounter - 1] = '{0:N2}' -f ([single]($bytes[$counter - 1]) * 1073741824)
                            $counter += 1
                        } else {
                            $tempByteArray[$tempByteArrayCounter] = $column
                            $counter += 1
                            $tempByteArrayCounter += 1
                        }
                    }
                    # Assign the appropriate column to values.
                    $TotalSize = Format-Size -SizeInBytes ([double]::Parse($tempByteArray[0]))
                    $CopiedSize = Format-Size -SizeInBytes ([double]::Parse($tempByteArray[1]))
                    $FailedSize = Format-Size -SizeInBytes ([double]::Parse($tempByteArray[4]))
                    # array columns 2,3, and 5 are available, but not being used currently.
                }
                # Speed metrics
                '^\s+Speed\s:.*sec.$' {
                    # Example:   Speed :             120.816 Bytes/min.
                    $speed = $_.Replace('Speed :', '').Trim()
                    $speed = $speed.Replace('Bytes/sec.', '').Trim()
                    # Remove any dots in the number
                    $speed = $speed.Replace('.', '').Trim()
                    # Assign the appropriate column to values.
                    $speed = Format-Size -SizeInBytes $speed
                }
            }
        }

        # Log results
        if ($FailedDirs -gt 0 -or $FailedFiles -gt 0) {
            Write-HTMLLog -Column1 'Dirs' -Column2 "$TotalDirs Total" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Dirs' -Column2 "$FailedDirs Failed" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Files:' -Column2 "$TotalFiles Total" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Files:' -Column2 "$FailedFiles Failed" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Size:' -Column2 "$TotalSize Total" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Size:' -Column2 "$FailedSize Failed" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            Stop-Script -ExitReason "Copy Error: $DownloadLabel - $DownloadName" 
        } else {
            Write-HTMLLog -Column1 'Dirs:' -Column2 "$CopiedDirs Copied"
            Write-HTMLLog -Column1 'Files:' -Column2 "$CopiedFiles Copied"
            Write-HTMLLog -Column1 'Size:' -Column2 "$CopiedSize"
            Write-HTMLLog -Column1 'Throughput:' -Column2 "$Speed/s"
            Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
        }
    } else {
        # RoboCopy not available, fallback to PowerShell copy
        Write-HTMLLog -Column1 'Starting:' -Column2 'Copy files using PowerShell Copy'
        try {
            # Start the Stopwatch
            $copyResultsStopWatch = [system.diagnostics.stopwatch]::startNew()
            $copyResults = Copy-Item -Path $Source -Destination $Destination -Filter $File -Recurse -PassThru -ErrorAction Stop
            # Stop the Stopwatch
            $copyResultsStopWatch.Stop() 
        
            # Log results for Copy-Item
            $totalFiles = ($copyResults | Where-Object { -not $_.PSIsContainer }).Count
            $totalDirs = ($copyResults | Where-Object { $_.PSIsContainer }).Count
            $totalSize = ($copyResults | Measure-Object Length -Sum).Sum
        
            Write-HTMLLog -Column1 'Dirs:' -Column2 $totalDirs
            Write-HTMLLog -Column1 'Files:' -Column2 "$totalFiles Copied"
            Write-HTMLLog -Column1 'Size:' -Column2 (Format-Size -SizeInBytes $totalSize)
            Write-HTMLLog -Column1 'Throughput:' -Column2 "$(Format-Size -SizeInBytes ($TotalSize/$copyResultsStopWatch.Elapsed.TotalSeconds))/s"
            Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
        } catch {
            # Log error for Copy-Item
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            Stop-Script -ExitReason "Copy Error: $DownloadLabel - $DownloadName"
        }
    }
}