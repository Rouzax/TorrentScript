function Start-FileCopy {
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
        Specifies the filter for files to be copied. Use '*' for all files (avoid '*.*').
    .PARAMETER DownloadLabel
        Specifies the label for the download operation (for logging purposes).
    .PARAMETER DownloadName
        Specifies the name of the download operation (for logging purposes).
    .OUTPUTS 
        None
    .EXAMPLE
        Start-FileCopy -Source "C:\Source" -Destination "D:\Destination" -File '*' -DownloadLabel "Label" -DownloadName "Download"
        Copies all files from C:\Source to D:\Destination and logs the results.
    .EXAMPLE
        Start-FileCopy -Source "C:\Source" -Destination "D:\Destination" -File 'example.txt' -DownloadLabel "Label" -DownloadName "Download"
        Copies a single file named 'example.txt' from C:\Source to D:\Destination and logs the results.
    #>
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
            } catch {
                Write-Error "Failed to import $functionName function: $_"
                exit 1
            }
        }
    }

    # --- Local helpers ------------------------------------------------------
    function Convert-ToDoubleInvariant([string]$s) {
        if (-not $s) {
            return 0 
        }
        if ($s -match '[\.,].*,') {
            $s = $s -replace '\.', ''; $s = $s -replace ',', '.' 
        } elseif ($s -match ',') {
            $s = $s -replace '\.', ''; $s = $s -replace ',', '.' 
        } else {
            $s = $s -replace '(?<=\d)\.(?=\d{3}\b)', '' 
        }
        return [double]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture)
    }
    function Convert-ToBytes([double]$num, [string]$unit) {
        switch ($unit.ToLower()) {
            'k' {
                return $num * 1KB 
            }
            'm' {
                return $num * 1MB 
            }
            'g' {
                return $num * 1GB 
            }
            default {
                return $num 
            }
        }
    }
    function Get-RoboCopyExitDescription {
        param([int]$RC)
        $flags = @()
        if ($RC -band 1) {
            $flags += "Copied" 
        }
        if ($RC -band 2) {
            $flags += "Extras" 
        }
        if ($RC -band 4) {
            $flags += "Mismatched" 
        }
        if ($RC -band 8) {
            $flags += "FailedCopies" 
        }
        if ($RC -band 16) {
            $flags += "FatalError" 
        }
        if ($flags.Count -eq 0) {
            return "0 = No files copied / everything up to date" 
        }
        return "$RC = " + ($flags -join " + ")
    }

    if (Test-Path "$env:SystemRoot\system32\Robocopy.exe") {

        # Normalize file mask: '*' truly means "all files"
        $mask = if ($File -eq '*.*' -or [string]::IsNullOrWhiteSpace($File)) {
            '*' 
        } else {
            $File 
        }

        # Set RoboCopy options
        $options = @('/R:1', '/W:1', '/J', '/NP', '/NJH', '/NFL', '/NDL', '/MT:8')
        if ($mask -eq '*') {
            $options += '/E' 
        }  # include empty dirs when copying all

        # Build arguments (let PowerShell do the quoting)
        $cmdArgs = @($Source, $Destination, $mask) + $options

        Write-HTMLLog -Column1 'Starting:' -Column2 'RoboCopy files'
        try {
            # Capture stdout + stderr so PS7 parsing works
            $Output = & robocopy @cmdArgs 2>&1
            $rc = $LASTEXITCODE
        } catch {
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed (robocopy launch error)' -ColorBg 'Error'
            Stop-Script -ExitReason "Copy Error: $DownloadLabel - $DownloadName (robocopy launch failed)"
        }

        $rcDesc = Get-RoboCopyExitDescription $rc

        # Initialize metrics
        $TotalDirs = 0; $CopiedDirs = 0; $FailedDirs = 0
        $TotalFiles = 0; $CopiedFiles = 0; $FailedFiles = 0
        $TotalSize = '0 B'; $CopiedSize = '0 B'; $FailedSize = '0 B'
        $Speed = '0 B'

        # Parse RoboCopy output
        foreach ($line in $Output) {
            switch -Regex ($line) {
                '^\s+Dirs\s:\s*' {
                    $dirs = ($_.Replace('Dirs :', '').Trim() -split '\s+')
                    if ($dirs.Count -ge 6) {
                        $TotalDirs = [int]$dirs[0]; $CopiedDirs = [int]$dirs[1]; $FailedDirs = [int]$dirs[4] 
                    }
                }
                '^\s+Files\s:\s[^*]' {
                    $files = ($_.Replace('Files :', '').Trim() -split '\s+')
                    if ($files.Count -ge 6) {
                        $TotalFiles = [int]$files[0]; $CopiedFiles = [int]$files[1]; $FailedFiles = [int]$files[4] 
                    }
                }
                '^\s+Bytes\s:\s*' {
                    $t = ($_.Replace('Bytes :', '').Trim() -split '\s+')
                    try {
                        $totalBytes = Convert-ToBytes (Convert-ToDoubleInvariant $t[0]) $t[1]
                        $copiedBytes = Convert-ToBytes (Convert-ToDoubleInvariant $t[2]) $t[3]
                        $failedBytes = if ($t.Count -ge 10) {
                            Convert-ToBytes (Convert-ToDoubleInvariant $t[8]) $t[9] 
                        } else {
                            0 
                        }
                        $TotalSize = Format-Size -SizeInBytes $totalBytes
                        $CopiedSize = Format-Size -SizeInBytes $copiedBytes
                        $FailedSize = Format-Size -SizeInBytes $failedBytes
                    } catch { 
                    }
                }
                '^\s+Speed\s:.*sec\.$' {
                    $s = $_.Replace('Speed :', '').Replace('Bytes/sec.', '').Trim()
                    $s = $s -replace '[^\d,\.]', ''
                    if ($s.Contains(',')) {
                        $s = $s.Replace(',', '.') 
                    }
                    try {
                        $Speed = Format-Size -SizeInBytes ([double]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture)) 
                    } catch {
                        $Speed = '0 B' 
                    }
                }
            }
        }

        # Fallback if parsing didn't populate numbers (localization etc.)
        if ($TotalFiles -eq 0 -and $CopiedFiles -eq 0) {
            $destItems = Get-ChildItem -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue
            $TotalFiles = ($destItems | Where-Object { -not $_.PSIsContainer }).Count
            $TotalDirs = ($destItems | Where-Object { $_.PSIsContainer }).Count
            $CopiedFiles = $TotalFiles
            $CopiedDirs = $TotalDirs
            $CopiedSize = Format-Size -SizeInBytes (($destItems | Where-Object { -not $_.PSIsContainer } | Measure-Object Length -Sum).Sum)
        }

        # Decide status by RC and parsed failures
        if ($rc -ge 8) {
            Write-HTMLLog -Column1 'Dirs' -Column2 "$TotalDirs Total" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Files:' -Column2 "$TotalFiles Total" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Size:' -Column2 "$TotalSize Total" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Result:' -Column2 "Failed ($rcDesc)" -ColorBg 'Error'
            Stop-Script -ExitReason "Copy Error: $DownloadLabel - $DownloadName ($rcDesc)"
        } elseif ($FailedDirs -gt 0 -or $FailedFiles -gt 0) {
            Write-HTMLLog -Column1 'Dirs' -Column2 "$TotalDirs Total / $FailedDirs Failed" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Files:' -Column2 "$TotalFiles Total / $FailedFiles Failed" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Size:' -Column2 "$TotalSize Total / $FailedSize Failed" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Result:' -Column2 "Partial ($rcDesc)" -ColorBg 'Error'
            Stop-Script -ExitReason "Copy Error: $DownloadLabel - $DownloadName (Partial: $rcDesc)"
        } else {
            Write-HTMLLog -Column1 'Dirs:' -Column2 "$CopiedDirs Copied"
            Write-HTMLLog -Column1 'Files:' -Column2 "$CopiedFiles Copied"
            Write-HTMLLog -Column1 'Size:' -Column2 "$CopiedSize"
            Write-HTMLLog -Column1 'Throughput:' -Column2 "$Speed/s"
            Write-HTMLLog -Column1 'Result:' -Column2 "Successful ($rcDesc)" -ColorBg 'Success'
        }

    } else {
        # RoboCopy not available, fallback to PowerShell copy
        Write-HTMLLog -Column1 'Starting:' -Column2 'Copy files using PowerShell Copy'
        try {
            $copyResultsStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
            $copyResults = Copy-Item -LiteralPath $Source -Destination $Destination -Filter $File -Recurse -PassThru -ErrorAction Stop
            $copyResultsStopWatch.Stop() 
        
            $totalFiles = ($copyResults | Where-Object { -not $_.PSIsContainer }).Count
            $totalDirs = ($copyResults | Where-Object { $_.PSIsContainer }).Count
            $totalSize = ($copyResults | Measure-Object Length -Sum).Sum
        
            Write-HTMLLog -Column1 'Dirs:' -Column2 $totalDirs
            Write-HTMLLog -Column1 'Files:' -Column2 "$totalFiles Copied"
            Write-HTMLLog -Column1 'Size:' -Column2 (Format-Size -SizeInBytes $totalSize)
            Write-HTMLLog -Column1 'Throughput:' -Column2 "$(Format-Size -SizeInBytes ($totalSize / [Math]::Max(0.001, $copyResultsStopWatch.Elapsed.TotalSeconds)))/s"
            Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
        } catch {
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            Stop-Script -ExitReason "Copy Error: $DownloadLabel - $DownloadName"
        }
    }
}
