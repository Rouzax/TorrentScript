function Start-MP4-2-MKV-Remux {
    <#
    .SYNOPSIS
    Converts MP4 video files to MKV format using mkvmerge.

    .DESCRIPTION
    This function remuxes MP4 video files to MKV format by extracting available tracks using mkvmerge. It deletes the original MP4 file upon successful conversion and saves the resulting MKV file.

    .PARAMETER Source
    Specifies the MP4 video file(s) to be converted to MKV format.

    .EXAMPLE
    Start-MP4-2-MKV-Remux -VideoFileObjects "C:\Videos\example.mp4"
    Remuxes the specified MP4 file to MKV format.

    .OUTPUTS
    This function does not return any specific output. It performs the conversion process and logs the result in an HTML log file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true
        )]
        $VideoFileObjects
    )

    # Make sure needed functions are available otherwise try to load them.
    $commands = 'Write-HTMLLog'
    foreach ($commandName in $commands) {
        if (!($command = Get-Command $commandName -ErrorAction SilentlyContinue)) {
            Try {
                . $PSScriptRoot\$commandName.ps1
                Write-Host "$commandName Function loaded." -ForegroundColor Green
            } Catch {
                Write-Error -Message "Failed to import $commandName function: $_"
                exit 1
            }
        }
    }
    
    # Start
    Write-HTMLLog -Column1 '***  Remux MP4 to MKV  ***' -Header
    $VideoFileObjects | ForEach-Object {
        Get-ChildItem -LiteralPath $_.FullName | ForEach-Object {
            $fileName = $_.BaseName
            $filePath = $_.FullName
            $fileRoot = $_.Directory

            # Start the json export with MKVMerge on the available tracks
            $TmpFileName = $fileName + '.tmp'
            $TmpMkvPath = Join-Path $fileRoot $TmpFileName
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVMergePath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            $StartInfo.Arguments = @("-o `"$TmpMkvPath`"", "`"$filePath`"")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            # $stderr = $Process.StandardError.ReadToEnd()
            # Write-Host "stdout: $stdout"
            # Write-Host "stderr: $stderr"
            $Process.WaitForExit()
            if ($Process.ExitCode -eq 2) {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            } elseif ($Process.ExitCode -eq 1) {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
            } elseif ($Process.ExitCode -eq 0) {
                # Delete original MP4
                Remove-Item -Force -LiteralPath $filePath
                # Overwrite original mkv after successful remux
                Rename-Item -LiteralPath $TmpMkvPath -NewName $($fileName + ".mkv")
                # Write-HTMLLog -Column1 "Removed:" -Column2 "$($SubIDsToRemove.Count) unwanted subtitle languages"
            } else {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
            }


        }
    }
}
