<#
.SYNOPSIS
    Remuxes MP4 files to MKV using MKVMerge.
.DESCRIPTION
    This function remuxes MP4 files to MKV format using MKVMerge. 
    It deletes the original MP4 file after successful remux and 
    overwrites the original MKV file.
.PARAMETER Source
    Specifies the path to the directory containing MP4 files to remux.
.PARAMETER MKVMergePath
    Specifies the path to the MKVMerge executable.
.OUTPUTS 
    Outputs the remuxed MKV files to the specified directory.
.EXAMPLE
    Start-MP4ToMKVRemux -Source "C:\Videos" -MKVMergePath "C:\Program Files\MKVToolNix\mkvmerge.exe"
    # Remuxes all MP4 files in the "C:\Videos" directory using MKVMerge.
#>
function Start-MP4ToMKVRemux {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$MKVMergePath
    )

    # Make sure needed functions are available otherwise try to load them.
    $functionsToLoad = @('Write-HTMLLog')
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
    
    # Start
    Write-HTMLLog -Column1 '***  Remux MP4 to MKV  ***' -Header

    Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.mp4' | Where-Object { $_.DirectoryName -notlike "*\Sample" } | ForEach-Object { 
        $fileName = $_.BaseName
        $filePath = $_.FullName
        $fileRoot = $_.Directory

        # Start MKVMerge to remux MP4 to MKV
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
        $Process.WaitForExit()

        switch ($process.ExitCode) {
            0 {
                # Delete original MP4
                Remove-Item -Force -LiteralPath $filePath
    
                # Overwrite original mkv after successful remux
                Rename-Item -LiteralPath $tmpMkvPath -NewName "$fileName.mkv"
                Write-HTMLLog -Column1 'Remuxed:' -Column2 "$fileName.mp4"
            }
            1 {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $process.ExitCode -ColorBg 'Warning'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Warning'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Warning'
            }
            default {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $process.ExitCode -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            }
        }
    }
    Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
}
