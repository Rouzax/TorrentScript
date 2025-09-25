function Start-UnRar {
    <#
    .SYNOPSIS
        Start-UnRar function extracts RAR files using WinRAR.
    .DESCRIPTION
        This function extracts RAR files specified by the UnRarSourcePath to the UnRarTargetPath.
        It logs the extraction process, handles errors, and stops the script if extraction fails.
    .PARAMETER UnRarSourcePath
        Specifies the path of the RAR file to be extracted.
    .PARAMETER UnRarTargetPath
        Specifies the target path where the contents of the RAR file will be extracted.
    .PARAMETER DownloadLabel
        Specifies the label associated with the download.
    .PARAMETER DownloadName
        Specifies the name of the download.
    .OUTPUTS
        Outputs the result of the extraction process.
    .EXAMPLE
        Start-UnRar -UnRarSourcePath "C:\Downloads\File.rar" -UnRarTargetPath "C:\Extracted" -DownloadLabel "TV" -DownloadName "Show1"
        # Extracts File.rar to C:\Extracted and logs the process.
    #>
    [CmdletBinding()]
    param( 
        [Parameter(Mandatory = $true)] 
        [string]$UnRarSourcePath, 
       
        [Parameter(Mandatory = $true)] 
        [string]$UnRarTargetPath,

        [Parameter(Mandatory = $true)] 
        [string]$DownloadLabel,

        [Parameter(Mandatory = $true)] 
        [string]$DownloadName
    )

    # Make sure needed functions are available otherwise try to load them.
    $functionsToLoad = @('Write-HTMLLog', 'Stop-Script')
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
    $RarFile = Split-Path -Path $UnRarSourcePath -Leaf
  
    # executing unrar command
    Write-HTMLLog -Column1 'File:' -Column2 "$RarFile"
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $WinRarPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @('x', "`"$UnRarSourcePath`"", "`"$UnRarTargetPath`"", '-y', '-idq')
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()

    $Process.WaitForExit()
    if ($Process.ExitCode -gt 0) {
        Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
        Write-HTMLLog -Column1 'Error:' -Column2 $stderr -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Unrar Error: $DownloadLabel - $DownloadName"
    } else {
        Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
    }
}