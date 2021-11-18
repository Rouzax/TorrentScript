function Start-UnRar
{
    <#
    .SYNOPSIS
    Unrar file
     .DESCRIPTION
    Takes rar file and unrar them to target
    .PARAMETER UnRarSourcePath
    Path of rar file to extract
    .PARAMETER UnRarTargetPath
    Destination folder path
    .EXAMPLE
    Start-UnRar -UnRarSourcePath 'C:\Temp\Source\file.rar' -UnRarTargetPath 'C:\Temp\Destination'
    #>
    Param( 
        [Parameter(
            Mandatory = $true
        )] 
        [string]    $UnRarSourcePath, 
       
        [Parameter(
            Mandatory = $true
        )] 
        [string]    $UnRarTargetPath
    )
 
    $RarFile = Split-Path -Path $UnRarSourcePath -Leaf
  
    #executing unrar command
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
    # $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    # Write-Host "stdout: $stdout"
    # Write-Host "stderr: $stderr"
    $Process.WaitForExit()
    if ($Process.ExitCode -gt 0)
    {
        Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
        Write-HTMLLog -Column1 'Error:' -Column2 $stderr -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Unrar Error: $DownloadLabel - $DownloadName"
    }
    else
    {
        Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
    }
}