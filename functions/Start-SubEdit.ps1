function Start-SubEdit
{
    <#
    .SYNOPSIS
    Start Subtitle Edit
    
    .DESCRIPTION
    Start Subtitle Edit to clean up subtitles
    
    .PARAMETER Source
    Path to process and clean subtitles
    
    .PARAMETER Files
    The files to process, this will be *.srt typically
    
    .EXAMPLE
    Start-SubEdit -File '*.srt' -Source 'C:\Temp\Episode'
    
    .NOTES
    General notes
    #>
    param (
        [Parameter(
            Mandatory = $true
        )] 
        [string]    $Source,

        [Parameter(
            Mandatory = $true
        )] 
        [string]    $Files
    )

    # Make sure needed functions are available otherwise try to load them.
    $commands = 'Write-HTMLLog'
    foreach ($commandName in $commands)
    {
        if (!($command = Get-Command $commandName -ErrorAction SilentlyContinue))
        {
            Try
            {
                . $PSScriptRoot\$commandName.ps1
                Write-Host "$commandName Function loaded." -ForegroundColor Green
            }
            Catch
            {
                Write-Error -Message "Failed to import $commandName function: $_"
                exit 1
            }
        }
    }
    # Start

    Write-HTMLLog -Column1 '***  Clean up Subtitles  ***' -Header
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $SubtitleEditPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @('/convert', "$Files", 'subrip', "/inputfolder`:`"$Source`"", '/overwrite', '/fixcommonerrors', '/removetextforhi', '/fixcommonerrors', '/fixcommonerrors')
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    $Process.WaitForExit()
    if ($Process.ExitCode -gt 1)
    {
        Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
        Write-HTMLLog -Column1 'Error:' -Column2 $stderr -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "SubEdit Error: $DownloadLabel - $DownloadName"
    }
    else
    {
        Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
    }
}