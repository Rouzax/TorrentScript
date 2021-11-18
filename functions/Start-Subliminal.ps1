function Start-Subliminal
{
    <#
    .SYNOPSIS
    Start Subliminal to download subs
    
    .DESCRIPTION
    Start Subliminal to download needed subtitles
    
    .PARAMETER Source
    Path to files that need subtitles
    
    .EXAMPLE
    Start-Subliminal -Source 'C:\Temp\Episode'
    
    .NOTES
    General notes
    #>
    param (
        [Parameter(
            Mandatory = $true
        )] 
        [string]    $Source
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

    Write-HTMLLog -Column1 '***  Download missing Subtitles  ***' -Header
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $SubliminalPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @('--opensubtitles', $OpenSubUser, $OpenSubPass, "--omdb $omdbAPI", 'download', '-r omdb', '-p opensubtitles', '-l eng', '-l nld', "`"$Source`"")
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    $Process.WaitForExit()
    # Write-Host $stdout
    # Write-Host $stderr
    if ($stdout -match '(\d+)(?=\s*video collected)')
    {
        $VideoCollected = $Matches.0
    }
    if ($stdout -match '(\d+)(?=\s*video ignored)')
    {
        $VideoIgnored = $Matches.0
    }
    if ($stdout -match '(\d+)(?=\s*error)')
    {
        $VideoError = $Matches.0
    }
    if ($stdout -match '(\d+)(?=\s*subtitle)')
    {
        $SubsDownloaded = $Matches.0
    }
    if ($stdout -match 'Some providers have been discarded due to unexpected errors')
    {
        $SubliminalExitCode = 1
    }
    if ($SubliminalExitCode -gt 0)
    {
        Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
        Write-HTMLLog -Column1 'Error:' -Column2 $stderr -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
    }
    else
    {
        if ($SubsDownloaded -gt 0)
        {
            # Write-HTMLLog -Column1 "Downloaded:" -Column2 "$SubsDownloaded Subtitles"
            Write-HTMLLog -Column1 'Collected:' -Column2 "$VideoCollected Videos"
            Write-HTMLLog -Column1 'Ignored:' -Column2 "$VideoIgnored Videos"
            Write-HTMLLog -Column1 'Error:' -Column2 "$VideoError Videos"
            Write-HTMLLog -Column1 'Downloaded:' -Column2 "$SubsDownloaded Subtitles"
            Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
        }
        else
        {
            Write-HTMLLog -Column1 'Result:' -Column2 'No subs downloaded with Subliminal'
        }
    }
}