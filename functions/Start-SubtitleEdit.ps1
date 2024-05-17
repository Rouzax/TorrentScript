<#
.SYNOPSIS
    Starts Subtitle Edit to clean up subtitles.
.DESCRIPTION
    This function initiates Subtitle Edit, a free and open-source subtitle editor,
    to clean up subtitles for specified files.
.PARAMETER Source
    Specifies the source directory of the subtitle files.
.PARAMETER Files
    Specifies the file or wildcard pattern for subtitle files to process.
.PARAMETER SubtitleEditPath
    Specifies the path to the Subtitle Edit executable.
.PARAMETER DownloadLabel
    Specifies the label associated with the download.
.PARAMETER DownloadName
    Specifies the name of the downloaded content.
.OUTPUTS
None.
.EXAMPLE
    Start-SubtitleEdit -Source "C:\Subtitles" -Files "*.srt" -SubtitleEditPath "C:\SubtitleEdit.exe" -DownloadLabel "TV" -DownloadName "Episode1"
    # Initiates Subtitle Edit to clean up subtitles for TV Episode1.
.EXAMPLE
    Start-SubtitleEdit -Source "C:\Movies" -Files "*.sub" -SubtitleEditPath "C:\SubtitleEdit.exe" -DownloadLabel "Movie" -DownloadName "FilmA"
    # Initiates Subtitle Edit to clean up subtitles for Movie FilmA.
#>
function Start-SubtitleEdit {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] 
        [string]$Source,

        [Parameter(Mandatory = $true)] 
        [string]$Files,
    
        [Parameter(Mandatory = $true)] 
        [string]$SubtitleEditPath,

        [Parameter(Mandatory = $true)] 
        [string]$DownloadLabel,

        [Parameter(Mandatory = $true)] 
        [string]$DownloadName
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

    Write-HTMLLog -Column1 '***  Clean up Subtitles  ***' -Header
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $SubtitleEditPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @('/convert', "$Files", 'subrip', "/inputfolder`:`"$Source`"", '/overwrite', '/MergeSameTexts', '/fixcommonerrors', '/removetextforhi', '/fixcommonerrors', '/fixcommonerrors')
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    $stdout = $Process.StandardOutput.ReadToEnd()
    # $stderr = $Process.StandardError.ReadToEnd()
    $Process.WaitForExit()

    # Regular expressions to extract information
    $versionRegex = "Subtitle Edit (\d+\.\d+\.\d+)"
    $filesConvertedRegex = "(\d+) file\(s\) converted"
    $timeTakenRegex = "(\d+:\d+:\d+\.\d+)"

    # Extract information using regular expressions
    $matchResult = $stdout -match $versionRegex
    if ($matchResult) {
        $version = $matches[1]
    }

    $matchResult = $stdout -match $filesConvertedRegex
    if ($matchResult) {
        $filesConverted = $matches[1]
    }

    $matchResult = $stdout -match $timeTakenRegex
    if ($matchResult) {
        $timeTaken = $matches[1]
        # Convert the string to a TimeSpan object
        $timeSpan = [TimeSpan]::Parse($timeTaken)
        # Format the TimeSpan object as a string in a user-friendly way
        $userFriendlyTime = $timeSpan.ToString("hh\:mm\:ss")
    }

    switch ($process.ExitCode) {
        0 {
            Write-HTMLLog -Column1 'SubtitleEdit:' -Column2 "V $version"
            Write-HTMLLog -Column1 'Subtitles:' -Column2 "$filesConverted Converted"
            Write-HTMLLog -Column1 'Time Taken:' -Column2 $userFriendlyTime
            Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
        }
        default {
            Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
            Write-HTMLLog -Column1 'Error:' -Column2 $stdout -ColorBg 'Error'
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            Stop-Script -ExitReason "SubEdit Error: $DownloadLabel - $DownloadName"
        }
    }
}