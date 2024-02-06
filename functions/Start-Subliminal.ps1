<#
.SYNOPSIS
	Starts the Subliminal process to download missing subtitles for videos.
.DESCRIPTION
	This function initiates the Subliminal process to download subtitles for the specified video source.
.PARAMETER Source
	The path or URL of the video source for which subtitles are to be downloaded.
.PARAMETER OpenSubUser
	The username for accessing OpenSubtitles.org.
.PARAMETER OpenSubPass
	The password for accessing OpenSubtitles.org.
.PARAMETER omdbAPI
	The API key for accessing the OMDB API.
.PARAMETER WantedLanguages
	An array of language codes (e.g., "eng", "dut") for the desired subtitles.
.PARAMETER SubliminalPath
	The path to the Subliminal executable.
.OUTPUTS 
	Outputs a log with details of the Subliminal process, including the number of videos collected,
	ignored, errors, and subtitles downloaded.
.EXAMPLE
	Start-Subliminal -Source "C:\Videos\Sample.mp4" -OpenSubUser "user" -OpenSubPass "Password123"
		-omdbAPI "2028C39D" -WantedLanguages @("eng", "dut") -SubliminalPath "C:\Subliminal\subliminal.exe"
#>
function Start-Subliminal {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] 
        [string]$Source,

        [Parameter(Mandatory = $true)] 
        [string]$OpenSubUser,

        [Parameter(Mandatory = $true)] 
        [string]$OpenSubPass,

        [Parameter(Mandatory = $true)] 
        [string]$omdbAPI,

        [Parameter(Mandatory = $true)]
        [array]$WantedLanguages,  

        [Parameter(Mandatory = $true)] 
        [string]$SubliminalPath

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

    # Start Subliminal process
    Write-HTMLLog -Column1 '***  Download missing Subtitles  ***' -Header

    # Initialize arguments with common parameters
    $arguments = @('--opensubtitles', $OpenSubUser, $OpenSubPass, "--omdb $omdbAPI", 'download', '-r omdb', '-p opensubtitles')

    # Add language parameters to the arguments array
    foreach ($lang in $WantedLanguages) {
        $arguments += '-l', $lang
    }

    # Add the source parameter at the end
    $arguments += "`"$Source`""
    
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $SubliminalPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = $StartInfo.Arguments = $arguments
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    $Process.WaitForExit()

    # Process Subliminal output
    if ($stdout -match '(\d+)(?=\s*video collected)') {
        $VideoCollected = $Matches.0
    }
    if ($stdout -match '(\d+)(?=\s*video ignored)') {
        $VideoIgnored = $Matches.0
    }
    if ($stdout -match '(\d+)(?=\s*error)') {
        $VideoError = $Matches.0
    }
    if ($stdout -match '(\d+)(?=\s*subtitle)') {
        $SubsDownloaded = $Matches.0
    }
    if ($stdout -match 'Some providers have been discarded due to unexpected errors') {
        $SubliminalExitCode = 1
    }

    # Check for errors and log results
    if ($SubliminalExitCode -gt 0) {
        Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
        Write-HTMLLog -Column1 'Error:' -Column2 $stderr -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
    } else {
        if ($SubsDownloaded -gt 0) {
            # Write-HTMLLog -Column1 "Downloaded:" -Column2 "$SubsDownloaded Subtitles"
            Write-HTMLLog -Column1 'Collected:' -Column2 "$VideoCollected Videos"
            Write-HTMLLog -Column1 'Ignored:' -Column2 "$VideoIgnored Videos"
            Write-HTMLLog -Column1 'Error:' -Column2 "$VideoError Videos"
            Write-HTMLLog -Column1 'Downloaded:' -Column2 "$SubsDownloaded Subtitles"
            Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
        } else {
            Write-HTMLLog -Column1 'Result:' -Column2 'No subs downloaded with Subliminal'
        }
    }
}