param(
    [Parameter(mandatory = $false)]
    [string] $DownloadPath, 
    [Parameter(mandatory = $false)]
    [string] $DownloadLabel,
    [Parameter(mandatory = $false)]
    [string] $TorrentHash,       
    [Parameter(Mandatory = $false)]
    [switch] $NoCleanUp
)

# User Variables
try
{
    $configPath = Join-Path $PSScriptRoot 'config.json'
    $Config = Get-Content $configPath -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}
catch
{
    Write-Host 'Exception:' $_.Exception.Message -ForegroundColor Red
    Write-Host 'Invalid config.json file' -ForegroundColor Red
    exit 1
}

# Log Date format
$LogFileDateFormat = Get-Date -Format $Config.DateFormat

# Script settings
# Temporary location of the files that are being processed, will be appended by the label and torrent name
$ProcessPath = $Config.ProcessPath

# Archive location of the log files of handeled donwloads
$LogArchivePath = $Config.LogArchivePath

# Default download root path
$DownloadRootPath = $Config.DownloadRootPath

# Label for TV Shows and Movies
$TVLabel = $Config.Label.TV
$MovieLabel = $Config.Label.Movie

# Additional tools that are needed
$WinRarPath = $Config.Tools.WinRarPath
$MKVMergePath = $Config.Tools.MKVMergePath
$MKVExtractPath = $Config.Tools.MKVExtractPath
$SubtitleEditPath = $Config.Tools.SubtitleEditPath
$SubliminalPath = $Config.Tools.SubliminalPath
$MailSendPath = $Config.Tools.MailSendPath

# Import Medusa Settings
$MedusaHost = $Config.Medusa.Host
$MedusaPort = $Config.Medusa.Port
$MedusaApiKey = $Config.Medusa.APIKey
$MedusaTimeOutMinutes = $Config.Medusa.TimeOutMinutes

# Import Radarr Settings
$RadarrHost = $Config.Radarr.Host
$RadarrPort = $Config.Radarr.Port
$RadarrApiKey = $Config.Radarr.APIKey
$RadarrTimeOutMinutes = $Config.Radarr.TimeOutMinutes

# Mail Settings
$MailTo = $Config.Mail.To
$MailFrom = $Config.Mail.From
$MailFromName = $Config.Mail.FromName
$SMTPserver = $Config.Mail.SMTPserver
$SMTPport = $Config.Mail.SMTPport
$SMTPuser = $Config.Mail.SMTPuser
$SMTPpass = $Config.Mail.SMTPpass

# OpenSubtitle User and Pass for Subliminal
$OpenSubUser = $Config.OpenSub.User
$OpenSubPass = $Config.OpenSub.Password
$omdbAPI = $Config.OpenSub.omdbAPI

# Language codes of subtitles to keep
$WantedLanguages = $Config.WantedLanguages
$SubtitleNamesToRemove = $Config.SubtitleNamesToRemove

# Get function definition files.
$functions = @( Get-ChildItem -Path $PSScriptRoot\functions\*.ps1  -ErrorAction SilentlyContinue )

# Dot source the files
ForEach ($import in @($functions))
{
    Try
    {
        # Lightweight alternative to dotsourcing a function script
        . ([ScriptBlock]::Create([System.Io.File]::ReadAllText($import)))
    }
    Catch
    {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}

# Function to Clean up subtitles
function Start-SubEdit
{
    param (
        [Parameter(Mandatory = $true)] 
        $Source,
        [Parameter(Mandatory = $true)] 
        $Files
    )
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

function Start-Subliminal
{
    param (
        [Parameter(Mandatory = $true)] 
        $Source
    )
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

# Fuction to Process Medusa
function Import-Medusa
{
    param (
        [Parameter(Mandatory = $true)] 
        $Source
    )

    $body = @{
        'proc_dir'       = $Source
        'resource'       = ''
        'process_method' = 'move'
        'force'          = $true
        'is_priority'    = $false
        'delete_on'      = $false
        'failed'         = $false
        'proc_type'      = 'manual'
        'ignore_subs'    = $false
    } | ConvertTo-Json

    $headers = @{
        'X-Api-Key' = $MedusaApiKey
    }
    Write-HTMLLog -Column1 '***  Medusa Import  ***' -Header
    try
    {
        $response = Invoke-RestMethod -Uri "http://$MedusaHost`:$MedusaPort/api/v2/postprocess" -Method Post -Body $Body -Headers $headers
    }
    catch
    {
        Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
    }
    if ($response.status -eq 'success')
    {
        $timeout = New-TimeSpan -Minutes $MedusaTimeOutMinutes
        $endTime = (Get-Date).Add($timeout)
        do
        {
            try
            {
                $status = Invoke-RestMethod -Uri "http://$MedusaHost`:$MedusaPort/api/v2/postprocess/$($response.queueItem.identifier)" -Method Get -Headers $headers
            }
            catch
            {
                Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
            }
            Start-Sleep 1
        }
        until ($status.success -or ((Get-Date) -gt $endTime))
        if ($status.success)
        {
            $ValuesToFind = 'Processing failed', 'aborting post-processing'
            $MatchPattern = ($ValuesToFind | ForEach-Object { [regex]::Escape($_) }) -join '|'
            if ($status.output -match $MatchPattern)
            {
                $ValuesToFind = 'Retrieving episode object for', 'Current quality', 'New quality', 'Old size', 'New size', 'Processing failed', 'aborting post-processing'
                $MatchPattern = ($ValuesToFind | ForEach-Object { [regex]::Escape($_) }) -join '|'
                foreach ($line in $status.output )
                {
                    if ($line -match $MatchPattern)
                    {
                        Write-HTMLLog -Column1 'Medusa:' -Column2 $line -ColorBg 'Error' 
                    }       
                }
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
                Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
            }
            else
            {
                Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'    
            }
        }
        else
        {
            Write-HTMLLog -Column1 'Medusa:' -Column2 $status.success -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Medusa:' -Column2 "Import Timeout: ($MedusaTimeOutMinutes) minutes" -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
        }
    }
    else
    {
        Write-HTMLLog -Column1 'Medusa:' -Column2 $response.status -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
    }
}

# Function to Process Radarr
function Import-Radarr
{
    param (
        [Parameter(Mandatory = $true)] 
        $Source
    )

    $body = @{
        'name'             = 'DownloadedMoviesScan'
        'downloadClientId' = $TorrentHash
        'importMode'       = 'Move'
        'path'             = $Source
    } | ConvertTo-Json

    $headers = @{
        'X-Api-Key' = $RadarrApiKey
    }
    Write-HTMLLog -Column1 '***  Radarr Import  ***' -Header
    try
    {
        $response = Invoke-RestMethod -Uri "http://$RadarrHost`:$RadarrPort/api/v3/command" -Method Post -Body $Body -Headers $headers
    }
    catch
    {
        Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
    }
    if ($response.status -eq 'queued' -or $response.status -eq 'started' -or $response.status -eq 'completed')
    {
        $timeout = New-TimeSpan -Minutes $RadarrTimeOutMinutes
        $endTime = (Get-Date).Add($timeout)
        do
        {
            try
            {
                $status = Invoke-RestMethod -Uri "http://$RadarrHost`:$RadarrPort/api/v3/command/$($response.id)" -Method Get -Headers $headers
            }
            catch
            {
                Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
            }
            Start-Sleep 1
        }
        until ($status.status -ne 'started' -or ((Get-Date) -gt $endTime) )
        if ($status.status -eq 'completed')
        {
            if ($status.duration -gt '00:00:05.0000000')
            {
                Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'         
            }
            else
            {
                Write-HTMLLog -Column1 'Radarr:' -Column2 'Completed but failed' -ColorBg 'Error' 
                Write-HTMLLog -Column1 'Radarr:' -Column2 'Radarr has no failed handling see: https://github.com/Radarr/Radarr/issues/5539' -ColorBg 'Error' 
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
                Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
            }
        }
        if ($status.status -eq 'failed')
        {
            Write-HTMLLog -Column1 'Radarr:' -Column2 $status.status -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Radarr:' -Column2 $status.exception -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
        }
        if ((Get-Date) -gt $endTime)
        {
            Write-HTMLLog -Column1 'Radarr:' -Column2 $status.status -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Radarr:' -Column2 "Import Timeout: ($RadarrTimeOutMinutes) minutes" -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
        }
    }
    else
    {
        Write-HTMLLog -Column1 'Radarr:' -Column2 $response.status -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
    }
}

# Function to close the log and send out mail
function Send-Mail
{
    param (
        [Parameter(Mandatory = $true)] 
        $MailSubject
    )
    # Close log file
    # Add-Content -LiteralPath  $LogFilePath -Value '</pre>'

    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $MailSendPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @("-smtp $SMTPserver", "-port $SMTPport", "-domain $SMTPserver", "-t $MailTo", "-f $MailFrom", "-fname `"$MailFromName`"", "-sub `"$MailSubject`"", 'body', "-file `"$LogFilePath`"", "-mime-type `"text/html`"", '-ssl', "auth -user $SMTPuser -pass $SMTPpass")
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    # $stdout = $Process.StandardOutput.ReadToEnd()
    # $stderr = $Process.StandardError.ReadToEnd()
    # Write-Host "stdout: $stdout"
    # Write-Host "stderr: $stderr"
    $Process.WaitForExit()
    # Write-Host "exit code: " + $p.ExitCode
    # return $stdout
}



# Fuction to clean up process folder 
function CleanProcessPath
{

    if ($NoCleanUp)
    {
        Write-HTMLLog -Column1 'Cleanup' -Column2 'NoCleanUp switch was given at command line, leaving files'
    }
    else
    {
        try
        {
            If (Test-Path -LiteralPath  $ProcessPathFull)
            {
                Remove-Item -Force -Recurse -LiteralPath $ProcessPathFull
            }
        }
        catch
        {
            Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        }
    }  
}


# Fuction to stop the script and send out the mail
function Stop-Script
{
    Param(
        [Parameter(Mandatory = $true)]
        [string] $ExitReason
    )         
    # Stop the Stopwatch
    $StopWatch.Stop()

    Write-HTMLLog -Column1 '***  Script Exection time  ***' -Header
    Write-HTMLLog -Column1 'Time Taken:' -Column2 $($StopWatch.Elapsed.ToString('mm\:ss'))
      
    Format-Table
    Write-Log -LogFile $LogFilePath
    Send-Mail -MailSubject $ExitReason
    # Clean up the Mutex

    Remove-Mutex -MutexObject $ScriptMutex
    Exit
}

function Test-Variable-Path
{
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path,
        [Parameter(Mandatory = $true)]
        [string] $Name
    )
    if (!(Test-Path -LiteralPath  $Path))
    {
        Write-Host "Cannot find: $Path" -ForegroundColor Red
        Write-Host "As defined in variable: $Name" -ForegroundColor Red
        Write-Host 'Will now exit!' -ForegroundColor Red
        Exit 1
    } 
}

# Test additional programs
Test-Variable-Path -Path $WinRarPath -Name 'WinRarPath'
Test-Variable-Path -Path $MKVMergePath -Name 'MKVMergePath'
Test-Variable-Path -Path $MKVExtractPath -Name 'MKVExtractPath'
Test-Variable-Path -Path $SubtitleEditPath -Name 'SubtitleEditPath'
Test-Variable-Path -Path $SubliminalPath -Name 'SubliminalPath'
Test-Variable-Path -Path $MailSendPath -Name 'MailSendPath'


# Get input if no parameters defined
# Build the download Location, this is the Download Root Path added with the Download name
if ( ($Null -eq $DownloadPath) -or ($DownloadPath -eq '') )
{
    $DownloadPath = Get-Input -Message 'Download Name' -Required
    $DownloadPath = Join-Path -Path $DownloadRootPath -ChildPath $DownloadPath 
}

# Download Name
$DownloadName = Split-Path -Path $DownloadPath -Leaf

# Download Label
if ( ($Null -eq $DownloadLabel) -or ($DownloadLabel -eq '') )
{
    $DownloadLabel = Get-Input -Message 'Download Label' 
}
# Torrent Hash (only needed for Radarr)
if ( ($Null -eq $TorrentHash) -or ($TorrentHash -eq '') )
{
    $TorrentHash = Get-Input -Message 'Torrent Hash' 
}

# Uppercase TorrentHash
$TorrentHash = $TorrentHash.ToUpper()

# Handle NoProcess Torrent Label
if ($DownloadLabel -eq 'NoProcess')
{
    Write-Host 'Do nothing'
    Exit
}

# Handle empty Torrent Label
if ($DownloadLabel -eq '')
{
    $DownloadLabel = 'NoLabel'
}

# Create Log file
# Log file of current processing file (will be used to send out the mail)
$LogFilePath = Join-Path -Path $LogArchivePath -ChildPath "$LogFileDateFormat-$DownloadName.html"

# Log Header
Format-Table -Start
Write-HTMLLog -Column1 '***  Information  ***' -Header
Write-HTMLLog -Column1 'Start:' -Column2 "$(Get-Date -Format 'yyyy-MM-dd') at $(Get-Date -Format 'HH:mm:ss')"
Write-HTMLLog -Column1 'Label:' -Column2 $DownloadLabel
Write-HTMLLog -Column1 'Name:' -Column2 $DownloadName
Write-HTMLLog -Column1 'Hash:' -Column2 $TorrentHash

# Test File Paths
If (!(Test-Path -LiteralPath  $ProcessPath))
{
    New-Item -ItemType Directory -Force -Path $ProcessPath | Out-Null
}
If (!(Test-Path -LiteralPath  $LogArchivePath))
{
    New-Item -ItemType Directory -Force -Path $LogArchivePath | Out-Null
}

# Start of script
$ScriptMutex = New-Mutex -MutexName 'DownloadScript'

# Start Stopwatch
$StopWatch = [system.diagnostics.stopwatch]::startNew()

# Check paths from Parameters
If (!(Test-Path -LiteralPath  $DownloadPath))
{
    Write-Host "$DownloadPath - Not valid location"
    Write-HTMLLog -Column1 'Path:' -Column2 "$DownloadPath - Not valid location" -ColorBg 'Error'
    Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
    Stop-Script -ExitReason "Path Error: $DownloadLabel - $DownloadName"
}

# Check if Single file or Folder 
$SingleFile = (Get-Item -LiteralPath $DownloadPath) -is [System.IO.FileInfo]
$Folder = (Get-Item -LiteralPath $DownloadPath) -is [System.IO.DirectoryInfo]

# Set Source and Destination paths and get Rar paths
if ($Folder)
{
    $ProcessPathFull = Join-Path -Path $ProcessPath -ChildPath $DownloadLabel | Join-Path -ChildPath $DownloadName
    $RarFilePaths = (Get-ChildItem -LiteralPath $DownloadPath -Recurse -Filter '*.rar').FullName
}
elseif ($SingleFile)
{
    $ProcessPathFull = Join-Path -Path $ProcessPath -ChildPath $DownloadLabel | Join-Path -ChildPath $DownloadName.Substring(0, $DownloadName.LastIndexOf('.'))
    $DownloadRootPath = Split-Path -Path $DownloadPath
    if ([IO.Path]::GetExtension($DownloadPath) -eq '.rar')
    {
        $RarFilePaths = (Get-Item -LiteralPath $DownloadPath).FullName
    } 
}

# Find rar files
$RarCount = $RarFilePaths.Count
if ($RarCount -gt 0)
{
    $RarFile = $true 
}
else
{
    $RarFile = $false 
}

# Check is destination folder exists otherwise create it
If (!(Test-Path -LiteralPath  $ProcessPathFull))
{
    New-Item -ItemType Directory -Force -Path $ProcessPathFull | Out-Null
}

if ($RarFile)
{
    Write-HTMLLog -Column1 '***  Unrar Download  ***' -Header
    Write-HTMLLog -Column1 'Starting:' -Column2 'Unpacking files'
    foreach ($Rar in $RarFilePaths)
    {
        Start-UnRar -UnRarSourcePath $Rar -UnRarTargetPath $ProcessPathFull
    }  
}
elseif (-not $RarFile -and $SingleFile)
{
    Write-HTMLLog -Column1 '***  Single File  ***' -Header
    Start-RoboCopy -Source $DownloadRootPath -Destination $ProcessPathFull -File $DownloadName
}
elseif (-not $RarFile -and $Folder)
{
    Write-HTMLLog -Column1 '***  Folder  ***' -Header
    Start-RoboCopy -Source $DownloadPath -Destination $ProcessPathFull -File '*.*'
}


# Starting Post Processing for Movies and TV Shows
if ($DownloadLabel -eq $TVLabel)
{
    $MKVFiles = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -Filter '*.mkv'
    $MKVCount = $MKVFiles.Count
    if ($MKVCount -gt 0)
    {
        $MKVFile = $true 
    }
    else
    {
        $MKVFile = $false 
    }
    if ($MKVFile)
    {
        # Download any missing subs
        Start-Subliminal -Source $ProcessPathFull
        
        # Remove unwanted subtitle languages and extract wanted subtitles and rename
        Start-MKV-Subtitle-Strip $ProcessPathFull
  
        # Clean up Subs
        Start-SubEdit -File '*.srt' -Source $ProcessPathFull
      
        Write-HTMLLog -Column1 '***  MKV Files  ***' -Header
        foreach ($Mkv in $MKVFiles)
        {
            Write-HTMLLog -Column1 ' ' -Column2 $Mkv.name
        }
        $SrtFiles = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -Filter '*.srt'
        if ($SrtFiles.Count -gt 0)
        {
            Write-HTMLLog -Column1 '***  Subtitle Files  ***' -Header
            foreach ($Srt in $SrtFiles)
            {
                Write-HTMLLog -Column1 ' ' -Column2 $srt.name
            }
        }
    }
    else
    {
        Write-HTMLLog -Column1 '***  Files  ***' -Header
        $Files = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -Filter '*.*'
        foreach ($File in $Files)
        {
            Write-HTMLLog -Column1 'File:' -Column2 $File.name
        }
    }

    # Call Medusa to Post Process
    Import-Medusa -Source $ProcessPathFull
    CleanProcessPath
    Stop-Script -ExitReason "$DownloadLabel - $DownloadName"
}
elseif ($DownloadLabel -eq $MovieLabel)
{
    $MKVFiles = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -Filter '*.mkv'
    $MKVCount = $MKVFiles.Count
    if ($MKVCount -gt 0)
    {
        $MKVFile = $true 
    }
    else
    {
        $MKVFile = $false 
    }
    if ($MKVFile)
    {
        # Download any missing subs
        Start-Subliminal -Source $ProcessPathFull

        # Extract wanted subtitles and rename
        Start-MKV-Subtitle-Strip $ProcessPathFull
          
        # Clean up Subs
        Start-SubEdit -File '*.srt' -Source $ProcessPathFull
        
        Write-HTMLLog -Column1 '***  MKV Files  ***' -Header
        foreach ($Mkv in $MKVFiles)
        {
            Write-HTMLLog -Column1 ' ' -Column2 $Mkv.name
        }
        Write-HTMLLog -Column1 '***  Subtitle Files  ***' -Header
        $SrtFiles = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -Filter '*.srt'
        foreach ($Srt in $SrtFiles)
        {
            Write-HTMLLog -Column1 ' ' -Column2 $srt.name
        }
    }
    else
    {
        Write-HTMLLog -Column1 '***  Files  ***' -Header
        $Files = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -Filter '*.*'
        foreach ($File in $Files)
        {
            Write-HTMLLog -Column1 'File:' -Column2 $File.name
        }
    }

    # Call Radarr to Post Process
    Import-Radarr -Source $ProcessPathFull
    CleanProcessPath
    Stop-Script -ExitReason "$DownloadLabel - $DownloadName"
}

# Reached the end of script
Write-HTMLLog -Column1 '***  Post Process General Download  ***' -Header
$Files = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -Filter '*.*'
foreach ($File in $Files)
{
    Write-HTMLLog -Column1 'File:' -Column2 $File.name
}
Stop-Script -ExitReason "$DownloadLabel - $DownloadName"