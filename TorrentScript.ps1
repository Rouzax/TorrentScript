param(
    [Parameter(
        mandatory = $false
    )]
    [string]    $DownloadPath, 

    [Parameter(
        mandatory = $false
    )]
    [string]    $DownloadLabel,

    [Parameter(
        mandatory = $false
    )]
    [string]    $TorrentHash,       

    [Parameter(
        Mandatory = $false
    )]
    [switch]    $NoCleanUp
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
$Functions = @( Get-ChildItem -Path $PSScriptRoot\functions\*.ps1  -ErrorAction SilentlyContinue )
# Dot source the files
ForEach ($import in @($Functions))
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
    $TotalSize = (Get-ChildItem -LiteralPath $DownloadPath -Recurse | Measure-Object -Property Length -sum).Sum
    $UnRarStopWatch = [system.diagnostics.stopwatch]::startNew()
    foreach ($Rar in $RarFilePaths)
    {
        Start-UnRar -UnRarSourcePath $Rar -UnRarTargetPath $ProcessPathFull
    }
    # Stop the Stopwatch
    $UnRarStopWatch.Stop() 
    Write-HTMLLog -Column1 'Size:' -Column2 (Format-Size -SizeInBytes $TotalSize)
    Write-HTMLLog -Column1 'Throughput:' -Column2 "$(Format-Size -SizeInBytes ($TotalSize/$UnRarStopWatch.Elapsed.TotalSeconds))/s"
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
    CleanProcessPath -Path $ProcessPathFull -NoCleanUp $NoCleanUp
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