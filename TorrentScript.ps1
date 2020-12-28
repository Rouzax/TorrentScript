param(
    [Parameter(mandatory = $false)]
    [string] $DownloadPath, 
    [Parameter(mandatory = $false)]
    [string] $DownloadName,
    [Parameter(mandatory = $false)]
    [string] $DownloadLabel,
    [Parameter(mandatory = $false)]
    [string] $TorrentHash
)

# User Variables
try {
    $configPath = Join-Path $PSScriptRoot "config.json"
    $Config = Get-Content $configPath -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}
catch {
    Write-Host "Exception:" $_.Exception.Message -ForegroundColor Red
    Write-Host "Invalid config.json file" -ForegroundColor Red
    exit 1
}

# Log Date format
$LogFileDateFormat = Get-Date -Format $Config.DateFormat

# Script settings
# Temporary location of the files that are being processed, will be appended by the label and torrent name
$ProcessPath = $Config.ProcessPath

# Archive location of the log files of handeled donwloads
$LogArchivePath = $Config.LogArchivePath

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

# Functions
Function Get-Input {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Message,
        [Parameter(Mandatory = $false)]
        [switch] $Required
    )
    if ($Required) {
        While ( ($Null -eq $Variable) -or ($Variable -eq '') ) {
            $Variable = Read-Host -Prompt "$Message"
            $Variable = $Variable.Trim()
        }
    }
    else {
        $Variable = Read-Host -Prompt "$Message"
        $Variable = $Variable.Trim()
    }
    Return $Variable
}

function New-Mutex {
    <#
	.SYNOPSIS
	Create a Mutex
	.DESCRIPTION
	This function attempts to get a lock to a mutex with a given name. If a lock
	cannot be obtained this function waits until it can.

	Using mutexes, multiple scripts/processes can coordinate exclusive access
	to a particular work product. One script can create the mutex then go about
	doing whatever work is needed, then release the mutex at the end. All other
	scripts will wait until the mutex is released before they too perform work
	that only one at a time should be doing.

	This function outputs a PSObject with the following NoteProperties:

		Name
		Mutex

	Use this object in a followup call to Remove-Mutex once done.
	.PARAMETER MutexName
	The name of the mutex to create.
	.INPUTS
	None. You cannot pipe objects to this function.
	.OUTPUTS
	PSObject
	#Requires -Version 2.0
	#>
    [CmdletBinding()][OutputType([PSObject])]
    Param ([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$MutexName)
    $MutexWasCreated = $false
    $Mutex = $Null
    Write-Host "Waiting to acquire lock [$MutexName]..." -ForegroundColor DarkGray
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Threading')
    try {
        $Mutex = [System.Threading.Mutex]::OpenExisting($MutexName)
    }
    catch {
        $Mutex = New-Object System.Threading.Mutex($true, $MutexName, [ref]$MutexWasCreated)
    }
    try { if (!$MutexWasCreated) { $Mutex.WaitOne() | Out-Null } } catch { }
    Write-Host "Lock [$MutexName] acquired. Executing..." -ForegroundColor DarkGray
    Write-Output ([PSCustomObject]@{ Name = $MutexName; Mutex = $Mutex })
} # New-Mutex

function Remove-Mutex {
    <#
	.SYNOPSIS
	Removes a previously created Mutex
	.DESCRIPTION
	This function attempts to release a lock on a mutex created by an earlier call
	to New-Mutex.
	.PARAMETER MutexObject
	The PSObject object as output by the New-Mutex function.
	.INPUTS
	None. You cannot pipe objects to this function.
	.OUTPUTS
	None.
	#Requires -Version 2.0
	#>
    [CmdletBinding()]
    Param ([Parameter(Mandatory)][ValidateNotNull()][PSObject]$MutexObject)
    # $MutexObject | fl * | Out-String | Write-Host
    Write-Host "Releasing lock [$($MutexObject.Name)]..." -ForegroundColor DarkGray
    try { [void]$MutexObject.Mutex.ReleaseMutex() } catch { }
} # Remove-Mutex

# Robocopy function
function Start-RoboCopy {
    <#
    .SYNOPSIS
    RoboCopy wrapper
    .DESCRIPTION
    Wrapper for RoboCopy since it it way faster to copy that way
    .PARAMETER Source
    Source path
    .PARAMETER Destination
    Destination path
    .PARAMETER File
    File patern to copy *.* or name.ext
    This allows for folder or single filecopy 
    .EXAMPLE
    Start-RoboCopy -Source 'C:\Temp\Source' -Destination 'C:\Temp\Destination -File '*.*'
    Start-RoboCopy -Source 'C:\Temp\Source' -Destination 'C:\Temp\Destination -File 'file.ext'
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $Source,
        [Parameter(Mandatory = $true)]
        [string] $Destination,
        [Parameter(Mandatory = $true)]
        [string] $File
    )

    if ($File -ne '*.*') {
        $options = @("/R:1", "/W:1", "/J", "/NP", "/NP", "/NJH", "/NFL", "/NDL", "/MT8")
    }
    elseif ($File -eq '*.*') {
        $options = @("/R:1", "/W:1", "/E", "/J", "/NP", "/NJH", "/NFL", "/NDL", "/MT8")
    }
    
    $cmdArgs = @("`"$Source`"", "`"$Destination`"", "`"$File`"", $options)
 
    #executing unrar command
    Write-HTMLLog -LogFile $LogFilePath -Column1 "Starting:" -Column2 "Copy files"
    #executing Robocopy command
    $Output = robocopy @cmdArgs
  
    foreach ($line in $Output) {
        switch -Regex ($line) {
            #Dir metrics
            '^\s+Dirs\s:\s*' {
                #Example:  Dirs :        35         0         0         0         0         0
                $dirs = $_.Replace('Dirs :', '').Trim()
                #Now remove the white space between the values.'
                $dirs = $dirs -split '\s+'
    
                #Assign the appropriate column to values.
                $TotalDirs = $dirs[0]
                $CopiedDirs = $dirs[1]
                $FailedDirs = $dirs[4]
            }
            #File metrics
            '^\s+Files\s:\s[^*]' {
                #Example:  Files :      8318         0      8318         0         0         0
                $files = $_.Replace('Files :', '').Trim()
                #Now remove the white space between the values.'
                $files = $files -split '\s+'
    
                #Assign the appropriate column to values.
                $TotalFiles = $files[0]
                $CopiedFiles = $files[1]
                $FailedFiles = $files[4]
            }
            #Byte metrics
            '^\s+Bytes\s:\s*' {
                #Example:   Bytes :   1.607 g         0   1.607 g         0         0         0
                $bytes = $_.Replace('Bytes :', '').Trim()
                #Now remove the white space between the values.'
                $bytes = $bytes -split '\s+'
    
                #The raw text from the log file contains a k,m,or g after the non zero numers.
                #This will be used as a multiplier to determin the size in MB.
                $counter = 0
                $tempByteArray = 0, 0, 0, 0, 0, 0
                $tempByteArrayCounter = 0
                foreach ($column in $bytes) {
                    if ($column -eq 'k') {
                        $tempByteArray[$tempByteArrayCounter - 1] = "{0:N2}" -f ([single]($bytes[$counter - 1]) / 1024)
                        $counter += 1
                    }
                    elseif ($column -eq 'm') {
                        $tempByteArray[$tempByteArrayCounter - 1] = "{0:N2}" -f $bytes[$counter - 1]
                        $counter += 1
                    }
                    elseif ($column -eq 'g') {
                        $tempByteArray[$tempByteArrayCounter - 1] = "{0:N2}" -f ([single]($bytes[$counter - 1]) * 1024)
                        $counter += 1
                    }
                    else {
                        $tempByteArray[$tempByteArrayCounter] = $column
                        $counter += 1
                        $tempByteArrayCounter += 1
                    }
                }
                #Assign the appropriate column to values.
                $TotalMBytes = $tempByteArray[0]
                $CopiedMBytes = $tempByteArray[1]
                $FailedMBytes = $tempByteArray[4]
                #array columns 2,3, and 5 are available, but not being used currently.
            }
            #Speed metrics
            '^\s+Speed\s:.*sec.$' {
                #Example:   Speed :             120.816 MegaBytes/min.
                $speed = $_.Replace('Speed :', '').Trim()
                $speed = $speed.Replace('Bytes/sec.', '').Trim()
                #Assign the appropriate column to values.
                $speed = $speed / 1048576
                $SpeedMBSec = [math]::Round($speed, 2)
            }
        }
    }

    if ($FailedDirs -gt 0 -or $FailedFiles -gt 0) {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Dirs" -Column2 "$TotalDirs Total" -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Dirs" -Column2 "$FailedDirs Failed" -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Files:" -Column2 "$TotalFiles Total" -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Files:" -Column2 "$FailedFiles Failed" -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Size:" -Column2 "$TotalMBytes MB Total" -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Size:" -Column2 "$FailedMBytes MB Failed" -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
        Stop-Script -ExitReason "Copy Error: $DownloadLabel - $DownloadName" 
    }
    else {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Dirs:" -Column2 "$CopiedDirs Copied"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Files:" -Column2 "$CopiedFiles Copied"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Size:" -Column2 "$CopiedMBytes MB"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Throughput:" -Column2 "$SpeedMBSec MB/s"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Successful" -ColorBg "Success"
    }
}

# Function to unrar
function Start-UnRar {
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
        [Parameter(Mandatory = $true)] 
        $UnRarSourcePath, 
        [Parameter(Mandatory = $true)] 
        $UnRarTargetPath
    )
 
    $RarFile = split-path -path $UnRarSourcePath -Leaf
  
    #executing unrar command
    Write-HTMLLog -LogFile $LogFilePath -Column1 "File:" -Column2 "$RarFile"
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $WinRarPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @("x", "`"$UnRarSourcePath`"", "`"$UnRarTargetPath`"", "-y", "-idq")
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    # $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    # Write-Host "stdout: $stdout"
    # Write-Host "stderr: $stderr"
    $Process.WaitForExit()
    if ($Process.ExitCode -gt 0) {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Error:" -Column2 $stderr -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
        Stop-Script -ExitReason "Unrar Error: $DownloadLabel - $DownloadName"
    }
    else {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Successful" -ColorBg "Success"
    }
}

function Start-MKV-Subtitle-Strip {
    <#
    .SYNOPSIS
    Extract wanted SRT subtitles from MKVs in root folder and remux MKVs to strip out unwanted subtitle languages
    .DESCRIPTION
    Searches for all MKV files in root folder and extracts the SRT that are defined in $WantedLanguages 
    Remux any MKV that has subtitles of unwanted languages or Track_name in $SubtitleNamesToRemove.
    Rename srt subtitle files based on $LanguageCodes
    .PARAMETER Source
    Defines the root folder to start the search for MKV files 
    .EXAMPLE
    Start-MKV-Subtitle-Strip 'C:\Temp\Source'
	.OUTPUTS
    SRT files of the desired languages
    file.en.srt
    file.2.en.srt
    file.nl.srt
    file.de.srt
    MKV files without the unwanted subtitles
    file.mkv
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $Source
    )
    $episodes = @()
    $SubsExtracted = $false
    $TotalSubsToKeep = 0
    $TotalSubsToRemove = 0

    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Extract srt files from MKV  ***" -Header
    Get-ChildItem -LiteralPath $Source -Recurse -filter "*.mkv" | ForEach-Object {
        Get-ChildItem -LiteralPath $_.FullName | ForEach-Object {
            $fileName = $_.BaseName
            $filePath = $_.FullName
            $fileRoot = $_.Directory

            # Start the json export with MKVMerge on the available tracks
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVMergePath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            $StartInfo.Arguments = @("-J", "`"$filePath`"")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            # $stderr = $Process.StandardError.ReadToEnd()
            # Write-Host "stdout: $stdout"
            # Write-Host "stderr: $stderr"
            $Process.WaitForExit()
            if ($Process.ExitCode -eq 2) {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvmerge:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
            }
            elseif ($Process.ExitCode -eq 1) {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvmerge:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Warning" -ColorBg "Error"
            }
            elseif ($Process.ExitCode -eq 0) {
                $fileMetadata = $stdout | ConvertFrom-Json
            }
            else {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvmerge:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Warning" -ColorBg "Error"
            }

            $file = @{
                FileName        = $fileName
                FilePath        = $filePath
                FileRoot        = $fileRoot
                FileTracks      = $fileMetadata.tracks
                FileAttachments = $fileMetadata.attachments
            }

            $episodes += New-Object PSObject -Property $file
        }
    }

    # Exctract wanted SRT subtitles
    $episodes | ForEach-Object {
        $episode = $_
        $SubIDsToKeep = @()
        $SubIDsToRemove = @()
        $SubsToExtract = @()
        $SubNamesToKeep = @()

        $episode.FileTracks | ForEach-Object {
            $FileTrack = $_
            if ($FileTrack.id) {
                # Check if subtitle is srt
                if ($FileTrack.type -eq "subtitles" -and $FileTrack.codec -eq "SubRip/SRT") {
                    
                    # Check to see if track_name is part of $SubtitleNamesToRemove list
                    if ($null -ne ($SubtitleNamesToRemove | Where-Object { $FileTrack.properties.track_name -match $_ })) {
                        $SubIDsToRemove += $FileTrack.id
                    }
                    # Check is subtitle is in $WantedLanguages list
                    elseif ($FileTrack.properties.language -in $WantedLanguages) {

                        # Handle multiple subtitles of same language, if exist append ID to file 
                        if ("$($episode.FileName).$($FileTrack.properties.language).srt" -in $SubNamesToKeep) {
                            $prefix = "$($FileTrack.id).$($FileTrack.properties.language)"
                        }
                        else {
                            $prefix = "$($FileTrack.properties.language)"
                        }
    
                        # Add Subtitle name and ID to be extracted
                        $SubsToExtract += "`"$($FileTrack.id):$($episode.FileRoot)\$($episode.FileName).$($prefix).srt`""
                        
                        # Keep track of subtitle file names that will be extracted to handle possible duplicates
                        $SubNamesToKeep += "$($episode.FileName).$($prefix).srt"

                        # Add subtitle ID to for MKV remux
                        $SubIDsToKeep += $FileTrack.id
                    }
                    else {
                        $SubIDsToRemove += $FileTrack.id
                    }
                }
            }
        }

        # Count all subtitles to keep and remove of logging
        $TotalSubsToKeep = $TotalSubsToKeep + $SubIDsToKeep.count
        $TotalSubsToRemove = $TotalSubsToRemove + $SubIDsToRemove.count
        
        # Extract the wanted subtitle languages
        if ($SubIDsToKeep.count -gt 0) {
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVExtractPath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            $StartInfo.Arguments = @("`"$($episode.FilePath)`"", "tracks", "$SubsToExtract")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            # $stderr = $Process.StandardError.ReadToEnd()
            # Write-Host "stdout: $stdout"
            # Write-Host "stderr: $stderr"
            $Process.WaitForExit()
            if ($Process.ExitCode -eq 2) {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvextract:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
            }
            elseif ($Process.ExitCode -eq 1) {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvextract:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Warning" -ColorBg "Error"
            }
            elseif ($Process.ExitCode -eq 0) {
                $SubsExtracted = $true
                # Write-HTMLLog -LogFile $LogFilePath -Column1 "Extracted:" -Column2 "$($SubsToExtract.count) Subtitles"
            }
            else {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvextract:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Unknown" -ColorBg "Error"
            }
        }

        # Remux and strip out all unwanted subtitle languages
        if ($SubIDsToRemove.Count -gt 0) {
            $TmpFileName = $Episode.FileName + ".tmp"
            $TmpMkvPath = Join-Path $episode.FileRoot $TmpFileName
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVMergePath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            $StartInfo.Arguments = @("-o `"$TmpMkvPath`"", "-s $($SubIDsToKeep -join ",")", "`"$($episode.FilePath)`"")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            # $stderr = $Process.StandardError.ReadToEnd()
            # Write-Host "stdout: $stdout"
            # Write-Host "stderr: $stderr"
            $Process.WaitForExit()
            if ($Process.ExitCode -eq 2) {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvmerge:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
            }
            elseif ($Process.ExitCode -eq 1) {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvmerge:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Warning" -ColorBg "Error"
            }
            elseif ($Process.ExitCode -eq 0) {
                # Overwrite original mkv after successful remux
                Move-Item -Path $TmpMkvPath -Destination $($episode.FilePath) -Force
                # Write-HTMLLog -LogFile $LogFilePath -Column1 "Removed:" -Column2 "$($SubIDsToRemove.Count) unwanted subtitle languages"
            }
            else {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "mkvmerge:" -Column2 $stdout -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Warning" -ColorBg "Error"
            }
        }
    }
   
    # Rename extracted subs to correct 2 county code based on $LanguageCodes
    if ($SubsExtracted) {
        $SrtFiles = Get-ChildItem -LiteralPath $Source -Recurse -filter "*.srt"
        foreach ($srt in $SrtFiles) {
            $FileDirectory = $srt.Directory
            $FilePath = $srt.FullName
            $FileName = $srt.Name
            foreach ($LanguageCode in $Config.LanguageCodes) {
                $FileNameNew = $FileName.Replace(".$($LanguageCode.alpha3).", ".$($LanguageCode.alpha2).") 
                $ReplacementWasMade = $FileName -cne $FileNameNew
                if ($ReplacementWasMade) {
                    $Destination = Join-Path -Path $FileDirectory -ChildPath $FileNameNew
                    Move-Item -Path $FilePath -Destination $Destination -Force
                    break
                }
            }
        }
        if ($TotalSubsToKeep -gt 0) {
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Subtitles:" -Column2 "$TotalSubsToKeep Extracted"
        }
        if ($TotalSubsToRemove -gt 0) {
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Subtitles:" -Column2 "$TotalSubsToRemove Removed"
        }
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Successful" -ColorBg "Success"
    }
    else {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "No SRT subs found in MKV"
    }
}

# Function to Clean up subtitles
function Start-SubEdit {
    param (
        [Parameter(Mandatory = $true)] 
        $Source,
        [Parameter(Mandatory = $true)] 
        $Files
    )
    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Clean up Subtitles  ***" -Header
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $SubtitleEditPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @("/convert", "$Files", "subrip", "/inputfolder`:`"$Source`"", "/overwrite", "/fixcommonerrors", "/removetextforhi", "/fixcommonerrors")
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    $Process.WaitForExit()
    if ($Process.ExitCode -gt 1) {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Error:" -Column2 $stderr -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
        Stop-Script -ExitReason "SubEdit Error: $DownloadLabel - $DownloadName"
    }
    else {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Successful" -ColorBg "Success"
    }
}

function Start-Subliminal {
    param (
        [Parameter(Mandatory = $true)] 
        $Source
    )
    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Download missing Subtitles  ***" -Header
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $StartInfo.FileName = $SubliminalPath
    $StartInfo.RedirectStandardError = $true
    $StartInfo.RedirectStandardOutput = $true
    $StartInfo.UseShellExecute = $false
    $StartInfo.Arguments = @("--opensubtitles", $OpenSubUser, $OpenSubPass, "--omdb $omdbAPI", "download", "-r omdb", "-p opensubtitles", "-l eng", "-l nld", "`"$Source`"")
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    $Process.Start() | Out-Null
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    $Process.WaitForExit()
    # Write-Host $stdout
    # Write-Host $stderr
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
        $Process.ExitCode = 1
    }
    if ($Process.ExitCode -gt 0) {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Exit Code:" -Column2 $($Process.ExitCode) -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Error:" -Column2 $stderr -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
    }
    else {
        if ($SubsDownloaded -gt 0) {
            # Write-HTMLLog -LogFile $LogFilePath -Column1 "Downloaded:" -Column2 "$SubsDownloaded Subtitles"
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Collected:" -Column2 "$VideoCollected Videos"
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Ignored:" -Column2 "$VideoIgnored Videos"
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Error:" -Column2 "$VideoError Videos"
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Downloaded:" -Column2 "$SubsDownloaded Subtitles"
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Successful" -ColorBg "Success"
        }
        else {
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "No subs downloaded with Subliminal"
        }
    }
}

# Fuction to Process Medusa
function Import-Medusa {
    param (
        [Parameter(Mandatory = $true)] 
        $Source
    )

    $body = @{
        'cmd'            = 'postprocess'
        'force_replace'  = 1
        'is_priority'    = 1
        'delete_files'   = 1
        'path'           = $Source
        'return_data'    = 1
        'process_method' = "move"
        'type'           = "manual"
    }
    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Medusa Import  ***" -Header
    try {
        $response = Invoke-RestMethod "http://$MedusaHost`:$MedusaPort/api/$MedusaApiKey" -Method Get -Body $Body
    }
    catch {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Exception:" -Column2 $_.Exception.Message -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
        Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
    }
    Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Successful" -ColorBg "Success"  
}

# Fuction to Process Radarr
function Import-Radarr {
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
    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Radarr Import  ***" -Header
    try {
        $response = Invoke-RestMethod -uri "http://$RadarrHost`:$RadarrPort/api/command" -Method Post -Body $Body -Headers $headers
    }
    catch {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Exception:" -Column2 $_.Exception.Message -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
        Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
    }
    if ($response.status -eq "queued" -or $response.status -eq "started" -or $response.status -eq "completed") {
        $timeout = New-TimeSpan -Minutes $RadarrTimeOutMinutes
        $endTime = (Get-Date).Add($timeout)
        do {
            try {
                $status = Invoke-RestMethod -uri "http://$RadarrHost`:$RadarrPort/api/command/$($response.id)" -Method Get -Headers $headers
            }
            catch {
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Exception:" -Column2 $_.Exception.Message -ColorBg "Error"
                Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
                Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
            }
            Start-Sleep 1
        }
        while ($status.status -ne "completed" -or ((Get-Date) -gt $endTime))
        if ($status.status -eq "completed") {
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Successful" -ColorBg "Success"         
        }
        else {
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Radarr:" -Column2 $status.status -ColorBg "Error" 
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Radarr:" -Column2 "Import Timeout: ($RadarrTimeOutMinutes) minutes" -ColorBg "Error" 
            Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error" 
        }
    }
    else {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Radarr:" -Column2 $response.status -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
    }
}

# Fuction to close the log and send out mail
function Send-Mail {
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
    $StartInfo.Arguments = @("-smtp $SMTPserver", "-port $SMTPport", "-domain $SMTPserver", "-t $MailTo", "-f $MailFrom", "-fname `"$MailFromName`"", "-sub `"$MailSubject`"", "body", "-file `"$LogFilePath`"", "-mime-type `"text/html`"", "-ssl", "auth -user $SMTPuser -pass $SMTPpass")
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
    Move-Item -Path $LogFilePath -Destination $LogArchivePath

}

Function Write-HTMLLog {
    Param(
        [Parameter(Mandatory = $false)]
        [string] $LogFile,
        [Parameter(Mandatory = $true)]
        [string] $Column1,
        [Parameter(Mandatory = $false)]
        [string] $Column2,
        [Parameter(Mandatory = $false)]
        [switch] $Header,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Success", "Error")]
        [string] $ColorBg
    )

    If ($LogFile) {
        Add-Content -LiteralPath  $LogFile -Value "<tr>"
        if ($Header) {
            Add-Content -LiteralPath  $LogFile -Value "<td colspan=`"2`" style=`"background-color:#398AA4;text-align:center;font-size:10pt`"><b>$Column1</b></td>"
        }
        else {
            if ($ColorBg -eq "") {
                Add-Content -LiteralPath  $LogFile -Value "<td style=`"vertical-align:top;padding: 0px 10px;`"><b>$Column1</b></td>"
                Add-Content -LiteralPath  $LogFile -Value "<td style=`"vertical-align:top;padding: 0px 10px;`">$Column2</td>"
                Add-Content -LiteralPath  $LogFile -Value "</tr>"
            }
            elseif ($ColorBg -eq "Success") {
                Add-Content -LiteralPath  $LogFile -Value "<td style=`"vertical-align:top;padding: 0px 10px;`"><b>$Column1</b></td>"
                Add-Content -LiteralPath  $LogFile -Value "<td style=`"vertical-align:top;padding: 0px 10px;background-color:#555000`">$Column2</td>"
                Add-Content -LiteralPath  $LogFile -Value "</tr>"  
            }
            elseif ($ColorBg -eq "Error") {
                Add-Content -LiteralPath  $LogFile -Value "<td style=`"vertical-align:top;padding: 0px 10px;`"><b>$Column1</b></td>"
                Add-Content -LiteralPath  $LogFile -Value "<td style=`"vertical-align:top;padding: 0px 10px;background-color:#550000`">$Column2</td>"
                Add-Content -LiteralPath  $LogFile -Value "</tr>"  
            }
        }
        Write-Output "$Column1 $Column2"
    }
    Else {
        Write-Output "$Column1 $Column2"
    }
}

function Format-Table {
    param (
        [Parameter(Mandatory = $false)]
        [switch] $Start,
        [Parameter(Mandatory = $false)]
        [string] $LogFile
    )
    if ($Start) {
        Add-Content -LiteralPath  $LogFile -Value "<table border=`"0`" align=`"center`" cellspacing=`"0`""
        Add-Content -LiteralPath  $LogFile -Value "style=`"border-collapse:collapse;background-color:#555555;color:#FFFFFF;font-family:arial,helvetica,sans-serif;font-size:10pt;`">"
        Add-Content -LiteralPath  $LogFile -Value "<col width=`"125`">"
        Add-Content -LiteralPath  $LogFile -Value "<col width=`"500`">"
        Add-Content -LiteralPath  $LogFile -Value "<tbody>"
    }
    else {
        Add-Content -LiteralPath  $LogFile -Value "</tbody>"
        Add-Content -LiteralPath  $LogFile -Value "</table>"
    }
}


# Fuction to stop the script and send out the mail
function Stop-Script {
    Param(
        [Parameter(Mandatory = $true)]
        [string] $ExitReason
    )         
    
    # Stop the Stopwatch
    $StopWatch.Stop()

    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Script Exection time  ***" -Header
    Write-HTMLLog -LogFile $LogFilePath -Column1 "Time Taken:" -Column2 $($StopWatch.Elapsed.ToString('mm\:ss'))
        
    # Clean up process folder 
    try {
        If (Test-Path -LiteralPath  $ProcessPathFull) {
            Remove-Item -Force -Recurse -LiteralPath $ProcessPathFull
        }
    }
    catch {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Exception:" -Column2 $_.Exception.Message -ColorBg "Error"
        Write-HTMLLog -LogFile $LogFilePath -Column1 "Result:" -Column2 "Failed" -ColorBg "Error"
    }
    
    Format-Table -LogFile $LogFilePath
    Send-Mail -MailSubject $ExitReason
    # Clean up the Mutex

    Remove-Mutex -MutexObject $ScriptMutex
    Exit
}

function Test-Variable-Path {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path,
        [Parameter(Mandatory = $true)]
        [string] $Name
    )
    if (!(Test-Path -LiteralPath  $Path)) {
        Write-Host "Cannot find: $Path" -ForegroundColor Red
        Write-Host "As defined in variable: $Name" -ForegroundColor Red
        Write-Host "Will now exit!" -ForegroundColor Red
        Exit 1
    } 
}

# Test additional programs
Test-Variable-Path -Path $WinRarPath -Name "WinRarPath"
Test-Variable-Path -Path $MKVMergePath -Name "MKVMergePath"
Test-Variable-Path -Path $MKVExtractPath -Name "MKVExtractPath"
Test-Variable-Path -Path $SubtitleEditPath -Name "SubtitleEditPath"
Test-Variable-Path -Path $SubliminalPath -Name "SubliminalPath"
Test-Variable-Path -Path $MailSendPath -Name "MailSendPath"


# Get input if no parameters defined
# Download Location
if ( ($Null -eq $DownloadPath) -or ($DownloadPath -eq '') ) {
    $DownloadPath = Get-Input -Message "Download Path" -Required 
}
# Download Name
if ( ($Null -eq $DownloadName) -or ($DownloadName -eq '') ) {
    $DownloadName = Get-Input -Message "Download Name" -Required 
}
# Download Label
if ( ($Null -eq $DownloadLabel) -or ($DownloadLabel -eq '') ) {
    $DownloadLabel = Get-Input -Message "Download Label" 
}
# Torrent Hash (only needed for Radarr)
if ( ($Null -eq $TorrentHash) -or ($TorrentHash -eq '') ) {
    $TorrentHash = Get-Input -Message "Torrent Hash" 
}

# Uppercase TorrentHash
$TorrentHash = $TorrentHash.ToUpper()

# Handle empty Torrent Label or NoProcess
if ($DownloadLabel -eq "" -or $DownloadLabel -eq "NoProcess") {
    write-host "Do nothing"
    Exit
}

# Check paths from Parameters
$DownloadPathFull = Join-Path -Path $DownloadPath -ChildPath $DownloadName
If (!(Test-Path -LiteralPath  $DownloadPath)) {
    Write-Host "$DownloadPath - Not valid location"
    Exit 1
}
If (!(Test-Path -LiteralPath  $DownloadPathFull)) {
    Write-Host "$DownloadPathFull - Not valid location"
    Exit 1
}

# Test File Paths
If (!(Test-Path -LiteralPath  $ProcessPath)) {
    New-Item -ItemType Directory -Force -Path $ProcessPath | Out-Null
}
If (!(Test-Path -LiteralPath  $LogArchivePath)) {
    New-Item -ItemType Directory -Force -Path $LogArchivePath | Out-Null
}

# Start of script
$ScriptMutex = New-Mutex -MutexName 'DownloadScript'

# Start Stopwatch
$StopWatch = [system.diagnostics.stopwatch]::startNew()

# Create Log file
# Log file of current processing file (will be used to send out the mail)
$LogFilePath = Join-Path -Path $ProcessPath -ChildPath "$LogFileDateFormat-$DownloadName.html"

# Log Header
Format-Table -LogFile $LogFilePath -Start
Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Information  ***" -Header
Write-HTMLLog -LogFile $LogFilePath -Column1 "Start:" -Column2 "$(Get-Date -Format "yyyy-MM-dd") at $(Get-Date -Format "HH:mm:ss")"
Write-HTMLLog -LogFile $LogFilePath -Column1 "Label:" -Column2 $DownloadLabel
Write-HTMLLog -LogFile $LogFilePath -Column1 "Name:" -Column2 $DownloadName
Write-HTMLLog -LogFile $LogFilePath -Column1 "Hash:" -Column2 $TorrentHash

# Check if Single file or Folder 
$SingleFile = (Get-Item -LiteralPath $DownloadPathFull) -is [System.IO.FileInfo]
$Folder = (Get-Item -LiteralPath $DownloadPathFull) -is [System.IO.DirectoryInfo]

# Set Source and Destination paths and get Rar paths
if ($Folder) {
    $ProcessPathFull = Join-Path -Path $ProcessPath -ChildPath $DownloadLabel | Join-Path -ChildPath $DownloadName
    $RarFilePaths = (Get-ChildItem -LiteralPath $DownloadPathFull -Recurse -filter "*.rar").FullName
}
elseif ($SingleFile) {
    $ProcessPathFull = Join-Path -Path $ProcessPath -ChildPath $DownloadLabel | Join-Path -ChildPath $DownloadName.Substring(0, $DownloadName.LastIndexOf('.'))
    if ([IO.Path]::GetExtension($downloadPathFull) -eq '.rar') {
        $RarFilePaths = (Get-Item -LiteralPath $DownloadPathFull).FullName
    } 
}

# Find rar files
$RarCount = $RarFilePaths.Count
if ($RarCount -gt 0) { $RarFile = $true } else { $RarFile = $false }

# Check is destination folder exists otherwise create it
If (!(Test-Path -LiteralPath  $ProcessPathFull)) {
    New-Item -ItemType Directory -Force -Path $ProcessPathFull | Out-Null
}

if ($RarFile) {
    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Unrar Download  ***" -Header
    Write-HTMLLog -LogFile $LogFilePath -Column1 "Starting:" -Column2 "Unpacking files"
    foreach ($Rar in $RarFilePaths) {
        Start-UnRar -UnRarSourcePath $Rar -UnRarTargetPath $ProcessPathFull
    }  
}
elseif (-not $RarFile -and $SingleFile) {
    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Single File  ***" -Header
    Start-RoboCopy -Source $DownloadPath -Destination $ProcessPathFull -File $DownloadName
}
elseif (-not $RarFile -and $Folder) {
    Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Folder  ***" -Header
    Start-RoboCopy -Source $DownloadPathFull -Destination $ProcessPathFull -File '*.*'
}


# Starting Post Processing for Movies and TV Shows
if ($DownloadLabel -eq $TVLabel) {
    $MKVFiles = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -filter "*.mkv"
    $MKVCount = $MKVFiles.Count
    if ($MKVCount -gt 0) { $MKVFile = $true } else { $MKVFile = $false }
    if ($MKVFile) {
        # Download any missing subs
        Start-Subliminal -Source $ProcessPathFull
        
        # Remove unwanted subtitle languages and extract wanted subtitles and rename
        Start-MKV-Subtitle-Strip $ProcessPathFull
  
        # Clean up Subs
        Start-SubEdit -File "*.srt" -Source $ProcessPathFull
      
        Write-HTMLLog -LogFile $LogFilePath -Column1 "***  MKV Files  ***" -Header
        foreach ($Mkv in $MKVFiles) {
            Write-HTMLLog -LogFile $LogFilePath -Column1 " " -Column2 $Mkv.name
        }
        Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Subtitle Files  ***" -Header
        $SrtFiles = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -filter "*.srt"
        foreach ($Srt in $SrtFiles) {
            Write-HTMLLog -LogFile $LogFilePath -Column1 " " -Column2 $srt.name
        }
    }
    else {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Files  ***" -Header
        $Files = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -filter "*.*"
        foreach ($File in $Files) {
            Write-HTMLLog -LogFile $LogFilePath -Column1 "File:" -Column2 $File.name
        }
    }

    # Call Medusa to Post Process
    Import-Medusa -Source $ProcessPathFull
    Stop-Script -ExitReason "$DownloadLabel - $DownloadName"
}
elseif ($DownloadLabel -eq $MovieLabel) {
    $MKVFiles = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -filter "*.mkv"
    $MKVCount = $MKVFiles.Count
    if ($MKVCount -gt 0) { $MKVFile = $true } else { $MKVFile = $false }
    if ($MKVFile) {
        # Download any missing subs
        Start-Subliminal -Source $ProcessPathFull

        # Extract wanted subtitles and rename
        Start-MKV-Subtitle-Strip $ProcessPathFull
          
        # Clean up Subs
        Start-SubEdit -File "*.srt" -Source $ProcessPathFull
        
        Write-HTMLLog -LogFile $LogFilePath -Column1 "***  MKV Files  ***" -Header
        foreach ($Mkv in $MKVFiles) {
            Write-HTMLLog -LogFile $LogFilePath -Column1 " " -Column2 $Mkv.name
        }
        Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Subtitle Files  ***" -Header
        $SrtFiles = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -filter "*.srt"
        foreach ($Srt in $SrtFiles) {
            Write-HTMLLog -LogFile $LogFilePath -Column1 " " -Column2 $srt.name
        }
    }
    else {
        Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Files  ***" -Header
        $Files = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -filter "*.*"
        foreach ($File in $Files) {
            Write-HTMLLog -LogFile $LogFilePath -Column1 "File:" -Column2 $File.name
        }
    }

    # Call Radarr to Post Process
    Import-Radarr -Source $ProcessPathFull
    Stop-Script -ExitReason "$DownloadLabel - $DownloadName"
}

# Reached the end of script
Write-HTMLLog -LogFile $LogFilePath -Column1 "***  Post Process General Download  ***" -Header
$Files = Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -filter "*.*"
foreach ($File in $Files) {
    Write-HTMLLog -LogFile $LogFilePath -Column1 "File:" -Column2 $File.name
}
Stop-Script -ExitReason "$DownloadLabel - $DownloadName"