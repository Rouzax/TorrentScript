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
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true
        )]
        [string]$Source
    )

    # Make sure needed functions are available otherwise try to load them.
    $commands = 'Write-HTMLLog'
    foreach ($commandName in $commands) {
        if (!($command = Get-Command $commandName -ErrorAction SilentlyContinue)) {
            Try {
                . $PSScriptRoot\$commandName.ps1
                Write-Host "$commandName Function loaded." -ForegroundColor Green
            } Catch {
                Write-Error -Message "Failed to import $commandName function: $_"
                exit 1
            }
        }
    }
    # Start
    
    $episodes = @()
    $SubsExtracted = $false
    $TotalSubsToExtract = 0
    $TotalSubsToRemove = 0

    Write-HTMLLog -Column1 '***  Extract srt files from MKV  ***' -Header
    Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.mkv' | ForEach-Object {
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
            $StartInfo.Arguments = @('-J', "`"$filePath`"")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            # $stderr = $Process.StandardError.ReadToEnd()
            # Write-Host "stdout: $stdout"
            # Write-Host "stderr: $stderr"
            $Process.WaitForExit()
            if ($Process.ExitCode -eq 2) {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            } elseif ($Process.ExitCode -eq 1) {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
            } elseif ($Process.ExitCode -eq 0) {
                $fileMetadata = $stdout | ConvertFrom-Json
            } else {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
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
        $SubIDsToExtract = @()
        $SubIDsToRemove = @()
        $SubsToExtract = @()
        $SubNamesToKeep = @()

        $episode.FileTracks | ForEach-Object {
            $FileTrack = $_
            if ($FileTrack.id) {
                # Check if subtitle is srt
                if ($FileTrack.type -eq 'subtitles' -and $FileTrack.codec -eq 'SubRip/SRT') {
                    
                    # Check to see if track_name is part of $SubtitleNamesToRemove list
                    if ($null -ne ($SubtitleNamesToRemove | Where-Object { $FileTrack.properties.track_name -match $_ })) {
                        $SubIDsToRemove += $FileTrack.id
                    }
                    # Check is subtitle is in $WantedLanguages list
                    elseif ($FileTrack.properties.language -in $WantedLanguages) {

                        # Handle multiple subtitles of same language, if exist append ID to file 
                        if ("$($episode.FileName).$($FileTrack.properties.language).srt" -in $SubNamesToKeep) {
                            $prefix = "$($FileTrack.id).$($FileTrack.properties.language)"
                        } else {
                            $prefix = "$($FileTrack.properties.language)"
                        }
    
                        # Add Subtitle name and ID to be extracted
                        $SubsToExtract += "`"$($FileTrack.id):$($episode.FileRoot)\$($episode.FileName).$($prefix).srt`""
                        
                        # Keep track of subtitle file names that will be extracted to handle possible duplicates
                        $SubNamesToKeep += "$($episode.FileName).$($prefix).srt"

                        # Add subtitle ID to for MKV remux
                        $SubIDsToExtract += $FileTrack.id
                    } else {
                        $SubIDsToRemove += $FileTrack.id
                    }
                }
            }
        }

        # Count all subtitles to keep and remove of logging
        $TotalSubsToExtract = $TotalSubsToExtract + $SubIDsToExtract.count
        $TotalSubsToRemove = $TotalSubsToRemove + $SubIDsToRemove.count
        
        # Extract the wanted subtitle languages
        if ($SubIDsToExtract.count -gt 0) {
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVExtractPath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            $StartInfo.Arguments = @("`"$($episode.FilePath)`"", 'tracks', "$SubsToExtract")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            # $stderr = $Process.StandardError.ReadToEnd()
            # Write-Host "stdout: $stdout"
            # Write-Host "stderr: $stderr"
            $Process.WaitForExit()
            if ($Process.ExitCode -eq 2) {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvextract:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            } elseif ($Process.ExitCode -eq 1) {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvextract:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
            } elseif ($Process.ExitCode -eq 0) {
                $SubsExtracted = $true
                # Write-HTMLLog -Column1 "Extracted:" -Column2 "$($SubsToExtract.count) Subtitles"
            } else {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvextract:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Unknown' -ColorBg 'Error'
            }
        }

        # Remux and strip out all unwanted subtitle languages
        if ($SubIDsToRemove.Count -gt 0) {
            $TmpFileName = $Episode.FileName + '.tmp'
            $TmpMkvPath = Join-Path $episode.FileRoot $TmpFileName
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVMergePath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            $StartInfo.Arguments = @("-o `"$TmpMkvPath`"", "-s !$($SubIDsToRemove -join ',')", "`"$($episode.FilePath)`"")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            # $stderr = $Process.StandardError.ReadToEnd()
            # Write-Host "stdout: $stdout"
            # Write-Host "stderr: $stderr"
            $Process.WaitForExit()
            if ($Process.ExitCode -eq 2) {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
            } elseif ($Process.ExitCode -eq 1) {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
            } elseif ($Process.ExitCode -eq 0) {
                # Overwrite original mkv after successful remux
                Move-Item -Path $TmpMkvPath -Destination $($episode.FilePath) -Force
                # Write-HTMLLog -Column1 "Removed:" -Column2 "$($SubIDsToRemove.Count) unwanted subtitle languages"
            } else {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
            }
        }
    }
   
    # Rename extracted subs to correct 2 county code based on $LanguageCodes
    if ($SubsExtracted) {
        $SrtFiles = Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.srt'
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
        if ($TotalSubsToExtract -gt 0) {
            Write-HTMLLog -Column1 'Subtitles:' -Column2 "$TotalSubsToExtract Extracted"
        }
        if ($TotalSubsToRemove -gt 0) {
            Write-HTMLLog -Column1 'Subtitles:' -Column2 "$TotalSubsToRemove Removed"
        }
        Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
    } else {
        Write-HTMLLog -Column1 'Result:' -Column2 'No SRT subs found in MKV'
    }
}