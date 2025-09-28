function Start-MKVSubtitleStrip {
    <#
    .SYNOPSIS
        Extracts wanted text (SRT) subtitles and removes unwanted subtitles (both text and image-based) from MKV files.
    .DESCRIPTION
        For each MKV in -Source, this function:
          1) Identifies subtitle tracks via `mkvmerge -J`.
          2) Extracts SRT-compatible text subtitle tracks whose language is in -WantedLanguages
             (skipping duplicates per language and names matching -SubtitleNamesToRemove).
          3) Removes ALL unwanted subtitle tracks (text or image-based like PGS/VobSub) by remuxing.
             - "Unwanted" means: language NOT in -WantedLanguages OR track_name matches -SubtitleNamesToRemove.
          4) Renames extracted .srt files from 3-letter ISO 639-2 codes to 2-letter ISO 639-1 codes
             using the provided -LanguageCodeLookup hashtable (prepared by your main script).

        Edge cases handled:
          - No subtitles at all: logs and skips extraction/remux.
          - Only text subtitles: extracts wanted SRTs, removes the rest.
          - Only image-based subtitles: removes those not in -WantedLanguages, keeps wanted ones.
          - Mixed text + image-based: extracts wanted text SRTs, removes all other subs.
    .PARAMETER Source
        Path to the directory containing MKV files.
    .PARAMETER MKVMergePath
        Path to mkvmerge.exe.
    .PARAMETER MKVExtractPath
        Path to mkvextract.exe.
    .PARAMETER WantedLanguages
        Array of 3-letter ISO 639-2 language codes to KEEP (and extract if text-based SRT).
        Example: @("eng","dut")
    .PARAMETER SubtitleNamesToRemove
        Array of patterns; if a track_name matches any, that track is removed
        (applies to both text and image-based). Example: @("Forced","SDH")
    .PARAMETER LanguageCodeLookup
        Hashtable mapping 639-2 (and optionally 639-2/B) -> 639-1 codes,
        e.g. @{ "eng"="en"; "nld"="nl"; "dut"="nl" }
    .OUTPUTS
        None
    .EXAMPLE
        Start-MKVSubtitleStrip -Source "C:\MKV" `
          -MKVMergePath "C:\Program Files\MKVToolNix\mkvmerge.exe" `
          -MKVExtractPath "C:\Program Files\MKVToolNix\mkvextract.exe" `
          -WantedLanguages @("eng","dut") `
          -SubtitleNamesToRemove @("Forced","SDH") `
          -LanguageCodeLookup @{ "eng"="en"; "nld"="nl"; "dut"="nl" }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$MKVMergePath,
        
        [Parameter(Mandatory = $true)]
        [string]$MKVExtractPath,

        [Parameter(Mandatory = $true)]
        [array]$WantedLanguages,

        [Parameter(Mandatory = $true)]
        [array]$SubtitleNamesToRemove,

        [Parameter(Mandatory = $true)]
        [hashtable]$LanguageCodeLookup
    )

    # Ensure helper functions are available
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

    # Init
    $MkvFiles = @()
    $SubsExtracted = $false
    $TotalSubsToExtract = 0
    $TotalSubsToRemove = 0

    Write-HTMLLog -Column1 '***  Extract SRT files from MKV & Strip Unwanted Subs  ***' -Header

    # Collect MKV metadata
    Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.mkv' |
        Where-Object { $_.DirectoryName -notlike "*\Sample" } | ForEach-Object {
            $MkvFileInfo = $_

            # mkvmerge -J "file"
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVMergePath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            $StartInfo.Arguments = @('-J', "`"$($MkvFileInfo.FullName)`"")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            $Process.WaitForExit()

            switch ($Process.ExitCode) {
                0 {
                    $MkvFileMetadata = $stdout | ConvertFrom-Json
                }
                default {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
                    return
                }
            }

            $file = @{
                FileName        = $MkvFileInfo.BaseName
                FilePath        = $MkvFileInfo.FullName
                FileRoot        = $MkvFileInfo.Directory
                FileTracks      = $MkvFileMetadata.tracks
                FileAttachments = $MkvFileMetadata.attachments
            }
            $MkvFiles += New-Object PSObject -Property $file
        }

    if ($MkvFiles.Count -eq 0) {
        Write-HTMLLog -Column1 'Result:' -Column2 'No MKV files found'
        return
    }

    # Process each MKV
    $MkvFiles | ForEach-Object {
        $MkvFile = $_
        $SubIDsToExtract = @()   # text SRT tracks to extract
        $SubIDsToRemove = @()    # any subs (text or image) to strip during remux
        $SubsToExtract = @()    # mkvextract "id:outfile" tuples
        $SubNamesToKeep = @()    # prevent duplicate language outputs (one SRT per language)
        $HasAnySubs = $false

        $MkvFile.FileTracks | ForEach-Object {
            $FileTrack = $_
            if (-not $FileTrack.id) {
                return 
            }

            if ($FileTrack.type -eq 'subtitles') {
                $HasAnySubs = $true

                # Determine text vs image-based
                $codecId = $FileTrack.properties.codec_id
                $codec = $FileTrack.codec
                $isText =
                ($FileTrack.properties.text_subtitles -eq $true) -or
                ($codec -in @('SubRip/SRT', 'Timed Text')) -or
                ($codecId -like 'S_TEXT/*')
                $isImage = -not $isText

                $lang3 = $FileTrack.properties.language
                $name = $FileTrack.properties.track_name

                # If name matches a removal pattern, mark for removal regardless of language/type
                $matchesRemoveName = $false
                if ($SubtitleNamesToRemove -and $name) {
                    $matchesRemoveName = $null -ne ($SubtitleNamesToRemove | Where-Object { $name -match $_ })
                }

                if ($matchesRemoveName) {
                    $SubIDsToRemove += $FileTrack.id
                    return
                }

                # Keep criteria: language is in WantedLanguages
                $isWantedLanguage = $lang3 -in $WantedLanguages

                if ($isText) {
                    # TEXT: Extract only if wanted language; otherwise, remove.
                    if ($isWantedLanguage) {
                        # Avoid duplicate SRT per language
                        $targetFile = "$($MkvFile.FileName).$($lang3).srt"
                        if ($targetFile -notin $SubNamesToKeep) {
                            $SubsToExtract += "`"$($FileTrack.id):$($MkvFile.FileRoot)\$targetFile`""
                            $SubNamesToKeep += $targetFile
                            $SubIDsToExtract += $FileTrack.id
                        } else {
                            # Duplicate text sub for same language -> remove
                            $SubIDsToRemove += $FileTrack.id
                        }
                    } else {
                        # Text sub in an unwanted language -> remove
                        $SubIDsToRemove += $FileTrack.id
                    }
                } else {
                    # IMAGE-BASED (e.g., PGS/VobSub): remove if not in wanted languages; keep if wanted
                    if (-not $isWantedLanguage) {
                        $SubIDsToRemove += $FileTrack.id
                    }
                    # If wanted language, we keep the image-based track in the MKV (no extraction).
                }
            }
        }

        # If no subtitle tracks in this file
        if (-not $HasAnySubs) {
            Write-HTMLLog -Column1 'File:' -Column2 $MkvFile.FileName
            Write-HTMLLog -Column1 'Subtitles:' -Column2 'None present'
            return
        }

        # Count for logging
        $TotalSubsToExtract += $SubIDsToExtract.Count
        $TotalSubsToRemove += $SubIDsToRemove.Count

        # Extract wanted SRT tracks
        if ($SubIDsToExtract.Count -gt 0) {
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVExtractPath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            # "mkvextract file tracks id:outfile id:outfile ..."
            $StartInfo.Arguments = @("`"$($MkvFile.FilePath)`"", 'tracks', "$SubsToExtract")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            $Process.WaitForExit()

            switch ($Process.ExitCode) {
                0 {
                    $SubsExtracted = $true
                }
                default {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $Process.ExitCode -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvextract:' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                }
            }
        }

        # Remux to strip all unwanted subs (text or image-based)
        if ($SubIDsToRemove.Count -gt 0) {
            $TmpFileName = $MkvFile.FileName + '.tmp'
            $TmpMkvPath = Join-Path $MkvFile.FileRoot $TmpFileName

            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVMergePath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            # -s !id,id,id -> exclude these subtitle track IDs
            $StartInfo.Arguments = @("-o `"$TmpMkvPath`"", "-s !$($SubIDsToRemove -join ',')", "`"$($MkvFile.FilePath)`"")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            $Process.WaitForExit()

            switch ($Process.ExitCode) {
                0 {
                    Move-Item -LiteralPath $TmpMkvPath -Destination $($MkvFile.FilePath) -Force
                }
                default {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                }
            }
        } else {
            # Nothing to strip in this MKV
            Write-HTMLLog -Column1 'File:' -Column2 $MkvFile.FileName
            Write-HTMLLog -Column1 'Subtitles:' -Column2 'No unwanted subs to remove'
        }
    }

    # Rename extracted SRTs to 2-letter language codes using lookup (only for 3-letter codes present in filenames)
    if ($SubsExtracted) {
        $SrtFiles = Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.srt'
        foreach ($srt in $SrtFiles) {
            $SrtDirectory = $srt.Directory
            $SrtPath = $srt.FullName
            $SrtName = $srt.Name

            # Extract trailing language code between basename and .srt
            $languageCode = $srt -replace '.*\.([a-zA-Z]+)\.srt$', '$1'

            if ($languageCode.Length -eq 3) {
                if ($LanguageCodeLookup.ContainsKey($languageCode)) {
                    $SrtNameNew = $SrtName -replace "\.$languageCode\.srt$", ".$($LanguageCodeLookup[$languageCode]).srt"
                    $Destination = Join-Path -Path $SrtDirectory -ChildPath $SrtNameNew
                    Move-Item -LiteralPath $SrtPath -Destination $Destination -Force
                } else {
                    Write-HTMLLog -Column2 "Language code [$languageCode] not found in the 3 letter lookup table." -ColorBg 'Warning'
                }
            }
        }

        if ($TotalSubsToExtract -gt 0) {
            Write-HTMLLog -Column1 'Subtitles:' -Column2 "$TotalSubsToExtract Extracted (text SRT)"
        }
        if ($TotalSubsToRemove -gt 0) {
            Write-HTMLLog -Column1 'Subtitles:' -Column2 "$TotalSubsToRemove Removed (text/image)"
        }
        Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
    } else {
        # No SRTs were extracted. Still report removals if any.
        if ($TotalSubsToRemove -gt 0) {
            Write-HTMLLog -Column1 'Subtitles:' -Column2 "$TotalSubsToRemove Removed (text/image)"
            Write-HTMLLog -Column1 'Result:' -Column2 'Successful (no SRT extraction)'
        } else {
            Write-HTMLLog -Column1 'Result:' -Column2 'No subtitles to extract or remove'
        }
    }
}
