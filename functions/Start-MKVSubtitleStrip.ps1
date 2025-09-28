function Start-MKVSubtitleStrip {
    <#
    .SYNOPSIS
        Strips unwanted subtitles from MKV files and extracts desired SRT subtitles.
    .DESCRIPTION
        This function extracts SRT subtitles from MKV files based on specified criteria,
        and removes unwanted subtitles. It also remuxes the MKV file to exclude undesired
        subtitle tracks.
    .PARAMETER Source
        Specifies the path to the directory containing MKV files.
    .PARAMETER MKVMergePath
        Specifies the path to the MKVMerge executable.
    .PARAMETER MKVExtractPath
        Specifies the path to the MKVExtract executable.
    .PARAMETER WantedLanguages
        Specifies an array of language codes for the desired subtitles.
        The function extracts subtitles for the specified languages.
        Example: @("eng", "dut")
    .PARAMETER SubtitleNamesToRemove
        Specifies an array of subtitle names to be removed.
        Subtitles with matching names will be excluded from the extraction.
        Example: @("Forced")
    .PARAMETER LanguageCodeLookup
        Specifies an Hashtable of language code mappings.
        Each item in the array should be a hashtable with 'alpha3' and 'alpha2' keys
        representing the three-letter and two-letter language codes, respectively.
        Example: @(@{alpha3="eng"; alpha2="en"}, @{alpha3="dut"; alpha2="nl"})
    .OUTPUTS
        None
    .EXAMPLE
        Start-MKVSubtitleStrip -Source "C:\Path\To\MKVFiles" -MKVMergePath "C:\Program Files\MKVToolNix\mkvmerge.exe"
        -MKVExtractPath "C:\Program Files\MKVToolNix\mkvextract.exe" -WantedLanguages @("eng", "dut") 
        -SubtitleNamesToRemove @("Forced") -LanguageCodes @{"en"=@{"639-1"="en";"639-2"="eng";"name"="English"}; "nl"=@{"639-1"="nl";"639-2"="nld";"639-2/B"="dut";"name"="Dutch"}}
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

    # Initialize variables
    $MkvFiles = @()
    $SubsExtracted = $false
    $TotalSubsToExtract = 0
    $TotalSubsToRemove = 0

    Write-HTMLLog -Column1 '***  Extract srt files from MKV  ***' -Header

    # Enumerate MKV files (skip folders named "Sample") and collect their track metadata via mkvmerge -J
    Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.mkv' |
        Where-Object { $_.DirectoryName -notlike "*\Sample" } |
        ForEach-Object {
            $MkvFileInfo = $_
            $MkvFileMetadata = $null

            # Query mkvmerge for track/attachment info in JSON
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
            $stderr = $Process.StandardError.ReadToEnd()
            $Process.WaitForExit()

            switch ($Process.ExitCode) {
                0 {
                    # mkvmerge exited OK; parse JSON (guard against malformed output)
                    try {
                        $MkvFileMetadata = $stdout | ConvertFrom-Json
                    } catch {
                        Write-HTMLLog -Column1 'mkvmerge JSON parse error:' -Column2 $_.Exception.Message -ColorBg 'Error'
                    }
                }
                1 {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $Process.ExitCode -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvmerge (stderr):' -Column2 $stderr -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvmerge (stdout):' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
                }
                default {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $Process.ExitCode -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvmerge (stderr):' -Column2 $stderr -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvmerge (stdout):' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
                }
            }

            # Guard: if we couldn’t get valid metadata, skip this file cleanly
            if (-not $MkvFileMetadata -or -not $MkvFileMetadata.tracks) {
                Write-HTMLLog -Column1 'Skipping file:' -Column2 $MkvFileInfo.FullName -ColorBg 'Warning'
                return
            }

            # Collect just what we need for the later extract/remux steps
            $file = @{
                FileName        = $MkvFileInfo.BaseName
                FilePath        = $MkvFileInfo.FullName
                FileRoot        = $MkvFileInfo.Directory
                FileTracks      = $MkvFileMetadata.tracks
                FileAttachments = $MkvFileMetadata.attachments
            }

            $MkvFiles += New-Object PSObject -Property $file
        }

    # Extract wanted SRT subtitles
    $MkvFiles | ForEach-Object {
        $MkvFile = $_
        $SubIDsToExtract = @()
        $SubIDsToRemove = @()
        $SubsToExtract = @()
        $SubNamesToKeep = @()

        # Iterate all tracks in the MKV and decide what to extract/remove.
        $MkvFile.FileTracks | ForEach-Object {
            $FileTrack = $_

            # IMPORTANT: track id can be 0, so check for $null (not truthiness).
            if ($null -ne $FileTrack.id) {

                # Only care about subtitle tracks here.
                if ($FileTrack.type -eq 'subtitles') {

                    # Normalize identifiers the JSON may expose.
                    $codec = $FileTrack.codec
                    $codec_id = $FileTrack.properties.codec_id

                    # We only extract *text* subtitles: SRT (SubRip) and Timed Text.
                    # Everything else (e.g., PGS/HDMV, VobSub, DVD Sub) is removed.
                    $isSrtOrTimedText =
                    ($codec -eq 'SubRip/SRT') -or
                    ($codec -eq 'Timed Text') -or
                    ($codec_id -eq 'S_TEXT/UTF8')   # sometimes reported via codec_id

                    if (-not $isSrtOrTimedText) {
                        # Non-SRT (image-based or other) → flag for removal during remux.
                        $SubIDsToRemove += $FileTrack.id
                        continue
                    }

                    # From here on we only handle SRT/Timed Text tracks.

                    # If a track name exists and matches any removal pattern (e.g., 'Forced'), remove it.
                    if ($FileTrack.properties.track_name -and
                        ($SubtitleNamesToRemove | Where-Object { $FileTrack.properties.track_name -match $_ })) {
                        $SubIDsToRemove += $FileTrack.id
                        continue
                    }

                    # Keep/extract only the wanted languages; remove the rest.
                    if ($FileTrack.properties.language -in $WantedLanguages) {
                        # Avoid extracting duplicate SRTs for the same language per file.
                        $targetName = "$($MkvFile.FileName).$($FileTrack.properties.language).srt"
                        if ($targetName -notin $SubNamesToKeep) {
                            # Queue extraction for mkvextract as id:path (keep quotes for safety).
                            $SubsToExtract += "`"$($FileTrack.id):$($MkvFile.FileRoot)\$targetName`""
                            $SubNamesToKeep += $targetName
                            $SubIDsToExtract += $FileTrack.id
                        }
                    } else {
                        # SRT/Timed Text but not in the wanted language list → remove.
                        $SubIDsToRemove += $FileTrack.id
                    }
                }
            }
        }


        # Count all subtitles to keep and remove for logging
        $TotalSubsToExtract = $TotalSubsToExtract + $SubIDsToExtract.count
        $TotalSubsToRemove = $TotalSubsToRemove + $SubIDsToRemove.count
        
        # Extract the wanted subtitle languages
        if ($SubIDsToExtract.count -gt 0) {
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVExtractPath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
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
                1 {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvextract:' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
                }
                default {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $process.ExitCode -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvextract:' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                }
            }
        }

        # Remux and strip out all unwanted subtitle languages
        if ($SubIDsToRemove.Count -gt 0) {
            $TmpFileName = $MkvFile.FileName + '.tmp'
            $TmpMkvPath = Join-Path $MkvFile.FileRoot $TmpFileName
            $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $StartInfo.FileName = $MKVMergePath
            $StartInfo.RedirectStandardError = $true
            $StartInfo.RedirectStandardOutput = $true
            $StartInfo.UseShellExecute = $false
            $StartInfo.Arguments = @("-o `"$TmpMkvPath`"", "-s !$($SubIDsToRemove -join ',')", "`"$($MkvFile.FilePath)`"")
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $StartInfo
            $Process.Start() | Out-Null
            $stdout = $Process.StandardOutput.ReadToEnd()
            $Process.WaitForExit()

            switch ($process.ExitCode) {
                0 {
                    Move-Item -LiteralPath $TmpMkvPath -Destination $($MkvFile.FilePath) -Force
                }
                1 {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
                }
                default {
                    Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                    Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                }
            }
        }
    }
   
    # Rename extracted subs to correct 2 letter language code code based on $LanguageCodes
    if ($SubsExtracted) {
        $SrtFiles = Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.srt'

        # Rename extracted SRTs from ISO-639-2 (3-letter) to ISO-639-1 (2-letter) codes
        # Example: My.Movie.eng.srt  →  My.Movie.en.srt
        # NOTE: LanguageCodeLookup must be a flat 3→2 map, e.g. @{ 'eng'='en'; 'nld'='nl'; 'dut'='nl' }
        foreach ($srt in $SrtFiles) {
            $SrtDirectory = $srt.Directory
            $SrtPath = $srt.FullName
            $SrtName = $srt.Name

            # Extract the trailing language code from the file name (before .srt), e.g. ".eng.srt" → "eng"
            # BUGFIX: operate on $SrtName (string), not on the FileInfo object $srt.
            $languageCode = $SrtName -replace '.*\.([A-Za-z]+)\.srt$', '$1'

            # Only convert when a 3-letter code is present (skip already-correct 2-letter files)
            if ($languageCode.Length -eq 3) {

                # Check if a mapping for this 3-letter code exists (e.g., 'eng' → 'en', 'dut' → 'nl')
                if ($LanguageCodeLookup.ContainsKey($languageCode) -and $LanguageCodeLookup[$languageCode]) {
                    $twoLetter = $LanguageCodeLookup[$languageCode]

                    # Build the new filename by replacing ".<3letter>.srt" with ".<2letter>.srt"
                    $SrtNameNew = $SrtName -replace "\.$languageCode\.srt$", ".$twoLetter.srt"
                    $Destination = Join-Path -Path $SrtDirectory -ChildPath $SrtNameNew

                    # Only move if the destination path is actually different
                    if ($Destination -ne $SrtPath) {
                        try {
                            Move-Item -LiteralPath $SrtPath -Destination $Destination -Force
                        } catch {
                            Write-HTMLLog -Column1 'Rename failed:' -Column2 "$SrtPath → $Destination :: $($_.Exception.Message)" -ColorBg 'Error'
                        }
                    }
                } else {
                    Write-HTMLLog -Column2 "Language code [$languageCode] not found in the 3→2 lookup table." -ColorBg 'Warning'
                }
            }
            # else: language is not 3 letters (probably already 2-letter); leave as-is.
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