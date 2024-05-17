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

function Start-MKVSubtitleStrip {
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

    Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.mkv' | Where-Object { $_.DirectoryName -notlike "*\Sample" } | ForEach-Object {
        $MkvFileInfo = $_

        # Start the json export with MKVMerge on the available tracks
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

        switch ($process.ExitCode) {
            0 {
                $MkvFileMetadata = $stdout | ConvertFrom-Json
            }
            1 {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
            }
            default {
                Write-HTMLLog -Column1 'Exit Code:' -Column2 $($Process.ExitCode) -ColorBg 'Error'
                Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Warning' -ColorBg 'Error'
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

    # Extract wanted SRT subtitles
    $MkvFiles | ForEach-Object {
        $MkvFile = $_
        $SubIDsToExtract = @()
        $SubIDsToRemove = @()
        $SubsToExtract = @()
        $SubNamesToKeep = @()

        $MkvFile.FileTracks | ForEach-Object {
            $FileTrack = $_
            if ($FileTrack.id) {
                # Check if subtitle is srt
                if ($FileTrack.type -eq 'subtitles' -and ($FileTrack.codec -eq 'SubRip/SRT' -or $FileTrack.codec -eq 'Timed Text')) {
                    
                    # Check to see if track_name is part of $SubtitleNamesToRemove list
                    if ($null -ne ($SubtitleNamesToRemove | Where-Object { $FileTrack.properties.track_name -match $_ })) {
                        $SubIDsToRemove += $FileTrack.id
                    }
                    # Check is subtitle is in $WantedLanguages list
                    elseif ($FileTrack.properties.language -in $WantedLanguages) {

                        # Handle multiple subtitles of same language, if exist skip duplicates of same language 
                        if ("$($MkvFile.FileName).$($FileTrack.properties.language).srt" -notin $SubNamesToKeep) {
                            $prefix = "$($FileTrack.properties.language)"

                            # Add Subtitle name and ID to be extracted
                            $SubsToExtract += "`"$($FileTrack.id):$($MkvFile.FileRoot)\$($MkvFile.FileName).$($prefix).srt`""
                            
                            # Keep track of subtitle file names that will be extracted to handle possible duplicates
                            $SubNamesToKeep += "$($MkvFile.FileName).$($prefix).srt"
    
                            # Add subtitle ID to for MKV remux
                            $SubIDsToExtract += $FileTrack.id
                        }                        
                    } else {
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

            switch ($process.ExitCode) {
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
                    Write-HTMLLog -Column1 'mkvmerge:' -Column2 $stdout -ColorBg 'Error'
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
   
    # Rename extracted subs to correct 2 county code based on $LanguageCodes
    if ($SubsExtracted) {
        $SrtFiles = Get-ChildItem -LiteralPath $Source -Recurse -Filter '*.srt'

        foreach ($srt in $SrtFiles) {
            $SrtDirectory = $srt.Directory
            $SrtPath = $srt.FullName
            $SrtName = $srt.Name
           
            # Extract the language code from the file name
            $languageCode = $srt -replace '.*\.([a-zA-Z]+)\.srt$', '$1' 

            # Check if the language code exists in the lookup table
            if ($LanguageCodeLookup.ContainsKey($languageCode)) {
                $SrtNameNew = $SrtName -replace "\.$languageCode\.srt$", ".$($LanguageCodeLookup[$languageCode]).srt"
                $Destination = Join-Path -Path $SrtDirectory -ChildPath $SrtNameNew
                Move-Item -LiteralPath $SrtPath -Destination $Destination -Force
            } else {
                Write-HTMLLog -Column2 "Language code $languageCode not found in the 3 letter lookup table." -ColorBg 'Error'
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