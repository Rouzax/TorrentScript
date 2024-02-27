<#
.SYNOPSIS
    Downloads missing subtitles from OpenSubtitle.com.
.DESCRIPTION
    This function searches for and downloads missing subtitles for specified video files
    from OpenSubtitle.com. It supports downloading subtitles in multiple languages.
.PARAMETER Source
    The source directory containing the video files for which subtitles are to be downloaded.
.PARAMETER OpenSubUser
    The username for accessing the OpenSubtitle API.
.PARAMETER OpenSubPass
    The password for accessing the OpenSubtitle API.
.PARAMETER OpenSubAPI
    The API key for accessing the OpenSubtitle API.
.PARAMETER OpenSubHearing_impaired
    Indicates whether to include subtitles for the hearing impaired (include/exclude/only).
.PARAMETER OpenSubForeign_parts_only
    Indicates whether to include subtitles for foreign parts only (include/exclude/only).
.PARAMETER OpenSubMachine_translated
    Indicates whether to include machine-translated subtitles (include/exclude/only).
.PARAMETER OpenSubAI_translated
    Indicates whether to include AI-translated subtitles (include/exclude/only).
.PARAMETER WantedLanguages
    An array of language codes specifying the desired subtitle languages.
.PARAMETER Type
    The type of the video files (e.g., movie, episode).

.OUTPUTS 
    This function outputs a log of downloaded and failed subtitles along with language-specific counts
    and the remaining downloads for the day.

.EXAMPLE
    Start-OpenSubtitlesDownload -Source "C:\Videos" -OpenSubUser "username" -OpenSubPass "password"
    -OpenSubAPI "APIKey" -OpenSubHearing_impaired "exclude" -OpenSubForeign_parts_only "exclude"
    -OpenSubMachine_translated "exclude" -OpenSubAI_translated "exclude" -WantedLanguages @("en", "fr")
    -Type "movie"
    Searches for and downloads missing subtitles for video files in the "C:\Videos" directory,
    excluding hearing impaired and foreign parts only subtitles, and including English and Spanish subtitles.
#>

function Start-OpenSubtitlesDownload {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$OpenSubUser,
           
        [Parameter(Mandatory = $true)]
        [string]$OpenSubPass,
           
        [Parameter(Mandatory = $true)]
        [string]$OpenSubAPI,
           
        [Parameter(Mandatory = $true)]
        [ValidateSet("include", "exclude", "only")]
        [string]$OpenSubHearing_impaired,
           
        [Parameter(Mandatory = $true)]
        [ValidateSet("include", "exclude", "only")]
        [string]$OpenSubForeign_parts_only,
           
        [Parameter(Mandatory = $true)]
        [ValidateSet("include", "exclude", "only")]
        [string]$OpenSubMachine_translated,
           
        [Parameter(Mandatory = $true)]
        [ValidateSet("include", "exclude", "only")]
        [string]$OpenSubAI_translated,
           
        [Parameter(Mandatory = $true)]
        [array]$WantedLanguages,
           
        [Parameter(Mandatory = $true)]
        [ValidateSet("episode", "movie")]
        [string]$Type
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

    # Function to connect to the OpenSubtitle API
    function Connect-OpenSubtitleAPI {
        param(
            [string]$username,
            [string]$password,
            [string]$APIKey
        )

        # Set headers
        $headers = @{
            "Content-Type" = "application/json"
            "User-Agent"   = "Torrentscript"
            "Accept"       = "application/json"
            "Api-Key"      = $APIKey
        }

        # Set body
        $body = @{
            username = $username
            password = $password
        } | ConvertTo-Json

        try {
            # Make requests
            $response = Invoke-RestMethod -Uri 'https://api.opensubtitles.com/api/v1/login' -Method POST -Headers $headers -ContentType 'application/json' -Body $body
            # Check for successful login
            if ($response.status -eq 200) {
                return $response.token
            } else {
                Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Login failed:" -ColorBg 'Error'
                Write-HTMLLog -Column2 "$($response.status)" -ColorBg 'Error'
                return $null
            }
        } catch {
            Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred while logging in:" -ColorBg 'Error'
            Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
            return $null
        }
    }

    # Function to disconnect from the OpenSubtitle API
    function Disconnect-OpenSubtitleAPI {
        param(
            [string]$APIKey,
            [string]$token
        )

        # Set headers
        $headers = @{
            "User-Agent"    = "Torrentscript"
            "Accept"        = "application/json"
            "Api-Key"       = $APIKey
            "Authorization" = "Bearer $token"
        }

        try {
            # Make request
            $response = Invoke-RestMethod -Uri 'https://api.opensubtitles.com/api/v1/logout' -Method DELETE -Headers $headers

            # Check for successful logout
            if ($response.status -eq 200) {
                Write-Host "Logout successful: $($response.message)"
            } else {
                Write-Host "Logout failed: $($response.status)"
            }
        } catch {
            Write-Host "Error occurred while logging out: $_"
        }
    }

    # Function to search for subtitles
    function Search-Subtitles {
        param(
            [string]$type,
            [string]$query,
            [string]$languages,
            [string]$moviehash,
            [string]$APIKey,
            [string]$hearing_impaired,
            [string]$foreign_parts_only,
            [string]$machine_translated,
            [string]$ai_translated
        )

        # Set headers
        $headers = @{
            "User-Agent" = "Torrentscript"
            "Api-Key"    = $APIKey
        }

        # Build the URI with query parameters
        $uri = "https://api.opensubtitles.com/api/v1/subtitles?type=$type&query=$query&languages=$languages&moviehash=$moviehash&hearing_impaired=$hearing_impaired&foreign_parts_only=$foreign_parts_only&machine_translated=$machine_translated&ai_translated=$ai_translated"

        try {
            # Make request
            $response = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers

            # Check for successful response
            if ($response) {
                # Initialize a hashtable to store subtitle IDs for each language
                $subtitleInfo = @{
                    VideoFileName = $query
                    SubtitleIds   = @{}
                }

                # Loop through the data to find the IDs for each language
                foreach ($subtitle in $response.data) {
                    $language = $subtitle.attributes.language

                    # Check if the language key exists in the hashtable
                    if (-not $subtitleInfo.SubtitleIds.ContainsKey($language)) {
                        $subtitleInfo.SubtitleIds[$language] = $subtitle.attributes.files[0].file_id
                    }
                }

                # Return the hashtable containing subtitle IDs for each language
                return $subtitleInfo
            } else {
                # No subtitles found for $query
                return $null
            }
        } catch {
            Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred while searching for subtitles:" -ColorBg 'Error'
            Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
            return $null
        }
    }

    # Function to download all subtitles missing
    function Save-AllSubtitles {
        param(
            [hashtable]$subtitleInfo,
            [string]$APIKey,
            [string]$token,
            [string]$baseDirectory
        )

        # Initialize counters
        $downloadedCount = 0
        $failedCount = 0
        $languageCounts = @{}

        # Set headers
        $headers = @{
            "User-Agent"    = "Torrentscript"
            "Content-Type"  = "application/json"
            "Accept"        = "application/json"
            "Api-Key"       = $APIKey
            "Authorization" = "Bearer $token"
        }

        # Iterate through each language in the subtitle info
        foreach ($language in $subtitleInfo.SubtitleIds.Keys) {
            $subtitleId = $subtitleInfo.SubtitleIds[$language]
            $videoFileName = $subtitleInfo.VideoFileName
            $subtitleFileName = "$baseDirectory\$videoFileName.$language.srt"

            # Check if the subtitle file already exists
            if (-not (Test-Path $subtitleFileName)) {
                # Build the body for the request
                $body = @{
                    file_id = $subtitleId
                } | ConvertTo-Json

                try {
                    # Make request to download the subtitle
                    $response = Invoke-RestMethod -Uri 'https://api.opensubtitles.com/api/v1/download' -Method POST -Headers $headers -ContentType 'application/json' -Body $body

                    # Check if the link is present in the response
                    if ($response.link) {
                        # Download the subtitle file
                        Invoke-WebRequest -Uri $response.link -OutFile $subtitleFileName

                        # Increment downloaded count
                        $downloadedCount++

                        # Increment language-specific count
                        if (-not $languageCounts.ContainsKey($language)) {
                            $languageCounts[$language] = 1
                        } else {
                            $languageCounts[$language]++
                        }
                    } else {
                        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Failed to download subtitle for language: $language" -ColorBg 'Error'
                        # Increment failed count
                        $failedCount++
                    }
                } catch {
                    Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred while downloading subtitle for language $($language):" -ColorBg 'Error'
                    Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
                    # Increment failed count
                    $failedCount++
                }
            }
        }

        # Return the counts and language-specific counts
        return @{
            Downloaded         = $downloadedCount
            Failed             = $failedCount
            LanguageCounts     = $languageCounts
            RemainingDownloads = $response.remaining
        }
    }


    # Function to get video hash
    function Get-VideoHash([string]$path) {
        $dataLength = 65536

        function LongSum([UInt64]$a, [UInt64]$b) { 
            [UInt64](([Decimal]$a + $b) % ([Decimal]([UInt64]::MaxValue) + 1)) 
        }

        function StreamHash([IO.Stream]$stream) {
            $hashLength = 8
            [UInt64]$lhash = 0
            [byte[]]$buffer = New-Object byte[] $hashLength
            $i = 0
            while ( ($i -lt ($dataLength / $hashLength)) -and ($stream.Read($buffer, 0, $hashLength) -gt 0) ) {
                $i++
                $lhash = LongSum $lhash ([BitConverter]::ToUInt64($buffer, 0))
            }
            $lhash
        }

        try { 
            $stream = [IO.File]::OpenRead($path) 
            [UInt64]$lhash = $stream.Length
            $lhash = LongSum $lhash (StreamHash $stream)
            $stream.Position = [Math]::Max(0L, $stream.Length - $dataLength)
            $lhash = LongSum $lhash (StreamHash $stream)
            $hash = "{0:X}" -f $lhash
            if ($hash.Length -lt 16) {
                $hash = ("0" * (16 - $hash.Length)) + $hash
            }
            $hash
        } finally {
            $stream.Close() 
        }
    }

    #* Start the search and download of the subtitles
    try {
        Write-HTMLLog -Column1 '***  Download missing subs from OpenSubtitle.com  ***' -Header
        $token = Connect-OpenSubtitleAPI -username $OpenSubUser -password $OpenSubPass -APIKey $OpenSubAPI

        if ($token) {
            $videoFiles = @(Get-ChildItem -LiteralPath $ProcessPathFull -Recurse -Filter '*.mkv' | Where-Object { $_.DirectoryName -notlike "*\Sample" })

            foreach ($videoFile in $videoFiles) {
                $videoHash = Get-VideoHash $videoFile.FullName

                $languageString = $WantedLanguages -join ','

                $queryParams = @{
                    query              = $videoFile.BaseName
                    APIKey             = $OpenSubAPI
                    type               = $Type
                    moviehash          = $videoHash
                    hearing_impaired   = $OpenSubHearing_impaired
                    foreign_parts_only = $OpenSubForeign_parts_only
                    machine_translated = $OpenSubMachine_translated
                    ai_translated      = $OpenSubAI_translated
                    languages          = $languageString
                }
                $subtitleInfo = Search-Subtitles @queryParams

                # Call Save-AllSubtitles and capture the counts
                $subtitlesCounts = Save-AllSubtitles -subtitleInfo $subtitleInfo -APIKey $OpenSubAPI -token $token -baseDirectory $videoFile.DirectoryName
            }

            if ($($subtitlesCounts.Downloaded) -gt 0) {
                # Log language-specific counts
                foreach ($language in $subtitlesCounts.LanguageCounts.Keys) {
                    Write-HTMLLog -Column1 "Downloaded:" -Column2 "$($subtitlesCounts.LanguageCounts[$language]) in $($language.ToUpper())"
                }
                # Log the counts
                Write-HTMLLog -Column1 "Downloaded:"  -Column2 "$($subtitlesCounts.Downloaded) Total"
                Write-HTMLLog -Column2 "Remaining downloads today: $($subtitlesCounts.RemainingDownloads)"

                if ($($subtitlesCounts.Failed) -gt 0) {
                    Write-HTMLLog -Column1 "Failed:" -Column2 "$($subtitlesCounts.Failed) failed to download" -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                } else {
                    Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
                }
            } else {
                Write-HTMLLog -Column1 'Result:' -Column2 'No downloads found or needed' -ColorBg 'Success'
            }

            Disconnect-OpenSubtitleAPI -APIKey $OpenSubAPI -token $token
        }
    } catch {
        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred:" -ColorBg 'Error'
        Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
    }
}