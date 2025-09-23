<#
.SYNOPSIS
    Download missing subtitles from OpenSubtitles.com with token caching and rate-limit handling.

.DESCRIPTION
    Scans a source directory for video files and downloads missing subtitles via the OpenSubtitles API.  
    Designed for unattended runs (e.g., post-download), with support for caching tokens, handling rate 
    limits (HTTP 429), and distinguishing between “already present” and “not found” subtitles.  

    Key features:
      • Token cache per user/API key (Windows: %LOCALAPPDATA%\TorrentScript\OpenSubtitles,
        Linux/macOS: $XDG_CACHE_HOME or ~/.cache/TorrentScript/OpenSubtitles).  
        Cached tokens are reused; invalid tokens (401/403) clear the cache.  
      • Login retried up to 5 times on 429, paced by Retry-After/ratelimit-reset headers.  
      • Search/download retried up to 3 times on 429.  
      • Logs out only if a fresh login was made (cached tokens remain valid).  
      • Logging via Write-HTMLLog (ColorBg supports 'Success' or 'Error').  
      • Computes a video hash (length + first/last 64 KiB) to improve matching.  
      • Ignores “Sample” folders. Supports filtering for episode/movie type and subtitle attributes.  

.PARAMETER Source
    Directory to scan recursively for video files (.mkv, .mp4, .avi).  
    Sample folders are skipped.  

.PARAMETER OpenSubUser
    OpenSubtitles username.  

.PARAMETER OpenSubPass
    OpenSubtitles password.  

.PARAMETER OpenSubAPI
    OpenSubtitles API key.  

.PARAMETER OpenSubHearing_impaired 
    Filter for hearing-impaired subtitles. 
    Accepted values: include, exclude, only 
    
.PARAMETER OpenSubForeign_parts_only
    Filter for “foreign parts only” subtitles. 
    Accepted values: include, exclude, only 

.PARAMETER OpenSubMachine_translated 
    Filter for machine-translated subtitles. 
    Accepted values: include, exclude, only 
    
.PARAMETER OpenSubAI_translated 
    Filter for AI-translated subtitles. 
    Accepted values: include, exclude, only

.PARAMETER WantedLanguages
    ISO language codes (e.g., 'en','nl','fr'). Downloads one per requested language per file.  

.PARAMETER Type
    Content type for the API: movie or episode.  

.INPUTS
    None.  

.OUTPUTS
    None. Writes progress and results using Write-HTMLLog, including:  
      - Downloaded per language & total  
      - Already present & failed counts  
      - Languages not found  
      - Remaining daily downloads (if reported by API)  

.EXAMPLE
    Start-OpenSubtitlesDownload `
      -Source "D:\Media\Movies" `
      -OpenSubUser "user" -OpenSubPass "pass" -OpenSubAPI "apikey" `
      -OpenSubHearing_impaired "exclude" -OpenSubForeign_parts_only "exclude" `
      -OpenSubMachine_translated "exclude" -OpenSubAI_translated "exclude" `
      -WantedLanguages @("en","nl") -Type "movie"

    Downloads English and Dutch subtitles for all movies under D:\Media\Movies.  

.NOTES
    • Safe for unattended execution; distinguishes “no subtitles needed” vs “none found.”  
    • Clears token cache automatically if authentication fails.  
    • Retries respect API rate-limit headers; adds small jitter to avoid collisions.  
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

    # Ensure external dependencies
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

    # ---------------- Token cache helpers ----------------
    function Get-TSCacheRoot {
        if ($IsWindows) {
            return (Join-Path $env:LOCALAPPDATA 'TorrentScript\OpenSubtitles')
        } else {
            $base = $env:XDG_CACHE_HOME
            if (-not $base -or $base -eq '') {
                $base = (Join-Path $HOME '.cache') 
            }
            return (Join-Path $base 'TorrentScript/OpenSubtitles')
        }
    }
    function Get-TokenCachePath {
        param([string]$Username, [string]$APIKey)
        $root = Get-TSCacheRoot
        if (-not (Test-Path $root)) {
            New-Item -ItemType Directory -Path $root -Force | Out-Null 
        }
        $hashInput = "$($Username)`n$($APIKey)"
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hashInput))).Replace('-', '').Substring(0, 32)
        return (Join-Path $root "token-$hash.json")
    }
    function Read-TokenCache {
        param([string]$Username, [string]$APIKey)
        $path = Get-TokenCachePath -Username $Username -APIKey $APIKey
        if (Test-Path $path) {
            try {
                $data = Get-Content $path -Raw | ConvertFrom-Json
                if ($data -and $data.token -and $data.expires_at) {
                    if ((Get-Date) -lt ([datetime]$data.expires_at)) {
                        return $data.token 
                    }
                }
            } catch {
            }
        }
        return $null
    }
    function Write-TokenCache {
        param([string]$Username, [string]$APIKey, [string]$Token, [datetime]$ExpiresAt)
        $path = Get-TokenCachePath -Username $Username -APIKey $APIKey
        try {
            [pscustomobject]@{
                username   = $Username
                api_key    = '***' # do not persist API key
                token      = $Token
                expires_at = $ExpiresAt.ToString('o')
            } | ConvertTo-Json | Set-Content -Path $path -Encoding UTF8
        } catch {
        }
    }
    function Clear-TokenCache {
        param([string]$Username, [string]$APIKey)
        $path = Get-TokenCachePath -Username $Username -APIKey $APIKey
        if (Test-Path $path) {
            Remove-Item $path -Force -ErrorAction SilentlyContinue 
        }
    }

    # Keep creds accessible to clear cache on auth errors
    $script:OpenSubCreds = @{ Username = $OpenSubUser; APIKey = $OpenSubAPI }

    # ---------------- Request wrapper (Invoke-WebRequest) ----------------
    function Invoke-OpenSubs {
        param(
            [Parameter(Mandatory)] [string]$Uri,
            [Parameter(Mandatory)] [ValidateSet('GET', 'POST', 'DELETE')] [string]$Method,
            [Parameter(Mandatory)] [hashtable]$Headers,
            [string]$ContentType,
            $Body = $null,
            [int]$MaxAttempts = 3,        # For /login, we’ll use 5 as requested
            [switch]$StopOnAuthError,     # Stop immediately on 401/403
            [ref]$RespHeaders,            # headers returned here (if provided)
            [ref]$StatusOut               # <-- NEW: set to last HTTP status code (int) or $null
        )

        # default out
        if ($StatusOut) {
            $StatusOut.Value = $null 
        }

        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            try {
                $args = @{
                    Uri     = $Uri
                    Method  = $Method
                    Headers = $Headers
                }
                # Compatibility with Windows PowerShell 5.1
                if ($PSVersionTable.PSEdition -eq 'Desktop') {
                    $args.UseBasicParsing = $true 
                }
                if ($PSBoundParameters.ContainsKey('ContentType')) {
                    $args.ContentType = $ContentType 
                }
                if ($null -ne $Body) {
                    $args.Body = $Body 
                }

                $resp = Invoke-WebRequest @args

                if ($RespHeaders) {
                    $RespHeaders.Value = $resp.Headers 
                }
                if ($StatusOut) {
                    $StatusOut.Value = 200 
                }

                $content = $resp.Content
                if ([string]::IsNullOrWhiteSpace($content)) {
                    return $null 
                }

                try {
                    return ($content | ConvertFrom-Json) 
                } catch {
                    return $content 
                }
            } catch {
                $status = $null
                $hdrs = $null
                try {
                    $status = $_.Exception.Response.StatusCode.value__ 
                } catch {
                }
                try {
                    $hdrs = $_.Exception.Response.Headers 
                } catch {
                }
                if ($StatusOut) {
                    $StatusOut.Value = $status 
                }

                if ($StopOnAuthError -and ($status -in 401, 403)) {
                    Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Authentication failed (HTTP $status). Stopping." -ColorBg 'Error'
                    return $null
                }
                if ($status -in 401, 403) {
                    # Token invalid mid-run? Clear cache so next login will succeed
                    try {
                        Clear-TokenCache -Username $script:OpenSubCreds.Username -APIKey $script:OpenSubCreds.APIKey 
                    } catch {
                    }
                }

                if ($status -eq 429) {
                    $wait = $null
                    if ($hdrs -and $hdrs['Retry-After']) {
                        if (-not [int]::TryParse($hdrs['Retry-After'], [ref]$wait)) {
                            $retryAt = Get-Date $hdrs['Retry-After'] -ErrorAction SilentlyContinue
                            if ($retryAt) {
                                $delta = [int][Math]::Ceiling(($retryAt - (Get-Date)).TotalSeconds)
                                if ($delta -gt 0) {
                                    $wait = $delta 
                                }
                            }
                        }
                    }
                    if (-not $wait -and $hdrs -and $hdrs['ratelimit-reset']) {
                        [int]::TryParse($hdrs['ratelimit-reset'], [ref]$wait) | Out-Null
                    }
                    if (-not $wait -or $wait -lt 1) {
                        $wait = 1 
                    }
                    $wait += Get-Random -Minimum 0 -Maximum 2

                    if ($attempt -lt $MaxAttempts) {
                        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "429 received. Retrying in ${wait}s (attempt $attempt of $MaxAttempts)..." -ColorBg 'Error'
                        Start-Sleep -Seconds $wait
                        continue
                    }
                }

                Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Request failed (HTTP $status): $($_.Exception.Message)" -ColorBg 'Error'
                return $null
            }
        }
    }


    # ---------------- Auth helpers (token cache + conditional logout) ----------------
    $script:OpenSubsUsedCachedToken = $false

    function Connect-OpenSubtitleAPI {
        param(
            [string]$username,
            [string]$password,
            [string]$APIKey
        )

        # 1) Try cached token
        $cached = Read-TokenCache -Username $username -APIKey $APIKey
        if ($cached) {
            $script:OpenSubsUsedCachedToken = $true
            return $cached
        }

        # 2) Login (up to 5 attempts per your request; stop on 401/403)
        $headers = @{
            "Content-Type" = "application/json"
            "User-Agent"   = "Torrentscript"
            "Accept"       = "application/json"
            "Api-Key"      = $APIKey
        }
        $body = @{ username = $username; password = $password } | ConvertTo-Json

        $respHeaders = $null
        $response = Invoke-OpenSubs -Uri 'https://api.opensubtitles.com/api/v1/login' `
            -Method POST -Headers $headers `
            -ContentType 'application/json' -Body $body `
            -MaxAttempts 5 -StopOnAuthError `
            -RespHeaders ([ref]$respHeaders)

        if ($response -and $response.token) {
            # Cache for ~23h
            $expires = (Get-Date).AddHours(23)
            Write-TokenCache -Username $username -APIKey $APIKey -Token $response.token -ExpiresAt $expires
            $script:OpenSubsUsedCachedToken = $false
            Start-Sleep -Seconds 1 # small grace after login
            return $response.token
        } else {
            Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Login failed: token not returned." -ColorBg 'Error'
            return $null
        }
    }

    function Disconnect-OpenSubtitleAPI {
        param(
            [string]$APIKey,
            [string]$token
        )
        $headers = @{
            "User-Agent"    = "Torrentscript"
            "Accept"        = "application/json"
            "Api-Key"       = $APIKey
            "Authorization" = "Bearer $token"
        }
        try {
            [void](Invoke-OpenSubs -Uri 'https://api.opensubtitles.com/api/v1/logout' -Method DELETE -Headers $headers -MaxAttempts 1)
        } catch {
        }
    }

    # ---------------- API helpers ----------------
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

        $headers = @{
            "User-Agent" = "Torrentscript"
            "Api-Key"    = $APIKey
        }
        $encodedQuery = [System.Web.HttpUtility]::UrlEncode($query)

        $uri = "https://api.opensubtitles.com/api/v1/subtitles?type=$type&query=$encodedQuery&languages=$languages&moviehash=$moviehash&hearing_impaired=$hearing_impaired&foreign_parts_only=$foreign_parts_only&machine_translated=$machine_translated&ai_translated=$ai_translated"

        try {
            $respHeaders = $null
            $response = Invoke-OpenSubs -Uri $uri -Method GET -Headers $headers -MaxAttempts 3 -RespHeaders ([ref]$respHeaders)
            if (-not $response -or -not $response.data) {
                return $null 
            }

            $subtitleInfo = @{
                VideoFileName = $query
                SubtitleIds   = @{}
            }

            foreach ($subtitle in $response.data) {
                if (-not $subtitle -or -not $subtitle.attributes) {
                    continue 
                }
                $language = $subtitle.attributes.language
                if (-not $language) {
                    continue 
                }
                $file = $subtitle.attributes.files | Select-Object -First 1
                if (-not $file -or -not $file.file_id) {
                    continue 
                }
                if (-not $subtitleInfo.SubtitleIds.ContainsKey($language)) {
                    $subtitleInfo.SubtitleIds[$language] = $file.file_id
                }
            }

            if ($subtitleInfo.SubtitleIds.Count -gt 0) {
                return $subtitleInfo 
            }
            return $null
        } catch {
            Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred while searching for subtitles:" -ColorBg 'Error'
            Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
            return $null
        }
    }

    function Save-AllSubtitles {
        param(
            [hashtable]$subtitleInfo,
            [string]$APIKey,
            [string]$token,
            [string]$baseDirectory,
            [array]$WantedLanguages
        )

        $downloadedCount = 0
        $failedCount = 0
        $alreadyPresentCount = 0
        $notFoundLanguages = @()  # languages requested but not available in API result
        $languageCounts = @{}
        $lastDailyRemaining = $null

        $headers = @{
            "User-Agent"    = "Torrentscript"
            "Content-Type"  = "application/json"
            "Accept"        = "application/json"
            "Api-Key"       = $APIKey
            "Authorization" = "Bearer $token"
        }

        # Decide per-language action based on "wanted" vs "available" vs "already present"
        $availableLangs = @($subtitleInfo.SubtitleIds.Keys)
        foreach ($lang in $WantedLanguages) {
            $langLower = $lang.ToLower()
            $targetPath = Join-Path $baseDirectory "$($subtitleInfo.VideoFileName).$langLower.srt"

            if ($availableLangs -notcontains $langLower) {
                $notFoundLanguages += $langLower
                continue
            }

            if (Test-Path $targetPath) {
                $alreadyPresentCount++
                continue
            }

            # Download it
            $subtitleId = $subtitleInfo.SubtitleIds[$langLower]
            $body = @{ file_id = $subtitleId } | ConvertTo-Json
            try {
                $respHeaders = $null
                $status = $null
                $response = Invoke-OpenSubs -Uri 'https://api.opensubtitles.com/api/v1/download' `
                    -Method POST -Headers $headers -ContentType 'application/json' `
                    -Body $body -MaxAttempts 3 -RespHeaders ([ref]$respHeaders) -StatusOut ([ref]$status)

                # If unauthorized, refresh token once and retry download
                if (-not $response -and ($status -in 401, 403)) {
                    # Re-login (will use cache or do a fresh login if cache was cleared)
                    $newToken = Connect-OpenSubtitleAPI -username $script:OpenSubCreds.Username -password $OpenSubPass -APIKey $script:OpenSubCreds.APIKey
                    if ($newToken) {
                        # Update Authorization header and retry once
                        $headers["Authorization"] = "Bearer $newToken"
                        $respHeaders = $null
                        $status = $null
                        $response = Invoke-OpenSubs -Uri 'https://api.opensubtitles.com/api/v1/download' `
                            -Method POST -Headers $headers -ContentType 'application/json' `
                            -Body $body -MaxAttempts 3 -RespHeaders ([ref]$respHeaders) -StatusOut ([ref]$status)
                    }
                }

                # If the API exposes daily "remaining" in response body, capture it
                if ($response -and ($response.PSObject.Properties.Name -contains 'remaining')) {
                    $lastDailyRemaining = $response.remaining
                }

                if ($response -and $response.link) {
                    Invoke-WebRequest -Uri $response.link -OutFile $targetPath
                    $downloadedCount++
                    if (-not $languageCounts.ContainsKey($langLower)) {
                        $languageCounts[$langLower] = 1 
                    } else {
                        $languageCounts[$langLower]++ 
                    }
                } else {
                    Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Failed to download subtitle for language: $langLower" -ColorBg 'Error'
                    $failedCount++
                }
            } catch {
                Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred while downloading subtitle for language $($langLower):" -ColorBg 'Error'
                Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
                $failedCount++
            }
        }

        return @{
            Downloaded        = $downloadedCount
            Failed            = $failedCount
            AlreadyPresent    = $alreadyPresentCount
            NotFoundLanguages = $notFoundLanguages
            LanguageCounts    = $languageCounts
            DailyRemaining    = $lastDailyRemaining
        }
    }

    # ---------------- Video hash helper ----------------
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
            $hash = $hash.ToLower()
            if ($hash.Length -lt 16) {
                $hash = ("0" * (16 - $hash.Length)) + $hash 
            }
            $hash
        } finally {
            $stream.Close() 
        }
    }

    # ---------------- Main flow ----------------
    try {
        Write-HTMLLog -Column1 '***  Download missing subs from OpenSubtitle.com  ***' -Header

        if (-not $WantedLanguages -or $WantedLanguages.Count -eq 0) {
            Write-HTMLLog -Column1 'OpenSubs:' -Column2 "No languages requested." -ColorBg 'Error'
            return
        }
        $wantedLangs = ($WantedLanguages | ForEach-Object { $_.ToString().Trim().ToLower() } | Where-Object { $_ -ne '' } | Select-Object -Unique)
        $languageString = $wantedLangs -join ','

        $token = Connect-OpenSubtitleAPI -username $OpenSubUser -password $OpenSubPass -APIKey $OpenSubAPI
        if ($token) {
            $videoFiles = @(Get-ChildItem -LiteralPath $Source -Recurse -Include *.mkv, *.mp4, *.avi | Where-Object { $_.PSIsContainer -eq $false -and $_.DirectoryName -notlike "*\Sample" })
            if ($videoFiles.Count -eq 0) {
                Write-HTMLLog -Column1 'Result:' -Column2 'No video files found' -ColorBg 'Success'
                if (-not $script:OpenSubsUsedCachedToken) {
                    Disconnect-OpenSubtitleAPI -APIKey $OpenSubAPI -token $token 
                }
                return
            }

            $totalDownloaded = 0
            $totalFailed = 0
            $totalAlreadyPresent = 0
            $notFoundByLang = @{}   # lang -> count across files
            $aggregateLangCounts = @{}
            $lastDailyRemaining = $null

            foreach ($videoFile in $videoFiles) {
                # Hash is best-effort; continue even if hashing fails
                $videoHash = $null
                try {
                    $videoHash = Get-VideoHash $videoFile.FullName 
                } catch {
                }

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
                if ($null -eq $subtitleInfo -or -not $subtitleInfo.SubtitleIds -or $subtitleInfo.SubtitleIds.Count -eq 0) {
                    # No matching subtitles found at all for this file
                    foreach ($lang in $wantedLangs) {
                        if (-not $notFoundByLang.ContainsKey($lang)) {
                            $notFoundByLang[$lang] = 0 
                        }
                        $notFoundByLang[$lang]++
                    }
                    continue
                }

                $subtitlesCounts = Save-AllSubtitles -subtitleInfo $subtitleInfo -APIKey $OpenSubAPI -token $token -baseDirectory $videoFile.DirectoryName -WantedLanguages $wantedLangs
                if ($subtitlesCounts) {
                    $totalDownloaded += [int]$subtitlesCounts.Downloaded
                    $totalFailed += [int]$subtitlesCounts.Failed
                    $totalAlreadyPresent += [int]$subtitlesCounts.AlreadyPresent
                    if ($subtitlesCounts.DailyRemaining) {
                        $lastDailyRemaining = $subtitlesCounts.DailyRemaining 
                    }

                    foreach ($lang in $subtitlesCounts.LanguageCounts.Keys) {
                        if (-not $aggregateLangCounts.ContainsKey($lang)) {
                            $aggregateLangCounts[$lang] = 0 
                        }
                        $aggregateLangCounts[$lang] += $subtitlesCounts.LanguageCounts[$lang]
                    }

                    foreach ($nf in $subtitlesCounts.NotFoundLanguages) {
                        if (-not $notFoundByLang.ContainsKey($nf)) {
                            $notFoundByLang[$nf] = 0 
                        }
                        $notFoundByLang[$nf]++
                    }
                }
            }

            # ------- Reporting -------
            if ($totalDownloaded -gt 0) {
                foreach ($language in $aggregateLangCounts.Keys) {
                    Write-HTMLLog -Column1 "Downloaded:" -Column2 "$($aggregateLangCounts[$language]) in $($language.ToUpper())"
                }
                Write-HTMLLog -Column1 "Downloaded:" -Column2 "$totalDownloaded Total"
                if ($totalAlreadyPresent -gt 0) {
                    Write-HTMLLog -Column1 "Already present:" -Column2 "$totalAlreadyPresent (skipped)"
                }
                if ($lastDailyRemaining -ne $null) {
                    Write-HTMLLog -Column2 "Downloads remaining today (last seen): $lastDailyRemaining"
                }
                if ($totalFailed -gt 0) {
                    Write-HTMLLog -Column1 "Failed:" -Column2 "$totalFailed failed to download" -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                } else {
                    Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
                }
            } else {
                # No downloads performed; clarify *why*
                if ($totalAlreadyPresent -gt 0 -and ($notFoundByLang.Keys.Count -eq 0 -or ($notFoundByLang.Values | Measure-Object -Sum).Sum -eq 0)) {
                    Write-HTMLLog -Column1 'Result:' -Column2 'No downloads needed (subtitles already present)' -ColorBg 'Success'
                } else {
                    # Summarize languages not found across files (if any)
                    if ($notFoundByLang.Keys.Count -gt 0) {
                        $summary = ($notFoundByLang.GetEnumerator() | ForEach-Object { "$($_.Key.ToUpper()): $($_.Value)" }) -join ', '
                        Write-HTMLLog -Column1 'Not found:' -Column2 $summary -ColorBg 'Error'
                    }
                    Write-HTMLLog -Column1 'Result:' -Column2 'No suitable subtitles found' -ColorBg 'Error'
                }
            }

            # Logout only if we logged in freshly; keep cached tokens alive for reuse
            if (-not $script:OpenSubsUsedCachedToken) {
                Disconnect-OpenSubtitleAPI -APIKey $OpenSubAPI -token $token
            }
        }
    } catch {
        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred:" -ColorBg 'Error'
        Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
    }
}

