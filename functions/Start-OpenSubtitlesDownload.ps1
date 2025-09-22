<#
.SYNOPSIS
    Searches for and downloads missing subtitles from OpenSubtitles.com with rate-limit aware retries and token caching.

.DESCRIPTION
    Start-OpenSubtitlesDownload scans a source directory for video files and downloads missing subtitles
    from the OpenSubtitles API. It is designed to run unattended (e.g., after a download completes) and
    handles common edge cases:

      • Auth: Uses a per-user token cache (Windows: %LOCALAPPDATA%\TorrentScript\OpenSubtitles,
              Linux/macOS: $XDG_CACHE_HOME or ~/.cache/TorrentScript/OpenSubtitles) to avoid repeated
              /login calls. If any API call returns 401/403, the cached token is cleared automatically.
      • Rate limits: On HTTP 429, retries are paced strictly by response headers (Retry-After or ratelimit-reset).
        /login is attempted at most twice and stops immediately on 401/403.
      • Logout: If a fresh login occurred during this run, the function logs out on completion.
                If a cached token was used, logout is skipped to keep the token usable for subsequent runs.
      • Logging: Uses Write-HTMLLog for structured status lines (only ColorBg 'Success' or 'Error').

    The function supports recursive scanning and ignores common "Sample" subfolders. It can filter by
    episode/movie type and by subtitle attributes such as hearing-impaired, foreign-parts-only, machine-
    translated, and AI-translated. Only subtitles in the requested languages are considered.

.PARAMETER Source
    Root directory to scan for video files. The search is recursive.
    Video extensions supported: .mkv, .mp4, .avi
    Sample folders (e.g., "*\Sample") are skipped.

.PARAMETER OpenSubUser
    OpenSubtitles account username used for authentication.

.PARAMETER OpenSubPass
    OpenSubtitles account password used for authentication.

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
    One or more ISO language codes indicating the desired subtitle languages (e.g. 'en','nl','fr').
    The function will attempt to download one matching subtitle per requested language, per file.

.PARAMETER Type
    Content type used by the OpenSubtitles search API.
    Accepted values: episode, movie

.INPUTS
    None. All inputs are provided via parameters.

.OUTPUTS
    None (no objects returned). Writes progress and results using Write-HTMLLog, including:
      - Per-language and total download counts
      - Failed download counts
      - Last-seen per-second remaining allowance (if available from headers)

.EXAMPLE
    Start-OpenSubtitlesDownload `
      -Source "D:\Media\Movies" `
      -OpenSubUser "myuser" `
      -OpenSubPass "mypassword" `
      -OpenSubAPI  "myapikey" `
      -OpenSubHearing_impaired  "exclude" `
      -OpenSubForeign_parts_only "exclude" `
      -OpenSubMachine_translated "exclude" `
      -OpenSubAI_translated "exclude" `
      -WantedLanguages @("en","nl") `
      -Type "movie"

    Recursively scans D:\Media\Movies for video files and downloads English and Dutch subtitles
    while excluding hearing-impaired, foreign-only, machine- and AI-translated variants.

.EXAMPLE
    Start-OpenSubtitlesDownload `
      -Source "E:\TV" `
      -OpenSubUser "myuser" `
      -OpenSubPass "mypassword" `
      -OpenSubAPI  "myapikey" `
      -OpenSubHearing_impaired  "include" `
      -OpenSubForeign_parts_only "include" `
      -OpenSubMachine_translated "only" `
      -OpenSubAI_translated "exclude" `
      -WantedLanguages @("en") `
      -Type "episode"

    Searches for episode-type subtitles under E:\TV with specified attribute filters.

.NOTES
    • Authentication:
        - The function first tries a cached token (per user+API key). If valid, it is reused.
        - If login is needed, /login is attempted at most once with one retry on 429 (MaxAttempts=2),
          and will not retry on 401/403 (bad credentials).
        - If a cached token was used, logout is skipped to keep the token valid for later runs.
          If a fresh login occurred, the function logs out on completion.
        - If any request returns 401/403, the cached token is cleared for the next run.

    • Rate Limiting:
        - Retries on 429 are driven strictly by API response headers (Retry-After or ratelimit-reset),
          with a small randomized jitter to avoid synchronized retries.
        - Search and download requests use up to 3 attempts; login uses up to 2 attempts.

    • Logging:
        - Requires Write-HTMLLog (ColorBg supports only 'Success' or 'Error').

    • Hashing:
        - A video hash is computed (first/last 64 KiB + file length) to improve match accuracy.
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

    # --- Token cache helpers (per user+API key) ---
    function Get-TSCacheRoot {
        if ($IsWindows) {
            return (Join-Path $env:LOCALAPPDATA 'TorrentScript\OpenSubtitles')
        } else {
            $base = $env:XDG_CACHE_HOME
            if (-not $base -or $base -eq '') { $base = (Join-Path $HOME '.cache') }
            return (Join-Path $base 'TorrentScript/OpenSubtitles')
        }
    }
    function Get-TokenCachePath {
        param([string]$Username, [string]$APIKey)
        $root = Get-TSCacheRoot
        if (-not (Test-Path $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }
        $hashInput = "$($Username)`n$($APIKey)"
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hashInput))).Replace('-','').Substring(0,32)
        return (Join-Path $root "token-$hash.json")
    }
    function Read-TokenCache {
        param([string]$Username, [string]$APIKey)
        $path = Get-TokenCachePath -Username $Username -APIKey $APIKey
        if (Test-Path $path) {
            try {
                $data = Get-Content $path -Raw | ConvertFrom-Json
                if ($data -and $data.token -and $data.expires_at) {
                    if ((Get-Date) -lt ([datetime]$data.expires_at)) { return $data.token }
                }
            } catch {}
        }
        return $null
    }
    function Write-TokenCache {
        param([string]$Username, [string]$APIKey, [string]$Token, [datetime]$ExpiresAt)
        $path = Get-TokenCachePath -Username $Username -APIKey $APIKey
        try {
            [pscustomobject]@{
                username   = $Username
                api_key    = '***' # don't persist API key
                token      = $Token
                expires_at = $ExpiresAt.ToString('o')
            } | ConvertTo-Json | Set-Content -Path $path -Encoding UTF8
        } catch {}
    }
    function Clear-TokenCache {
        param([string]$Username, [string]$APIKey)
        $path = Get-TokenCachePath -Username $Username -APIKey $APIKey
        if (Test-Path $path) { Remove-Item $path -Force -ErrorAction SilentlyContinue }
    }

    # Keep creds in script scope so lower helpers can clear cache on 401/403
    $script:OpenSubCreds = @{ Username = $OpenSubUser; APIKey = $OpenSubAPI }

    # --- Header-aware, rate-limit-respecting request wrapper ---
 function Invoke-OpenSubs {
    param(
        [Parameter(Mandatory)] [string]$Uri,
        [Parameter(Mandatory)] [ValidateSet('GET','POST','DELETE')] [string]$Method,
        [Parameter(Mandatory)] [hashtable]$Headers,
        [string]$ContentType,
        $Body = $null,
        [int]$MaxAttempts = 3,        # For /login, use 2
        [switch]$StopOnAuthError,     # Stop immediately on 401/403 (invalid creds)
        [ref]$RespHeaders             # returned via .Value = header dictionary
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $args = @{
                Uri     = $Uri
                Method  = $Method
                Headers = $Headers
            }
            # Ensure compatibility with Windows PowerShell 5.1 (no IE dependency)
            if ($PSVersionTable.PSEdition -eq 'Desktop') { $args.UseBasicParsing = $true }
            if ($PSBoundParameters.ContainsKey('ContentType')) { $args.ContentType = $ContentType }
            if ($null -ne $Body) { $args.Body = $Body }

            $resp = Invoke-WebRequest @args

            # Capture headers for rate-limit accounting
            if ($RespHeaders) { $RespHeaders.Value = $resp.Headers }

            # DELETE endpoints may return no content
            $content = $resp.Content
            if ([string]::IsNullOrWhiteSpace($content)) { return $null }

            # Try to parse JSON; if not JSON, return raw content
            try {
                return ($content | ConvertFrom-Json)
            } catch {
                return $content
            }
        } catch {
            $status = $null
            $hdrs = $null
            try { $status = $_.Exception.Response.StatusCode.value__ } catch {}
            try { $hdrs = $_.Exception.Response.Headers } catch {}

            if ($StopOnAuthError -and ($status -in 401,403)) {
                Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Authentication failed (HTTP $status). Stopping." -ColorBg 'Error'
                return $null
            }

            # Token invalid mid-run? Clear cache so next run relogs.
            if ($status -in 401,403) {
                try { Clear-TokenCache -Username $script:OpenSubCreds.Username -APIKey $script:OpenSubCreds.APIKey } catch {}
            }

            if ($status -eq 429) {
                # Derive wait strictly from headers
                $wait = $null
                if ($hdrs -and $hdrs['Retry-After']) {
                    if (-not [int]::TryParse($hdrs['Retry-After'], [ref]$wait)) {
                        $retryAt = Get-Date $hdrs['Retry-After'] -ErrorAction SilentlyContinue
                        if ($retryAt) {
                            $delta = [int][Math]::Ceiling(($retryAt - (Get-Date)).TotalSeconds)
                            if ($delta -gt 0) { $wait = $delta }
                        }
                    }
                }
                if (-not $wait -and $hdrs -and $hdrs['ratelimit-reset']) {
                    [int]::TryParse($hdrs['ratelimit-reset'], [ref]$wait) | Out-Null
                }
                if (-not $wait -or $wait -lt 1) { $wait = 1 }
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


    # --- OpenSubtitles auth helpers (with token cache + conditional logout) ---
    $script:OpenSubsUsedCachedToken = $false

    function Connect-OpenSubtitleAPI {
        param(
            [string]$username,
            [string]$password,
            [string]$APIKey
        )

        # 1) Try cache first
        $cached = Read-TokenCache -Username $username -APIKey $APIKey
        if ($cached) {
            $script:OpenSubsUsedCachedToken = $true
            return $cached
        }

        # 2) Login (max 2 attempts, stop on 401/403)
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
                                    -MaxAttempts 2 -StopOnAuthError `
                                    -RespHeaders ([ref]$respHeaders)

        if ($response -and $response.token) {
            # Cache for ~23h (refresh ahead of any TTL)
            $expires = (Get-Date).AddHours(23)
            Write-TokenCache -Username $username -APIKey $APIKey -Token $response.token -ExpiresAt $expires
            $script:OpenSubsUsedCachedToken = $false
            Start-Sleep -Seconds 1 # gentle pause after login
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
            # Use wrapper; no retries needed, just a best-effort cleanup
            [void](Invoke-OpenSubs -Uri 'https://api.opensubtitles.com/api/v1/logout' -Method DELETE -Headers $headers -MaxAttempts 1)
        } catch {}
    }

    # --- Search & download helpers ---
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
            if (-not $response -or -not $response.data) { return $null }

            $subtitleInfo = @{
                VideoFileName = $query
                SubtitleIds   = @{}
            }

            foreach ($subtitle in $response.data) {
                if (-not $subtitle -or -not $subtitle.attributes) { continue }
                $language = $subtitle.attributes.language
                if (-not $language) { continue }
                $file = $subtitle.attributes.files | Select-Object -First 1
                if (-not $file -or -not $file.file_id) { continue }
                if (-not $subtitleInfo.SubtitleIds.ContainsKey($language)) {
                    $subtitleInfo.SubtitleIds[$language] = $file.file_id
                }
            }

            if ($subtitleInfo.SubtitleIds.Count -gt 0) { return $subtitleInfo }
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
            [string]$baseDirectory
        )

        $downloadedCount = 0
        $failedCount = 0
        $languageCounts = @{}
        $lastRemaining = $null

        $headers = @{
            "User-Agent"    = "Torrentscript"
            "Content-Type"  = "application/json"
            "Accept"        = "application/json"
            "Api-Key"       = $APIKey
            "Authorization" = "Bearer $token"
        }

        foreach ($language in $subtitleInfo.SubtitleIds.Keys) {
            $subtitleId = $subtitleInfo.SubtitleIds[$language]
            $videoFileName = $subtitleInfo.VideoFileName
            $subtitleFileName = Join-Path $baseDirectory "$videoFileName.$language.srt"

            if (-not (Test-Path $subtitleFileName)) {
                $body = @{ file_id = $subtitleId } | ConvertTo-Json
                try {
                    $respHeaders = $null
                    $response = Invoke-OpenSubs -Uri 'https://api.opensubtitles.com/api/v1/download' `
                                                -Method POST -Headers $headers -ContentType 'application/json' `
                                                -Body $body -MaxAttempts 3 -RespHeaders ([ref]$respHeaders)
                    if ($respHeaders -and $respHeaders['x-ratelimit-remaining-second']) {
                        $lastRemaining = $respHeaders['x-ratelimit-remaining-second']
                    } elseif ($respHeaders -and $respHeaders['ratelimit-remaining']) {
                        $lastRemaining = $respHeaders['ratelimit-remaining']
                    }

                    if ($response -and $response.link) {
                        Invoke-WebRequest -Uri $response.link -OutFile $subtitleFileName
                        $downloadedCount++
                        if (-not $languageCounts.ContainsKey($language)) { $languageCounts[$language] = 1 } else { $languageCounts[$language]++ }
                    } else {
                        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Failed to download subtitle for language: $language" -ColorBg 'Error'
                        $failedCount++
                    }
                } catch {
                    Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred while downloading subtitle for language $($language):" -ColorBg 'Error'
                    Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
                    $failedCount++
                }
            }
        }

        return @{
            Downloaded         = $downloadedCount
            Failed             = $failedCount
            LanguageCounts     = $languageCounts
            RemainingDownloads = $lastRemaining
        }
    }

    # --- Video hash helper (unchanged) ---
    function Get-VideoHash([string]$path) {
        $dataLength = 65536
        function LongSum([UInt64]$a, [UInt64]$b) { [UInt64](([Decimal]$a + $b) % ([Decimal]([UInt64]::MaxValue) + 1)) }
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
            if ($hash.Length -lt 16) { $hash = ("0" * (16 - $hash.Length)) + $hash }
            $hash
        } finally { $stream.Close() }
    }

    #* Start the search and download of the subtitles
    try {
        Write-HTMLLog -Column1 '***  Download missing subs from OpenSubtitle.com  ***' -Header

        # Build language string once
        if (-not $WantedLanguages -or $WantedLanguages.Count -eq 0) {
            Write-HTMLLog -Column1 'OpenSubs:' -Column2 "No languages requested." -ColorBg 'Error'
            return
        }
        $languageString = ($WantedLanguages | ForEach-Object { $_.ToString().Trim().ToLower() } | Where-Object { $_ -ne '' } | Select-Object -Unique) -join ','

        $token = Connect-OpenSubtitleAPI -username $OpenSubUser -password $OpenSubPass -APIKey $OpenSubAPI
        if ($token) {
            $videoFiles = @(Get-ChildItem -LiteralPath $Source -Recurse -Include *.mkv,*.mp4,*.avi | Where-Object { $_.PSIsContainer -eq $false -and $_.DirectoryName -notlike "*\Sample" })
            if ($videoFiles.Count -eq 0) {
                Write-HTMLLog -Column1 'Result:' -Column2 'No video files found' -ColorBg 'Success'
                if (-not $script:OpenSubsUsedCachedToken) { Disconnect-OpenSubtitleAPI -APIKey $OpenSubAPI -token $token }
                return
            }

            $totalDownloaded = 0
            $totalFailed = 0
            $aggregateLangCounts = @{}
            $lastRemaining = $null

            foreach ($videoFile in $videoFiles) {
                $videoHash = $null
                try { $videoHash = Get-VideoHash $videoFile.FullName } catch {}

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
                    continue
                }

                $subtitlesCounts = Save-AllSubtitles -subtitleInfo $subtitleInfo -APIKey $OpenSubAPI -token $token -baseDirectory $videoFile.DirectoryName
                if ($subtitlesCounts) {
                    $totalDownloaded += ($subtitlesCounts.Downloaded   | ForEach-Object { [int]$_ } | Measure-Object -Sum).Sum
                    $totalFailed     += ($subtitlesCounts.Failed       | ForEach-Object { [int]$_ } | Measure-Object -Sum).Sum
                    $lastRemaining    = $subtitlesCounts.RemainingDownloads

                    foreach ($lang in $subtitlesCounts.LanguageCounts.Keys) {
                        if (-not $aggregateLangCounts.ContainsKey($lang)) { $aggregateLangCounts[$lang] = 0 }
                        $aggregateLangCounts[$lang] += $subtitlesCounts.LanguageCounts[$lang]
                    }
                }
            }

            if ($totalDownloaded -gt 0) {
                foreach ($language in $aggregateLangCounts.Keys) {
                    Write-HTMLLog -Column1 "Downloaded:" -Column2 "$($aggregateLangCounts[$language]) in $($language.ToUpper())"
                }
                Write-HTMLLog -Column1 "Downloaded:"  -Column2 "$totalDownloaded Total"
                if ($lastRemaining) { Write-HTMLLog -Column2 "Remaining per-second allowance (last seen): $lastRemaining" }

                if ($totalFailed -gt 0) {
                    Write-HTMLLog -Column1 "Failed:" -Column2 "$totalFailed failed to download" -ColorBg 'Error'
                    Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                } else {
                    Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
                }
            } else {
                Write-HTMLLog -Column1 'Result:' -Column2 'No downloads found or needed' -ColorBg 'Success'
            }

            # Logout only if we logged in freshly; keep cached tokens alive for next runs
            if (-not $script:OpenSubsUsedCachedToken) {
                Disconnect-OpenSubtitleAPI -APIKey $OpenSubAPI -token $token
            }
        }
    } catch {
        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error occurred:" -ColorBg 'Error'
        Write-HTMLLog -Column2 "$($_.Exception.Message)" -ColorBg 'Error'
    }
}
