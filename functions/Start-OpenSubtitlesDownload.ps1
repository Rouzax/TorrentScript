function Start-OpenSubtitlesDownload {
    <#
    .SYNOPSIS
        Download missing subtitles from OpenSubtitles.com with token caching and rate-limit handling.
    
    .DESCRIPTION
        Scans a source directory for video files and downloads missing subtitles via the OpenSubtitles API.
        Designed for unattended runs (e.g., post-download), with support for caching tokens, handling rate limits (HTTP 429),
        and distinguishing between “already present” and “not found” subtitles.
    
        Key features:
          • Token cache per user/API key (Windows: %LOCALAPPDATA%\TorrentScript\OpenSubtitles,
            Linux/macOS: $XDG_CACHE_HOME or ~/.cache/TorrentScript/OpenSubtitles).
            Cached tokens are reused; invalid tokens (401/403) clear the cache.
          • Login retried up to 5 times on 429, paced by Retry-After/ratelimit-reset headers.
          • Search/download retried up to 3 times on 429.
          • Logs out only if a fresh login was made (cached tokens remain valid).
          • Logging via Write-HTMLLog (supports 'Success', 'Warning', 'Error').
          • Computes a video hash (length + first/last 64 KiB) to improve matching.
          • Ignores “Sample” folders. Supports filtering for episode/movie type and subtitle attributes.
    
    .PARAMETER Source
        Directory to scan recursively for video files (.mkv, .mp4, .avi).
        "Sample" folders are skipped.
    
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
        ISO 639-1 language codes (e.g., 'en','nl','fr'). Downloads one per requested language per file.
    
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
          -Source "D:\Media\TV\Show\Season 01" `
          -OpenSubUser "user" -OpenSubPass "pass" -OpenSubAPI "apikey" `
          -OpenSubHearing_impaired "exclude" -OpenSubForeign_parts_only "exclude" `
          -OpenSubMachine_translated "exclude" -OpenSubAI_translated "exclude" `
          -WantedLanguages @("en","nl") -Type "episode"
    
    .NOTES
        • Safe for unattended execution; distinguishes “no subtitles needed” vs “none found.”
        • Clears token cache automatically if authentication fails.
        • Retries respect API rate-limit headers; adds small jitter to avoid collisions.
    #>
    [CmdletBinding()]
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

    # Ensure external dependency
    $functionsToLoad = @('Write-HTMLLog')
    foreach ($functionName in $functionsToLoad) {
        if (-not (Get-Command $functionName -ErrorAction SilentlyContinue)) {
            try {
                . "$PSScriptRoot\$functionName.ps1"
                Write-Host "$functionName function loaded." -ForegroundColor Green
            } catch {
                Write-Error "Failed to import $functionName function: $_"
                return
            }
        }
    }

    # Normalize WantedLanguages (639-1, lower)
    $WantedLanguages = @($WantedLanguages | ForEach-Object {
            if ($_ -is [string]) {
                $_.ToLower().Trim() 
            } else {
                $_ 
            }
        })

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
    $script:OpenSubsUsedCachedToken = $false

    # ---------------- Request wrapper ----------------
    function Invoke-OpenSubs {
        param(
            [Parameter(Mandatory)] [string]$Uri,
            [Parameter(Mandatory)] [ValidateSet('GET', 'POST', 'DELETE')] [string]$Method,
            [Parameter(Mandatory)] [hashtable]$Headers,
            [string]$ContentType,
            $Body = $null,
            [int]$MaxAttempts = 3,        # For /login, we’ll use 5
            [switch]$StopOnAuthError,
            [ref]$RespHeaders,
            [ref]$StatusOut,
            [int[]]$SuppressLogForStatus = @()
        )
        if ($StatusOut) {
            $StatusOut.Value = $null 
        }

        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            try {
                $args = @{ Uri = $Uri; Method = $Method; Headers = $Headers }
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
                $status = $null; $hdrs = $null
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
                        Start-Sleep -Seconds $wait
                        continue
                    }
                }

                # If caller asked to suppress logging for this status, just return quietly.
                if ($SuppressLogForStatus -and ($SuppressLogForStatus -contains $status)) {
                    return $null
                }

                Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Request failed (HTTP $status): $($_.Exception.Message)" -ColorBg 'Error'
                return $null
            }
        }
    }

    # ---------------- Auth helpers ----------------
    function Connect-OpenSubtitleAPI {
        param([string]$username, [string]$password, [string]$APIKey)

        # Try cached token first
        $cached = Read-TokenCache -Username $username -APIKey $APIKey
        if ($cached) {
            $script:OpenSubsUsedCachedToken = $true
            return $cached
        }

        # Fresh login (up to 5 attempts, stop on auth error)
        $headers = @{
            "Content-Type" = "application/json"
            "User-Agent"   = "Torrentscript"
            "Accept"       = "application/json"
            "Api-Key"      = $APIKey
        }
        $body = @{ username = $username; password = $password } | ConvertTo-Json

        $respHeaders = $null
        $Parameters = @{
            Uri             = 'https://api.opensubtitles.com/api/v1/login'
            Method          = 'POST'
            Headers         = $headers
            ContentType     = 'application/json'
            Body            = $body
            MaxAttempts     = 5
            StopOnAuthError = $true
            RespHeaders     = ([ref]$respHeaders)
        }
        $response = Invoke-OpenSubs @Parameters

        if ($response -and $response.token) {
            $expires = (Get-Date).AddHours(23)
            Write-TokenCache -Username $username -APIKey $APIKey -Token $response.token -ExpiresAt $expires
            $script:OpenSubsUsedCachedToken = $false
            Start-Sleep -Seconds 1
            return $response.token
        } else {
            Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Login failed: token not returned." -ColorBg 'Error'
            return $null
        }
    }

    function Disconnect-OpenSubtitleAPI {
        param([string]$APIKey, [string]$token)
        if (-not $token) {
            return 
        }
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
            [string]$languagesCsv,
            [string]$moviehash,
            [string]$APIKey,
            [string]$hearing_impaired,
            [string]$foreign_parts_only,
            [string]$machine_translated,
            [string]$ai_translated
        )

        $headers = @{ "User-Agent" = "Torrentscript"; "Api-Key" = $APIKey }
        $encodedQuery = [System.Web.HttpUtility]::UrlEncode($query)

        $uri = "https://api.opensubtitles.com/api/v1/subtitles?type=$type&query=$encodedQuery&languages=$languagesCsv&moviehash=$moviehash&hearing_impaired=$hearing_impaired&foreign_parts_only=$foreign_parts_only&machine_translated=$machine_translated&ai_translated=$ai_translated"

        try {
            $respHeaders = $null
            $response = Invoke-OpenSubs -Uri $uri -Method GET -Headers $headers -MaxAttempts 3 -RespHeaders ([ref]$respHeaders)
            if (-not $response -or -not $response.data) {
                return $null 
            }

            $subtitleInfo = @{
                VideoFileName = $query
                SubtitleIds   = @{}   # language -> file_id
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
                $lng = $language.ToLower()
                if (-not $subtitleInfo.SubtitleIds.ContainsKey($lng)) {
                    $subtitleInfo.SubtitleIds[$lng] = $file.file_id
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
            [ref]$token,
            [string]$baseDirectory,
            [string]$videoBaseName,
            [array]$WantedLanguages,
            [System.Collections.IDictionary]$aggregate # running totals (by ref)
        )

        $tokenRefreshed = $false
    
        $headers = @{
            "User-Agent"    = "Torrentscript"
            "Content-Type"  = "application/json"
            "Accept"        = "application/json"
            "Api-Key"       = $APIKey
            "Authorization" = "Bearer $($Token.Value)"
        }
    
        $availableLangs = @($subtitleInfo.SubtitleIds.Keys)
        $downloadedThisFile = New-Object System.Collections.Generic.List[string]
        $notFoundThisFile = New-Object System.Collections.Generic.List[string]
    
        foreach ($lang in $WantedLanguages) {
            $langLower = ($lang -as [string]).ToLower()
            if (-not $langLower) {
                continue 
            }
    
            $targetPath = Join-Path $baseDirectory "$($videoBaseName).$langLower.srt"
    
            # 1) Already present on disk -> count & skip
            if (Test-Path $targetPath) {
                $aggregate.AlreadyPresent++
                continue
            }
    
            # 2) Not available in API result -> count as "not found"
            if ($availableLangs -notcontains $langLower) {
                if (-not $aggregate.NotFoundByLang.ContainsKey($langLower)) {
                    $aggregate.NotFoundByLang[$langLower] = 1 
                } else {
                    $aggregate.NotFoundByLang[$langLower]++ 
                }
                [void]$notFoundThisFile.Add($langLower)
                continue
            }
    
            # 3) Download
            $subtitleId = $subtitleInfo.SubtitleIds[$langLower]
            $body = @{ file_id = $subtitleId } | ConvertTo-Json
    
            try {
                $respHeaders = $null
                $status = $null
                # First attempt: suppress the noisy 401/403 log; we’ll handle them.
                $Parameters = @{
                    Uri                  = 'https://api.opensubtitles.com/api/v1/download'
                    Method               = 'POST'
                    Headers              = $headers
                    ContentType          = 'application/json'
                    Body                 = $body
                    MaxAttempts          = 3
                    RespHeaders          = ([ref]$respHeaders)
                    StatusOut            = ([ref]$status)
                    SuppressLogForStatus = @(401, 403)
                }
                $response = Invoke-OpenSubs @Parameters

                # If unauthorized, refresh token once and retry (this time DO log if it still fails)
                if (-not $response -and ($status -in 401, 403)) {
                    $newToken = Connect-OpenSubtitleAPI -username $script:OpenSubCreds.Username -password $OpenSubPass -APIKey $script:OpenSubCreds.APIKey
                    if ($newToken) {
                        $Token.Value = $newToken
                        $tokenRefreshed = $true                
                        $headers["Authorization"] = "Bearer $newToken"
                        $respHeaders = $null
                        $status = $null
                        $Parameters = @{
                            Uri         = 'https://api.opensubtitles.com/api/v1/download'
                            Method      = 'POST'
                            Headers     = $headers
                            ContentType = 'application/json'
                            Body        = $body
                            MaxAttempts = 3
                            RespHeaders = ([ref]$respHeaders)
                            StatusOut   = ([ref]$status)
                        }
                        $response = Invoke-OpenSubs @Parameters
                    }
                }

                # If still nothing after the retry, log a proper error and count a failure
                if (-not $response) {
                    Write-HTMLLog -Column2 "Failed to download subtitle for language: $($langLower.ToUpperInvariant()) (HTTP $status)" -ColorBg 'Error'
                    $aggregate.Failed++
                    continue
                }

    
                if ($response -and ($response.PSObject.Properties.Name -contains 'remaining')) {
                    $aggregate.LastDailyRemaining = $response.remaining
                }
    
                if ($response -and $response.link) {
                    Invoke-WebRequest -Uri $response.link -OutFile $targetPath
                    $aggregate.Downloaded++
                    if (-not $aggregate.LanguageCounts.ContainsKey($langLower)) {
                        $aggregate.LanguageCounts[$langLower] = 1 
                    } else {
                        $aggregate.LanguageCounts[$langLower]++ 
                    }
                    [void]$downloadedThisFile.Add($langLower)
                } else {
                    Write-HTMLLog -Column2 "Failed to download subtitle for language: $langLower" -ColorBg 'Error'
                    $aggregate.Failed++
                }
            } catch {
                Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Error while downloading subtitle for $($langLower.ToUpperInvariant()): $($_.Exception.Message)" -ColorBg 'Error'
                $aggregate.Failed++
            }
        }
    
        # Return a concise per-file result so caller can log one line
        return @{
            Downloaded     = $downloadedThisFile.ToArray()
            NotFound       = $notFoundThisFile.ToArray()
            TokenRefreshed = $tokenRefreshed
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
                $lhash = [UInt64](([Decimal]$lhash + [BitConverter]::ToUInt64($buffer, 0)) % ([Decimal]([UInt64]::MaxValue) + 1))
            }
            $lhash
        }
        $stream = $null
        try {
            $stream = [IO.File]::OpenRead($path)
            [UInt64]$lhash = $stream.Length
            $lhash = [UInt64](([Decimal]$lhash + (StreamHash $stream)) % ([Decimal]([UInt64]::MaxValue) + 1))
            $stream.Position = [Math]::Max(0L, $stream.Length - $dataLength)
            $lhash = [UInt64](([Decimal]$lhash + (StreamHash $stream)) % ([Decimal]([UInt64]::MaxValue) + 1))
            $hash = ("{0:X}" -f $lhash).ToLower()
            if ($hash.Length -lt 16) {
                $hash = ("0" * (16 - $hash.Length)) + $hash 
            }
            $hash
        } finally {
            if ($stream) {
                $stream.Close() 
            } 
        }
    }

    # ---------------- Main flow ----------------
    Write-HTMLLog -Column1 '***  Download missing subs from OpenSubtitle.com  ***' -Header

    if (-not (Test-Path $Source)) {
        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Source path not found: $Source" -ColorBg 'Error'
        return
    }

    # Connect
    $token = Connect-OpenSubtitleAPI -username $OpenSubUser -password $OpenSubPass -APIKey $OpenSubAPI
    if (-not $token) {
        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "Aborting: no API token." -ColorBg 'Error'
        return
    }

    # Auth summary line (cached vs new)
    if ($script:OpenSubsUsedCachedToken) {
        Write-HTMLLog -Column1 'Auth:' -Column2 'Using cached OpenSubtitles token'
    } else {
        Write-HTMLLog -Column1 'Auth:' -Column2 'Logged in to OpenSubtitles (new token cached)'
    }

    # Gather video files, skipping any path containing '\Sample\' (case-insensitive)
    $videoExts = @('*.mkv', '*.mp4', '*.avi')
    $allVideos = @()
    foreach ($ext in $videoExts) {
        $allVideos += Get-ChildItem -Path $Source -Recurse -File -Filter $ext -ErrorAction SilentlyContinue
    }
    $videos = $allVideos | Where-Object { $_.FullName -notmatch '(?i)(\\|/)Sample(\\|/)' }

    if (-not $videos -or $videos.Count -eq 0) {
        Write-HTMLLog -Column1 'OpenSubs:' -Column2 "No video files found under $Source" -ColorBg 'Warning'
        if (-not $script:OpenSubsUsedCachedToken) {
            Disconnect-OpenSubtitleAPI -APIKey $OpenSubAPI -token $token 
        }
        return
    }

    # Aggregates
    $aggregate = @{
        Downloaded         = 0
        Failed             = 0
        AlreadyPresent     = 0
        NotFoundByLang     = @{}     # lang -> count
        LanguageCounts     = @{}     # lang -> count
        LastDailyRemaining = $null
    }

    foreach ($video in $videos) {
        $dir = $video.DirectoryName
        $name = [IO.Path]::GetFileNameWithoutExtension($video.Name)

        Write-HTMLLog -Column1 'File:' -Column2 $($video.Name)

        # Determine which languages are missing on disk (per file)
        $missing = @()
        foreach ($lang in $WantedLanguages) {
            $l = ($lang -as [string]).ToLower()
            if (-not $l) {
                continue 
            }
            $expected = Join-Path $dir "$name.$l.srt"
            if (-not (Test-Path $expected)) {
                $missing += $l 
            } else {
                $aggregate.AlreadyPresent++ 
            }
        }

        if ($missing.Count -eq 0) {
            Write-HTMLLog -Column2 "All wanted subtitles already present for this file."
            continue
        }

        # Compute moviehash (best-effort)
        $hash = $null
        try {
            $hash = Get-VideoHash -path $video.FullName 
        } catch {
            $hash = $null 
        }

        # Search for all missing languages in one query to reduce API calls
        $languagesCsv = ($missing -join ',')
        $Parameters = @{
            type               = $Type
            query              = $name
            languagesCsv       = $languagesCsv
            moviehash          = $hash
            APIKey             = $OpenSubAPI
            hearing_impaired   = $OpenSubHearing_impaired
            foreign_parts_only = $OpenSubForeign_parts_only
            machine_translated = $OpenSubMachine_translated
            ai_translated      = $OpenSubAI_translated
        }
        $subtitleInfo = Search-Subtitles @Parameters

        if (-not $subtitleInfo) {
            # Count all missing as "not found"
            foreach ($m in $missing) {
                if (-not $aggregate.NotFoundByLang.ContainsKey($m)) {
                    $aggregate.NotFoundByLang[$m] = 1 
                } else {
                    $aggregate.NotFoundByLang[$m]++ 
                }
            }
            Write-HTMLLog -Column2 "No suitable subtitles found." -ColorBg 'Warning'
            continue
        }

        # Download only the missing ones for this file
        $Parameters = @{
            subtitleInfo    = $subtitleInfo
            APIKey          = $OpenSubAPI
            Token           = ([ref]$token)  
            baseDirectory   = $dir
            videoBaseName   = $name
            WantedLanguages = $missing
            aggregate       = $aggregate
        }
        $fileResult = Save-AllSubtitles @Parameters

        # Concise per-file summary lines
        if ($fileResult -and $fileResult.Downloaded.Count -gt 0) {
            Write-HTMLLog -Column2 ("Downloaded: " + ( ($fileResult.Downloaded | ForEach-Object { $_.ToUpperInvariant() }) -join ', '))
        }
        if ($fileResult -and $fileResult.NotFound.Count -gt 0) {
            Write-HTMLLog -Column2 ("Not found: " + ( ($fileResult.NotFound | ForEach-Object { $_.ToUpperInvariant() }) -join ', ')) -ColorBg 'Warning'
        }

        # Backoff only if we refreshed the token for this file
        if ($fileResult -and $fileResult.TokenRefreshed) {
            Start-Sleep -Milliseconds (200 + (Get-Random -Minimum 0 -Maximum 400))
        }
    }

    # Disconnect only if we did a fresh login
    if (-not $script:OpenSubsUsedCachedToken) {
        Disconnect-OpenSubtitleAPI -APIKey $OpenSubAPI -token $token
    }

    # ------- Reporting -------
    $totalDownloaded = $aggregate.Downloaded
    $totalFailed = $aggregate.Failed
    $totalAlreadyPresent = $aggregate.AlreadyPresent
    $notFoundByLang = $aggregate.NotFoundByLang
    $aggregateLangCounts = $aggregate.LanguageCounts
    $lastDailyRemaining = $aggregate.LastDailyRemaining

    $nfTotal = 0
    if ($notFoundByLang.Keys.Count -gt 0) {
        $nfTotal = ($notFoundByLang.Values | Measure-Object -Sum).Sum
    }

    if ($totalDownloaded -gt 0) {
        foreach ($language in $aggregateLangCounts.Keys) {
            Write-HTMLLog -Column1 "Total:" -Column2 "$($aggregateLangCounts[$language]) in $($language.ToUpperInvariant())"
        }
        if ($totalAlreadyPresent -gt 0) {
            Write-HTMLLog -Column1 "Already present:" -Column2 "$totalAlreadyPresent (skipped)"
        }
        if ($lastDailyRemaining -ne $null) {
            Write-HTMLLog -Column2 "Downloads remaining today: $lastDailyRemaining"
        }
        if ($totalFailed -gt 0) {
            Write-HTMLLog -Column1 "Failed:" -Column2 "$totalFailed failed to download" -ColorBg 'Error'
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        } else {
            Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'
        }
    } else {
        if ($totalAlreadyPresent -gt 0) {
            Write-HTMLLog -Column1 'Already present:' -Column2 "$totalAlreadyPresent (skipped)"
        }
        if ($notFoundByLang.Keys.Count -gt 0) {
            # Note the parentheses: pipeline first, then -join
            $summary = ( $notFoundByLang.GetEnumerator() |
                    ForEach-Object { "$($_.Key.ToUpperInvariant()): $($_.Value)" } ) -join ', '
            Write-HTMLLog -Column1 'Not found:' -Column2 $summary -ColorBg 'Warning'
        }

        if ($nfTotal -eq 0) {
            Write-HTMLLog -Column1 'Result:' -Column2 'No downloads needed (subtitles already present)' -ColorBg 'Success'
        } else {
            Write-HTMLLog -Column1 'Result:' -Column2 'No suitable subtitles downloaded' -ColorBg 'Warning'
        }
    }
}
