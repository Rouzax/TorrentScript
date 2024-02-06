<#
.SYNOPSIS
    Imports downloaded movies into Radarr and monitors the import progress.
.DESCRIPTION
    This function imports downloaded movies into Radarr and monitors the import progress. It uses Radarr API 
    to initiate the import process and checks the status until completion or timeout.
.PARAMETER Source
    Specifies the path of the downloaded movie.
.PARAMETER RadarrApiKey
    Specifies the API key for Radarr.
.PARAMETER RadarrHost
    Specifies the host (URL or IP) where Radarr is running.
.PARAMETER RadarrPort
    Specifies the port on which Radarr is listening.
.PARAMETER RadarrTimeOutMinutes
    Specifies the timeout duration (in minutes) for the Radarr import operation.
.PARAMETER TorrentHash
    Specifies the unique identifier (hash) of the downloaded torrent.
.PARAMETER DownloadLabel
    Specifies a label for the downloaded movie.
.PARAMETER DownloadName
    Specifies the name of the downloaded movie.
.OUTPUTS 
    None
.EXAMPLE
    Import-Radarr -Source "C:\Downloads\Movie1" -RadarrApiKey "yourApiKey" -RadarrHost "localhost" -RadarrPort 7878 
    -RadarrTimeOutMinutes 30 -TorrentHash "abc123" -DownloadLabel "Action" -DownloadName "Movie1"
    Initiates Radarr import for the specified movie and monitors the progress.
#>
function Import-Radarr {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] 
        [string]$Source,

        [Parameter(Mandatory = $true)] 
        [string]$RadarrApiKey,

        [Parameter(Mandatory = $true)] 
        [string]$RadarrHost,

        [Parameter(Mandatory = $true)] 
        [int]$RadarrPort,

        [Parameter(Mandatory = $true)] 
        [int]$RadarrTimeOutMinutes,

        [Parameter(Mandatory = $true)] 
        [string]$TorrentHash,

        [Parameter(Mandatory = $true)] 
        [string]$DownloadLabel,

        [Parameter(Mandatory = $true)] 
        [string]$DownloadName
    )

    # Make sure needed functions are available otherwise try to load them.
    $functionsToLoad = @('Write-HTMLLog', 'Stop-Script')
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

    $body = @{
        'name'             = 'DownloadedMoviesScan'
        'downloadClientId' = $TorrentHash
        'importMode'       = 'Move'
        'path'             = $Source
    } | ConvertTo-Json

    $headers = @{
        'X-Api-Key'    = $RadarrApiKey
        'Content-Type' = 'application/json'
    }
    Write-HTMLLog -Column1 '***  Radarr Import  ***' -Header
    try {
        $MoviesScanJob = Invoke-RestMethod -Uri "http://$RadarrHost`:$RadarrPort/api/v3/command" -Method Post -Body $Body -Headers $headers
    } catch {
        Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
    }
    if ($MoviesScanJob.status -eq 'queued' -or $MoviesScanJob.status -eq 'started' -or $MoviesScanJob.status -eq 'completed') {
        $timeout = New-TimeSpan -Minutes $RadarrTimeOutMinutes
        $endTime = (Get-Date).Add($timeout)
        do {
            try {
                $MoviesScanResult = Invoke-RestMethod -Uri "http://$RadarrHost`:$RadarrPort/api/v3/command/$($MoviesScanJob.id)" -Method Get -Headers $headers
            } catch {
                Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
            }
            Start-Sleep 1
        }
        until ($MoviesScanResult.status -ne 'started' -or ((Get-Date) -gt $endTime) )
        if ($MoviesScanResult.status -eq 'completed') {
            if ($MoviesScanResult.result -eq 'successful') {
                Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'         
            } else {
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
                Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
            }
        }
        if ($MoviesScanResult.status -eq 'failed') {
            Write-HTMLLog -Column1 'Radarr:' -Column2 $MoviesScanResult.status -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Radarr:' -Column2 $MoviesScanResult.exception -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
        }
        if ((Get-Date) -gt $endTime) {
            Write-HTMLLog -Column1 'Radarr:' -Column2 $MoviesScanResult.status -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Radarr:' -Column2 "Import Timeout: ($RadarrTimeOutMinutes) minutes" -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
        }
    } else {
        Write-HTMLLog -Column1 'Radarr:' -Column2 $MoviesScanJob.status -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
    }
}