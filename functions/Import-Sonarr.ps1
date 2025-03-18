function Import-Sonarr {
    <#
    .SYNOPSIS
        Imports downloaded episodes into Sonarr and monitors the import progress.
    .DESCRIPTION
        This function imports downloaded episodes into Sonarr and monitors the import progress. It uses Sonarr API 
        to initiate the import process and checks the status until completion or timeout.
    .PARAMETER Source
        Specifies the path of the downloaded episode.
    .PARAMETER SonarrApiKey
        Specifies the API key for Sonarr.
    .PARAMETER SonarrHost
        Specifies the host (URL or IP) where Sonarr is running.
    .PARAMETER SonarrPort
        Specifies the port on which Sonarr is listening.
    .PARAMETER SonarrTimeOutMinutes
        Specifies the timeout duration (in minutes) for the Sonarr import operation.
    .PARAMETER TorrentHash
        Specifies the unique identifier (hash) of the downloaded torrent.
    .PARAMETER DownloadLabel
        Specifies a label for the downloaded episode.
    .PARAMETER DownloadName
        Specifies the name of the downloaded episode.
    .OUTPUTS 
        None
    .EXAMPLE
        Import-Sonarr -Source "C:\Downloads\Episode1" -SonarrApiKey "yourApiKey" -SonarrHost "localhost" -SonarrPort 8989 
        -SonarrTimeOutMinutes 30 -TorrentHash "abc123" -DownloadLabel "Action" -DownloadName "Episode1"
        Initiates Sonarr import for the specified episode and monitors the progress.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] 
        [string]$Source,

        [Parameter(Mandatory = $true)] 
        [string]$SonarrApiKey,

        [Parameter(Mandatory = $true)] 
        [string]$SonarrHost,

        [Parameter(Mandatory = $true)] 
        [int]$SonarrPort,

        [Parameter(Mandatory = $true)] 
        [int]$SonarrTimeOutMinutes,

        [Parameter(Mandatory = $false)] 
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
        'name'             = 'DownloadedEpisodesScan'
        'downloadClientId' = $TorrentHash
        'importMode'       = 'Move'
        'path'             = $Source
    } | ConvertTo-Json

    $headers = @{
        'X-Api-Key'    = $SonarrApiKey
        'Content-Type' = 'application/json'
    }
    Write-HTMLLog -Column1 '***  Sonarr Import  ***' -Header
    try {
        $EpisodesScanJob = Invoke-RestMethod -Uri "http://$SonarrHost`:$SonarrPort/api/v3/command" -Method Post -Body $body -Headers $headers
    } catch {
        Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Sonarr Error: $DownloadLabel - $DownloadName"
    }
    if ($EpisodesScanJob.status -eq 'queued' -or $EpisodesScanJob.status -eq 'started' -or $EpisodesScanJob.status -eq 'completed') {
        $timeout = New-TimeSpan -Minutes $SonarrTimeOutMinutes
        $endTime = (Get-Date).Add($timeout)
        do {
            try {
                $EpisodesScanResult = Invoke-RestMethod -Uri "http://$SonarrHost`:$SonarrPort/api/v3/command/$($EpisodesScanJob.id)" -Method Get -Headers $headers
            } catch {
                Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                Stop-Script -ExitReason "Sonarr Error: $DownloadLabel - $DownloadName"
            }
            Start-Sleep 1
        }
        until ($EpisodesScanResult.status -ne 'started' -or ((Get-Date) -gt $endTime) )
        if ($EpisodesScanResult.status -eq 'completed') {
            if ($EpisodesScanResult.result -eq 'successful') {
                Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'         
            } else {
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
                Stop-Script -ExitReason "Sonarr Error: $DownloadLabel - $DownloadName"
            }
        }
        if ($EpisodesScanResult.status -eq 'failed') {
            Write-HTMLLog -Column1 'Sonarr:' -Column2 $EpisodesScanResult.status -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Sonarr:' -Column2 $EpisodesScanResult.exception -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Sonarr Error: $DownloadLabel - $DownloadName"
        }
        if ((Get-Date) -gt $endTime) {
            Write-HTMLLog -Column1 'Sonarr:' -Column2 $EpisodesScanResult.status -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Sonarr:' -Column2 "Import Timeout: ($SonarrTimeOutMinutes) minutes" -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Sonarr Error: $DownloadLabel - $DownloadName"
        }
    } else {
        Write-HTMLLog -Column1 'Sonarr:' -Column2 $EpisodesScanJob.status -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Sonarr Error: $DownloadLabel - $DownloadName"
    }
}