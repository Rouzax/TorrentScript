function Import-Radarr {
    <#
    .SYNOPSIS
    Start Radarr movie import
    
    .DESCRIPTION
    STart the Radarr movie import with error handling
    
    .PARAMETER Source
    Path to movie to import
    
    .EXAMPLE
    Import-Radarr -Source 'C:\Temp\Movie'
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true
        )] 
        [string]$Source
    )

    # Make sure needed functions are available otherwise try to load them.
    $commands = 'Write-HTMLLog', 'Stop-Script'
    foreach ($commandName in $commands) {
        if (!($command = Get-Command $commandName -ErrorAction SilentlyContinue)) {
            Try {
                . $PSScriptRoot\$commandName.ps1
                Write-Host "$commandName Function loaded." -ForegroundColor Green
            }
            Catch {
                Write-Error -Message "Failed to import $commandName function: $_"
                exit 1
            }
        }
    }
    # Start

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
    }
    catch {
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
            }
            catch {
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
            }
            else {
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
    }
    else {
        Write-HTMLLog -Column1 'Radarr:' -Column2 $MoviesScanJob.status -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Radarr Error: $DownloadLabel - $DownloadName"
    }
}