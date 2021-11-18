
function Import-Medusa
{
    <#
    .SYNOPSIS
    Start import in Medusa
    
    .DESCRIPTION
    Start import to Medusa with timeout and error handeling
    
    .PARAMETER Source
    Path to episode to import

    .EXAMPLE
    Import-Medusa -Source 'C:\Temp\Episode'
    
    .NOTES
    General notes
    #>
    param (
        [Parameter(Mandatory = $true)] 
        $Source
    )

    # Make sure needed functions are available otherwise try to load them.
    $commands = 'Write-HTMLLog', 'Stop-Script'
    foreach ($commandName in $commands)
    {
        if (!($command = Get-Command $commandName -ErrorAction SilentlyContinue))
        {
            Try
            {
                . $PSScriptRoot\$commandName.ps1
                Write-Host "$commandName Function loaded." -ForegroundColor Green
            }
            Catch
            {
                Write-Error -Message "Failed to import $commandName function: $_"
                exit 1
            }
        }
    }
    # Start

    $body = @{
        'proc_dir'       = $Source
        'resource'       = ''
        'process_method' = 'move'
        'force'          = $true
        'is_priority'    = $false
        'delete_on'      = $false
        'failed'         = $false
        'proc_type'      = 'manual'
        'ignore_subs'    = $false
    } | ConvertTo-Json

    $headers = @{
        'X-Api-Key' = $MedusaApiKey
    }
    Write-HTMLLog -Column1 '***  Medusa Import  ***' -Header
    try
    {
        $response = Invoke-RestMethod -Uri "http://$MedusaHost`:$MedusaPort/api/v2/postprocess" -Method Post -Body $Body -Headers $headers
    }
    catch
    {
        Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
    }
    if ($response.status -eq 'success')
    {
        $timeout = New-TimeSpan -Minutes $MedusaTimeOutMinutes
        $endTime = (Get-Date).Add($timeout)
        do
        {
            try
            {
                $status = Invoke-RestMethod -Uri "http://$MedusaHost`:$MedusaPort/api/v2/postprocess/$($response.queueItem.identifier)" -Method Get -Headers $headers
            }
            catch
            {
                Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
            }
            Start-Sleep 1
        }
        until ($status.success -or ((Get-Date) -gt $endTime))
        if ($status.success)
        {
            $ValuesToFind = 'Processing failed', 'aborting post-processing'
            $MatchPattern = ($ValuesToFind | ForEach-Object { [regex]::Escape($_) }) -join '|'
            if ($status.output -match $MatchPattern)
            {
                $ValuesToFind = 'Retrieving episode object for', 'Current quality', 'New quality', 'Old size', 'New size', 'Processing failed', 'aborting post-processing'
                $MatchPattern = ($ValuesToFind | ForEach-Object { [regex]::Escape($_) }) -join '|'
                foreach ($line in $status.output )
                {
                    if ($line -match $MatchPattern)
                    {
                        Write-HTMLLog -Column1 'Medusa:' -Column2 $line -ColorBg 'Error' 
                    }       
                }
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
                Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
            }
            else
            {
                Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'    
            }
        }
        else
        {
            Write-HTMLLog -Column1 'Medusa:' -Column2 $status.success -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Medusa:' -Column2 "Import Timeout: ($MedusaTimeOutMinutes) minutes" -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
        }
    }
    else
    {
        Write-HTMLLog -Column1 'Medusa:' -Column2 $response.status -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
    }
}