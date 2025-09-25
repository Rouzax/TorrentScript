function Import-Medusa {
    <#
    .SYNOPSIS
        Imports items into Medusa for post-processing.
    .DESCRIPTION
        This function initiates the import process in Medusa for post-processing. It takes various parameters 
        such as source directory, Medusa API key, host information, timeout settings, and download details.
    .PARAMETER Source
        Specifies the source directory for the items to be imported.
    .PARAMETER MedusaApiKey
        Specifies the API key for authentication with the Medusa server.
    .PARAMETER MedusaHost
        Specifies the host address of the Medusa server.
    .PARAMETER MedusaPort
        Specifies the port number for communication with the Medusa server.
    .PARAMETER MedusaTimeOutMinutes
        Specifies the timeout duration (in minutes) for the Medusa import operation.
    .PARAMETER DownloadLabel
        Specifies the label for the download being imported.
    .PARAMETER DownloadName
        Specifies the name of the download being imported.
    .OUTPUTS 
        None
    .EXAMPLE
        Import-Medusa -Source "C:\Downloads\Show" -MedusaApiKey "123456" -MedusaHost "medusa.example.com" 
        -MedusaPort 8081 -MedusaTimeOutMinutes 30 -DownloadLabel "TV" -DownloadName "Show"
    
        Initiates the Medusa import for a episode named "Show" in the "C:\Downloads" directory.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] 
        [string]$Source,

        [Parameter(Mandatory = $true)] 
        [string]$MedusaApiKey,

        [Parameter(Mandatory = $true)] 
        [string]$MedusaHost,

        [Parameter(Mandatory = $true)] 
        [int]$MedusaPort,

        [Parameter(Mandatory = $true)] 
        [int]$MedusaTimeOutMinutes,

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

    # Start Medusa import.
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

    # Doing the API call to Medusa
    try {
        $Parameters = @{
            Uri     = "http://$MedusaHost`:$MedusaPort/api/v2/postprocess"
            Method  = 'Post'
            Body    = $Body
            Headers = $headers
        }
        $response = Invoke-RestMethod @Parameters
    } catch {
        Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
    }

    if ($response.status -eq 'success') {
        $timeout = New-TimeSpan -Minutes $MedusaTimeOutMinutes
        $endTime = (Get-Date).Add($timeout)

        # Check progress of Import Job in Medusa and wait till success or Time Out
        do {
            try {
                $Parameters = @{
                    Uri     = "http://$MedusaHost`:$MedusaPort/api/v2/postprocess/$($response.queueItem.identifier)"
                    Method  = 'Get'
                    Headers = $headers
                }
                $status = Invoke-RestMethod @Parameters
            } catch {
                Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
                Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
            }
            Start-Sleep 1
        }
        until ($status.success -or ((Get-Date) -gt $endTime))

        if ($status.success) {
            # Find if there were any errors posted back by Medusa
            $ValuesToFind = 'Processing failed', 'aborting post-processing', 'Unable to figure out what folder to process'
            $MatchPattern = ($ValuesToFind | ForEach-Object { [regex]::Escape($_) }) -join '|'

            if ($status.output -match $MatchPattern) {
                $ValuesToFind = 'Retrieving episode object for', 'Current quality', 'New quality', 'Old size', 'New size', 'Processing failed', 'aborting post-processing', 'Unable to figure out what folder to process'
                $MatchPattern = ($ValuesToFind | ForEach-Object { [regex]::Escape($_) }) -join '|'

                foreach ($line in $status.output ) {
                    if ($line -match $MatchPattern) {
                        Write-HTMLLog -Column1 'Medusa:' -Column2 $line -ColorBg 'Warning' 
                    }       
                }

                Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
                Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
            } else {
                Write-HTMLLog -Column1 'Result:' -Column2 'Successful' -ColorBg 'Success'    
            }
        } else {
            Write-HTMLLog -Column1 'Medusa:' -Column2 $status.success -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Medusa:' -Column2 "Import Timeout: ($MedusaTimeOutMinutes) minutes" -ColorBg 'Error' 
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error' 
            Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
        }
    } else {
        Write-HTMLLog -Column1 'Medusa:' -Column2 $response.status -ColorBg 'Error'
        Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        Stop-Script -ExitReason "Medusa Error: $DownloadLabel - $DownloadName"
    }
}