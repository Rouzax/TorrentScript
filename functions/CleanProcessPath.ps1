function CleanProcessPath {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true
        )] 
        [string]$Path,

        [Parameter(
            Mandatory = $false
        )]
        [bool]$NoCleanUp
    )

    # Make sure needed functions are available otherwise try to load them.
    $commands = 'Write-HTMLLog'
    foreach ($commandName in $commands) {
        if (!($command = Get-Command $commandName -ErrorAction SilentlyContinue)) {
            Try {
                . $PSScriptRoot\$commandName.ps1
                Write-Host "$commandName Function loaded." -ForegroundColor Green
            } Catch {
                Write-Error -Message "Failed to import $commandName function: $_"
                exit 1
            }
        }
    }
    # Start

    if ($NoCleanUp) {
        Write-HTMLLog -Column1 'Cleanup' -Column2 'NoCleanUp switch was given at command line, leaving files'
    } else {
        try {
            If (Test-Path -LiteralPath $ProcessPathFull) {
                Remove-Item -Force -Recurse -LiteralPath $ProcessPathFull
            }
        } catch {
            Write-HTMLLog -Column1 'Exception:' -Column2 $_.Exception.Message -ColorBg 'Error'
            Write-HTMLLog -Column1 'Result:' -Column2 'Failed' -ColorBg 'Error'
        }
    }  
}