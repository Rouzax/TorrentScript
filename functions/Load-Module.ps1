<#
.SYNOPSIS
    Loads a PowerShell module and ensures it is installed.

.DESCRIPTION
    This function checks if a specified PowerShell module is already loaded. If not, it checks if the module is installed and installs it if necessary.

.PARAMETER ModuleName
    Specifies the name of the module to load.

.NOTES
    File Name      : Load-Module.ps1
    Prerequisite   : PowerShell V5

.LINK
    https://docs.microsoft.com/en-us/powershell/
#>

function Load-Module {
    [CmdletBinding()]
    param ([Parameter(Mandatory = $true)] 
        [string]$ModuleName
    )

    # Check if Module is already loaded and active
    if (-not (Get-Module -Name $ModuleName)) {
        # If Module not yet loaded, check to see if it is installed
        Write-Host "Check if $ModuleName Module is installed" -ForegroundColor DarkYellow -NoNewline

        if (-not (Get-Module -Name $ModuleName -ListAvailable)) {
            Write-Host " -> FAILED" -ForegroundColor DarkRed

            # Check if NuGet is available as Package Provider
            if (-not (Get-PackageProvider -Name "NuGet")) {
                Write-Host "Check if NuGet is available as Package Provider" -ForegroundColor DarkYellow -NoNewline

                if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                    Write-Host " -> FAILED" -ForegroundColor DarkRed

                    # Install NuGet Package Provider if not available
                    if (Find-PackageProvider -Name NuGet) {
                        Write-Host "    We need to install NuGet first, this might take a while" -ForegroundColor DarkRed
                        Install-PackageProvider -Name NuGet -Force -Confirm:$false -Scope CurrentUser | Out-Null
                        Write-Host "Check if NuGet is available as Package Provider" -ForegroundColor DarkYellow -NoNewline
                    } else {
                        Write-Host "NuGet not imported, not available and not in an online gallery, exiting."
                        Exit 1
                    }
                } else {
                    Write-Host " -> DONE" -ForegroundColor DarkGreen
                }
            }

            # Check if the module is available for installation
            if (Find-Module -Name $ModuleName -ErrorAction SilentlyContinue) {
                Write-Host "    We need to install $ModuleName PowerShell Modules first, this might take a while" -ForegroundColor DarkYellow
                Write-Host "    Installing $ModuleName PowerShell Modules" -ForegroundColor DarkYellow -NoNewline

                # Install the module
                Install-Module -Name $ModuleName -AllowClobber -Force -Confirm:$false -Scope CurrentUser
                Write-Host " -> DONE" -ForegroundColor DarkGreen

                Write-Host "    Loading $ModuleName Module" -ForegroundColor Cyan -NoNewline

                # Import the module
                Import-Module -Name $ModuleName    
                Write-Host " -> DONE" -ForegroundColor DarkGreen
            } else {
                Write-Host "Module $ModuleName not imported, not available and not in an online gallery, exiting."
                Exit 1
            }
        } else {
            Write-Host " -> DONE" -ForegroundColor DarkGreen
            Write-Host "    Loading $ModuleName Module" -ForegroundColor Cyan -NoNewline

            # Import the module
            Import-Module -Name $ModuleName    
            Write-Host " -> DONE" -ForegroundColor DarkGreen
        }
    }
}
