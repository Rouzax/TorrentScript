<#
.SYNOPSIS
    Installs and imports a PowerShell module, checking if it's already loaded.
.DESCRIPTION
    This function installs a PowerShell module specified by the $ModuleName parameter.
    It checks if the module is already loaded and active. If not, it verifies if the
    module is installed. If not installed, it checks for the availability of NuGet as
    the Package Provider and installs it if necessary. Finally, it installs and loads
    the specified module.
.PARAMETER ModuleName
    Specifies the name of the PowerShell module to be installed and imported.
.INPUTS
    String - You can pipeline the name of the module as a string to this function.
.OUTPUTS
    None - This function does not return any objects.
.EXAMPLE
    Install-PSModule -ModuleName "ExampleModule"
    This example installs and imports the "ExampleModule" PowerShell module.
#>
function Install-PSModule {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true
        )] 
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
