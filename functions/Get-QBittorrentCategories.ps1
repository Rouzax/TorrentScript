<#
.SYNOPSIS
    Retrieves qBittorrent categories using REST API.
.DESCRIPTION
    This function connects to a qBittorrent server via REST API, authenticates using provided
    credentials, and retrieves the list of categories.
.PARAMETER qBittorrentUrl
    Specifies the URL of the qBittorrent server.
.PARAMETER qBittorrentPort
    Specifies the port number of the qBittorrent server.
.PARAMETER username
    Specifies the username for authenticating to the qBittorrent server.
.PARAMETER password
    Specifies the password for authenticating to the qBittorrent server.
.OUTPUTS 
    Returns an array of qBittorrent categories.
.EXAMPLE
    Get-QBittorrentCategories -qBittorrentUrl "http://localhost" -qBittorrentPort 8080 -username "admin" -password "admin123"
    # Retrieves and returns the list of qBittorrent categories.
.NOTES
    This function uses qBittorrent's API v2.
#>
function Get-QBittorrentCategories {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$qBittorrentUrl,

        [Parameter(Mandatory = $true)]
        [int]$qBittorrentPort,

        [Parameter(Mandatory = $true)]
        [string]$username,

        [Parameter(Mandatory = $true)]
        [string]$password
    )

    # Construct the base URI for API calls
    $baseUri = "$qBittorrentUrl`:$qBittorrentPort/api/v2"

    try {
        # Create a session to persist the authentication
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

        # Log in to qBittorrent
        $loginParams = @{
            username = $username
            password = $password
        }
        $Parameters = @{
            Uri         = "$baseUri/auth/login"
            Method      = 'Post'
            Body        = $loginParams
            WebSession  = $session
            ErrorAction = 'Stop'
        }
        Invoke-RestMethod @Parameters | Out-Null

        # Get categories
        $Parameters = @{
            Uri         = "$baseUri/torrents/categories"
            WebSession  = $session
            ErrorAction = 'Stop'
        }
        $categoriesResponse = Invoke-RestMethod @Parameters
        
        # Log out from qBittorrent
        $Parameters = @{
            Uri         = "$baseUri/auth/logout"
            Method      = 'Post'
            WebSession  = $session
            ErrorAction = 'Stop'
        }
        Invoke-RestMethod @Parameters | Out-Null
        
        # Check if categories are retrieved
        if ($null -ne $categoriesResponse) {
            $Categories = $($categoriesResponse.PSObject.Properties.Value.name)
            return $Categories
        } else {
            Write-Host "Categories not found or empty response."
        }
    } catch [System.Net.WebException] {
        Write-Host "WebException occurred: $_"
        Write-Host "Please check if the qBittorrent URL is correct or accessible."
    } catch [System.Management.Automation.ItemNotFoundException] {
        Write-Host "ItemNotFoundException occurred: $_"
        Write-Host "The requested item was not found. Please check the URL or endpoint."
    } catch {
        Write-Host "An unexpected error occurred: $_"
    }
}