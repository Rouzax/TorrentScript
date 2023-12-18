function Get-QBittorrentCategories {
    <#
    .SYNOPSIS
    Retrieves categories from qBittorrent API.

    .DESCRIPTION
    This function retrieves the categories available in qBittorrent using its API.

    .PARAMETER qBittorrentUrl
    The URL of the qBittorrent WebUI.

    .PARAMETER username
    The username for authentication in the qBittorrent WebUI.

    .PARAMETER password
    The password for authentication in the qBittorrent WebUI.

    .EXAMPLE
    $categories = Get-QBittorrentCategories -qBittorrentUrl "http://localhost:8080" -username "admin" -password "password"
    Retrieves categories from qBittorrent using specified credentials.

    .NOTES
    This function uses qBittorrent's API v2.
    #>
    param (
        [string]$qBittorrentUrl,
        [string]$username,
        [string]$password
    )

    # Construct the base URI for API calls
    $baseUri = "$qBittorrentUrl/api/v2"

    try {
        # Create a session to persist the authentication
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

        # Log in to qBittorrent
        $loginParams = @{
            username = $username
            password = $password
        }
        Invoke-RestMethod -Uri "$baseUri/auth/login" -Method Post -Body $loginParams -WebSession $session -ErrorAction Stop | Out-Null

        # Get categories
        $categoriesResponse = Invoke-RestMethod -Uri "$baseUri/torrents/categories" -WebSession $session -ErrorAction Stop

        # Check if categories are retrieved
        if ($null -ne $categoriesResponse) {
            # Log out from qBittorrent
            Invoke-RestMethod -Uri "$baseUri/auth/logout" -Method Post -WebSession $session -ErrorAction Stop | Out-Null
            return $categoriesResponse
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
        # Additional handling for specific errors can be added here
    }
}