function Compare-JsonKeys {
    <#
    .SYNOPSIS
        Compares two JSON objects recursively for structural differences.
    
    .DESCRIPTION
        This function checks if config.json is missing any keys from config-sample.json
        and also detects any extra keys that exist in config.json but not in config-sample.json.
    
    .PARAMETER ReferenceJson
        The reference JSON object (config-sample.json).
    
    .PARAMETER TestJson
        The user-defined JSON object to validate (config.json).
    
    .PARAMETER Path
        (Optional) A string used for tracking nested paths.
    
    .OUTPUTS
        A hashtable containing:
        - Missing keys from config.json
        - Extra keys in config.json
    
    .EXAMPLE
        Compare-JsonKeys -ReferenceJson $SampleConfig -TestJson $UserConfig
    
        This will compare the user's config.json with the sample and return any discrepancies.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$ReferenceJson,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$TestJson,
        
        [string]$Path = ""
    )

    # Initialize result containers
    $missingKeys = @()
    $extraKeys = @()

    # Check for missing keys
    foreach ($key in $ReferenceJson.PSObject.Properties.Name) {
        if (-not $TestJson.PSObject.Properties[$key]) {
            $missingKeys += "$Path$key"
        } elseif ($ReferenceJson.$key -is [PSCustomObject]) {
            # If key exists and is an object, recurse into it
            $subPath = if ($Path) {
                "$Path$key." 
            } else {
                "$key." 
            }
            $subResult = Compare-JsonKeys -ReferenceJson $ReferenceJson.$key -TestJson $TestJson.$key -Path $subPath
            $missingKeys += $subResult.Missing
            $extraKeys += $subResult.Extra
        }
    }

    # Check for extra keys
    foreach ($key in $TestJson.PSObject.Properties.Name) {
        if (-not $ReferenceJson.PSObject.Properties[$key]) {
            $extraKeys += "$Path$key"
        }
    }

    return @{
        Missing = $missingKeys
        Extra   = $extraKeys
    }
}
