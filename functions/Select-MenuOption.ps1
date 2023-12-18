function Select-MenuOption {
    <#
    .SYNOPSIS
    This function creates a menu with options provided in the $MenuOptions parameter and prompts the user to select an option. The selected option is returned as output.
    
    .DESCRIPTION
    The Select-MenuOption function is used to create a menu with options provided as an array in the $MenuOptions parameter. The function prompts the user to select an option by displaying the options with their corresponding index numbers. The user must enter the index number of the option they wish to select. The function checks if the entered number is within the range of the available options and returns the selected option.
    
    .PARAMETER MenuOptions
    An array of options to be presented in the menu. The options must be of the same data type.
    
    .PARAMETER MenuQuestion
    A string representing the question to be asked when prompting the user for input.
    
    .EXAMPLE
    $Options = @("Option 1","Option 2","Option 3")
    $Question = "an option"
    $SelectedOption = Select-MenuOption -MenuOptions $Options -MenuQuestion $Question
    This example creates a menu with three options "Option 1", "Option 2", and "Option 3". The user is prompted to select an option by displaying the options with their index numbers. The function returns the selected option.
    
    .NOTES
    #>
    param (
        [Parameter(Mandatory = $true)]
        [Object]$MenuOptions,

        [Parameter(Mandatory = $true)]
        [string]$MenuQuestion
    )
    if ($MenuOptions.Count -eq 1) {
        Return $MenuOptions
    } else {
        Write-Host "`nSelect the correct $MenuQuestion" -ForegroundColor DarkCyan
        $menu = @{}
        $maxWidth = [math]::Ceiling([math]::Log10($MenuOptions.Count + 1))
        for ($i = 1; $i -le $MenuOptions.count; $i++) { 
            $indexDisplay = "$i.".PadRight($maxWidth + 2)
            Write-Host "$indexDisplay" -ForegroundColor Magenta -NoNewline
            Write-Host "$($MenuOptions[$i - 1])" -ForegroundColor White 
            $menu.Add($i, ($MenuOptions[$i - 1]))
        }
        do {
            try {
                $numOk = $true
                [int]$ans = Read-Host "Enter $MenuQuestion number to select"
                if ($ans -lt 1 -or $ans -gt $MenuOptions.Count) {
                    $numOK = $false
                    Write-Host 'Not a valid selection' -ForegroundColor DarkRed
                }
            } catch {
                $numOK = $false
                Write-Host 'Please enter a number' -ForegroundColor DarkRed
            }
        } # end do 
        until (($ans -ge 1 -and $ans -le $MenuOptions.Count) -and $numOK)
        Return $MenuOptions[$ans - 1]
    }
}