function Select-MenuOption {
    <#
    .SYNOPSIS
        Presents a menu of options to the user and allows them to select one.
    .DESCRIPTION
        The Select-MenuOption function is used to create a menu with options provided 
        as an array in the $MenuOptions parameter. 
        The function prompts the user to select an option by displaying the options 
        with their corresponding index numbers. 
        The user must enter the index number of the option they wish to select. 
        The function checks if the entered number is within the range of the available options and returns the selected option.
    .PARAMETER MenuOptions
        An array of options to be presented in the menu. The options must be of the same data type.
    .PARAMETER MenuQuestion
        A string representing the question to be asked when prompting the user for input.
    .OUTPUTS 
        The selected menu option.
    .EXAMPLE
        $Options = @("Option 1","Option 2","Option 3")
        $Question = "an option"
        $SelectedOption = Select-MenuOption -MenuOptions $Options -MenuQuestion $Question
            This example creates a menu with three options "Option 1", "Option 2", and "Option 3". 
            The user is prompted to select an option by displaying the options with their index numbers. 
            The function returns the selected option.
    #>
    param (
        [CmdletBinding()]
        [Parameter(Mandatory = $true)]
        [Object]$MenuOptions,

        [Parameter(Mandatory = $true)]
        [string]$MenuQuestion
    )

    # Check if there is only one option, return it directly
    if ($MenuOptions.Count -eq 1) {
        return $MenuOptions
    } 

    Write-Host "`nSelect the correct $MenuQuestion" -ForegroundColor DarkCyan

    $menu = @{}
    $maxWidth = [math]::Ceiling([math]::Log10($MenuOptions.Count + 1))

    # Display menu options with index numbers
    for ($i = 1; $i -le $MenuOptions.count; $i++) { 
        $indexDisplay = "$i.".PadRight($maxWidth + 2)
        Write-Host "$indexDisplay" -ForegroundColor Magenta -NoNewline
        Write-Host "$($MenuOptions[$i - 1])" -ForegroundColor White 
        $menu.Add($i, ($MenuOptions[$i - 1]))
    }

    # Prompt the user for input and validate the selection
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

    # Return the selected option
    return $MenuOptions[$ans - 1]

}