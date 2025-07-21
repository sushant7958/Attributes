Function New-UserPrincipalName {

<#
.SYNOPSIS
    Generates a unique username (UserPrincipalName) based on the SSC naming standards.
.DESCRIPTION
    Generates a unique username (UserPrincipalName) based on the SSC naming standards.
.PARAMETER GivenName
    The first name of the user.
.PARAMETER Surname
    The surname of the user.
.PARAMETER BusinessUnitShortcode
    The business unit shortcode.
.EXAMPLE
    New-UserPrincipalName -GivenName 'John' -Surname 'Smith'
    Returns a username for a standard user account, e.g. John.Smith@bsg.local.
.OUTPUTS
    System.String.  New-UserPrincipalName returns the username as a string.
#>
    
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [String]$GivenName,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [String]$Surname,
        
        [Parameter(Mandatory = $False)]
        [ValidateNotNullOrEmpty()]
        [String]$BusinessUnitShortcode
    )

    Write-Verbose "Entered function $($MyInvocation.MyCommand)."

    switch ($BusinessUnitshortcode) {
        'TWO_UK_CAP' {
            $Domain = '@twinings.com'
        }
        default {
            $Domain = '@bsg.local'
        }
    }

    # Set the maximum length of the user name element.
    $MaxLength = 64

    Write-Information 'Discovering unused user principal name (UPN).'
    $ValidUserName = $False

    $regex = '\&\@\"|\[|\]|\:|\;|\||\=|\+|\*|\?|\<|\>|\/|\\|\,|\s|\.$'
    $BaseUserName = "$($GivenName -replace $regex).$($Surname -replace $regex)"
    $UserPrincipalName = $BaseUserName + $Domain
    $UserNameIncrement = 1

    while ($ValidUserName -eq $False) {
        if ($BaseUserName.Length -gt $MaxLength) {
            $BaseUserName = $BaseUserName.SubString(0, $MaxLength)
            $UserPrincipalName = $BaseUserName + $Domain
        }

        try {
            $AdUser = Get-ADUser -Filter { UserPrincipalName -eq $UserPrincipalName } -ErrorAction 'STOP'
            if ($null -eq $AdUser) {
                $ValidUserName = $true
                $AdUser = $null
            }
        }
        catch {
            throw($_)
        }

        if ($null -ne $AdUser) {
            if ($BaseUserName.Length -ge $MaxLength -and $UserNameIncrement -lt 10) {
                $BaseUserName = $BaseUserName.Substring(0, ($MaxLength - 1))
            }
            elseIf ($BaseUserName.Length -ge ($MaxLength - 1) -and $UserNameIncrement -ge 10) {
                $BaseUserName = $BaseUserName.Substring(0, ($MaxLength - 2))
            }
            $UserPrincipalName = $BaseUserName + $UserNameIncrement + $Domain
            $UserNameIncrement++
        }
    }
    
    $UserPrincipalName
}