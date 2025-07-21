Function New-UserName {

<#
.SYNOPSIS
    Generates a unique username (SamAccountName) based on the SSC naming standards.
.DESCRIPTION
    Generates a unique username (SamAccountName) based on the SSC naming standards.
.PARAMETER GivenName
    The first name of the user.
.PARAMETER Surname
    The surname of the user.
.PARAMETER AccountType
    The type of account that's being created.
.EXAMPLE
    New-UserName -GivenName 'John' -Surname 'Smith'
    Returns a username for a standard user account, e.g. JSmith.
.EXAMPLE
    New-UserName -GivenName 'John' -Surname 'Smith' -AccountType 'TPA'
    Returns a username for a 3rd Party account, e.g. JSmith.tpa.
.OUTPUTS
    System.String.  New-Username returns the username as a string.
#>

    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [String]$GivenName,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [String]$Surname,

        [Parameter(Mandatory = $False)]
        [ValidateSet('TPA', 'Standard', 'Admin')]
        [String]$AccountType
    )

    Write-Verbose "Entered function $($MyInvocation.MyCommand)."

    # Set maximum length of the SamAccountName
    $MaxLength = 20

    switch ($AccountType) {
        'TPA' {
            $Prefix = $Null
            $Suffix = '.tpa'
        }
        'Admin' {
            $Prefix = 'Admin_'
            $Suffix = $Null
        }
        default {
            $Prefix = $Null
            $Suffix = $Null
        }
    }

    # Calculate maximum length when suffix is specified
    $Length = $MaxLength - $Suffix.Length

    Write-Information -MessageData 'Discovering unused username.'
    $ValidUserName = $False

    $regex = '\"|\[|\]|\:|\;|\||\=|\+|\*|\?|\<|\>|\/|\\|\,|\s|\.$'
    $BaseUserName = $Prefix + ($GivenName -replace $regex)[0] + ($Surname -replace $regex)
    $UserName = $BaseUserName + $Suffix
    $UserNameIncrement = 1

    While ($ValidUserName -eq $False) {
        If ($BaseUserName.Length -gt $Length) {
            $BaseUserName = $BaseUserName.SubString(0, $Length)
            $UserName = $BaseUserName + $Suffix
        }

        Try {
            $AdUser = Get-ADUser -Identity $UserName -ErrorAction 'STOP'
        }
        Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $ValidUserName = $True
            $AdUser = $Null
        }
        Catch {
            Throw($_)
        }

        If ($Null -ne $AdUser) {
            If ($BaseUserName.Length -ge ($Length - 1) -and $UserNameIncrement -lt 10) {
                $BaseUserName = $BaseUserName.Substring(0, ($BaseUserName.Length - 1))
            }
            ElseIf ($BaseUserName.Length -ge ($Length - 1) -and $UserNameIncrement -ge 10) {
                $BaseUserName = $BaseUserName.Substring(0, ($BaseUserName.Length - 2))
            }
            $UserName = $BaseUserName + $UserNameIncrement + $Suffix
            $UserNameIncrement++
        }
    }
    
    $UserName
}