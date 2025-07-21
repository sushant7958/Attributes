Function Import-AssystUserJson
{
    <#
    .SYNOPSIS
    Imports required configuration file 

    .DESCRIPTION
    Imports the required configuration file for the script to run from the path specified in the -ConfigPath parameter

    .OUTPUTS
    $Hashtable containing the keys and values from the Json input
    
    .EXAMPLE
    Import-AssystUserJson -JsonPath "C:\Users\khinton$\Documents\GitHub\SSC-ServerAccess\Config\TestConfig.json"
    #>

    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory=$True)][String]$JsonPath
    )

    Write-Information (("#" * 5) + "IMPORT-JSONCONFIG" + ("#" * 5))

    $Json = Get-Content $JsonPath -ErrorAction "Stop" | ConvertFrom-Json -ErrorAction "Stop"
    Write-Information "New user json imported"

    # The below may make Get-HashtableFromCsv irrelevant as I've taken the code from there and converted it into this function for the outputting of a hashtable.
    $Headers = ($Json | Get-Member -MemberType NoteProperty).Name
    $HashTable = @{ }

    Foreach ($Header in $Headers)
    {
        $HashTable.Add($Header, $Json.$Header)
    }
    Write-Information "Hashtable built from imported json"

    Return $HashTable
}