# Description: Checks that the required attributes are popuated in the hashtable
# Accepts: [HashTable]$Data - The hashtable with the populated attributes
# Returns: [Bool] True - All required attributes populated, False - One or more required attributes
#           not populated
Function Test-RequiredAttributes
{
    Param(
        [Parameter(Mandatory = $True)][HashTable]$Data
    )

    Write-Information (("#"*5) + "Test-RequiredAttributes" + ("#"*5) )

    $DataAssystRef = $Data["AssystRef"]
    $DataTaskRef = $Data["TaskRef"]
    $DataBusinessUnit = $Data["BusinessUnit"]
    $DataGivenName = $Data["GivenName"]
    $DataSurname = $Data["Surname"]

    Write-Information -MessageData ("AssystRef: $DataAssystRef" + [System.Environment]::NewLine + `
            "TaskRef: $DataTaskRef" + [System.Environment]::NewLine + `
            "BusinessUnit: $DataBusinessUnit" + [System.Environment]::NewLine + `
            "GivenName: $DataGivenName" + [System.Environment]::NewLine + `
            "Surname: $DataSurname")

    If (
        #[String]::IsNullOrEmpty($Data["AssystRef"]) -Or
        #[String]::IsNullOrEmpty($Data["TaskRef"]) -Or
        [String]::IsNullOrEmpty($Data["BusinessUnit"]) -Or
        [String]::IsNullOrEmpty($Data["GivenName"]) -Or
        [String]::IsNullOrEmpty($Data["Surname"])
    )
    {
        Write-Information "Required attributes NOT present"
        Return $False
    }

    Write-Information "Required attributes present"

    Return $True
}