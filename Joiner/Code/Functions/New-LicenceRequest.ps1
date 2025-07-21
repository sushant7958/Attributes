function New-LicenceRequest {
    <#
    .SYNOPSIS
        Generates a PSCustomObject representing a licence request.
    .DESCRIPTION
        Generates a PSCustomObject representing a licence request.
    .NOTES
        The properties of the object represent the columns on the licence request form supplied by Bytes, who manage licence requests.
    .EXAMPLE
        New-LicenceRequest | Export-CSV E:\Temp\NewLicenceRequest.CSV
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$Action,
        [Parameter(Mandatory = $true)]
        [String]$Forename,
        [Parameter(Mandatory = $true)]
        [String]$Surname,
        [Parameter(Mandatory = $true)]
        [String]$ADUserName,
        [Parameter(Mandatory = $false)]
        [String]$CorporateEmail,
        [Parameter(Mandatory = $true)]
        [String]$BusinessUnitShortCode,
        [Parameter(Mandatory = $false)]
        [String]$Office,
        [Parameter(Mandatory = $false)]
        [String]$DepartmentKey,
        [Parameter(Mandatory = $false)]
        [String]$Title,
        [Parameter(Mandatory = $false)]
        [String]$SKUToBeApplied,
        [Parameter(Mandatory = $false)]
        [String]$SpecialConditions = 'None.'
    )

    Write-Verbose "Entered function $($MyInvocation.MyCommand)."

    $JsonBuConfigurationPath = "D:\Automation\PowerShellModuleConfigurations\SSC-ActiveDirectory\$BusinessUnitShortCode.json"

    try {
        Write-Verbose "Importing configuration file : $JsonBuConfigurationPath"
        $JsonBuConfiguration = Get-Content $JsonBuConfigurationPath | ConvertFrom-Json
        Write-Verbose 'Configuration imported successfully'
    }
    catch {
        Write-Error "Unable to import configuration file.  Exiting script."
    }

    if ($JsonBuConfiguration.Global.LicencesToApply.LicenceName -eq 'Unknown') {
        $LicenceRequest = [PSCustomObject] @{
            Unknown         = $true
            BUApprovalQueue = $JsonBuConfiguration.Global.LicencesToApply.BUApprovalQueue | Where-Object { $Null -ne $_}
        }
    }
    else {
        $LicencesToApply = @()

        $LicencesToApply += $JsonBuConfiguration.Global.LicencesToApply.LicenceName | Where-Object { $Null -ne $_}
        $LicencesToApply += ( $JsonBuConfiguration.Sites | Where-Object { $_.Name -like $Office } ).LicencesToApply.LicenceName
        $LicencesToApply += ( $JsonBuConfiguration.Department | Where-Object { $_.Name -like $DepartmentKey -and $_.Sites.Name -contains $Office } ).LicencesToApply.LicenceName

        $LicencesToApply = $LicencesToApply | Select-Object -Unique

        $LicenceRequest = foreach ($Licence in $LicencesToApply) {

            [PSCustomObject][Ordered] @{
                'ACTION'               = $Action
                'FORENAME'             = $Forename	
                'SURNAME'              = $Surname
                'AD USER NAME'         = $ADUserName
                'CORPORATE EMAIL'      = $CorporateEmail
                'BUSINESS UNIT ABBRV.' = $BusinessUnitShortCode
                'SKU TO BE APPLIED'    = $Licence	
                'SPECIAL CONDITIONS'   = $SpecialConditions
            }
        }
    }

    $LicenceRequest    

}