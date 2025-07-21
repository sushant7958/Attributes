Function Export-CoreData
{
    <#
    .SYNOPSIS
    Exports the CoreData Object.

    .DESCRIPTION
    Provides you with the choice to export a CoreData class object to a file or attach to an Assyst ticket.

    When using the Assyst parameters the path parameter will be used to temporarily store the file. The file will be deleted afterwards.
    
    .EXAMPLE
    Export-CoreData -CoreDataObject $DataObject -Path "C:\Temp\Output"

    Export-CoreData -CoreDataObject $DataObject -Path "C:\Temp\Output" -AssystReference 123456 -Environment "DEV" -AssystDescription "CoreData attached"
    #>
    Param(
        [CmdletBinding(DefaultParameterSetName="Default")]
        [Parameter(ParameterSetName="Default",Mandatory=$True)]
        [Parameter(ParameterSetName="Assyst",Mandatory=$True)]
        [CoreData]$CoreDataObject,

        [Parameter(ParameterSetName="Default",Mandatory=$True)]
        [Parameter(ParameterSetName="Assyst",Mandatory=$True)]
        [String]$Path,

        [Parameter(ParameterSetName="Default",Mandatory=$True)]
        [Parameter(ParameterSetName="Assyst",Mandatory=$True)]
        [String]$TranscriptingGuid,

        [Parameter(ParameterSetName="Default",Mandatory=$True)]
        [Parameter(ParameterSetName="Assyst",Mandatory=$True)]
        [Int]$AssystReference,

        [Parameter(ParameterSetName="Assyst",Mandatory=$True)]
        [ValidateSet("LIVE", "DEV")]
        [String]$Environment,

        [Parameter(ParameterSetName="Assyst",Mandatory=$True)]
        [String]$AssystDescription
    )

    $CoreDataPath = "$Path\DataFile-$AssystReference-$TranscriptingGuid.json"

    ConvertTo-Json -InputObject $CoreDataObject | Out-File $CoreDataPath

    If($PsCmdlet.ParameterSetName -like "Assyst")
    {
        Add-AssystAttachment -IncidentId $AssystReference -Path $CoreDataPath -AssystEnvironment $Environment -Description "New Start Provisioning Log for attempt $ProvisionAttemptGuid"
        Remove-Item -Path $CoreDataPath
    }

}