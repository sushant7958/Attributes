function Update-AutomationTaskList {
    <#
    .SYNOPSIS
        Updates the fields in the automation task list.
    .DESCRIPTION
        The variable $AutomationTaskList represents the list of tasks to be performed by the automation.  This function, which is a helper function,
        can be used to update the fields.  If a field is not included in the list of fields to be updated, any existing data will be preserved.
    .EXAMPLE
        Update-AutomationTaskList -TaskName 'Generate Username' -Status 'Completed' -Result 'Success'
        Updates the Status field and the Result field in the Generate Username task with 'Completed' and 'Success' respectively.  Information in the Detail
        field remains unchanged.
    #>
    
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $True)]
            [String]$TaskName,
            [Parameter(Mandatory = $False)]
            [String]$Status,
            [Parameter(Mandatory = $False)]
            [String]$Result,
            [Parameter(Mandatory = $False)]
            [String]$Detail
        )
        
        $FieldsToUpdate = $PSBoundParameters.Keys -ne 'TaskName'
    
        foreach ($Field in $FieldsToUpdate) {
            ($AutomationTaskList | Where-Object {$_.TaskName -eq $TaskName}).$Field = $PSBoundParameters.$Field
        }
    
    }