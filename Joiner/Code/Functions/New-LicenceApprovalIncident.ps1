function New-LicenceApprovalIncident {
    <#
    .SYNOPSIS
        Raises a new Assyst incident for a BU to process the joiner's licence requirements.
    .DESCRIPTION
        Raises a new Assyst incident for a BU to process the joiner's licence requirements and links it to the original Assyst reference.
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [String]
        $AssystReference,
        [Parameter(Mandatory = $True)]
        [PSCustomObject]
        $Configuration,
        [Parameter(Mandatory = $True)]
        [String]
        $AssystQueue,
        [Parameter(Mandatory = $True)]
        [String]
        $Environment,
        [Parameter(Mandatory = $True)]
        [Hashtable]
        $Request,
        [Parameter(Mandatory = $False)]
        [PSObject[]]
        $AutomationTaskList,
        [Parameter(Mandatory = $False)]
        [PSObject[]]
        $BusinessUnitData
    )
    
    Write-Verbose "Entered function $($MyInvocation.MyCommand)."
        
    $NewAssystIncident = @{
        Summary                = "Joiner Automation: New Licence Request."
        AffectedUserShortCode  = 'SSCAUTOMATIONSERVICE'
        SystemServiceShortCode = 'LICENSING SERVICES'
        CategoryShortCode      = 'REQS - LICENSING'
        Priority               = 'SR2 - 4 DAYS'
        Seriousness            = 'SR2'
        AssignedResolverGroup  = $AssystQueue
        ErrorAction            = "Stop"
        AssystEnvironment      = $Environment
        Description            = @"
This Incident was logged by the New Starter Automation process.            

Please determine the user's licence requirements and then assign this incident to the SSC License Admin Assyst queue.

The joiner's details are:
Name: $($Request.Forename) $($Request.Surname)
AD Username: $($Request.ADUserName)
E-mail Address: $($Request.CorporateEmail)
Business Unit: $($Request.BusinessUnitShortcode)
Office: $($Request.Office)
Department: $($Request.DepartmentKey)
Job Title: $($Request.Title)

"@
    }
    try {
        $LicenceIncident = New-AssystIncident @NewAssystIncident -Verbose
        $UpdateTaskListParams = @{
            TaskName = 'Provision Licence' 
            Status   = 'Completed' 
            Result   = 'Success'
        }
        Update-AutomationTaskList @UpdateTaskListParams
    }
    catch {
        $UpdateTaskListParams = @{
            TaskName = 'Provision Licence' 
            Status   = 'Completed' 
            Result   = 'Failed' 
            Detail   = 'An error occured while raising the licence request.'
        }
        Update-AutomationTaskList @UpdateTaskListParams
        Write-Warning 'Error raising Assyst ticket for licence requirements.'
    }

    if ($LicenceIncident.ResponseCode -ne '200') {
        $UpdateTaskListParams = @{
            TaskName = 'Provision Licence'
            Status   = 'Completed' 
            Result   = 'Failed' 
            Detail   = 'An error occured while raising the licence request.'
        }
        Update-AutomationTaskList @UpdateTaskListParams

        $AssystAnalysisInfo = @{
            IncidentID        = $LinkedTaskReference
            Description       = 'New Starter Automation: Unable to raise ticket for licence requirements.'
            AssystEnvironment = $Environment
            ErrorAction       = 'CONTINUE'
        }
        Add-AssystAnalysisInformation @AssystAnalysisInfo

        $AssystAssignmentInfo = @{
            IncidentID                = $LinkedTaskReference
            AssignedServiceDepartment = $Configuration.ResolverGroups.IAM
            AssystEnvironment         = $Environment
            AssigningNotes            = 'New Starter Automation: Unable to raise ticket for licence requirements.'
            ErrorAction               = 'CONTINUE'
        }
        Set-AssystIncidentAssignee @AssystAssignmentInfo

        $InvokeScriptFailureParam = @{
            ResolverGroup         = $GenericConfig.ResolverGroups.AutomationTeam
            FallbackEmail         = @($GenericConfig.Generic.FallbackEmail)
            Summary               = 'New Starter Automation: Unable to raise ticket for licence requirements.'
            AffectedUserShortCode = $GenericConfig.Generic.AffectedUserShortCode
            AssystEnvironment     = $Environment
            SmtpServer            = $GenericConfig.Generic.SmtpServer
            BaseConfigPath        = $BaseConfigPath
            JsonPath              = $JsonPath
            AssystRef             = $AssystReference
            CoreDataObject        = $CoreData
            TranscriptingGuid     = $TranscriptingGuid
            AutomationTaskList    = $AutomationTaskList
            BusinessUnitData      = $BusinessUnitData
            AzAppEmailAddress 	  = $GenericConfig.EmailAzAppReg.AppEmailAddress
            Exit                  = $False
        }
        Invoke-ScriptFailure @InvokeScriptFailureParam
    }
    else {
        Write-Verbose "Assyst reference for licence processing is $($LicenceIncident.IncidentRef)"
        $IncidentLinkParam = @{
            IncidentIDs       = @($AssystReference, $LicenceIncident.IncidentRef)
            LinkDescription   = 'Licence Request'
            AssystEnvironment = $Environment
            ErrorAction       = 'STOP'
        }
        try {
            Write-Verbose "Linking $($LicenceIncident.IncidentRef) and $AssystReference."
            $LinkedResult = Add-AssystIncidentLink @IncidentLinkParam
        }
        catch {
            Write-Warning "Unable to link Assyst incidents $($LicenceIncident.IncidentRef) and $AssystReference."
        }
    
        if ($LinkedResult.ResponseCode -ne '200') {
            $AssystAnalysisInfo = @{
                IncidentID        = $LinkedTaskReference
                AssystEnvironment = $Environment
                ErrorAction       = 'Continue'
                Description       = @"
New Starter Automation: Unable to link tickets for new starter and licence processing.
Please see Assyst incident ref #: $($LicenceIncident.IncidentRef).
"@
            }
            Add-AssystAnalysisInformation @AssystAnalysisInfo

            $AssystAssignmentInfo = @{
                IncidentID                = $LinkedTaskReference
                AssignedServiceDepartment = $Configuration.ResolverGroups.IAM
                AssystEnvironment         = $Environment
                AssigningNotes            = 'New Starter Automation: Unable to link tickets for new starter and licence processing.'
                ErrorAction               = 'Continue'
            }
            Set-AssystIncidentAssignee @AssystAssignmentInfo

            $InvokeScriptFailureParam = @{
                ResolverGroup         = $GenericConfig.ResolverGroups.AutomationTeam
                FallbackEmail         = @($GenericConfig.Generic.FallbackEmail)
                Summary               = 'New Starter Automation: Unable to link tickets for new starter and licence processing.'
                AffectedUserShortCode = $GenericConfig.Generic.AffectedUserShortCode
                AssystEnvironment     = $Environment
                SmtpServer            = $GenericConfig.Generic.SmtpServer
                BaseConfigPath        = $BaseConfigPath
                JsonPath              = $JsonPath
                AssystRef             = $AssystReference
                CoreDataObject        = $CoreData
                TranscriptingGuid     = $TranscriptingGuid
                AutomationTaskList    = $AutomationTaskList
                BusinessUnitData      = $BusinessUnitData
                AzAppEmailAddress 	  = $GenericConfig.EmailAzAppReg.AppEmailAddress
                Exit                  = $False
            }
            Invoke-ScriptFailure @InvokeScriptFailureParam
        }
        else {
            Write-Verbose "Assyst incidents $($LicenceIncident.IncidentRef) and $AssystReference were linked."
        }                
    }
}