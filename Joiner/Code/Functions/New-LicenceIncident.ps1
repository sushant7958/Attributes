function New-LicenceIncident {
    <#
    .SYNOPSIS
        Raises a new Assyst incident for processing a joiner's licence requirements.
    .DESCRIPTION
        Raises a new Assyst incident for processing a joiner's licence requirements and links it to the original Assyst reference.
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
        $LicenceRequestPath,
        [Parameter(Mandatory = $True)]
        [String]
        $Environment,
        [Parameter(Mandatory = $False)]
        [PSObject[]]
        $AutomationTaskList,
        [Parameter(Mandatory = $False)]
        [PSObject[]]
        $BusinessUnitData
    )
    
    Write-Verbose "Entered function $($MyInvocation.MyCommand)."
    
    $CurrentDate, $CurrentTime = ((Get-Date -f 'yyyy-MM-dd HH:mm') -split ' ')[0, 1]
    $LaterDate, $LaterTime = (((Get-Date).AddMinutes(45) | Get-Date -f 'yyyy-MM-dd HH:mm') -split ' ')[0, 1]

    $DescriptionPlaceHolder = @(
        $CurrentDate
        $CurrentTime
        $LaterTime
        $LaterDate
    )
        
    $NewAssystIncident = @{
        Summary                = 'Joiner Automation: New Licence Request.'
        AffectedUserShortCode  = 'SSCAUTOMATIONSERVICE'
        SystemServiceShortCode = 'LICENSING SERVICES'
        CategoryShortCode      = 'REQS - LICENSING'
        Priority               = 'SR2 - 4 DAYS'
        Seriousness            = 'SR2'
        AssignedResolverGroup  = $Configuration.ResolverGroups.LicenceTeam
        ErrorAction            = 'STOP'
        AssystEnvironment      = $Environment
        Description            = @'
This incident was logged by the Joiner Automation process.            
Please refer to the attached CSV file which contains the user's licence requirements.

New accounts may take up to 45 minutes to appear in Entra ID.
This account was created on {0} at {1}.  If you cannot find the account, please check again
after {2} on {3}.
The ticket should only be assigned to the service desk queue if the account cannot be found and
more than 45 minutes have elapsed since the account was created.
'@ -f $DescriptionPlaceHolder
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
            ErrorAction       = 'Continue'
        }
        Add-AssystAnalysisInformation @AssystAnalysisInfo

        $AssystAssignmentInfo = @{
            IncidentID                = $LinkedTaskReference
            AssignedServiceDepartment = $Configuration.ResolverGroups.IAM
            AssystEnvironment         = $Environment
            AssigningNotes            = 'New Starter Automation: Unable to raise ticket for licence requirements.'
            ErrorAction               = 'Continue'
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
            AzAppEmailAddress     = $GenericConfig.EmailAzAppReg.AppEmailAddress
            Exit                  = $False
        }
        Invoke-ScriptFailure @InvokeScriptFailureParam
    }
    else {
        Write-Verbose "Assyst reference for licence processing is $($LicenceIncident.IncidentRef)"
        $AttachmentParam = @{
            IncidentId        = $($LicenceIncident.IncidentRef)
            Path              = $LicenceRequestPath
            AssystEnvironment = $Environment
            Description       = 'Licence Request Form.'
        }
        try {
            Write-Verbose "Attaching $LicenceRequestPath to $($LicenceIncident.IncidentRef)."
            Add-AssystAttachment @AttachmentParam
        }
        catch {
            Write-Warning 'Error Adding Attachment.'
        }

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
                AzAppEmailAddress     = $GenericConfig.EmailAzAppReg.AppEmailAddress
                Exit                  = $False
            }
            Invoke-ScriptFailure @InvokeScriptFailureParam
        }
        else {
            Write-Verbose "Assyst incidents $($LicenceIncident.IncidentRef) and $AssystReference were linked."
        }                
    }
}