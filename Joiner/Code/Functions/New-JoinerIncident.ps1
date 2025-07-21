function New-JoinerIncident {
    <#
    .SYNOPSIS
        Raises a new Assyst incident for processing a joiner's account if the account
		is not auto enabled, or if no mailboxes are provided and emailing the line manager
		is set to No.
    .DESCRIPTION
        Raises a new Assyst incident for processing a joiner's account and links it
		to the original Assyst reference.
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$AssystReference,
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Configuration,
        [Parameter(Mandatory = $true)]
        [String]$Environment,
        [Parameter(Mandatory = $true)]
        [String]$IncidentReason
    )
    
    Write-Verbose "Entered function $($MyInvocation.MyCommand)."

    switch ($IncidentReason) {
        'AccountExpirationDate' {
            $IncidentDescription = "The requester has provided the account expiration date to be set but did not provide a valid, future date. Please can IAM reach out to the requester to retrieve a valid account expiration date and set it on the new starter's account."
            $Summary = 'Joiner Automation: Account Expiration Date failed validation'
            $AssystReferenceText = 'Account Expiration Date is not set for a future date'
            $AssignedResolverGroup = $Configuration.ResolverGroups.IAM
            $RaiseTicketCatchWarning = 'Error raising Assyst ticket for incident: Account Expiration Date validation failure.'
            $Not200JoinerIncidentDescription = 'New Starter Automation: Unable to raise ticket for Account Expiration Date validation failure incident.'
            $Not200LinkedResultDescription = 'New Starter Automation: Unable to link the ticket Account Expiration Date validation failure to the ticket New Starter automation.'
            $LinkDescription = 'Account Expiration Date Validation Failure'
        }
        'AccountExpirationAndStartDate' {
            $IncidentDescription = "The requester has provided the account expiration date and user's start date, but the start date is after the expiration date. Please can IAM reach out to the requester to retrieve a valid account expiration date and set it on the new starter's account."
            $Summary = 'Joiner Automation: Account Expiration and Start Date failed validation'
            $AssystReferenceText = 'Start Date is later than Account Expiration Date'
            $AssignedResolverGroup = $Configuration.ResolverGroups.IAM
            $RaiseTicketCatchWarning = 'Error raising Assyst ticket for incident: Start Date is later than Account Expiration Date.'
            $Not200JoinerIncidentDescription = 'New Starter Automation: Unable to raise ticket for Start Date is later than Account Expiration Date validation failure incident.'
            $Not200LinkedResultDescription = 'New Starter Automation: Unable to link the ticket Start Date is later than Account Expiration Date validation failure to the ticket New Starter automation.'
            $LinkDescription = 'Start Date is later than Account Expiration Date Validation Failure'
        }
        'AutoEnable' {
            $IncidentDescription = "The business unit has requested the account NOT be enabled by New Starter Automation. Please can IAM enable the account and contact the business unit to provide the new starter's credentials."
            $Summary = 'Joiner Automation: Auto Enable Account is Disabled'
            $AssystReferenceText = 'auto enable is disabled'
            $AssignedResolverGroup = $Configuration.ResolverGroups.IAM
            $RaiseTicketCatchWarning = 'Error raising Assyst ticket for incident: auto enabling of the account by New Starter Automation is set to false.'
            $Not200JoinerIncidentDescription = 'New Starter Automation: Unable to raise ticket for auto enable account set to false incident.'
            $Not200LinkedResultDescription = 'New Starter Automation: Unable to link the ticket auto enable account set to false to the ticket new starter automation.'
            $LinkDescription = 'Enable Account Request'
        }
        'NoNotificationMailboxes' {
            $IncidentDescription = 'The business unit has declined for new starter credentials be sent to the line manager and have provided no alternative mailbox. Please can IAM enable the account and reset the password, then contact the business unit with the credentials.'
            $Summary = 'Joiner Automation: Enable Account and Reset Password.'
            $AssystReferenceText = 'No available mailbox addresses to send credentials to'
            $AssignedResolverGroup = $Configuration.ResolverGroups.IAM
            $RaiseTicketCatchWarning = 'Error raising Assyst ticket for incident: no mailbox addresses have provided to send new starter credentials to.'
            $Not200JoinerIncidentDescription = 'New Starter Automation: Unable to raise ticket for no mailbox addresses provided incident.'
            $Not200LinkedResultDescription = 'New Starter Automation: Unable to link ticket no mailboxes provided to ticket new starter automation.'
            $LinkDescription = 'Reset Account and Send New Starter Credentials to BU Request'
        }
        'NoLineManager' {
            $IncidentDescription = 'The business unit has requested new starter credentials be sent to the line manager, but no line manager address was provided in the JSON. Please can IAM enable the account and reset the password, then contact the business unit with the credentials.'
            $Summary = 'Joiner Automation: Enable Account and Reset Password.'
            $AssystReferenceText = 'No line manager address to send credentials to'
            $AssignedResolverGroup = $Configuration.ResolverGroups.IAM
            $RaiseTicketCatchWarning = 'Error raising Assyst ticket for incident: no line manager address to send credentials to.'
            $Not200JoinerIncidentDescription = 'New Starter Automation: Unable to raise ticket for no line manager address incident.'
            $Not200LinkedResultDescription = 'New Starter Automation: Unable to link ticket no line manager address to ticket new starter automation.'
            $LinkDescription = 'Reset Account and Send New Starter Credentials to BU Request'
        }
        'NoThirdPartyEmailAddress' {
            $IncidentDescription = 'The business unit sends credentials to a third party email address, but no third party email address was provided in the request. Please can IAM enable the account and reset the password, then contact the third party with the credentials.'
            $Summary = 'Joiner Automation: Enable Account and Reset Password.'
            $AssystReferenceText = 'No third party email address to send credentials to'
            $AssignedResolverGroup = $Configuration.ResolverGroups.IAM
            $RaiseTicketCatchWarning = 'Error raising Assyst ticket for incident: no third party email address to send credentials to.'
            $Not200JoinerIncidentDescription = 'New Starter Automation: Unable to raise ticket for no third party email address incident.'
            $Not200LinkedResultDescription = 'New Starter Automation: Unable to link ticket no third party email address to ticket new starter automation.'
            $LinkDescription = 'Reset Account and Send New Starter Credentials to Third Party Request'
        }
    }
        
    $NewAssystIncident = @{
        Summary                = $Summary
        AffectedUserShortCode  = 'SSCAUTOMATIONSERVICE'
        SystemServiceShortCode = 'JOINER / MOVER / LEAVER SERVICE'
        CategoryShortCode      = 'REQS - JOINER'
        Priority               = 'SR2 - 4 DAYS'
        Seriousness            = 'SR2'
        AssignedResolverGroup  = $AssignedResolverGroup
        ErrorAction            = 'STOP'
        AssystEnvironment      = $Environment
        Description            = @"
This Incident was logged by the New Starter Automation process.            
$IncidentDescription
"@
    }

    try {
        $JoinerIncident = New-AssystIncident @NewAssystIncident -Verbose
    }
    catch {
        Write-Warning $RaiseTicketCatchWarning
    }

    if ($JoinerIncident.ResponseCode -ne '200') {
        $AssystAnalysisInfo = @{
            IncidentID        = $LinkedTaskReference
            Description       = $Not200JoinerIncidentDescription
            AssystEnvironment = $Environment
            ErrorAction       = 'Continue'
        }
        Add-AssystAnalysisInformation @AssystAnalysisInfo

        $AssystAssignmentInfo = @{
            IncidentID                = $LinkedTaskReference
            AssignedServiceDepartment = $Configuration.ResolverGroups.IAM
            AssystEnvironment         = $Environment
            AssigningNotes            = $Not200JoinerIncidentDescription
            ErrorAction               = 'CONTINUE'
        }
        Set-AssystIncidentAssignee @AssystAssignmentInfo

        $InvokeScriptFailureParam = @{
            ResolverGroup         = $GenericConfig.ResolverGroups.AutomationTeam
            FallbackEmail         = @($GenericConfig.Generic.FallbackEmail)
            Summary               = "New Starter Automation: Error Processing Request: $AssystReference."
            AffectedUserShortCode = $GenericConfig.Generic.AffectedUserShortCode
            AssystEnvironment     = $Environment
            SmtpServer            = $GenericConfig.Generic.SmtpServer
            BaseConfigPath        = $BaseConfigPath
            JsonPath              = $JsonPath
            AssystRef             = $AssystReference
            CoreDataObject        = $CoreData
            TranscriptingGuid     = $TranscriptingGuid
            Exit                  = $False
        }
        Invoke-ScriptFailure @InvokeScriptFailureParam

    }
    else {
        Write-Verbose "Assyst reference for $AssystReferenceText is $($JoinerIncident.IncidentRef)"       
        $IncidentLinkParam = @{
            IncidentIDs       = @($AssystReference, $JoinerIncident.IncidentRef)
            LinkDescription   = $LinkDescription
            AssystEnvironment = $Environment
            ErrorAction       = 'STOP'
        }
        try {
            Write-Verbose "Linking $($JoinerIncident.IncidentRef) and $AssystReference."
            $LinkedResult = Add-AssystIncidentLink @IncidentLinkParam
        }
        catch {
            Write-Warning "Unable to link Assyst incidents $($JoinerIncident.IncidentRef) and $AssystReference."
        }
    
        if ($LinkedResult.ResponseCode -ne '200') {
            $AssystAnalysisInfo = @{
                IncidentID        = $LinkedTaskReference
                AssystEnvironment = $Environment
                ErrorAction       = 'CONTINUE'
                Description       = @"
$Not200LinkedResultDescription
Please see Assyst incident ref #: $($JoinerIncident.IncidentRef).
"@
            }
            Add-AssystAnalysisInformation @AssystAnalysisInfo

            $AssystAssignmentInfo = @{
                IncidentID                = $LinkedTaskReference
                AssignedServiceDepartment = $Configuration.ResolverGroups.IAM
                AssystEnvironment         = $Environment
                AssigningNotes            = $Not200LinkedResultDescription
                ErrorAction               = 'Continue'
            }
            Set-AssystIncidentAssignee @AssystAssignmentInfo

            $InvokeScriptFailureParam = @{
                ResolverGroup         = $GenericConfig.ResolverGroups.AutomationTeam
                FallbackEmail         = @($GenericConfig.Generic.FallbackEmail)
                Summary               = "New Starter Automation: Error Processing Request: $AssystReference."
                AffectedUserShortCode = $GenericConfig.Generic.AffectedUserShortCode
                AssystEnvironment     = $Environment
                SmtpServer            = $GenericConfig.Generic.SmtpServer
                BaseConfigPath        = $BaseConfigPath
                JsonPath              = $JsonPath
                AssystRef             = $AssystReference
                CoreDataObject        = $CoreData
                TranscriptingGuid     = $TranscriptingGuid
                Exit                  = $False
            }
            Invoke-ScriptFailure @InvokeScriptFailureParam
        }
        else {
            Write-Verbose "Assyst incidents $($JoinerIncident.IncidentRef) and $AssystReference were linked."
        }                
    }
}