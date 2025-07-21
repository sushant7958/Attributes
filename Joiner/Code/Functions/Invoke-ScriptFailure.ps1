# Description: Used to handle errors within the script
# Accepts: 
#           [Parameter(Mandatory = $True)]$ResolverGroup - Where to assign the ticket to
#           [Parameter(Mandatory = $True)]$FallbackEmail - Email to send failure to if unable to talk to Assyst
#           [Parameter(Mandatory = $True)]$Summary - Summary of the ticket to log
#           [Parameter(Mandatory = $True)]$AffectedUserShortCode - Who to log the ticket in the name of
#           [Parameter(Mandatory = $True)]$AssystEnvironment - Which Assyst environment to log the ticket in (Dev/Live)
#           [Parameter(Mandatory = $True)]$SmtpServer - Details of the Smtp server, likely smtp.bsg.local
# Returns: ????
Function Invoke-ScriptFailure
{
    Param(
        [Parameter(Mandatory = $True)][String]$ResolverGroup,
        [Parameter(Mandatory = $True)][Array]$FallbackEmail,
        [Parameter(Mandatory = $True)][String]$Summary,
        [Parameter(Mandatory = $True)][String]$AffectedUserShortCode,
        [Parameter(Mandatory = $True)][String]$AssystEnvironment,
        [Parameter(Mandatory = $True)][String]$SmtpServer,
        [Parameter(Mandatory = $True)][String]$BaseConfigPath,
        [Parameter(Mandatory = $True)][String]$JsonPath,
        [Parameter(Mandatory = $True)][String]$TranscriptingGuid,
        [Parameter(Mandatory = $True)][CoreData]$CoreDataObject,
        [Parameter(Mandatory = $True)][String]$AssystRef,
        [Parameter(Mandatory = $False)][PSObject]$BusinessUnitData,
        [Parameter(Mandatory = $False)][String]$AzAppEmailAddress,
        [Parameter(Mandatory = $False)][PSObject[]]$AutomationTaskList,
        [Parameter(Mandatory = $True)][Switch]$Exit
    )

    if ($Summary -eq 'The requester and joiner names are the same.' -or
        $Summary -eq 'The user may already exist in Active Directory.') {
            $LogTicket = $false
            $BasePath = "$BaseConfigPath\DuplicateAccounts\$Environment"
    }
    else {
        $LogTicket = $true
        $BasePath = "$BaseConfigPath\FailedFiles\$Environment"
    }

    Try
    {
        if ($Exit) 
        {
            Move-Item -Path $JsonPath -Destination $BasePath -ErrorAction "STOP"
        }    
    }
    Catch
    {
        $ErrorMessage = $_.Exception.Message
        $Body = "Ref: $AssystRef" + [System.Environment]::NewLine + "Unable to move JSON file: $JsonPath" + [System.Environment]::NewLine + "Error: $ErrorMessage"
        Send-MailMessage -To $FallbackEmail -From "JML-JoinerAutomation@abfoods.com" -Subject "[JML-Joiner]Unable to move failed JSON file" -Body $Body -SmtpServer $SmtpServer
        Write-Warning ("[InvokeScriptFailure] Failed to move JSON" + [System.Environment]::NewLine + "Error: $ErrorMessage")
    }

    if ($BusinessUnitData.Global.FailureNotification.SendEmail) {
        $FailureEmailParams = @{
            AssystRef          = $AssystRef
            Summary            = $Summary
            AutomationTaskList = $AutomationTaskList
            Recipients         = $BusinessUnitData.Global.FailureNotification.MailRecipients
            AzAppEmailAddress  = $AzAppEmailAddress
        }
        Write-Information "Sending failure notification e-mail to the business unit."
        try {
            Send-FailureNotificationEmail @FailureEmailParams
        }
        catch {
            Write-Warning "Failed to send the e-mail notification. $($_.Exception.Message)"
        }
        
    }
    else {
        Write-Information "The business unit has not requested e-mail notification of failures."
    }

    if ($LogTicket) {
        Try
        {
            $NewAssystIncident = @{
                Summary = $Summary
                Description = "Please refer to the script logs and investigate further."
                AffectedUserShortCode = $AffectedUserShortCode
                SystemServiceShortCode = "SERVER SERVICES"
                CategoryShortCode = "REQS - ACCESS"
                AssignedResolverGroup = $ResolverGroup
                ErrorAction = "Stop"
                AssystEnvironment = $AssystEnvironment
            }

            $NewTicket = New-AssystIncident @NewAssystIncident
            
            If($NewTicket.ResponseCode -notlike "200")
            {
                Write-Verbose -Message "Assyst response was not HTTP 200. This means it errored."
                $AssystError = ("Unable to raise ticket" + [System.Environment]::NewLine + `
                    "ResponseCode: " + $NewTicket.ResponseCode + [System.Environment]::NewLine + `
                    "ResponseDescription: " + $NewTicket.ResponseDescription  + [System.Environment]::NewLine)
                Write-Error $AssystError
                $TicketData.NewTicketError = $AssystError
                Export-CoreData -CoreDataObject $CoreDataObject -Path $BasePath -TranscriptingGuid $TranscriptingGuid
            }
            Else
            {
                $IncidentId = $NewTicket.IncidentRef                
                Write-Verbose ("Raised ticket" + $IncidentId)
                Write-Verbose "Attaching core data file to the Assyst Ref"
                Export-CoreData -CoreDataObject $CoreDataObject -Path $BasePath -AssystReference $IncidentId -Environment $AssystEnvironment -AssystDescription "CoreData attached" -TranscriptingGuid $TranscriptingGuid
            }
        }
        Catch
        {
            Write-Verbose -Message "Assyst ticket failed, sending email"
            $Body = "Please investigate ticket $AssystRef as the automated script can't interact with Assyst. Is there a problem with Assyst?" + [Environment]::NewLine + "ERROR:$ErrorMessage"
            Send-MailMessage -To $FallbackEmail -From "JML-JoinerAutomation@abfoods.com" -Subject "Unable to raise linked tickets or interact with Assyst. Please investigate!" -Body $Body -SmtpServer $SmtpServer
        }
    }

    if ($Exit) {
        Exit
    }

}