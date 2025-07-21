<#
.SYNOPSIS
	Test if a user may have received an incorrect, default mailbox address from Exchange.
.DESCRIPTION
	Tests if the user's domain is set to britishsugar.com and if the user's business unit is set to
	BSU or BSUA. If false, Exchange may have used the default mail rule. An alert is sent to IAM via
	Assyst ticket and via Email using Graph.
.EXAMPLE
	Test-BritishSugarDomain
#>

function Test-BritishSugarDomain {
	
	[CmdletBinding()]
	param (
		
	)

	Write-Information "Entered function $($MyInvocation.MyCommand)."
	
	if ($null -eq $UserDetails.UserPrincipalName) {
		Write-Information "UserPrincipalName is null. Unable to test if the user has the correct domain."
	}
	else {
		Write-Information "User $($UserDetails.UserPrincipalName) belongs to $($CoreData.Attributes["BusinessUnit"])."
		
		if ($UserDetails.UserPrincipalName -notmatch "britishsugar.com" -or
			($UserDetails.UserPrincipalName -match "britishsugar.com" -and
			$CoreData.Attributes["BusinessUnit"] -eq "BSU")) {
			Write-Information "BritishSugar domain test complete. No action required."
		}
		else {
			Write-Warning "Potential issue detected. $($CoreData.Attributes["BusinessUnit"]) should not receive britishsugar.com domain."
			Write-Information "Raising a ticket for IAM for review."
	
			$AssystDescription = @"
This incident was logged by the New Starter Automation process.
Please review user account SamAccountName: $($UserDetails.SamAccountName) created via New Starter Automation - IncidentID $IncidentId
The user belongs to business unit $($CoreData.Attributes["BusinessUnit"]).
Exchange has set the user's domain to britishsugar.com.
This is indicative of Exchange using the default e-mail address policy, therefore the mailbox address may be incorrect.
Please review the account and amend if necessary to the correct domain.
Please ensure the necessary AD attributes are applied to the user based on their business unit.
"@
			$Summary = "New Starter Automation - $($CoreData.Attributes["BusinessUnit"]) user given britishsugar.com domain"

			$NewIncidentParams = @{
				Summary                = $Summary
				AffectedUserShortCode  = 'SSCAUTOMATIONSERVICE'
				SystemServiceShortCode = 'JOINER / MOVER / LEAVER SERVICE'
				CategoryShortCode      = 'REQS - JOINER'
				Priority               = 'SR2 - 4 DAYS'
				Seriousness            = 'SR2'
				AssignedResolverGroup  = $GenericConfig.ResolverGroups.IAM
				ErrorAction            = 'STOP'
				AssystEnvironment      = $Environment
				Description            = $AssystDescription
			}

			try {
				$Incident = New-AssystIncident @NewIncidentParams
				if ($Incident.ResponseCode -ne '200') {
					Write-Warning "Response code was $($Incident.ResponseCode)"
					throw $_
				}
				else {
					Write-Information "Assyst incident ref #: $($Incident.IncidentRef) has been raised."

					# Attach new incident to the joiner's parent ticket

					$IncidentLinkParam = @{
						IncidentIDs       = @($IncidentId, $Incident.IncidentRef)
						LinkDescription   = '[Joiner Automation] - Potential britishsugar.com domain error'
						AssystEnvironment = $Environment
						ErrorAction       = 'STOP'
					}
					try {
						Write-Information "Linking $($Incident.IncidentRef) and $IncidentId."
						$LinkedResult = Add-AssystIncidentLink @IncidentLinkParam
						if ($LinkedResult.ResponseCode -ne '200') {
							Write-Warning "Response code was $($Incident.ResponseCode)"
							throw $_
						}
					}
					catch {
						Write-Warning "Unable to link Assyst incidents $($Incident.IncidentRef) and $IncidentId. $($_.Exception.Message)."
					}
				}
			}
			catch {
				Write-Warning 'Failed to raise a new Assyst incident.'
			}

			$MessageSubject = "[Joiner Automation] - Assyst ref #: $($Incident.IncidentRef) - Potential britishsugar.com domain error"
			# Assyst shows the HTML tags in the description, but emails do not use the here string formatting
			# Creating a copy of the Assyst description with HTML tags for use in the email
			$MessageBody = @"
<p>This incident was logged by the New Starter Automation process.</br></br>
Please review user account SamAccountName: $($UserDetails.SamAccountName) created via New Starter Automation - IncidentID $IncidentId</br>
The user belongs to business unit $($CoreData.Attributes["BusinessUnit"]).</br>
Exchange has set the user's domain to britishsugar.com.</br>
This is indicative of Exchange using the default e-mail address policy, therefore the mailbox address may be incorrect.</br>
Please review the account and amend if necessary to the correct domain.</br>
Please ensure the necessary AD attributes are applied to the user based on their business unit.</br></p>
"@
	
			$MailParams = @{
				UserId          = $GenericConfig.EmailAzAppReg.AppEmailAddress
				SaveToSentItems = $false
				Message         = @{
					Subject = $MessageSubject
					Body    = @{
						ContentType = 'html'
						Content     = "<html><body>" + $MessageBody + "</body></html>"
					}
				ToRecipients    = @(
					@{ EmailAddress = @{ Address = $GenericConfig.ResolverGroupsEmail.IAM } } )
				}
			}

			try{
				Send-MgUserMail @MailParams -ErrorAction Stop
				Write-Information "Email successfully sent to $($GenericConfig.ResolverGroupsEmail.IAM) regarding domain of the user."
			}
			catch {
				Write-Warning "Error sending email to $($GenericConfig.ResolverGroupsEmail.IAM) regarding domain of the user."
			}
		}
	}
}