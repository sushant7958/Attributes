function Send-JoinerEmail {
    <#
    .SYNOPSIS
        Send credentials of the user to the specified email address
    .DESCRIPTION
       The purpose of the Send-JoinerEmail function is to send an email to a recipient's email address
       using an App Registration on Azure and Graph API, containing the New Starter's username, email address, and password. 
    .NOTES
        Function relies on Graph API. Function called from NewStarterProvisioning.ps1
    .EXAMPLE
        $NewStarter = [PSCustomObject]@{
            GivenName    = "Jon"
            Surname      = "Snow"
            Username     = "JSnow"
            Password     = ConvertTo-SecureString -AsPlainText -Force "PassW0rd"
            EmailAddress = "Jon.Snow@abfoods.com"
        }

        $RecipientList.Add([PSCustomObject]@{
            GivenName    = "Ned"
            Surname      = "Stark"
            EmailAddress = "Ned.Stark@abfoods.com"
        })

        $IncidentId = 0123456

        $AzAppEmailAddress = "automationservice@abfoods.com"
    
        Send-JoinerEmail -NewStarter $NewStarter -RecipientList $RecipientList -IncidentId $IncidentId -AzAppEmailAddress $AzAppEmailAddress
        
        The above example would send an email to Ned.Stark@abfoods.com, containing the New Starter user credentials of 
        Jon Snow, including the Username, Password, and Email Address. The email would also contain the general information
        of the new starter including their given name and surname, as well as the Assyst ticket number.
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [PSCustomObject]$NewStarter,
        [Parameter(Mandatory = $true)] [Collections.Generic.List[PSCustomObject]]$RecipientList,
        [Parameter(Mandatory = $true)] [string]$IncidentId,
        [Parameter(Mandatory = $true)] [string]$AzAppEmailAddress,
		[Parameter(Mandatory = $true)] [string]$BusinessUnitId
    )

	$Data = Import-PowerShellDataFile "D:\Automation\NewStarter\Scripts\DataFiles\Send-JoinerEmailDataFile.psd1"

    Write-Verbose "Entered function $($MyInvocation.MyCommand)."

	# Determine if at least one successful email was sent
	$AtLeastOneSuccessfulEmail = $false

    # $ExchangeOnlineMailFlowRule ensures email is encrypted
    $ExchangeOnlineMailFlowRule = '#Secure'
    # Used to set uppercase names, as parameters for recipient givenname and surname come through lowercase
    $UpperCaseName = (Get-Culture).TextInfo

	foreach($Recipient in $RecipientList) {
		# Prepared hashtable in case email fails to send.
		$AnalysisInformation = @{
			IncidentId        = $IncidentId
			Description       = "Failed to send email containing credentials to $($Recipient.EmailAddress)"
			AssystEnvironment = $Environment 
			ErrorAction       = 'CONTINUE'
		}

		if (($Recipient.EmailAddress -as [System.Net.Mail.MailAddress])) {
			$Subject = "$($ExchangeOnlineMailFlowRule) New Starter - $($NewStarter.GivenName) $($NewStarter.Surname) - $($IncidentId)"
			if (!($Recipient.PsObject.Properties.name -match "GivenName") -or !($Recipient.PsObject.Properties.name -match "Surname") -or
				[String]::IsNullOrEmpty($Recipient.GivenName) -or [String]::IsNullOrEmpty($Recipient.Surname)) {
				$Greetings = "<p>Dear colleague</p>"
			}
			else {
				$Greetings = "<p>Dear $($UpperCaseName.ToTitleCase($($Recipient.GivenName))) $($UpperCaseName.ToTitleCase($($Recipient.Surname)))</p>"
			}

		switch ($BusinessUnitId) {
			'TWO_UK_CAP' {
				$MessageText = @"
<p><b>This is an unmonitored mailbox. Please do not reply.</b></p><br>

<p>We have now set up an account for you, please see your login details below.</p>
UPN: $($NewStarter.UserPrincipalName)<br>
Username: $($NewStarter.Username)<br>
Password: <code>$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($NewStarter.Password))))</code></p>
<p>Please find attached instructions for setting up multi-factor authentication (MFA.docx) and registering for self-service password reset (SSPR Registration.docx).</p>
<p>For any queries please contact the Service Desk at  +44 (0) 1733 39 7530 quoting Assyst Reference Number: $($IncidentID)</p>
"@				

			} # end TWO_UK_CAP switch block

			default {
				$MessageText = @"
<p><b>This is an unmonitored mailbox. Please do not reply.</b><br>
For any queries please contact the Service Desk at servicedesk@abfoods.com quoting Assyst Reference Number: $($IncidentID)</p>

<p>As requested, a new account has been created for $($NewStarter.GivenName) $($NewStarter.Surname), please find the login details for this below.<br>
Username: $($NewStarter.Username)<br>
Password: <code>$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($NewStarter.Password))))</code></p>
<p>When logging in for the first time $($NewStarter.GivenName) $($NewStarter.Surname) will be prompted to change their password to something new,
this will need to match the following criteria:<br>
•             Minimum of 14 characters long<br>
•             At least 1 upper case & 1 lower case letters<br>
•             At least 1 number or special character (!, $ etc)<br>
If this account will be used on a laptop, the initial login will need to be made whilst the laptop is connected via network cable to ensure the
account details are stored locally on the laptop.</p>
"@

			if (![bool]$NewStarter.psobject.Properties["EmailAddress"] -or ([String]::IsNullOrEmpty($NewStarter.EmailAddress))) {
				$EmailText = "<p>No mailbox was requested for this user, therefore no email address has been created.</p>"
			}
			else {
				$EmailText = @" 
<p>$($NewStarter.GivenName) $($NewStarter.Surname)'s mailbox has also been created, please find their email address below:<br>
Email Address - $($NewStarter.EmailAddress)<br>
It may take a few days for the new email address to appear in the Global Address list.</p>
"@      
			}

				$BookMyEngineer = @"
<p>As you may be aware, we offer a follow up service to all new starters. This is to ensure that all access is in place 
and to assist with initial setup. Please select the “SSC Book My Engineer (office365.com)” link below and then choose a date & time
that is convenient.<br>
Book My Engineer:<br>
<a href ="https://outlook.office365.com/owa/calendar/ABFSharedServiceCentre@ABFoods.onmicrosoft.com/bookings/">SSC Book My Engineer (office365.com)</a></p>
<p>Please note there may be additional steps required before account setup is finalised.</p>
"@
			} # end default switch block
		}
			$MailParams = @{
				UserId          = $AzAppEmailAddress
				SaveToSentItems = $false
				Message         = @{
					Subject = $Subject
					Body    = @{
						ContentType = 'html'
						Content     = "<html><body>" + $Greetings + $MessageText + $EmailText + $BookMyEngineer + $Data.$BusinessUnitId.CustomMessage + "</body></html>"
					}
				ToRecipients    = @(
					@{ EmailAddress = @{ Address = $Recipient.EmailAddress } } )
				}
			}

			if ($Data.$BusinessUnitId.HasAttachments) {
				[Array]$AttachmentList = foreach ($Path in $Data.$BusinessUnitId.Attachments) {

					# Dynamically set ContentType.
					# List of MIME Types can be found at https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types

					$ContentType = switch ((Get-Item -Path $Path).Extension) {
						'.docx' { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
						'.xlsx' { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
						'.txt'  { 'text/plain' }
						default { 'application/octet-stream' }
					}

					$EncodedFile = [Convert]::ToBase64String((Get-Content -Path $Path -Encoding Byte -ReadCount 0))
					@{
						"@odata.type" = "#microsoft.graph.fileAttachment"
						Name          = "$($Path | Split-Path -Leaf)"
						ContentType   = $ContentType
						ContentBytes  = $EncodedFile
					}
				}
				$MailParams.Message.Attachments = $AttachmentList
			}

			try {
				if ($MailParams.SaveToSentItems -ne $false) {
					Write-Warning 'SaveToSentItems is not equal to false.'
					Throw "SaveToSentItems must be false." 
				} 
				elseif ($ExchangeOnlineMailFlowRule -ne "#Secure" -or !($Subject.Contains($ExchangeOnlineMailFlowRule))) {
					Write-Warning '#Secure is not within the subject.'
					Throw "#Secure must be within the Subject."
				}

				Send-MgUserMail @MailParams -ErrorAction Stop
				Write-Verbose "Email successfully sent to $($Recipient.EmailAddress)."
				$AtLeastOneSuccessfulEmail = $true
			}
			catch {
				Write-Warning 'Error sending email.'
				Add-AssystAnalysisInformation @AnalysisInformation
			}
		}
		else {
			Write-Verbose "Recipient email address $($Recipient.EmailAddress) is invalid, therefore email was not sent to them."
			Add-AssystAnalysisInformation @AnalysisInformation
		}
	}

	if($AtLeastOneSuccessfulEmail) {
		Write-Verbose "At least 1 email address was valid, and an email containing credentials was sent successfully."
	}
	else {
		$Message = "No credential email(s) were sent successfully. IAM intervention required."
		Write-Warning $Message
		throw $Message
	}
}