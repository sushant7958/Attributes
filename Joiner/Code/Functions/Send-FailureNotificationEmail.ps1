function Send-FailureNotificationEmail {
<#
.SYNOPSIS
    Sends a failure notification e-mail to the specified recipients.
.DESCRIPTION
    Uses MS Graph Send-MgUserMail to send an e-mail to the specified recipients notifying them that the automation failed.
.EXAMPLE
Send-FailureNotificationEmail -AzAppEmailAddress 'automationservice@abfoods.com' -Recipients 'john.smith@abfoods.com'

Sends an e-mail from automationservice@abfoods.com to john.smith@abfoods.com.  As no other parameters are provided, the error message (Summary), 
the Assyst reference, and the Automation Tasklist, may not appear in the e-mail unless the variables used in the message are in the same scope as
the function.
.EXAMPLE
$FailureEmailParams = @{
    AssystRef          = '1234567'
    Summary            = 'An error occurred.'
    AutomationTaskList = $AutomationTaskList
    Recipients         = $BusinessUnitData.Global.FailureNotification.MailRecipients
    AzAppEmailAddress  = $AzAppEmailAddress
}     
Send-FailureNotificationEmail @FailureEmailParams

Sends an email from the address specified in the $AzAppEmailAddress variable to the the one or more recipients contained in the 
Failure Notification section of the configuration file ($BusinessUnitData is the object representation of the JSON configuration file).
The error message in the e-mail will be the Summary parameter 'An error occurred'.    
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [String]$AssystRef,
        [Parameter(Mandatory = $false)]
        [String]$Summary,
        [Parameter(Mandatory = $false)]
        [PSObject[]]$AutomationTaskList,
        [Parameter(Mandatory = $true)] 
        [string]$AzAppEmailAddress,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Recipients
    )

	$IAMSeniorsMailbox = $GenericConfig.ResolverGroupsEmail.IAM
    
    foreach ($MailRecipient in $Recipients) {
        $Subject = "[Joiner Automation] - Failure Notification for Assyst ref #: $AssystRef"

        $MessageText = @"
<p>This is an unmonitored mailbox. Please do not reply.</p>

<p><b>What happened?</b><br>
The automation failed to complete your request.  The error message was:<br>
$Summary</p>

<p><b>What do I need to do?</b></br>
You do not need to take any action because your request has been passed to the BTS IAM team to complete any outstanding tasks.<br>
However, if your request is urgent, you may wish to contact the Service Desk at servicedesk@abfoods.com quoting Assyst ref #: $AssystRef.</p>

<p><b>Automation Task Summary</b>
<pre>
$($AutomationTaskList | Out-String -Width 256)
</pre></p>
"@

        $MailParams = @{
            UserId          = $AzAppEmailAddress
            SaveToSentItems = $false
            Message         = @{
                Subject = $Subject
                Body    = @{
                    ContentType = 'html'
                    Content     = '<html><body>' + $MessageText + '</body></html>'
                }
            ToRecipients    = @(
                @{ EmailAddress = @{ Address = $MailRecipient } } )
			BCCRecipients   = @(
				@{ EmailAddress = @{ Address = $IAMSeniorsMailbox } } )
            }
        }

        try {
            if ($MailParams.SaveToSentItems -ne $false) {
                Write-Warning 'SaveToSentItems is not equal to false.'
                Throw "SaveToSentItems must be false." 
            } 

            Send-MgUserMail @MailParams -ErrorAction Stop
            Write-Information "Email successfully sent to $MailRecipient and $IAMSeniorsMailbox."
        }
        catch {
            Write-Warning 'Error sending email.'
            Throw $_
        }
    
    }

}
