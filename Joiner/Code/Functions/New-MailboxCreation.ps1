# Description: Enables a new mailbox on an existing AD account
# Accepts: [String]$SamAccountName - The username of the existing account
# Returns: [Bool] True = successfully created, False = Failed

#Requires -Modules SSC-Exchange

Function New-MailboxCreation
{
    Param(
        [Parameter(Mandatory = $True)][String]$SamAccountName,
        [Parameter(Mandatory = $True)][String]$BusinessUnitShortCode
    )

    Write-Information '##### New-MailboxCreation #####'

    $ConfigurationPaths = Get-SscAdConfigurationPath

    Try
    {
        $BuConfigurationPath = Join-Path -Path $ConfigurationPaths.BuConfiguration -ChildPath "$BusinessUnitShortCode.json"
        $BuConfiguration = (Get-Content $BuConfigurationPath | ConvertFrom-Json).Global
    }
    Catch
    {
        Write-Information "[New-MailboxCreation] unable to import Json Configuration"
        Write-Warning $_.Exception.Message
        Throw($_)
    }

    $MailboxEnvironment = $BuConfiguration.MailboxEnvironment
    $DomainController = Get-SscAdDomainController

    Write-Information "Selected domain controller is: $DomainController"

    If ($Null -eq $DomainController)
    {
        $DomainControllerErrorMessage = "New-MailboxCreation - No domain controller returned from Get-SscADDomainController"
        Write-Warning $DomainControllerErrorMessage
        Throw($DomainControllerErrorMessage)
    }

    Write-Information "Connecting to Exchange"

    try
    {
        Connect-SscExchange -Verbose
    }
    catch
    {
        Write-Information "Connection to Exchange: FAILED!"
        Throw($_)
    }

    $ExchangeQuery = @{
        Identity         = $SamAccountName
        DomainController = $DomainController
        Verbose          = $True
    }

    Write-Information -MessageData "Creating mailbox against account: $SamAccountName"

    Switch ($MailboxEnvironment)
    {
        "ExchangeOnline"
        {
            If ($BuConfiguration.ArchiveMailbox -eq "True")
            {
                $Query.Add("Archive", $True)
            }

            Enable-RemoteMailbox @ExchangeQuery
        }

        "Exchange2016"
        {
            Enable-Mailbox @ExchangeQuery
        }

        "ExchangeOnlineMigration"
        {
            # For potential future development. Would need to output the data so another script could pick it up for migration.
            # I would say call a script directly although then we would have potentially too many powershell processes running on the AOP servers.

            # Examples:
            # Enable-Mailbox $Username #Exchange On-Prem
            # Wait for Azure AD Sync
            # Connect-ExchangeOnline -Credentials $Creds
            # New-MoveRequest -Identity $User -Remote -RemoteHostName owamail.abfoods.com -TargetDeliveryDomain abfoods.mail.onmicrosoft.com -RemoteCredentials $OnPremCredentials
        }
    }

    Write-Information "Confirming domain controller: $DomainController"
    Start-Sleep -Seconds 5
    $UPN = (Get-Mailbox $SamAccountName -DomainController $DomainController -ErrorAction "STOP" -Verbose).PrimarySMTPAddress

    try {

        $SetADUserParam = @{
            Identity           = $SamAccountName
            UserPrincipalName  = $UPN
            Server             = $DomainController
            ErrorAction        = 'STOP' 
        }

        Set-ADUser @SetADUserParam

        $UpdateTaskListParams = @{
            TaskName = 'Set UserPrincipalName'
            Status   = 'Completed'
            Result   = 'Success'
            Detail   = "UserPrincipalName is $UPN."
        }
        Update-AutomationTaskList @UpdateTaskListParams
    }
    catch {
        Write-Warning 'Failed to set UserPrincipalName.'
        $UpdateTaskListParams = @{
            TaskName = 'Set UserPrincipalName'
            Status   = 'Completed'
            Result   = 'Failed'
            Detail   = "An error occurred when setting the UserPrincipalName."
        }
        Update-AutomationTaskList @UpdateTaskListParams
    }

    Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession
}