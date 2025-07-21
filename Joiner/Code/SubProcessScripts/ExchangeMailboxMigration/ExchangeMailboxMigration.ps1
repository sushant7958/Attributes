#SSCT:AzureADModule #SSCT:AzureAD #SSCT:ExchangeOnline #SSCT:SSC-Assyst #SSCT:Assyst

$ErrorActionPreference = "STOP"

Function Update-Json
{
    Param(
        [Parameter(Mandatory=$True)][PSCustomObject]$JsonObject,
        [Parameter(Mandatory=$True)][String]$Path
    )

    If(Test-Path $Path)
    {
        $JsonObject | ConvertTo-Json | Out-File $Path
    }
    Else
    {
        Write-Error ("[Update-Json] Invalid path. Error: {0}" -f $_.Exception.Message)
    }
}

#Folder Structure
# Root \
# Scripts \Scripts
# Input \Input
# Completed \Completed
# Failed \Failed
# InvalidJson \InvalidJson
# Transcripts \Logs

$ScriptRootPath = (Get-Item $PSScriptRoot).Parent.FullName

$ScriptPaths =@{
    Input = Join-Path -Path $ScriptRootPath -ChildPath "Input"
    Completed = Join-Path -Path $ScriptRootPath -ChildPath "Completed"
    Failed = Join-Path -Path $ScriptRootPath -ChildPath "Failed"
    InvalidJson = Join-Path -Path $ScriptRootPath -ChildPath "InvalidJson"
    Transcript = Join-Path -Path $ScriptRootPath -ChildPath "Logs"
    Configuration = Join-Path -Path $ScriptRootPath -Child "Configuration\MainConfiguration.json"
}

Start-Transcript -Path ($ScriptPaths["Transcript"] + "\$(Get-Date -f ddMMyy-HHmmss).txt")

Write-Information -MessageData "Script paths: $([System.Environment]::NewLine) $($ScriptPaths | Out-String)"

Write-Information -MessageData "Importing configuration file"
$Configuration = Get-Content $ScriptPaths["Configuration"] | ConvertFrom-Json

<#Configuration structure
AssystUsername
AssystEncryptedPasswordFile

ExchangeServiceAccount
ExchangeEncryptedPasswordFile
#>


# Set Assyst credentials
Set-AssystAPICredential -APIUser $Configuration.AssystUsername -PathToSecureStringFile $Configuration.AssystEncryptedPasswordFile

$JsonChildItems = Get-ChildItem $ScriptPaths["Input"] -File -Filter *.json

Write-Information -MessageData "Json files to process: $([System.Environment]::NewLine) $($JsonChildItems.Name | Out-String)"

<# JSON STRUCTURE
AssystRef
UserPrincipalName
FirstRun
JsonMoveStatus
JsonError
ExchangeMoveStatus
MoveBatchName
AssystEnvironment
#>

Foreach($Item in $JsonChildItems)
{
    Write-Information -MessageData "$("#" * 6;[System.Environment]::NewLine) Processing Json file: $($Item.Name)"
    $JsonFullPath = $Item.FullName

    Try
    {
        $Json = Get-Content $JsonFullPath | ConvertFrom-Json
        Write-Information -MessageData "Json imported successfully."
    }
    Catch
    {
        Move-Item -Path $JsonFullPath -Destination $ScriptPaths["InvalidJson"]
        Continue        
    }

    Write-Information -MessageData "JsonMoveStatus: $($Json.MoveStatus)"

    Switch($Json.JsonMoveStatus)
    {
        New
        {
            If([String]::IsNullOrEmpty($Json.FirstRun))
            {
                Write-Information -MessageData "First run detected."
                $Json.FirstRun = Get-Date -Format "dd/MM/yy HH:mm"
            }

            Try
            {
                $Account = Get-AzureADUser $Json.UserPrincipalName
                Write-Information "Account found in Azure AD"
            }
            Catch
            {
                If(($Json.FirstRun - (Get-Date).Hours) -gt 24)
                {
                    $FailureMessage = "Account hasn't synced up in over 24hrs"
                    Write-Warning -Message $FailureMessage
                    $Json.JsonMoveStatus = "Failed"
                    $Json.JsonError = $FailureMessage
                }
            }


            If($Null -ne $Account)
            {
                Try
                {
                    Write-Information -MessageData "Connecting to Exchange Online"
                    $ExchangeUserName = $Configuration.ExchangeServiceAccount
                    $ExchangePassword = Get-Content -Path $Configuration.ExchangeEncryptedPasswordFile | ConvertTo-SecureString
                    $ExchangeCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList $ExchangeUserName, $ExchangePassword
                    Remove-Variable -Name ExchangeUserName
                    Remove-Variable -Name ExchangePassword

                    Connect-ExchangeOnline -Credential $ExchangeCredentials
                    Remove-Variable -Name ExchangeCredentials

                    Write-Information -MessageData "Exchange Online connection successful. Initiating move request"
                    $NewMoveRequestParam =@{
                        Identity = $Account
                        Remote = $true
                        RemoteHostname = "owamail.abfoods.com"
                        TargetDeliveryDomain = "abfoods.mail.onmicrosoft.com"
                        RemoteCredential = $OnPremCredentials
                        BatchName = "NewStarterMove-$($Json.AssystRef)"
                    }

                    Write-Information -MessageData "Move parameters: $([System.Environment]::NewLine) $($NewMoveRequestParam | Out-String)"

                    $NewMoveRequesst = New-MoveRequest @NewMoveRequestParam
                    $Json.JsonMoveStatus = "Moving"
                    $Json.ExchangeMoveStatus = $NewMoveRequesst.Status
                    $Json.MoveBatchName = $NewMoveRequestParam.BatchName
                    Write-Information -MessageData "Move Initiated."
                }
                Catch
                {
                    If(($Json.FirstRun - (Get-Date).Hours) -gt 24)
                    {
                        $FailureMessage = ("Failed to start move request. ERROR: {0}" -f $_.Exception.Message)
                        Write-Warning -Message $FailureMessage
                        $Json.JsonMoveStatus = "Failed"
                        $Json.JsonError = $FailureMessage
                    }
                }
            }

            Update-Json -JsonObject $JsonObject -Path $JsonFullPath
        }

        Moving
        {
            $LiveMoveStatus = Get-MoveRequest -BatchName $Json.MoveBatchName
            $Json.ExchangeMoveStatus = $LiveMoveStatus.Status
            Switch($LiveMoveStatus)
            {
                {"InProgress" -or "Queued" -or "CompletionInProgress" -or "Retrying" -or "AutoSuspended"}
                {
                    $FirstRun = Get-Date $Json.FirstRun
                    If(($FirstRun - (Get-Date).Hours) -gt 24)
                    {
                        $Json.JsonMoveStatus = "Failed"
                        $Json.JsonError = "Failed to move in over 24hrs"
                    }
                }

                {"Completed" -or "CompletedWithWarning"}
                {
                    $Json.JsonMoveStatus = "Completed"

                    Update-Json -JsonObject $JsonObject -Path $JsonFullPath
                    Move-Item -Path $JsonFullPath -Destination $ScriptPaths["Completed"]
                    Exit
                }

                {"Failed" -or "Suspended"}
                {
                    $Json.JsonMoveStatus = "Failed"
                }
            }

            Update-Json -JsonObject $JsonObject -Path $JsonFullPath
        }

        Failed
        {
            $NewAssystIncident = @{
                Summary = "New Starter mailbox migration failed"
                Description = "The mail migration as part of reference $($Json.AssystRef) has failed. Please migrate the mailbox $($Json.UserPrincipalName) manually."
                Priority = "P3 - 16HR"
                Seriousness = "P3_INC_BU"
                AffectedUserShortCode = "SSCOPERATIONS"
                SystemServiceShortCode = "SERVER SERVICES"
                CategoryShortCode = "FAULT DATA ISS"
                AssignedResolverGroup = "SSC DC OPS"
                SendEmailIfApiFails = $True
                AssystEnvironment = $Json.AssystEnvironment
            }
    
            New-AssystIncident @NewAssystIncident
            Update-Json -JsonObject $JsonObject -Path $JsonFullPath
            Move-Item -Path $JsonFullPath -Destination $ScriptPaths["Failed"]
        }
    }
}