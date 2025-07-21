#SSCT:ActiveDirectory, SSCT:Exchange2016, SSCT:SSC-Assyst, #SSCT:SSC-Exchange
using namespace System.Collections.Generic

param(
    [Parameter(Mandatory = $True)][String]$TranscriptingGuid,
    [Parameter(Mandatory = $True)][String]$Environment,
    [Parameter(Mandatory = $True)][String]$JsonPath,
    [Parameter(Mandatory = $True)][Int]$IncidentId,
    [Parameter(Mandatory = $True)][String]$BaseConfigPath
)

$ErrorActionPreference = 'STOP'

Get-ChildItem "$BaseConfigPath\Scripts" | Unblock-File

$TranscriptPath = "$BaseConfigPath\Logs\$TranscriptingGuid.txt"

Start-Transcript $TranscriptPath

# Import Json
$GenericConfig = Get-Content "$BaseConfigPath\Config\Settings.json" | ConvertFrom-Json
# Set $TLS to higher versions for graph connection to work
$TLS = [System.Net.SecurityProtocolType] 'SSL3,TLS12'
[System.Net.ServicePointManager]::SecurityProtocol = $TLS

# Connect to Azure using Graph. Do not move this code. A portion of code lower in the order
# is affecting Graph's ability to connect consistently. Moving the connection here solves the
# issue.

Write-Information 'Connecting to Microsoft Graph.'

$CertificatePath = Join-Path -Path $GenericConfig.EmailAzAppReg.CertificatePath -ChildPath $GenericConfig.EmailAzAppReg.Thumbprint
$Certificate = Get-Item -Path $CertificatePath

$GraphConnectionParams = @{
    ClientId    = $GenericConfig.EmailAzAppReg.ClientId
    TenantId    = $GenericConfig.EmailAzAppReg.TenantId
    Certificate = $Certificate
}

try {
    Connect-MgGraph @GraphConnectionParams
    Write-Information 'Connected to Microsoft Graph.'
}
catch {
    Write-Warning 'Unable to call Connect-MgGraph. Credentials will need resetting and manually sending to recipient.'
}

#For testing via Centrify
Import-Module SSC-Assyst -Force
#Import-Module "$BaseConfigPath\TestFiles\SSC-ActiveDirectory.1.7.350\SSC-ActiveDirectory.psm1" -Force
Import-Module SSC-ActiveDirectory -Force
Import-Module SSC-Exchange -Force

# Set Assyst API Credentials
switch ($Environment) {
    'DEV' {
        $APICred = @{
            APIUser                = '16120200-ASSYSTREST'
            PathToSecureStringFile = 'D:\AssystAPISec\PREPROD-APISec'
            PathTokeyFile          = 'D:\AssystAPISec\PREPROD-APISec-Key'
        }
    }
    'LIVE' {
        $APICred = @{
            APIUser                = '161202-ASSYSTREST'
            PathToSecureStringFile = 'D:\AssystAPISec\PROD-APISec'
            PathTokeyFile          = 'D:\AssystAPISec\PROD-APISec-Key'
        }
    }
}

try {
    Set-AssystAPICredential @APICred -AssystEnvironment $Environment
}
catch {
    Write-Warning 'Unable to set Assyst credentials, exiting.'
    Write-Warning "$($_.Exception.Message)"
    Exit
}

#Class definitions
Class FunctionResults {
    [String]$FunctionName
    [Bool]$Successful
    [String]$FunctionError
}

Class CoreData {
    $FunctionResults = @()
    $Attributes = @{}

    [void]SetSuccessful($FunctionName) {
        $i = 0
        while ($i -lt $This.FunctionResults.count) {
            if ($This.FunctionResults[$i].FunctionName -like $FunctionName) {
                $This.FunctionResults[$i].Successful = $True
                $i = $This.FunctionResults.count
            }
            $i++
        }
    }

    [void]SetFunctionError($FunctionName, $FunctionError) {
        $i = 0
        while ($i -lt $This.FunctionResults.count) {
            if ($This.FunctionResults[$i].FunctionName -like $FunctionName) {
                $This.FunctionResults[$i].FunctionError = $FunctionError
                $i = $This.FunctionResults.count
            }
            $i++
        }
    }

    [FunctionResults]GetFunctionResult($Arg) {

        [FunctionResults]$Output = $This.FunctionResults | Where-Object { $_.FunctionName -like $Arg }

        return $Output
    }
}

Write-Information 'Creating CoreData Object'
$CoreData = New-Object CoreData

$InvokeScriptFailureParam = @{
    ResolverGroup         = $GenericConfig.ResolverGroups.AutomationTeam
    FallbackEmail         = @($GenericConfig.Generic.FallbackEmail)
    Summary               = $GenericConfig.Generic.Summary
    AffectedUserShortCode = $GenericConfig.Generic.AffectedUserShortCode
    AssystEnvironment     = $Environment
    SmtpServer            = $GenericConfig.Generic.SmtpServer
    BaseConfigPath        = $BaseConfigPath
    JsonPath              = $JsonPath
    AssystRef             = $IncidentId
    CoreDataObject        = $CoreData
    TranscriptingGuid     = $TranscriptingGuid
    AutomationTaskList    = $AutomationTaskList
    BusinessUnitData      = $BusinessUnitData
    AzAppEmailAddress     = $GenericConfig.EmailAzAppReg.AppEmailAddress
    Exit                  = $True
}

$CodeFiles = @(Get-ChildItem -Path "$BaseConfigPath\Scripts\Functions" -ErrorAction SilentlyContinue)
foreach ($import in $CodeFiles) {
    $FunctionObject = New-Object FunctionResults

    try {
        $ImportFullName = $Import.FullName
        Write-Information "Importing $($ImportFullName)"
        . $ImportFullName
        $FunctionObject.FunctionName = $Import.BaseName
    }
    catch {
        Write-Error "Failed to import function $($import.FullName): $_"

        $InvokeScriptFailureParam.Summary = 'Failed to import functions'
        $InvokeScriptFailureParam.CoreDataObject = $CoreData
        Invoke-ScriptFailure @InvokeScriptFailureParam
    }

    $CoreData.FunctionResults += $FunctionObject
}

Write-Information 'All functions imported successfully'

# Import the CSV or JSON
try {
    $CoreData.Attributes = Import-AssystUserJson -JsonPath $JsonPath
    $CoreData.SetSuccessful('Import-AssystUserJson')
    Write-Information 'Imported Assyst User Json configuration'
}
catch {
    $CoreData.SetFunctionError('Import-AssystUserJson', $_.Exception.Message)
    Write-Error ('Failed to import JSON/CSV ERROR:{0}' -f $_.Exception.Message)
    $InvokeScriptFailureParam.Summary = 'Failed to import JSON'
    $InvokeScriptFailureParam.CoreDataObject = $CoreData
    Invoke-ScriptFailure @InvokeScriptFailureParam
}

# An Assyst limitation means AccountExpirationDate may come through with a numerical value on the end.
# Code fixes the $CoreData.Attributes with a valid AD key for AccountExpirationDate.
$NumericalAccountExpirationDate = $false

foreach ($Key in $CoreData.Attributes.Keys) {
    if ($Key -match '^AccountExpirationDate\d+$') {
        $NumericalAccountExpirationDate = $true
        $AccountExpirationDateWithNumericValueKey = $Key
        Write-Information "$Key found. Numeric value on end means AD attribute will not be set unless fixed."
        Write-Information 'Setting NumericalAccountExpirationDate to true.'
    }
}

if ($NumericalAccountExpirationDate -and $null -ne $AccountExpirationDateWithNumericValueKey) {
    $CorrectKey = 'AccountExpirationDate'
    Write-Information 'Creating AccountExpirationDate key and settings its value.'
    $CoreData.Attributes[$CorrectKey] = $CoreData.Attributes[$AccountExpirationDateWithNumericValueKey]
    Write-Information "Removing $AccountExpirationDateWithNumericValueKey key."
    $CoreData.Attributes.Remove($AccountExpirationDateWithNumericValueKey)
}

$Results = Test-RequiredAttributes -Data $CoreData.Attributes

# Set AccountExpirationDate to be a date
if ($CoreData.Attributes['AccountExpirationDate']) {
    Write-Information 'Setting AccountExpirationDate to a date type.'
    try {
        $CoreData.Attributes['AccountExpirationDate'] = Get-Date $CoreData.Attributes['AccountExpirationDate']
        Write-Information 'AccountExpirationDate successfully set to a date type.'
    }
    catch {
        Write-Warning 'Unable to convert AccountExpirationDate to datetime. Account will not have its Expiry Date set.'
    }
}
else {
    Write-Information 'CoreData Attributes does not contain AccountExpirationDate, account will not have its Expiry Date set.'
}

# Set StartDate to be a date
if ($CoreData.Attributes['StartDate']) {
    Write-Information 'Setting StartDate to a date type.'
    try {
        $CoreData.Attributes['StartDate'] = Get-Date $CoreData.Attributes['StartDate']
        Write-Information 'StartDate successfully set to a date type.'
    }
    catch {
        Write-Warning 'Unable to convert StartDate to datetime.'
    }
}
else {
    Write-Information 'CoreData Attributes does not contain StartDate.'
}

# Test if the AccountExpirationDate is in the past (Assyst cannot validate on the front end)
$RemoveAccountExpirationDateFromCoreData = $false
if (!($CoreData.Attributes.AccountExpirationDate)) {
    Write-Information 'AccountExpirationDate attribute does not exist in CoreData Attributes. Will not compare against current date.'
}
elseif ($CoreData.Attributes.AccountExpirationDate -gt (Get-Date)) {
    Write-Information "AccountExpirationDate is in the future: $($CoreData.Attributes.AccountExpirationDate) Validation Complete."
}
else {
    Write-Information "AccountExpirationDate is not set in the future: $($CoreData.Attributes.AccountExpirationDate)"
    $RemoveAccountExpirationDateFromCoreData = $true
    Write-Information 'Raising new incident to IAM'
    $JoinerIncidentParam = @{
        AssystReference = $IncidentId
        Configuration   = $GenericConfig
        Environment     = $Environment
        IncidentReason  = 'AccountExpirationDate'
    }
    New-JoinerIncident @JoinerIncidentParam -Verbose
}

# Test if the Start Date is earlier than the AccountExpirationDate (Assyst cannot validate on the front end)
if (!($CoreData.Attributes.AccountExpirationDate)) {
    Write-Information "AccountExpirationDate attribute does not exist in CoreData Attributes. Will not compare against user's start date."
}
elseif (!($CoreData.Attributes.StartDate) -or $CoreData.Attributes.StartDate -isnot [System.DateTime]) {
    Write-Information "StartDate attribute either does not exist in CoreData Attributes or is not a DateTime object. Will not compare against user's AccountExpirationDate."
}
elseif ($CoreData.Attributes.AccountExpirationDate -gt $CoreData.Attributes.StartDate) {
    Write-Information 'AccountExpirationDate is further in future than the StartDate. Validation Complete.'
}
else {
    Write-Information "The StartDate $($CoreData.Attributes.StartDate) is further in the future than the AccountExpirationDate $($CoreData.Attributes.AccountExpirationDate)"
    $RemoveAccountExpirationDateFromCoreData = $true
    Write-Information 'Raising new incident to IAM'
    $JoinerIncidentParam = @{
        AssystReference = $IncidentId
        Configuration   = $GenericConfig
        Environment     = $Environment
        IncidentReason  = 'AccountExpirationAndStartDate'
    }
    New-JoinerIncident @JoinerIncidentParam -Verbose
}

if ($RemoveAccountExpirationDateFromCoreData) {
    Write-Information 'Removing AccountExpirationDate from CoreData, account will not have its Expiry Date set.'
    $CoreData.Attributes.Remove('AccountExpirationDate')
}

# Trim leading and trailing spaces from GivenName and Surname
$CoreData.Attributes['GivenName'] = $CoreData.Attributes['GivenName'].Trim()
$CoreData.Attributes['Surname'] = $CoreData.Attributes['Surname'].Trim()

# Assyst UPN output will contain AD_. Clean data then get copy of manager's email address before it is changed to a DN
if ($CoreData.Attributes.Keys -contains 'Manager' -and $CoreData.Attributes.Manager -like 'AD_*') {

    Write-Information 'Manager field is like AD_*. Removing AD_ prefix from the CoreData.Attributes.Manager field.'
    $CoreData.Attributes.Manager = $CoreData.Attributes.Manager -replace '^AD_'
    Write-Information "CoreData.Attributes.Manager has been updated to $($CoreData.Attributes.Manager)."

    $ManagerEmailAddress = $CoreData.Attributes['Manager'].Trim()
}
elseif ($CoreData.Attributes.Keys -contains 'Manager' -and $CoreData.Attributes.Manager -notlike 'AD_*') {
    $ManagerEmailAddress = $CoreData.Attributes['Manager'].Trim()
}

if ([string]::IsNullOrEmpty($ManagerEmailAddress) -or [string]::IsNullOrWhiteSpace($ManagerEmailAddress)) {
    $LineManagerProvided = $false
    Write-Information 'No Line Manager provided in the JSON file.'
}
else {
    $LineManagerProvided = $true
    Write-Information 'Line Manager has been provided in the JSON file.'
}

# Get a copy of Department before it is changed to its friendly name.
$DepartmentKey = $CoreData.Attributes['Department']

# Check if $CoreData.Attributes contains key SscAutoEnable from Assyst
# if no add key SscAutoEnable to $CoreData.Attributes and set to No
if (!($CoreData.Attributes.Contains('SscAutoEnable'))) {
    $CoreData.Attributes.Add('SscAutoEnable', 'No')
}

# Handle special characters and diacritics
# Store original names before any special characters or diacritics are potentially removed
if ($CoreData.Attributes['Office'] -eq 'ABEZ_Darmstadt-' -or
    $CoreData.Attributes['Office'] -eq 'ABEZ_Rajamaki-') {
    $CoreData.Attributes['OriginalGivenName'] = $CoreData.Attributes['GivenName']
    $CoreData.Attributes['OriginalSurname'] = $CoreData.Attributes['Surname']
}

switch ($CoreData.Attributes['Office']) {
    'ABEZ_Darmstadt-' {
        $CoreData.Attributes['GivenName'] = $CoreData.Attributes['GivenName'] -creplace 'ö', 'oe' -creplace 'Ö', 'Oe' -creplace 'ä', 'ae' -creplace 'Ä', 'Ae' -creplace 'ü', 'ue' -creplace 'Ü', 'Ue' -creplace 'ß', 'ss'
        $CoreData.Attributes['Surname'] = $CoreData.Attributes['Surname'] -creplace 'ö', 'oe' -creplace 'Ö', 'Oe' -creplace 'ä', 'ae' -creplace 'Ä', 'Ae' -creplace 'ü', 'ue' -creplace 'Ü', 'Ue' -creplace 'ß', 'ss'
    }
    'ABEZ_Rajamaki-' {
        $CoreData.Attributes['GivenName'] = $CoreData.Attributes['GivenName'] -creplace 'Ä', 'A' -creplace 'ä', 'a' -creplace 'Ö', 'O' -creplace 'ö', 'o' -creplace 'Å', 'A' -creplace 'å', 'a'
        $CoreData.Attributes['Surname'] = $CoreData.Attributes['Surname'] -creplace 'Ä', 'A' -creplace 'ä', 'a' -creplace 'Ö', 'O' -creplace 'ö', 'o' -creplace 'Å', 'A' -creplace 'å', 'a'
    }
}

# Update Business Unit
$CoreData.Attributes['BusinessUnit'] = $CoreData.Attributes['BusinessUnit'].replace('_CSG', '')

# Using the updated Business Unit value, get Business Unit data from BusinessUnitData JSON
$BusinessUnitShortCode = $CoreData.Attributes['BusinessUnit']
$JsonBuConfigurationPath = "D:\Automation\PowerShellModuleConfigurations\SSC-ActiveDirectory\$BusinessUnitShortCode.json"
$BusinessUnitData = Get-Content $JsonBuConfigurationPath | ConvertFrom-Json
# Replace $CoreData.Attributes["SscAutoEnable"] value to match value in $BusinessUnitData "AutoEnable" key
if ($CoreData.Attributes['SscAutoEnable'] -notlike 'Yes*') {
    $CoreData.Attributes['SscAutoEnable'] = $BusinessUnitData.Global.AutoEnable
}

# Trim Remote_ for remote users so we can find the correct site for their other settings, and store the original
# value to update the AD value later.

if ($CoreData.Attributes['Office'] -match '^Remote_') {
    $CoreData.Attributes.Add('RemoteOffice', $CoreData.Attributes['Office'])
    $CoreData.Attributes['Office'] = $CoreData.Attributes['Office'] -replace '^Remote_'
}

# BU's in this array have mail enabled by default, their LMT form does not output MailboxProvisioning
# so we add it to trigger mailbox creation.

$EnableMailByDefault = @(
    'ABA'
    'ABEZ'
    'ABI'
    'ABM_NA'
    'ABN'
    'ABT'
    'ALM'
    'CFM'
    'FYT'
    'GRO'
    'GSC'
    'HNH'
    'JRY'
    'MPL'
    'OHLY_EU'
    'SBA'
    'SIL'
    'TWO_AU'
    'TWO_CH'
    'TWO_CN'
    'TWO_FI'
    'TWO_IN'
    'TWO_NG'
    'TWO_PL'
    'TWO_TH'
    'TWO_UK'
    'TWO_US'
)

if ($CoreData.Attributes['BusinessUnit'] -in $EnableMailByDefault) {
    try {
        $CoreData.Attributes['MailboxProvisioning'] = 'Y'
        Write-Information "This BU has mail enabled by default, successfully set MailboxProvisioning requirement to 'Y'."
    }
    catch {
        Write-Information 'Failed to set MailboxProvisioning for this user.'
    }
}

# ABE's form does not have the Internet Access field which is required for all ABE users.

if ($CoreData.Attributes['BusinessUnit'] -eq 'ABEZ') {
    try {
        $CoreData.Attributes.Add('InternetAccess', 'YES')
        Write-Information "AB Enzymes user: successfully set InternetAccess requirement to 'YES'."
    }
    catch {
        Write-Information 'Failed to set default mailbox and/or Internet access requirement for AB Enzymes user.'
    }
}

try {
    $CoreData.Attributes.Add('AssystRef', $IncidentId)
}
catch {
    Write-Information 'Failed to add AssystRef to Core Data'
}

# Build Task List for Tracking and Reporting

$AutomationTasks = @(
    'Generate Username'
    'Duplicate Account Check'
    'Create User Account'
    'Set MFA Group'
    'Set Account Password'
    'Disable User Must Change Password At Next Logon'
    'Enable User Account'
    'Provision Mailbox'
    'Set UserPrincipalName'
    'Update extensionAttribute9'
    'Provision Licence'
    'E-mail Account Details'
    'Close Task Ticket'
)

$AutomationTaskList = foreach ($Task in $AutomationTasks) {
    [PSCustomObject] @{
        TaskName = $Task
        Status   = 'Not Started'
        Result   = 'N/A'
        Detail   = 'N/A'
    }
}

$UpdateHtmlData = @{}
if ($Results -eq $True) {
    $CoreData.SetSuccessful('Test-RequiredAttributes')

    Write-Information "Getting linked tasks for incident ref#: $IncidentID."
    $LinkedTaskList = (Get-AssystLinkedEventGroup -IncidentId $IncidentId -AssystEnvironment $Environment).LinkedEvents.formattedReference | Where-Object { $_ -notlike $IncidentId }

    Write-Information 'Processing linked tasks.'
    foreach ($TaskReference in $LinkedTaskList) {
        $TaskReference = $TaskReference -replace '[a-z]'
        $TaskDetails = Get-AssystIncident -IncidentId $TaskReference -AssystEnvironment $Environment
        if (($TaskDetails.Summary -like 'New Starter - AD Account and Mailbox Access*') -and ($TaskDetails.Status -eq 'OPEN')) {
            $AutomationTask = $TaskDetails.IncidentRef -replace '[a-z]'
            break
        }
    }
    if ($null -ne $AutomationTask) {
        $LinkedTaskReference = $AutomationTask
        Write-Information "Automation task incident ref#: $LinkedTaskReference."
    }
    else {
        Write-Information 'Unable to pull automation task from linked tasks.  Exiting script.'
        Write-Information "Linked task count is: $($LinkedTaskList.count)"
        Write-Information ($LinkedTaskList | Out-String)
        $InvokeScriptFailureParam.Summary = 'Unable to pull automation task from linked tasks.'
        $InvokeScriptFailureParam.CoreDataObject = $CoreData
        $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
        $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
        Invoke-ScriptFailure @InvokeScriptFailureParam
    }

    Write-Information "Checking the requester's name does not match the joiner's name."
    $Incident = Get-AssystIncident -IncidentId $IncidentId -AssystEnvironment $Environment
    $RequesterName = $Incident.ResponseXml.event.affectedUserName
    $JoinerName = "$($CoreData.Attributes['GivenName']) $($CoreData.Attributes['Surname'])"
    Write-Information "Requester: $RequesterName, Joiner: $JoinerName."
    if ($RequesterName -eq $JoinerName) {
        Write-Information 'The requester and joiner names are the same.  Exiting script.'
        $AnalysisInformation = @{
            IncidentId        = $LinkedTaskReference
            Description       = 'The requester and joiner names are the same.'
            AssystEnvironment = $Environment
            ErrorAction       = 'CONTINUE'
        }
        Add-AssystAnalysisInformation @AnalysisInformation
        $AssignmentInformation = @{
            IncidentId                = $LinkedTaskReference
            AssignedServiceDepartment = $GenericConfig.ResolverGroups.IAM
            AssystEnvironment         = $Environment
            AssigningNotes            = 'The requester and joiner names are the same. To avoid creating a duplicate account, the automation was stopped.  Please review the request.'
            ErrorAction               = 'CONTINUE'
        }
        Set-AssystIncidentAssignee @AssignmentInformation
        $InvokeScriptFailureParam.Summary = 'The requester and joiner names are the same.'
        $InvokeScriptFailureParam.CoreDataObject = $CoreData
        $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
        $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
        Invoke-ScriptFailure @InvokeScriptFailureParam
    }
    else {
        Write-Information 'The requester and joiner names are different.  The automation will continue.'
    }

    try {
        $NewUserNameParam = @{
            GivenName = $CoreData.Attributes['GivenName']
            Surname   = $CoreData.Attributes['Surname']
        }

        # Using a switch statement, we can handle other BU requirements and account types in future.

        switch ($BusinessUnitShortCode) {
            'TWO_UK_CAP' {
                $NewUserNameParam.Add('AccountType', 'TPA')
            }
        }

        $SamAccountName = New-UserName @NewUserNameParam

        $UpdateTaskListParams = @{
            TaskName = 'Generate Username'
            Status   = 'Completed'
            Result   = 'Success'
            Detail   = "SAMAccountName is $SamAccountName."
        }
        Update-AutomationTaskList @UpdateTaskListParams
        # Stop the automation if a duplicate account is suspected.  Duplicates of 'NStarterAuto' are permitted for testing.
        if ($SamAccountName -notlike 'NStarterAuto*') {
            Write-Information 'Checking Active Directory for existing users.'
            $TestSamAccountName = ($SamAccountName -replace '\d+$', '') + '*'
            $DuplicateCandidates = Get-ADUser -Filter { SamAccountName -like $TestSamAccountName } |
            Select-Object GivenName, Surname, SamAccountName, DistinguishedName
            $DuplicateAccounts = [List[Object]]::New()
            foreach ($Candidate in $DuplicateCandidates) {
                if ($Candidate.GivenName -eq $CoreData.Attributes['GivenName']) {
                    $DuplicateAccounts.Add($Candidate)
                }
            }

            if ($DuplicateAccounts.Count -gt 0) {
                $UpdateTaskListParams = @{
                    TaskName = 'Duplicate Account Check'
                    Status   = 'Completed'
                    Result   = 'Failure'
                    Detail   = 'Duplicate account suspected.'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                Write-Information 'This user may already exist in Active Directory. Existing users:'
                Write-Output ($DuplicateAccounts | Out-String -Width 256)
                Write-Information 'The ticket will be assigned to IAM for review. Exiting script.'
                $AnalysisInformation = @{
                    IncidentId        = $LinkedTaskReference
                    Description       = 'The user may already exist in Active Directory.'
                    AssystEnvironment = $Environment
                    ErrorAction       = 'CONTINUE'
                }
                Add-AssystAnalysisInformation @AnalysisInformation
                $AssigningNotes = @"
The user may already exist in Active Directory.
To avoid creating a duplicate account, the automation was stopped.  Please review the request.
Existing users:
$($DuplicateAccounts | Out-String -Width 256)
The automation failed at the 'Duplicate Account Check' stage.  Please review and complete the outstanding tasks:
$($AutomationTaskList | Format-List | Out-String -Width 256)
"@
                Write-Information $AssigningNotes
                $AssignmentInformation = @{
                    IncidentId                = $LinkedTaskReference
                    AssignedServiceDepartment = $GenericConfig.ResolverGroups.IAM
                    AssystEnvironment         = $Environment
                    AssigningNotes            = $AssigningNotes
                    ErrorAction               = 'CONTINUE'
                }
                Set-AssystIncidentAssignee @AssignmentInformation
                $InvokeScriptFailureParam.Summary = 'The user may already exist in Active Directory.'
                $InvokeScriptFailureParam.CoreDataObject = $CoreData
                $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
                $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
                Invoke-ScriptFailure @InvokeScriptFailureParam
            }
            else {
                $UpdateTaskListParams = @{
                    TaskName = 'Duplicate Account Check'
                    Status   = 'Completed'
                    Result   = 'Success'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                Write-Information 'No potential duplicates found; the account creation process will continue.'
                Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description "Username has been set to: $SamAccountName." -AssystEnvironment $Environment
                $CoreData.SetSuccessful('New-UserName')
                Write-Information "SamAccountName:$SamAccountName"
            }
        }
        else {
            $UpdateTaskListParams = @{
                TaskName = 'Duplicate Account Check'
                Status   = 'Skipped'
                Detail   = 'Test Account.'
            }
            Update-AutomationTaskList @UpdateTaskListParams
        }
    }
    catch {
        $UpdateTaskListParams = @{
            TaskName = 'Generate Username'
            Status   = 'Completed'
            Result   = 'Failure'
            Detail   = 'An error occured while generating the user name.'
        }
        Update-AutomationTaskList @UpdateTaskListParams
        $CoreData.SetFunctionError('New-Username', $_.Exception.Message)
        Write-Information ('[Username] Failed to generate new username: {0}' -f $_.Exception.Message)
        $InvokeScriptFailureParam.Summary = 'Failed to generate new username'
        $InvokeScriptFailureParam.CoreDataObject = $CoreData
        $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
        $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
        Invoke-ScriptFailure @InvokeScriptFailureParam
    }

    if ($Null -ne $SamAccountName) {
        $CoreData.Attributes.Add('SamAccountName', $SamAccountName)

        try {
            Write-Information '[Account Creation] Starting'
            $NewUser = New-SscAdUserFromHashTable -AttributeHashTable $CoreData.Attributes -BusinessUnitShortCode $CoreData.Attributes['BusinessUnit'] -Verbose
            $UpdateTaskListParams = @{
                TaskName = 'Create User Account'
                Status   = 'Completed'
                Result   = 'Success'
            }
            Update-AutomationTaskList @UpdateTaskListParams
            Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description "User account has been created with the username: $SamAccountName." -AssystEnvironment $Environment
            $CoreData.SetSuccessful('New-SscAdUserFromHashTable')
            Write-Information '[Account Creation] User account created successfully'
            $UpdateHtmlData += $NewUser
            $UpdateHtmlData.Add('Account creation', 'Success')
            # Ensure we use the same DC for updates as we used to create the account.
            $Server = $NewUser.Server
        }
        catch {
            $UpdateTaskListParams = @{
                TaskName = 'Create User Account'
                Status   = 'Completed'
                Result   = 'Failure'
                Detail   = 'An error occured while creating the AD account.'
            }
            Update-AutomationTaskList @UpdateTaskListParams
            Write-Warning -Message $_.Exception.Message
            $UpdateHtmlData.Add('Account creation', 'FAILED!')
            Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description "Account creation failed for: $SamAccountName." -AssystEnvironment $Environment -ErrorAction Continue
            $AssigningNotes = @"
The automation failed at the 'Create User Account' stage.  Please review and complete the outstanding tasks:
$($AutomationTaskList | Format-List | Out-String -Width 256)
"@
            $AssignmentInformation = @{
                IncidentId                = $LinkedTaskReference
                AssignedServiceDepartment = $GenericConfig.ResolverGroups.IAM
                AssystEnvironment         = $Environment
                AssigningNotes            = $AssigningNotes
                ErrorAction               = 'CONTINUE'
            }
            Set-AssystIncidentAssignee @AssignmentInformation
            $InvokeScriptFailureParam.Summary = 'Account creation failed'
            $InvokeScriptFailureParam.CoreDataObject = $CoreData
            $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
            $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
            Invoke-ScriptFailure @InvokeScriptFailureParam
        }

        if ($NewUser['Path'] -eq (Get-ADDomain).UsersContainer) {
            $UpdateTaskListParams = @{
                TaskName = 'Create User Account'
                Detail   = 'The account was created in the default Users container.'
            }
            Update-AutomationTaskList @UpdateTaskListParams
            $ErrorMessage = "The user account $SamAccountName was created in the default Users container: $($NewUser['Path'])"
            Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description $ErrorMessage -AssystEnvironment $Environment -ErrorAction Continue
            $InvokeScriptFailureParam.Summary = $ErrorMessage
            $InvokeScriptFailureParam.CoreDataObject = $CoreData
            $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
            $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
            Set-AssystIncidentAssignee -IncidentId $LinkedTaskReference -AssignedServiceDepartment $GenericConfig.ResolverGroups.IAM -AssystEnvironment $Environment -AssigningNotes "Please move the user account $SamAccountName to the correct OU." -ErrorAction Continue
            $InvokeScriptFailureParam.Exit = $false
            Invoke-ScriptFailure @InvokeScriptFailureParam
        }

        # Set MFA Allowed or Block External Access Group (Default is Block External Access except for SSC and TwO)

        $MFAAllowByDefault = @(
            'ABI'
            'ABF_INTRA'
            'ACH_MX'
            'ACH_US'
            'HNH'
            'FYT'
            'OHLY_EU'
            'OHLY_US'
            'SSC'
            'TWO_AU'
            'TWO_BE'
            'TWO_CH'
            'TWO_CN'
            'TWO_FI'
            'TWO_IN'
            'TWO_NG'
            'TWO_PL'
            'TWO_TH'
            'TWO_UK'
            'TWO_US'
        )

        $MFAParams = @{
            MFAAllowed            = $false
            BusinessUnitShortCode = $CoreData.Attributes['BusinessUnit']
            sAMAccountName        = $SamAccountName
            Verbose               = $true
        }

        if (($CoreData.Attributes['BusinessUnit'] -in @('ABA', 'ABN', 'ALM', 'GSC', 'GRO', 'SBA')) -and
            ($CoreData.Attributes['D365'] -eq 'y') -or
            ($CoreData.Attributes['MobileRequired'] -eq 'SMART PHONE') -or
            ($CoreData.Attributes['WorkerProfile'] -like 'REMOTE*')) {
            $MFAAllowByDefault += $CoreData.Attributes['BusinessUnit']
        }

        if (($CoreData.Attributes['MFARequired'] -like 'Yes*') -or
            ($CoreData.Attributes['BusinessUnit'] -in $MFAAllowByDefault) -or
            ($CoreData.Attributes['BusinessUnit'] -like 'ABEZ*') -and ($CoreData.Attributes['Department'] -ne 'Finland_Operator')) {
            $MFAParams.MFAAllowed = $true
        }

        if ($Null -ne $NewUser) {
            try {
                Write-Information 'Setting MFA access.'
                Set-MFAGroup @MFAParams
                $UpdateTaskListParams = @{
                    TaskName = 'Set MFA Group'
                    Status   = 'Completed'
                    Result   = 'Success'
                }
                Update-AutomationTaskList @UpdateTaskListParams
            }
            catch {
                $UpdateTaskListParams = @{
                    TaskName = 'Set MFA Group'
                    Status   = 'Completed'
                    Result   = 'Failure'
                    Detail   = 'An error occured while setting the MFA Group.'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                $ErrorMessage = "Failed to set an MFA Allowed or Block External Access group for $sAMAccountName."
                Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description $ErrorMessage -AssystEnvironment $Environment -ErrorAction Continue
                $InvokeScriptFailureParam.Summary = $ErrorMessage
                $InvokeScriptFailureParam.CoreDataObject = $CoreData
                $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
                $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
                Set-AssystIncidentAssignee -IncidentId $LinkedTaskReference -AssignedServiceDepartment $GenericConfig.ResolverGroups.IAM -AssystEnvironment $Environment -AssigningNotes "Please check and assign the user's MFA group." -ErrorAction Continue
                $InvokeScriptFailureParam.Exit = $false
                Invoke-ScriptFailure @InvokeScriptFailureParam
            }

        }

        if ($CoreData.Attributes['SscAutoEnable'] -like 'Yes*' -and ($Null -ne $NewUser)) {
            Write-Information 'SscAutoEnable value is Yes.  The account will be enabled.'
            try {
                $ADAccountPasswordData = Set-SSCADAccountPassword -SamAccountName $SamAccountName -Simple
                if ($ADAccountPasswordData.PasswordReset) {
                    if ($CoreData.Attributes['BusinessUnit'] -in @('ABF_INTRA', 'TWO_UK_CAP')) {
                        Write-Information "$($CoreData.Attributes['BusinessUnit']) account, disabling 'User must change password at next logon'."
                        try {
                            Set-ADUser -Identity $SamAccountName -ChangePasswordAtLogon:$False
                            Write-Information "'User must change password at next logon' was disabled."
                            $UpdateTaskListParams = @{
                                TaskName = 'Disable User Must Change Password At Next Logon'
                                Status   = 'Completed'
                                Result   = 'Success'
                            }
                            Update-AutomationTaskList @UpdateTaskListParams
                        }
                        catch {
                            Write-Warning "Unable to disable 'User must change password at next logon.'"
                            $ErrorMessage = "Unable to disable 'User must change password at next logon' for: $sAMAccountName."
                            Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description $ErrorMessage -AssystEnvironment $Environment -ErrorAction Continue
                            $UpdateTaskListParams = @{
                                TaskName = 'Disable User Must Change Password At Next Logon'
                                Status   = 'Completed'
                                Result   = 'Failed'
                            }
                            Update-AutomationTaskList @UpdateTaskListParams
                        }
                    }
                    else {
                        $UpdateTaskListParams = @{
                            TaskName = 'Disable User Must Change Password At Next Logon'
                            Status   = 'Skipped'
                            Detail   = 'Implemented only for ABF Intranet Only & TPA Accounts.'
                        }
                        Update-AutomationTaskList @UpdateTaskListParams
                    }
                    $UpdateTaskListParams = @{
                        TaskName = 'Set Account Password'
                        Status   = 'Completed'
                        Result   = 'Success'
                    }
                    Update-AutomationTaskList @UpdateTaskListParams
                    $UpdateHtmlData.Add('Password Reset', 'Success')
                    Enable-ADAccount -Identity $SamAccountName -ErrorAction Stop
                    $UpdateTaskListParams = @{
                        TaskName = 'Enable User Account'
                        Status   = 'Completed'
                        Result   = 'Success'
                    }
                    Update-AutomationTaskList @UpdateTaskListParams
                    $UpdateHtmlData.Add('Enable Account', 'Success')
                }
                else {
                    throw 'Failed to Set Password.'
                }
            }
            catch {
                if ($_.Exception.Message -eq 'Failed to Set Password.') {
                    $UpdateTaskListParams = @{
                        TaskName = 'Set Account Password'
                        Status   = 'Completed'
                        Result   = 'Failed'
                    }
                    Update-AutomationTaskList @UpdateTaskListParams
                    $ErrorMessage = "The password could not be set on account: $SamAccountName"
                    $UpdateHtmlData.Add('Password Reset', 'FAILED!')
                    $UpdateHtmlData.Add('Enable Account', 'Skipped')
                }
                else {
                    $UpdateTaskListParams = @{
                        TaskName = 'Enable User Account'
                        Status   = 'Skipped'
                    }
                    Update-AutomationTaskList @UpdateTaskListParams
                    $ErrorMessage = "Unable to enable account: $SamAccountName"
                    $UpdateHtmlData.Add('Enable Account', 'FAILED!')
                }
                Write-Information $ErrorMessage
                Write-Information "Error: $($_.Exception.Message)"
                Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description $ErrorMessage -AssystEnvironment $Environment -ErrorAction Continue
                $InvokeScriptFailureParam.Summary = $ErrorMessage
                $InvokeScriptFailureParam.CoreDataObject = $CoreData
                $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
                $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
                Set-AssystIncidentAssignee -IncidentId $LinkedTaskReference -AssignedServiceDepartment $GenericConfig.ResolverGroups.IAM -AssystEnvironment $Environment -AssigningNotes 'Automation failed to reset password or enable the account. Please review and complete.' -ErrorAction Continue
                $InvokeScriptFailureParam.Exit = $false
                Invoke-ScriptFailure @InvokeScriptFailureParam
            }
        }
        elseif ($CoreData.Attributes['SscAutoEnable'] -notlike 'Yes*') {
            Write-Information "SscAutoEnable attribute not found or has value 'No'. $SamAccountName account will need to have the password set and be enabled manually."
            $UpdateTaskListParams = @{
                TaskName = 'Set Account Password'
                Status   = 'Skipped'
                Detail   = 'Auto-enable was not requested.'
            }
            Update-AutomationTaskList @UpdateTaskListParams
            $UpdateTaskListParams = @{
                TaskName = 'Enable User Account'
                Status   = 'Skipped'
                Detail   = 'Auto-enable was not requested.'
            }
            Update-AutomationTaskList @UpdateTaskListParams

            $JoinerIncidentParam = @{
                AssystReference = $IncidentId
                Configuration   = $GenericConfig
                Environment     = $Environment
                IncidentReason  = 'AutoEnable'
            }

            New-JoinerIncident @JoinerIncidentParam -Verbose
        }

        # Update Office value for Germains - this ensures a consistent value in AD, when the same office
        # has multiple locations in the Sites section of the configuration file.

        if ($CoreData.Attributes['BusinessUnit'] -eq 'GER') {

            $GermainsOffices = @{
                'GER_AALTEN GENERAL-'                  = 'GER_AALTEN-'
                'GER_AALTEN OPSTEMP-'                  = 'GER_AALTEN-'
                'GER_AALTEN QATEMP-'                   = 'GER_AALTEN-'
                'GER_AALTEN SUPCTEMP-'                 = 'GER_AALTEN-'
                'GER_CASTELOLLI GENERAL-'              = 'GER_CASTELOLLI-'
                'GER_ENKHUIZEN GENERAL-'               = 'GER_ENKHUIZEN-'
                'GER_FARGO GENERAL-'                   = 'GER_FARGO-'
                'GER_GILROY GENERAL-'                  = 'GER_GILROY-'
                'GER_KINGSLYNN F1 GENERAL-'            = 'GER_KINGSLYNN-'
                'GER_KINGSLYNN F1 SEASONAL-'           = 'GER_KINGSLYNN-'
                'GER_KINGSLYNN F2 GENERAL-'            = 'GER_KINGSLYNN-'
                'GER_KINGSLYNN F2 SEASONAL-'           = 'GER_KINGSLYNN-'
                'GER_KINGSLYNN SAINT ANDREWS GENERAL-' = 'GER_KINGSLYNN-'
            }

            try {
                Write-Information 'Germains BU detected, updating Office attribute.'
                Set-ADUser -Identity $SamAccountName -Office $GermainsOffices.$($CoreData.Attributes['Office'])
            }
            catch {
                Write-Warning "Failed to update user's Office attribute."
                Write-Warning "$($_.Exception.Message)"
            }
        }

        # Update remote Office value for AB Agri - this is to ensure the Office locations in AD match the office location
        # in Aspire.

        if ($CoreData.Attributes['BusinessUnit'] -eq 'ABG' -and $CoreData.Attributes['RemoteOffice']) {

            $RemoteOffice = switch -Regex ($CoreData.Attributes['RemoteOffice']) {
                '_UK$' {
                    'ABG_Remote/Home_UK'
                }
                default {
                    'ABG_Remote/Home_International'
                }
            }

            try {
                Write-Information 'Updating Office for AB Agri remote user.'
                Set-ADUser -Identity $SamAccountName -Office $RemoteOffice
            }
            catch {
                Write-Warning "Failed to update user's Office attribute."
                Write-Warning "$($_.Exception.Message)"
            }
        }

        if ($CoreData.Attributes['MailboxProvisioning'] -eq 'Y' -and $Null -ne $NewUser) {
            try {
                New-MailboxCreation -SamAccountName $SamAccountName -BusinessUnitShortCode $CoreData.Attributes['BusinessUnit']
                $UpdateTaskListParams = @{
                    TaskName = 'Provision Mailbox'
                    Status   = 'Completed'
                    Result   = 'Success'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                Write-Information '[Mailbox Creation] Mailbox created successfully'
                $CoreData.SetSuccessful('New-MailboxCreation')
                Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description "Mailbox provisioning successful for: $SamAccountName." -AssystEnvironment $Environment
                $UpdateHtmlData.Add('Mailbox creation', 'Success')
            }
            catch {
                $UpdateTaskListParams = @{
                    TaskName = 'Provision Mailbox'
                    Status   = 'Completed'
                    Result   = 'Failed'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                $CoreData.SetFunctionError('New-MailboxCreation', $_.Exception.Message)
                Write-Warning -Message $_.Exception.Message
                $UpdateHtmlData.Add('Mailbox creation', 'FAILED!')
                Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description "Mailbox provisioning failed for: $SamAccountName." -AssystEnvironment $Environment -ErrorAction Continue
                $AssigningNotes = @"
The automation failed at the 'Provision Mailbox' stage.  Please review and complete the outstanding tasks:
$($AutomationTaskList | Format-List | Out-String -Width 256)
"@
                $AssignmentInformation = @{
                    IncidentId                = $LinkedTaskReference
                    AssignedServiceDepartment = $GenericConfig.ResolverGroups.IAM
                    AssystEnvironment         = $Environment
                    AssigningNotes            = $AssigningNotes
                    ErrorAction               = 'CONTINUE'
                }
                Set-AssystIncidentAssignee @AssignmentInformation
                $InvokeScriptFailureParam.Summary = 'Mailbox creation failed'
                $InvokeScriptFailureParam.CoreDataObject = $CoreData
                $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
                $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
                # Need to set Exit explicitly to True to ensure the script exits if a previous invocation set it to False.
                $InvokeScriptFailureParam.Exit = $true
                Invoke-ScriptFailure @InvokeScriptFailureParam
            }
        }
        else {
            $UpdateTaskListParams = @{
                TaskName = 'Provision Mailbox'
                Status   = 'Skipped'
                Detail   = 'A mailbox was not requested for this user.'
            }
            Update-AutomationTaskList @UpdateTaskListParams

            # If no mailbox was requested, we must set the UserPrincipalName.

            try {
                $NewUserPrincipalNameParam = @{
                    GivenName             = $CoreData.Attributes['GivenName']
                    Surname               = $CoreData.Attributes['Surname']
                    BusinessUnitShortCode = $BusinessUnitShortCode
                }

                $UserPrincipalName = New-UserPrincipalName @NewUserPrincipalNameParam

                $SetADUserParam = @{
                    Identity          = $SamAccountName
                    UserPrincipalName = $UserPrincipalName
                    Server            = $Server
                    ErrorAction       = 'STOP'
                }

                # If it's a TWO_UK_CAP account, the Mail attribute should be set to the same value as the UPN.

                if ($BusinessUnitShortCode -eq 'TWO_UK_CAP') {
                    $SetADUserParam.Add('EmailAddress', $UserPrincipalName)
                }

                Set-ADUser @SetADUserParam

                $UpdateTaskListParams = @{
                    TaskName = 'Set UserPrincipalName'
                    Status   = 'Completed'
                    Result   = 'Success'
                    Detail   = "UserPrincipalName is $UserPrincipalName."
                }
                Update-AutomationTaskList @UpdateTaskListParams
            }
            catch {
                Write-Warning 'Failed to set UserPrincipalName.'
                $UpdateTaskListParams = @{
                    TaskName = 'Set UserPrincipalName'
                    Status   = 'Completed'
                    Result   = 'Failed'
                    Detail   = 'An error occurred when setting the UserPrincipalName.'
                }
                Update-AutomationTaskList @UpdateTaskListParams
            }
        }

        # User information from the created account is required by the rest of the script

        $UserPropertiesToQuery = @(
            'CanonicalName'
            'DisplayName'
        )
        if ($CoreData.Attributes['MailboxProvisioning'] -eq 'Y') {
            $UserPropertiesToQuery += 'EmailAddress'
        }

        try {
            $UserDetails = Get-ADUser -Identity $SamAccountName -Properties $UserPropertiesToQuery -Server $Server
        }
        catch {
            Write-Warning 'Unable to get user account'
        }

        Test-BritishSugarDomain

        if ($null -ne $UserDetails -and $null -ne $NewUser) {
            Write-Information "Populating extensionAttribute9 with $($UserDetails.CanonicalName)"
            $NewProperties = @{
                extensionAttribute9 = $UserDetails.CanonicalName
            }
            try {
                Set-ADUser -Identity $SamAccountName -Add $NewProperties -Server $Server
                $UpdateTaskListParams = @{
                    TaskName = 'Update extensionAttribute9'
                    Status   = 'Completed'
                    Result   = 'Success'
                }
                Update-AutomationTaskList @UpdateTaskListParams
            }
            catch {
                $UpdateTaskListParams = @{
                    TaskName = 'Update extensionAttribute9'
                    Status   = 'Completed'
                    Result   = 'Failed'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                Write-Warning 'Failed to set extensionAttribute9.'
            }
        }

        if ($null -ne $UserDetails -and $null -ne $NewUser) {

            Write-Information "Processing user's licences."
            Write-Information "DepartmentKey is $DepartmentKey"

            if ($UserPropertiesToQuery -contains 'EmailAddress') {
                $CorporateEmail = $UserDetails.EmailAddress
            }
            else {
                $CorporateEmail = 'No e-mail.'
            }

            $LicenceRequestParam = @{
                Action                = 'Joiner'
                Forename              = $CoreData.Attributes['Givenname']
                Surname               = $CoreData.Attributes['Surname']
                ADUserName            = $SamAccountName
                CorporateEmail        = $CorporateEmail
                BusinessUnitShortcode = $CoreData.Attributes['BusinessUnit']
                Office                = $CoreData.Attributes['Office']
                DepartmentKey         = $DepartmentKey
                Title                 = $CoreData.Attributes['Title']
            }

            [Array]$LicenceRequests = New-LicenceRequest @LicenceRequestParam -Verbose

            if ($CoreData.Attributes['TelephonyRequired'] -like 'Y*') {
                $TelephonyLicenceRequest = [PSCustomObject][Ordered] @{
                    'ACTION'               = 'Joiner'
                    'FORENAME'             = $CoreData.Attributes['Givenname']
                    'SURNAME'              = $CoreData.Attributes['Surname']
                    'AD USER NAME'         = $SamAccountName
                    'CORPORATE EMAIL'      = $CorporateEmail
                    'BUSINESS UNIT ABBRV.' = $CoreData.Attributes['BusinessUnit']
                    'SKU TO BE APPLIED'    = 'Microsoft Teams Phone Standard'
                    'SPECIAL CONDITIONS'   = 'None.'
                }
                $LicenceRequests += $TelephonyLicenceRequest
            }

            $RequestCSVPath = "$($GenericConfig.LicenceRequests.RequestPath)\$IncidentId.csv"

            if (($null -ne $LicenceRequests) -and ($LicenceRequests[0].Unknown -eq $true)) {
                $LicenceApprovalIncidentParam = @{
                    AssystReference    = $IncidentId
                    AssystQueue        = $LicenceRequests[0].BUApprovalQueue
                    Configuration      = $GenericConfig
                    Environment        = $Environment
                    Request            = $LicenceRequestParam
                    AutomationTaskList = $AutomationTaskList
                    BusinessUnitData   = $BusinessUnitData
                }
                New-LicenceApprovalIncident @LicenceApprovalIncidentParam -Verbose
            }
            elseif ($LicenceRequests.Count -gt 0) {
                $LicenceRequests | Export-Csv -Path $RequestCSVPath -NoTypeInformation
                $LicenceIncidentParam = @{
                    AssystReference    = $IncidentId
                    Configuration      = $GenericConfig
                    LicenceRequestPath = $RequestCSVPath
                    Environment        = $Environment
                    AutomationTaskList = $AutomationTaskList
                    BusinessUnitData   = $BusinessUnitData
                }
                New-LicenceIncident @LicenceIncidentParam -Verbose
            }
            else {
                $UpdateTaskListParams = @{
                    TaskName = 'Provision Licence'
                    Status   = 'Skipped'
                    Detail   = 'No licence information found for this request.'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                Write-Information 'No licence information found for this request.'
            }
        }

        # Check if special characters need adding back to the name related properties of the user.
        # Be aware that from this point on, GivenName and Surname may contain special characters.
        Write-Information 'Checking if name properties of user need updating to include special characters.'
        if ($CoreData.Attributes['Office'] -eq 'ABEZ_Darmstadt-' -or
            $CoreData.Attributes['Office'] -eq 'ABEZ_Rajamaki-') {

            $CoreData.Attributes['GivenName'] = $CoreData.Attributes['OriginalGivenName']
            $CoreData.Attributes['Surname'] = $CoreData.Attributes['OriginalSurname']

            $SetADUserAttributes = @{
                Identity = $SamAccountName
                Server   = $Server
            }

            if ($UserDetails.GivenName -ne $CoreData.Attributes['GivenName']) {
                Write-Information "$($UserDetails.GivenName) requires changing to $($CoreData.Attributes['GivenName'])"
                $SetADUserAttributes['GivenName'] = $CoreData.Attributes['GivenName']
            }
            if ($UserDetails.Surname -ne $CoreData.Attributes['Surname']) {
                Write-Information "$($UserDetails.Surname) requires changing to $($CoreData.Attributes['Surname'])"
                $SetADUserAttributes['Surname'] = $CoreData.Attributes['Surname']
            }
            if ($SetADUserAttributes.ContainsKey('GivenName') -or $SetADUserAttributes.ContainsKey('Surname')) {
                Write-Information "$($UserDetails.DisplayName) requires changing to $($CoreData.Attributes['GivenName']) $($CoreData.Attributes['Surname'])"
                Write-Information "$($UserDetails.Name) requires changing to $($CoreData.Attributes['GivenName']) $($CoreData.Attributes['Surname'])"
                $SetADUserAttributes['DisplayName'] = $($CoreData.Attributes['GivenName']) + ' ' + $($CoreData.Attributes['Surname'])
            }

            try {
                Set-ADUser @SetADUserAttributes -ErrorAction Stop
                Write-Information "Successfully updated user's GivenName, Surname, and DisplayName attributes."
                Rename-ADObject -Identity $UserDetails.DistinguishedName -NewName $SetADUserAttributes['DisplayName']
                Write-Information "Successfully updated user's DistinguishedName, Name, and cn attributes."
            }
            catch {
                Write-Warning "Unable to update user account attributes. $($_.Exception.Message)"
            }
        }

        # Start of sending new starter credentials to recipient automatically via email

        # Setup $NewStarter with values required for Send-JoinerEmail script
        $NewStarter = [PSCustomObject]@{
            GivenName         = $CoreData.Attributes['GivenName']
            Surname           = $CoreData.Attributes['Surname']
            Username          = $SamAccountName
            UserPrincipalName = $UserDetails.UserPrincipalName
            Password          = $ADAccountPasswordData.Password
            EmailAddress      = $UserDetails.EmailAddress
        }

        if ($BusinessUnitData.Global.NotificationSettings.ThirdParty) {

            Write-Information 'Third Party notification setting is set to true.
            Checking if third party email address has been provided in the JSON file.'

            if ([string]::IsNullOrEmpty($CoreData.Attributes['ThirdPartyEmailAddress']) -or
                [string]::IsNullOrWhiteSpace($CoreData.Attributes['ThirdPartyEmailAddress'])) {

                $ThirdPartyEmailAddressProvided = $false
                Write-Information 'No Third Party email address has been provided in the JSON file.'
            }
            else {
                $ThirdPartyEmailAddressProvided = $true
                Write-Information 'Third Party email address has been provided in the JSON file.'
            }
        }

        $RecipientList = New-Object Collections.Generic.List[PsCustomObject]
        # As there can be multiple Sites in a single config, we only want to check for NotificationSettings and AdditionalMailbox
        # attributes in the site linked to the New Starter. Therefore, get Json attributes and values for the specific site only.
        $OfficeJsonData = $BusinessUnitData.Sites | Where-Object { $_.Name -eq $CoreData.Attributes['Office'] }

        if (!($ADAccountPasswordData.PasswordReset)) {
            Write-Information "Password was not set. Script will not attempt to send email containing new starter's credentials."
        }
        # If sending credentials to the linemanager is set to no, and there are no additional mailboxes to send credentials to,
        # and there are no AdditionalMailboxes within the new starter's site, and e-mails are not sent to a third party recipient
        # then raise an incident ticket.
        elseif ($BusinessUnitData.Global.NotificationSettings.LineManager -eq 'No' -and
            $BusinessUnitData.Global.NotificationSettings.AdditionalMailbox.Count -eq 0 -and
            $OfficeJsonData.NotificationSettings.AdditionalMailbox.Count -eq 0 -and
                (-not $BusinessUnitData.Global.NotificationSettings.ThirdParty -or
            $BusinessUnitData.Global.NotificationSettings.ThirdParty -eq 'No')) {

            $NoMailboxesForCredentials = $true
            Write-Information "No recipient settings found. IAM will need to reset credentials and contact the BU with $SamAccountName credentials."

            $JoinerIncidentParam = @{
                AssystReference = $IncidentId
                Configuration   = $GenericConfig
                Environment     = $Environment
                IncidentReason  = 'NoNotificationMailboxes'
            }

            New-JoinerIncident @JoinerIncidentParam -Verbose
        }
        elseif ($BusinessUnitData.Global.NotificationSettings.LineManager -eq 'Yes' -and
            -not($LineManagerProvided) -and
            $BusinessUnitData.Global.NotificationSettings.AdditionalMailbox.Count -eq 0 -and
            $OfficeJsonData.NotificationSettings.AdditionalMailbox.Count -eq 0 -and
                (-not $BusinessUnitData.Global.NotificationSettings.ThirdParty -or
            $BusinessUnitData.Global.NotificationSettings.ThirdParty -eq 'No')) {

            $NoMailboxesForCredentials = $true
            Write-Information 'LineManager notification is Yes, but no Line Manager is provided.'
            Write-Information "No recipient settings found. IAM will need to reset credentials and contact the BU with $SamAccountName credentials."

            $JoinerIncidentParam = @{
                AssystReference = $IncidentId
                Configuration   = $GenericConfig
                Environment     = $Environment
                IncidentReason  = 'NoLineManager'
            }

            New-JoinerIncident @JoinerIncidentParam -Verbose
        }
        elseif ($BusinessUnitData.Global.NotificationSettings.ThirdParty -eq 'Yes' -and
            -not($ThirdPartyEmailAddressProvided)) {

            Write-Information "$BusinessUnitShortCode JSON configuration 'ThirdParty' is set to Yes.
                However, no third party email address was provided in the JSON. Raising an incident."

            $JoinerIncidentParam = @{
                AssystReference = $IncidentId
                Configuration   = $GenericConfig
                Environment     = $Environment
                IncidentReason  = 'NoThirdPartyEmailAddress'
            }

            New-JoinerIncident @JoinerIncidentParam -Verbose
        }
        else {

            if ($BusinessUnitData.Global.NotificationSettings.LineManager -eq 'Yes' -and
                $LineManagerProvided) {

                Write-Information "$BusinessUnitShortCode JSON configuration LineManager is set to Yes.
                Credentials will be sent to Line Manager."

                $ManagerFullName = $ManagerEmailAddress.Split('@')[0]
                $ManagerGivenName = $ManagerFullName.Split('.')[0]
                $ManagerSurname = $ManagerFullName.Split('.')[1]
                $RecipientList.Add([PSCustomObject]@{
                        GivenName    = $ManagerGivenName
                        Surname      = $ManagerSurname
                        EmailAddress = $ManagerEmailAddress
                    })
            }

            if ($BusinessUnitData.Global.NotificationSettings.ThirdParty -eq 'Yes' -and
                $ThirdPartyEmailAddressProvided) {

                Write-Information "$BusinessUnitShortCode JSON configuration 'ThirdParty' is set to Yes.
                Credentials will be sent to the third party address entered on the form."

                $ThirdPartyEmailAddress = $CoreData.Attributes['ThirdPartyEmailAddress'].Trim()
                $ThirdPartyFullName = $ThirdPartyEmailAddress.Split('@')[0]
                $ThirdPartyGivenName = $ThirdPartyFullName.Split('.')[0]
                $ThirdPartySurname = $ThirdPartyFullName.Split('.')[1]
                $RecipientList.Add([PSCustomObject]@{
                        GivenName    = $ThirdPartyGivenName
                        Surname      = $ThirdPartySurname
                        EmailAddress = $ThirdPartyEmailAddress
                    })
            }

            if ($BusinessUnitData.Global.NotificationSettings.AdditionalMailbox.Count -gt 0) {

                Write-Information "$BusinessUnitShortCode JSON configuration AdditionalMailbox count is greater
                than 0. Additional Mailboxes will be sent credentials."

                foreach ($MailboxAddress in $BusinessUnitData.Global.NotificationSettings.AdditionalMailbox) {
                    if ($MailboxAddress -as [System.Net.Mail.MailAddress]) {
                        Write-Information "$MailboxAddress is a valid mailbox address."
                        $RecipientList.Add([PSCustomObject]@{
                                EmailAddress = $MailboxAddress
                            })
                    }
                    else {
                        Write-Warning "$MailboxAddress is an invalid mailbox address. Please review
                        $BusinessUnitShortCode JSON configuration."
                    }
                }
            }

            # Check if Sites attribute in JSON contains attribute NotificationSettings, then AdditionalMailbox
            # If both attributes are found, check AdditonalMailbox has a count greater than 0 (therefore contains addresses)
            # If counter greater than 0 the additional mailboxes need adding to the list for sending email credentials

            if (!($OfficeJsonData.PsObject.Properties.name -eq 'NotificationSettings')) {
                Write-Information "$BusinessUnitShortCode JSON configuration does not contain NotificationSettings attribute within Site $($CoreData.Attributes['Office'])."
            }
            elseif (!($OfficeJsonData.NotificationSettings.PsObject.Properties.name -eq 'AdditionalMailbox')) {
                Write-Warning "Office $($CoreData.Attributes['Office']) within $BusinessUnitShortCode JSON configuration does not contain AdditionalMailbox attribute within the NotificationSettings attribute."
            }
            elseif ($OfficeJsonData.NotificationSettings.AdditionalMailbox.Count -eq 0) {
                Write-Warning "Office $($CoreData.Attributes['Office']) within $BusinessUnitShortCode JSON configuration has 0 values within AdditionalMailboxes."
            }
            elseif ($OfficeJsonData.NotificationSettings.AdditionalMailbox.Count -gt 0) {

                Write-Information "$BusinessUnitShortCode JSON configuration contains AdditionalMailbox within Sites and the
                count is greater than 0. Site based additional mailboxes will be sent credentials."

                foreach ($MailboxAddress in $OfficeJsonData.NotificationSettings.AdditionalMailbox) {
                    if ($MailboxAddress -as [System.Net.Mail.MailAddress]) {
                        Write-Information "$MailboxAddress is a valid mailbox address."
                        $RecipientList.Add([PSCustomObject]@{
                                EmailAddress = $MailboxAddress
                            })
                    }
                    else {
                        Write-Warning "$MailboxAddress is an invalid mailbox address. Please review
                        $BusinessUnitShortCode JSON configuration."
                    }
                }
            }

            $JoinerEmailParams = @{
                NewStarter        = $NewStarter
                RecipientList     = $RecipientList
                IncidentId        = $IncidentId
                AzAppEmailAddress = $GenericConfig.EmailAzAppReg.AppEmailAddress
                BusinessUnitId    = $CoreData.Attributes['BusinessUnit']
            }

            try {
                Send-JoinerEmail @JoinerEmailParams -Verbose
                $UpdateTaskListParams = @{
                    TaskName = 'E-mail Account Details'
                    Status   = 'Completed'
                    Result   = 'Success'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                Write-Information 'New Starter credentials email successfully sent.'
                $CoreData.SetSuccessful('Send-JoinerEmail')
                Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description "New Starter credentials sent successfully for: $SamAccountName." -AssystEnvironment $Environment
                $UpdateHtmlData.Add('Credential Email Sending', 'Success')
                $EmailSent = $true
            }
            catch {
                $UpdateTaskListParams = @{
                    TaskName = 'E-mail Account Details'
                    Status   = 'Completed'
                    Result   = 'Failed'
                    Detail   = 'An error occured while e-mailing the account details.'
                }
                Update-AutomationTaskList @UpdateTaskListParams
                $EmailSent = $false
                $CoreData.SetFunctionError('Send-JoinerEmail', $_.Exception.Message)
                Write-Warning 'Unable to Send Joiner Email.'
                $UpdateHtmlData.Add('Credential Email Sending', 'FAILED!')
                Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description "New Starter credentials email sending failed for: $SamAccountName." -AssystEnvironment $Environment -ErrorAction Continue
                Set-AssystIncidentAssignee -IncidentId $LinkedTaskReference -AssignedServiceDepartment $GenericConfig.ResolverGroups.IAM -AssystEnvironment $Environment -AssigningNotes 'Automation failed at credential email stage. Please review and complete. This will require a password reset of the New Starter account' -ErrorAction Continue
                $InvokeScriptFailureParam.Summary = 'User credential email sending failed'
                $InvokeScriptFailureParam.CoreDataObject = $CoreData
                $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
                $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
                $InvokeScriptFailureParam.Exit = $false
                Invoke-ScriptFailure @InvokeScriptFailureParam
            }

        }

        $UpdateHtmlData.Add('Script status', 'Completed')
        try {
            Update-Html -Data $UpdateHtmlData
            $CoreData.SetSuccessful('Update-Html')
        }
        catch {
            Write-Information 'Unable to update HTML file'
            $CoreData.SetFunctionError('Update-Html', $_.Exception.Message)
        }
    }
}
else {
    $CoreData.SetFunctionError('Test-RequiredAttributes', 'Required attributes not present')
    $InvokeScriptFailureParam.Summary = 'Required attributes not present'
    $InvokeScriptFailureParam.CoreDataObject = $CoreData
    $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
    $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
    Invoke-ScriptFailure @InvokeScriptFailureParam
}

try {
    # If statement stops Assyst ticket closure so IAM know to send New Starter credentials manually
    if ($EmailSent -eq $true -or
        $CoreData.Attributes['SscAutoEnable'] -notlike 'Yes*' -or
        $NoMailboxesForCredentials -eq $true) {

        $UpdateTaskListParams = @{
            TaskName = 'Close Task Ticket'
            Status   = 'Completed'
            Result   = 'Success'
        }
        Update-AutomationTaskList @UpdateTaskListParams

        $Description = @"
Automation stage completed.

Please review the automation task list and complete any outstanding tasks:
$($AutomationTaskList | Format-List | Out-String -Width 256)
"@

        Write-Information 'Attempting to close Assyst incident.'
        $CloseIncidentParam = @{
            IncidentId        = $LinkedTaskReference
            Description       = $Description
            CauseItem         = 'AUTOMATION - JOINER'
            CauseCategory     = 'ACCOUNT CREATED'
            AssystEnvironment = $Environment
        }
        $Closure = Close-AssystIncident @CloseIncidentParam
        if ($Closure.ResponseCode -ne '200') {
            Write-Warning 'Closure Response Code was not 200. Moving file to FailedFiles and throwing.'
            Move-Item -Path $JsonPath -Destination "$BaseConfigPath\FailedFiles\$Environment" -ErrorAction 'STOP'
            throw
        }
        else {
            Write-Information 'Closure Response code was 200. Moving file to Completed.'
            Move-Item -Path $JsonPath -Destination "$BaseConfigPath\Completed\$Environment" -ErrorAction 'STOP'
        }
    }
    else {
        Write-Information 'Moving file to FailedFiles as it did not pass EmailSent / SscAutoEnable / NoMailboxesForCredentials checks.'
        Move-Item -Path $JsonPath -Destination "$BaseConfigPath\FailedFiles\$Environment" -ErrorAction 'STOP'
    }
}
catch {
    $UpdateTaskListParams = @{
        TaskName = 'Close Task Ticket'
        Status   = 'Completed'
        Result   = 'Failed'
    }
    Update-AutomationTaskList @UpdateTaskListParams

    Add-AssystAnalysisInformation -IncidentId $LinkedTaskReference -Description 'Failed to Close Task Ticket' -AssystEnvironment $Environment -ErrorAction Continue
    $AssigningNotes = @"
The automation failed at the 'Close Task Ticket' stage.  Please review and complete the outstanding tasks:
$($AutomationTaskList | Format-List | Out-String -Width 256)
"@
    $AssignmentInformation = @{
        IncidentId                = $LinkedTaskReference
        AssignedServiceDepartment = $GenericConfig.ResolverGroups.IAM
        AssystEnvironment         = $Environment
        AssigningNotes            = $AssigningNotes
        ErrorAction               = 'CONTINUE'
    }
    Set-AssystIncidentAssignee @AssignmentInformation

    Write-Information '[Assyst] Unable to close Assyst Incident'
    $InvokeScriptFailureParam.Summary = 'Unable to close Assyst incident'
    $InvokeScriptFailureParam.CoreDataObject = $CoreData
    $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
    $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
    # Need to set Exit explicitly to True to ensure the script exits if a previous invocation set it to False.
    $InvokeScriptFailureParam.Exit = $false
    Invoke-ScriptFailure @InvokeScriptFailureParam
}

# Add transcript to Assyst ticket
try {
    $AssystAttachmentsPath = "$BaseConfigPath\AssystAttachments"
    $TranscriptTempPath = "$AssystAttachmentsPath\Transcript-$IncidentId-$TranscriptingGuid.txt"
    $CoreDataPath = $AssystAttachmentsPath

    Copy-Item -Path $TranscriptPath -Destination $TranscriptTempPath
    Add-AssystAttachment -IncidentId $IncidentId -Path $TranscriptTempPath -AssystEnvironment $Environment -Description "New Starter Provisioning Log for attempt $ProvisionAttemptGuid"
    Remove-Item -Path $TranscriptTempPath

    Export-CoreData -CoreDataObject $CoreData -TranscriptingGuid $TranscriptingGuid -Path $CoreDataPath -AssystReference $IncidentId -Environment $Environment -AssystDescription "New Starter Provisioning Log for attempt $ProvisionAttemptGuid"
    $CoreData.SetSuccessful('Export-CoreData')
}
catch {
    $CoreData.SetFunctionError('Export-CoreData', $_.Exception.Message)
    Write-Information ('[Assyst] Unable to add attachment to Assyst Incident: {0}' -f $_.Exception.Message)
    $InvokeScriptFailureParam.Summary = 'Unable to add attachment to Assyst incident'
    $InvokeScriptFailureParam.CoreDataObject = $CoreData
    $InvokeScriptFailureParam.BusinessUnitData = $BusinessUnitData
    $InvokeScriptFailureParam.AutomationTaskList = $AutomationTaskList
    # Need to set Exit explicitly to True to ensure the script exits if a previous invocation set it to False.
    $InvokeScriptFailureParam.Exit = $false
    Invoke-ScriptFailure @InvokeScriptFailureParam
}

try {
    Disconnect-MgGraph
    Write-Information 'Successfully disconnected from Connect-MgGraph using Disconnect-MgGraph cmdlet.'
}
catch {
    Write-Warning 'Unable to disconnect from Graph.'
}

Write-Information "Task Summary:`r`n $($AutomationTaskList | Out-String -Width 256)"

Stop-Transcript