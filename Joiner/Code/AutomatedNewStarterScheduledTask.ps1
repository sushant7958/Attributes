# ----------------------------------------------------------------------------------------------------------
# PURPOSE:    To process requests for New Starter Automation
#
# VERSION     DATE         USER                DETAILS
# 1           13/05/2019   Craig Tolley        First version
# 1.1         20/05/2019   Craig Tolley        Added validation of input files
# 1.2         10/06/2019   Craig Tolley        Added support for both LIVE and DEV environments
#
# ----------------------------------------------------------------------------------------------------------

# Start transcript and remove old versions
$BaseConfigPath = '\\bsg.local\ssc\Automation\NewStarter'
$LogsPath = Join-Path -Path $BaseConfigPath -ChildPath 'Logs'
$InputFilesPath = Join-Path -Path $BaseConfigPath -ChildPath 'InputFiles'
$ScriptsPath = Join-Path -Path $BaseConfigPath -ChildPath 'Scripts'
$ScriptsPath | Unblock-File
$InvalidInputFilesPath = Join-Path -Path $BaseConfigPath -ChildPath 'InvalidInputFiles'

$MaximumConcurrentPowerShellSessions = 10

$TranscriptPath = (Join-Path -Path $LogsPath -ChildPath "Transcript-$(Get-Date -Format yyyy-MM-dd)-$($env:COMPUTERNAME).log")
Start-Transcript -Path $TranscriptPath -Append

Get-ChildItem -Path $LogsPath -Filter 'Transcript-*.log' -Depth 0 -File -ErrorAction SilentlyContinue | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-14) } | Remove-Item -ErrorAction SilentlyContinue

# ----------------------------------------------------------------------------------------------------------
# High Availability Checks
# We want to run this task on PXGBSSC1AOP001 if it is available, and only run on an AOP server
if ($PSSenderInfo) {
    Write-Error 'Script is being executed in a remote session. Not supported. Exiting' -ErrorAction Stop
    Stop-Transcript
    exit
}

switch ($env:COMPUTERNAME) {
    'PXGBSSC1AOP001' {
        "Running on the primary execution node. $env:COMPUTERNAME"
    }

    'PXGBSSC2AOP001' {
        try {
            Invoke-Command -ComputerName 'PXGBSSC1AOP001' -ScriptBlock { "Connected to $env:COMPUTERNAME at $(Get-Date)" } -ErrorAction Stop
            'Primary execution node is online, and this is not the primary node. Exiting'
            Stop-Transcript
            exit
        }
        catch {
            "No connection. Script will execute on $env:COMPUTERNAME"
        }
    }

    default {
        Write-Error "This script is only designed to run on the AOP servers. Computer name is $env:COMPUTERNAME. Exiting" -ErrorAction Stop
        Stop-Transcript
        exit
    }
}

# ----------------------------------------------------------------------------------------------------------

# Start processing for each new input file
$InputFilesToProcess = Get-ChildItem -Filter '*.json' -Path $InputFilesPath -Recurse
Write-Output ('Json Input Files found to process: {0}' -f @($InputFilesToProcess).Count)

foreach ($ReqToProcess in $InputFilesToProcess) {
    Write-Output ('-' * 100)
    Write-Output "    Processing: $($ReqToProcess.Fullname)"

    # Import configuration file and validate that it is in the correct format and importable.
    $InvalidFileDetectedReason = $null

    try {
        $ReqDetails = Get-Content -Path $ReqToProcess.FullName -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        Write-Output $ReqDetails        
    }
    catch {
        $InvalidFileDetectedReason = ('Invalid Json file found. Failed to import. Error: {0}' -f $_.Exception.Error)
    }

    if ($ReqDetails.Processed) {
        Write-Output "$($ReqToProcess.Basename) has already been processed and will be skipped."
        Continue
    }
    else {
        $ReqDetails | Add-Member -MemberType NoteProperty -Name 'Processed' -Value $True
        $ReqDetails | ConvertTo-Json | Set-Content $ReqToProcess.FullName
    }

    # Assyst unable to put the Incident ID in the Json. Read this from the filename, and insert to the file if not already present
    if ($ReqDetails.PsObject.Properties.Name -notcontains 'IncidentId') {
        Add-Member -InputObject $ReqDetails -MemberType NoteProperty -Name 'IncidentId' -Value $ReqToProcess.BaseName
    }

    if ($ReqToProcess.BaseName -ne $ReqDetails.IncidentId -and -not $InvalidFileDetectedReason) {
        $InvalidFileDetectedReason = 'Invalid Json file found. Filename does not match Incident ID contained within the file.'
    }

    if ($ReqDetails.PsObject.Properties.Name -notcontains 'GivenName' -and -not $InvalidFileDetectedReason) {
        $InvalidFileDetectedReason = 'Invalid Json file found. Property missing from JSON: GivenName.'
    }

    if ($ReqDetails.PsObject.Properties.Name -notcontains 'Surname' -and -not $InvalidFileDetectedReason) {
        $InvalidFileDetectedReason = 'Invalid Json file found. Property missing from JSON: Surname'
    }

    if ($ReqDetails.PsObject.Properties.Name -notcontains 'BusinessUnit' -and -not $InvalidFileDetectedReason) {
        $InvalidFileDetectedReason = 'Invalid Json file found.  Property missing from JSON: BusinessUnit'
    }

    if ('LIVE', 'DEV' -notcontains $ReqToProcess.Directory.Name) {
        $InvalidFileDetectedReason = 'Invalid Json file found. File has not been found within a valid environment directory.'
    }
    else {
        $TargetEnvironment = $ReqToProcess.Directory.Name
        Write-Output "Target Environment: $TargetEnvironment"
    }

    # Set Assyst API Credentials
    switch ($TargetEnvironment) {
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
        Set-AssystAPICredential @APICred -AssystEnvironment $TargetEnvironment
    }
    catch {
        Write-Warning 'Unable to set Assyst credentials, exiting.'
        Write-Warning "$($_.Exception.Message)"
        Exit
    }

    if ($InvalidFileDetectedReason) {
        Write-Output $InvalidFileDetectedReason
        if ($null -ne $TargetEnvironment) {
            $InvalidInputFilePath = (Join-Path -Path $InvalidInputFilesPath -ChildPath $TargetEnvironment)
        }
        $DestinationFileName = (Join-Path -Path $InvalidInputFilesPath -ChildPath ("$(Get-Date -Format yyyyMMdd-HHmm)-$($ReqToProcess.Name)"))
        Move-Item -Path $ReqToProcess.FullName -Destination $DestinationFileName -Force
        Write-Output 'Raising an incident as the input file was invalid.'
        $NewAssystIncident = @{
            Summary                = "Automated New Starter: Invalid Input File Detected: $($ReqToProcess.Name)"
            Description            = "An invalid input file has been detected. This file has been moved to $DestinationFileName. The reason it has been detected as invalid is: $InvalidFileDetectedReason"
            Priority               = 'P3 - 16HR'
            Seriousness            = 'P3_DIS_SGL_USR'
            AffectedUserShortCode  = 'SSCAUTOMATIONSERVICE'
            SystemServiceShortCode = 'AUTOMATION - JOINER'
            CategoryShortCode      = 'FAULT UNDEFINED'
            AssignedResolverGroup  = 'SSC AUTOMATION'
            SendEmailIfApiFails    = $True
            AssystEnvironment      = $TargetEnvironment
        }

        New-AssystIncident @NewAssystIncident
        continue
    }

    # Determine if the process is still running. Set the $RetryRequest value which will be passed to the processing script
    if ($ReqDetails.ProcessId -and (Get-Process -Id $ReqDetails.ProcessId -ComputerName $ReqDetails.ProcessHost -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq 'powershell' })) {
        Write-Output ('    Process ID {0} is still running on {1}' -f $ReqDetails.ProcessId, $ReqDetails.ProcessHost)
        continue
    }
    elseif ($ReqDetails.ProcessId) {
        Write-Output ('    Process ID {0} is no longer running on {1}. Request will be resubmitted.' -f $ReqDetails.ProcessId, $ReqDetails.ProcessHost)
        $RetryRequest = $true
    }
    else {
        Write-Output '    New request. Starting processing'
        $RetryRequest = $false
    }

    # Simplistic rate limiting
    while ((Get-Process -Name powershell -ErrorAction SilentlyContinue | Measure-Object).Count -ge $MaximumConcurrentPowerShellSessions) {
        Start-Sleep -Seconds 5 
    }

    # Call the provisioning script with all details
    if ($Null -eq $ReqDetails.RequestGuid) {
        $ProvisionAttemptGuid = (New-Guid).Guid
    }
    else {
        $ProvisionAttemptGuid = $ReqDetails.RequestGuid
    }

    $Arguments = "-Command ""$ScriptsPath\NewStarterProvisioning.ps1 -TranscriptingGuid $ProvisionAttemptGuid -JsonPath $($ReqToProcess.FullName) -Environment '$TargetEnvironment' -BaseConfigPath $BaseConfigPath -IncidentId $($ReqDetails.IncidentId)"""

    $JobProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList $Arguments -PassThru
    Write-Output '    Process Started.'
    Write-Output "    Provision Attempt Guid: $($ProvisionAttemptGuid)"
    Write-Output "    Process ID: $($JobProcess.Id)"
    Write-Output "    Process Working Directory: $($Pwd.Path)"
    Write-Output "    Process Arguments: $Arguments"

    # Without this, the last request is not processed. It never seems to get initialised and started before the parent script exits, and as such is never run.
    Start-Sleep -Seconds 15
}

Stop-Transcript