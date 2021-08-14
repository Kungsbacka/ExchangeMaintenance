$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\SqlDatabase.ps1"
. "$PSScriptRoot\ExchangeOnline.ps1"
. "$PSScriptRoot\LogItem.ps1"
. "$PSScriptRoot\MailboxTask.ps1"
. "$PSScriptRoot\Logger.ps1"

function RescheduleTask {
    # -Daily makes sure that the task will run even if the dispatcher failed to reschedule
    $trigger = New-ScheduledTaskTrigger -At (Get-Date).Add($Script:Config.BatchDelay) -Daily
    $null = Set-ScheduledTask -TaskName $Script:Config.ScheduledTaskName -Trigger $trigger
}

# Clean up default parameter values so it will not interfere
$default = $PSDefaultParameterValues.GetEnumerator() | Where-Object Name -like 'Connect-ExchangeOnline*'
foreach ($item in $default) {
    $PSDefaultParameterValues.Remove($item.Name)
}

$tasks = New-Object -TypeName 'System.Collections.ArrayList'
Get-ChildItem "$PSScriptRoot\Tasks\*.ps1" | ForEach-Object {
    . $_.FullName
    $task = New-Object -TypeName ([System.IO.Path]::GetFileNameWithoutExtension($_.Name))
    $null = $tasks.Add($task)
}
try {
    [ExchangeOnline]::Connect()
}
catch {
    Log -Task 'Dispatcher:Connect' -Message 'Failed to connect to Exchange' -ErrorRecord $_
    RescheduleTask
    exit
}
[ExchangeOnline]::Simulate = $Script:Config.Simulate

Log -Task 'Dispatcher:Start' -Message "Started new batch. Batch size is $($Script:Config.BatchSize)"

$startTime = Get-Date

$queue =  New-Object -TypeName 'System.Collections.Generic.Queue[object]'
try {
    Import-Csv "$PSScriptRoot\mailboxes.csv" -Encoding UTF8 | ForEach-Object {
        if (-not $_.ResourceType) {
            $_.ResourceType = $null
        }
        $queue.Enqueue($_)
    }
    Log -Task 'Dispatcher:ImportCsv' -Message "Loaded $($queue.Count) saved mailboxes"
}
catch {
    # Failed to load file
}

if ($queue.Count -eq 0) {
    Log -Task 'Dispatcher:GetMailbox' -Message 'No saved mailboxes found. Fetching all mailboxes'
    try {
        $mailboxes = [ExchangeOnline]::GetMailbox(@{
            ResultSize = 'Unlimited'
            Properties = @('AddressBookPolicy','RetentionPolicy','ForwardingAddress','ForwardingSmtpAddress','IsResource','IsShared','ResourceType')
        })
    }
    catch {
        # Getting all mailboxes in the tenant can fail sometimes, so we reschedule and try again
        RescheduleTask
        exit
    }
    foreach ($item in $mailboxes) {
        $obj = [PSCustomObject]@{
            AddressBookPolicy = $item.AddressBookPolicy
            ExternalDirectoryObjectId = $item.ExternalDirectoryObjectId
            ForwardingAddress = $item.ForwardingAddress
            ForwardingSmtpAddress = $item.ForwardingSmtpAddress
            Guid = $item.Guid
            Identity = $item.PrimarySmtpAddress
            IsResource = $item.IsResource
            IsShared = $item.IsShared
            PrimarySmtpAddress = $item.PrimarySmtpAddress
            ResourceType = $item.ResourceType
            RetentionPolicy = $item.RetentionPolicy
        }
        $queue.Enqueue($obj)
    }
}

# Execute initializers
foreach ($task in $tasks) {
    try {
        $task.Initialize()
    }
    catch {
        $tasks.Remove($task)
        Log -Task 'Dispatcher:Initialize' -Message "Initialization for task $($task.GetType().Name) failed with error: $($_.ToString())"
    }
}
if ($tasks.Count -eq 0) {
    # If all initializers failed, there must be something seriously wrong and running the
    # script again will probably not solve the problem.
    Log -Task 'Dispatcher:Initialize' -Message "Initializers for all tasks failed. Script is NOT rescheduled."
    exit 99
}

$reasonsToReconnect = @(
    'ConnectionFailedTransientException'
    'ADTopologyEndpointNotFoundException'
    'PSRemotingTransportException'
    'DatabaseUnavailableException'
    'MapiNetworkErrorException'
    'UnableToFindServerForDatabaseException'
    'CommandNotFoundException'
    'MethodInvocationException'
    'DCOverloadedException'
    'ADServerSettingsChangedException'
    'OverBudgetException'
)

function ShouldReconnect($err)
{
    if ($err.CateoryInfo.Reason -in $reasonsToReconnect) {
        $true
    }
    elseif ($err.Exception.Message -like '*max proxy depth limitation*') {
        $true
    }
    $false
}

# Call ProcessMailbox on mailboxes in batch
$mailboxCount = 0
$needToReconnect = $false
while ($mailboxCount -lt $Script:Config.BatchSize -and $queue.Count -gt 0) {
    $mailbox = $queue.Dequeue()
    foreach ($task in $tasks) {
        try {
            $task.ProcessMailbox($mailbox)
        }
        catch {
            if (ShouldReconnect $_) {
                $needToReconnect = $true
                Log -Task 'Dispatcher:Process' -Message "Connection to Exchange Online broke with error: $($_.ToString())"
                break
            }
            # Log all other exceptions except "mailbox not found". Since we are working with a cached list of mailboxes,
            # this is bound to happen now and then and will just clutter the log.
            if ($_.CategoryInfo.Reason -ne 'ManagementObjectNotFoundException') {
                $params = @{
                    Task = 'Dispatcher:Process'
                    Mailbox = $mailbox
                    Message = "Processing task $($task.GetType().Name) failed with error: $($_.ToString())"
                }
                if (-not $needToReconnect) {
                    # Only save error record if the error is not identified in ShouldReconnect()
                    $params.ErrorRecord = $_
                }
                Log @params
            }
        }
        $task.GetLog() | Log
    }
    if ($needToReconnect) {
        break # Stop processing and reschedule
    }
    $mailboxCount++
}

# Call Cleanup() on all tasks
foreach ($task in $tasks) {
    try {
        $task.Cleanup()
    }
    catch {
        Log -Task 'Dispatcher:Cleanup' -Message "Cleanup for task $($task.GetType().Name) failed with error: $($_.ToString())"
    }
}

if ($queue.Count -gt 0) {
    $queue | Export-Csv "$PSScriptRoot\mailboxes.csv" -NoTypeInformation -Encoding UTF8
}
else {
    Remove-Item -Path "$PSScriptRoot\mailboxes.csv" -Force -ErrorAction SilentlyContinue
}

$batchTime = (Get-Date) - $startTime
Log -Task 'Dispatcher:End' -Message "Batch ended after $($batchTime.ToString('hh\:mm\:ss'))"

RescheduleTask