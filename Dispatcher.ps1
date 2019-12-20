$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\SqlDatabase.ps1"
. "$PSScriptRoot\ExchangeOnline.ps1"
. "$PSScriptRoot\LogItem.ps1"
. "$PSScriptRoot\MailboxTask.ps1"
. "$PSScriptRoot\Logger.ps1"

$tasks = New-Object -TypeName 'System.Collections.ArrayList'
Get-ChildItem "$PSScriptRoot\Tasks\*.ps1" | ForEach-Object {
    . $_.FullName
    $task = New-Object -TypeName ([System.IO.Path]::GetFileNameWithoutExtension($_.Name))
    $null = $tasks.Add($task)
}

[ExchangeOnline]::Connect()
[ExchangeOnline]::Simulate = $Script:Config.Simulate

Log -Task 'Dispatcher:Start' -Message "Started new batch with $($Script:Config.BatchSize) mailboxes"

$startTime = Get-Date

$queue =  New-Object -TypeName 'System.Collections.Generic.Queue[object]'
try {
    Import-Csv "$PSScriptRoot\mailboxes.csv" -Encoding UTF8 | ForEach-Object {
        if (-not $_.ResourceType) {
            $_.ResourceType = $null
        }
        $queue.Enqueue($_)
    }
}
catch {
    # Failed to load file
}
if ($queue.Count -eq 0) {
    $mailboxes = [ExchangeOnline]::GetMailbox(@{ResultSize = 'Unlimited'})
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
    Log -Task 'Dispatcher:Initialize' -Message "Initializers for all tasks failed. Script is NOT rescheduled."
    exit 1
}

# Call ProcessMailbox on mailboxes in batch
$mailboxCount = 0
$connectionBroken = $false
while ($mailboxCount -lt $Script:Config.BatchSize -and $queue.Count -gt 0) {
    $mailbox = $queue.Dequeue()
    foreach ($task in $tasks) {
        try {
            $task.ProcessMailbox($mailbox)
        }
        catch {
            if ($_.CategoryInfo.Reason -eq 'ConnectionFailedTransientException') {
                $queue.Enqueue($mailbox) # if connection failed we want to process current mailbox on next run
                $connectionBroken = $true
                Log -Task 'Dispatcher:Process' -Message "Connection to Exchange Online broke with error: $($_.ToString())"
                break
            }
            Log -Task 'Dispatcher:Process' -Mailbox $mailbox -Message "Processing task $($task.GetType().Name) failed with error: $($_.ToString())" -StackTrace $_.Exception.StackTrace
        }
        $task.GetLog() | Log
    }
    if ($connectionBroken) {
        break # If connection is broken, we exit and reschedule
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

# Reschedule
$trigger = New-ScheduledTaskTrigger -At (Get-Date).Add($Script:Config.BatchDelay) -Once
$null = Set-ScheduledTask -TaskName $Script:Config.ScheduledTaskName -Trigger $trigger
