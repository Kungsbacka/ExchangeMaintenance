class MailboxTask
{
    hidden [bool]$_initialized
    hidden [System.Collections.Generic.List[psobject]]$_logItems

    MailboxTask() {
        $this._initialized = $false
        $this._logItems = New-Object -TypeName 'System.Collections.Generic.List[psobject]'
    }

    [void]Initialize() {
        if ($this._initialized) {
            return
        }
        $this._internalInitialize()
        $this._initialized = $true
    }

    [void]ProcessMailbox($mailbox) {
        if (-not $this._initialized) {
            throw 'Initialize must be called before the first call to ProcessMailbox'
        }
        if ($this._internalShouldProcess($mailbox)) {
            $this._internalProcessMailbox($mailbox)
        }
    }

    [void]Cleanup() {
        if (-not $this._initialized) {
            throw 'Initialize must be called before calling Cleanup'
        }
        $this._internalCleanup()
    }

    [LogItem[]]GetLog() {
        $logItems = $this._logItems.ToArray()
        $this._logItems.Clear()
        return $logItems
    }

    hidden [void]_addLogItem([LogItem]$logItem) {
        $this._logItems.Add($logItem)
    }

    hidden [void]_addLogItem([string]$task, [string]$message) {
        $this._addLogItem([LogItem]::new($task, $message))
    }

    hidden [void]_addLogItem([string]$task, [object]$mailbox, [string]$message) {
        $this._addLogItem([LogItem]::new($task, $mailbox, $message))
    }

    hidden [void]_internalInitialize() {
        # Override if initialization is needed before processing.
        # Called once before any processing starts.
    }

    hidden [bool]_internalShouldProcess($mailbox) {
        # Override to control which mailboxes should be processed.
        # Called before each mailbox is processed.
        return $true
    }

    hidden [psobject[]]_internalProcessMailbox($mailbox) {
        # This method does the actual processing.
        # Called once for each mailbox in the batch.
        throw 'Not implemented'
    }

    hidden [void]_internalCleanup() {
        # Override if cleanups is needed after processing.
        # Called once after all mailboxes in a batch are processed.
    }
}

<#
class TemplateTask {

    hidden [void]_internalInitialize() {

    }

    hidden [bool]_internalShouldProcess($identity) {

    }

    hidden [void]_internalProcessMailbox($mailbox) {

    }

    hidden [void]_internalCleanup() {

    }
}
#>
