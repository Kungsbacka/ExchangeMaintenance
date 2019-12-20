class LogItem {
    [string]$Task
    [object]$Mailbox
    [string]$Message
    [DateTime]$Time
    [bool]$Simulation

    LogItem([string]$task, [string]$message) {
        $this.Task = $task
        $this.Message = $message
        $this.Time = Get-Date
        $this.Simulation = [ExchangeOnline]::Simulate
    }

    LogItem([string]$task, [object]$mailbox, [string]$message) {
        $this.Task = $task
        $this.Mailbox = $mailbox
        $this.Message = $message
        $this.Time = Get-Date
        $this.Simulation = [ExchangeOnline]::Simulate
    }

    [string]ToString() {
        $target = ''
        if ($this.Mailbox) {
            $target = $this.Mailbox.PrimarySmtpAddress
        }
        $sb = New-Object -TypeName 'System.Text.StringBuilder'
        $null = $sb.Append(($this.Time.ToString('yyyy-MM-dd HH:mm:ss')))
        $null = $sb.Append(' [')
        $null = $sb.Append($this.Task)
        if ($this.Simulation) {
            $null = $sb.Append('-SIMULATION')
        }
        $null = $sb.Append(']')
        $null = $sb.Append(' ', [Math]::Max(0, 70 - $sb.Length))
        $null = $sb.Append('<')
        $null = $sb.Append($target)
        $null = $sb.Append('>')
        $null = $sb.Append(' ', [Math]::Max(0, 155 - $sb.Length))
        $null = $sb.Append($this.Message)
        return $sb.ToString()
    }
}
