﻿function Log
{
    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ParameterSetName='LogItem')]
        [LogItem]$LogItem,
        [Parameter(Mandatory=$true,ParameterSetName='Message')]
        [string]$Task,
        [Parameter(Mandatory=$true,ParameterSetName='Message')]
        [string]$Message,
        [Parameter(Mandatory=$false,ParameterSetName='Message')]
        $Mailbox,
        [Parameter(Mandatory=$false,ParameterSetName='Message')]
        [object]$ErrorRecord
    )
    process {
        if ($PSCmdlet.ParameterSetName -eq 'Message') {
            $LogItem = New-Object 'LogItem' -ArgumentList @($Task, $Mailbox, ($Message -replace "`r`n", ' '))
        }
        if ($LogItem) {
            if ($Script:Config.LogPath) {
                $filePath = Join-Path -Path $Script:Config.LogPath -ChildPath 'log.txt'
                $LogItem.ToString() | Out-File -FilePath $filePath -Encoding UTF8 -Append
            }
            if ($Script:Config.LogToDatabase) {
                $params = @{
                    logTime = $LogItem.Time
                    task = $LogItem.Task
                    simulation = $LogItem.Simulation
                    result = $LogItem.Message
                }
                if ($LogItem.Mailbox) {
                    $params.mailbox = $LogItem.Mailbox.PrimarySmtpAddress
                }
                [LogDb]::ExecuteOnly('dbo.spNewExchangeMaintenanceLogEntry', $params)
            }
        }
        if ($ErrorRecord) {
            if ($Script:Config.LogPath) {
                $fileName = 'error_' + (Get-Date -Format 'yyyyMMdd_HHmmss_mmm') + '.xml'
                $filePath = Join-Path -Path $Script:Config.LogPath -ChildPath $fileName
                $ErrorRecord | Export-Clixml -Path $filePath -Encoding UTF8
            }
        }
    }
}
