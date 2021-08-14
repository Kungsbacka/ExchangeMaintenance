class ExchangeOnline
{
    hidden static [bool]$_isConnected

    static [bool]$Simulate = $false

    static [void]Connect() {
        if ([ExchangeOnline]::_isConnected) {
            return
        }
        $params = @{
            CertificateFilePath = $Script:Config.AppCertificatePath
            CertificatePassword = ($Script:Config.AppCertificatePassword | ConvertTo-SecureString)
            AppId = $Script:Config.AppId
            Organization = $Script:Config.Organization
            CommandName = @(
                'Get-EXOMailbox'
                'Get-EXOMailboxFolderPermission'
                'Get-EXOMailboxFolderStatistics'
                'Get-EXOMailboxStatistics'
                'Remove-MailboxFolderPermission'
                'Set-Mailbox'
                'Set-MailboxFolderPermission'
            )
            PageSize = 5000
            ShowBanner = $false
            ShowProgress = $false
        }
        Connect-ExchangeOnline @params
        [ExchangeOnline]::_isConnected = $true
    }

    hidden static [void]_internalReconnect() {
        [ExchangeOnline]::_isConnected = $false
        [ExchangeOnline]::Connect()
    }

    hidden static [object]_internalExecuteCommand([string]$command, [hashtable]$params) {
        if (-not [ExchangeOnline]::_isConnected) {
            throw "Call Connect() before calling other methods"
        }
        $params.ErrorAction = 'Stop'
        if ($command -like 'Set-*' -or $command -like 'Remove-*') {
            if ([ExchangeOnline]::Simulate -or $params.Simulate) {
                return $null
            }
            $params.Confirm = $false
        }
        $result = & $command @params
        if ($result) {
            return $result
        }
        return $null
    }

    static [object]GetMailbox([hashtable]$params) {
        return [ExchangeOnline]::_internalExecuteCommand('Get-EXOMailbox', $params)
    }

    static [object]GetMailboxFolderPermission([hashtable]$params) {
        return [ExchangeOnline]::_internalExecuteCommand('Get-EXOMailboxFolderPermission', $params)
    }

    static [object]GetMailboxFolderStatistics([hashtable]$params) {
        return [ExchangeOnline]::_internalExecuteCommand('Get-EXOMailboxFolderStatistics', $params)
    }

    static [object]GetMailboxStatistics([hashtable]$params) {
        return [ExchangeOnline]::_internalExecuteCommand('Get-EXOMailboxStatistics', $params)
    }

    static [void]RemoveMailboxFolderPermission([hashtable]$params) {
        [ExchangeOnline]::_internalExecuteCommand('Remove-MailboxFolderPermission', $params)
    }

    static [void]SetMailbox([hashtable]$params) {
        [ExchangeOnline]::_internalExecuteCommand('Set-Mailbox', $params)
    }

    static [void]SetMailboxFolderPermission([hashtable]$params) {
        [ExchangeOnline]::_internalExecuteCommand('Set-MailboxFolderPermission', $params)
    }
}
