class ExchangeOnline
{
    hidden static [bool]$_isConnected
    hidden static [object]$_session
    hidden static [object]$_module

    static [bool]$Simulate = $false


    static [void]Connect() {
        if ([ExchangeOnline]::_isConnected) {
            return
        }
        $credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
            $Script:Config.ExchangeUser
            $Script:Config.ExchangePassword | ConvertTo-SecureString
        )
        $params = @{
            Name = 'EXO'
            ConfigurationName = 'Microsoft.Exchange'
            ConnectionUri = 'https://outlook.office365.com/powershell'
            Credential = $credential
            Authentication = 'Basic'
            AllowRedirection = $true
        }
        [ExchangeOnline]::_session = New-PSSession @params
        $params = @{
            Session = [ExchangeOnline]::_session
            AllowClobber = $true
            CommandName = @(
                'Get-Mailbox'
                'Get-MailboxFolderPermission'
                'Get-MailboxFolderStatistics'
                'Get-MailboxStatistics'
                'Remove-MailboxFolderPermission'
                'Set-Mailbox'
                'Set-MailboxFolderPermission'
            )
        }
        [ExchangeOnline]::_module = Import-PSSession @params
        [ExchangeOnline]::_isConnected = $true
    }

    hidden static [void]_internalReconnect() {
        [ExchangeOnline]::_isConnected = $false
        if ([ExchangeOnline]::_module) {
            Remove-Module -Name ([ExchangeOnline]::_module.Name)
        }
        if ([ExchangeOnline]::_session) {
            Remove-PSSession -Session ([ExchangeOnline]::_session)
        }
        [ExchangeOnline]::Connect()
    }

    hidden static [object]_internalExecuteCommand([string]$command, [hashtable]$params) {
        if (-not [ExchangeOnline]::_isConnected) {
            throw "Call Connect() before executing commands."
        }
        $params.ErrorAction = 'Stop'
        if ($command -like 'Set-*' -or $command -like 'Remove-*') {
            if ([ExchangeOnline]::Simulate -or $params.Simulate) {
                return $null
            }
            $params.Confirm = $false
        }
        try {
            return (& $command @params)
        }
        catch {
            if ($_.CategoryInfo.Reason -eq 'ConnectionFailedTransientException') {
                [ExchangeOnline]::_internalReconnect()
                return (& $command @params)
            }
            else {
                throw
            }
        }
    }

    static [object]GetMailbox([hashtable]$params) {
        return [ExchangeOnline]::_internalExecuteCommand('Get-Mailbox', $params)
    }

    static [object]GetMailboxFolderPermission([hashtable]$params) {
        return [ExchangeOnline]::_internalExecuteCommand('Get-MailboxFolderPermission', $params)
    }

    static [object]GetMailboxFolderStatistics([hashtable]$params) {
        return [ExchangeOnline]::_internalExecuteCommand('Get-MailboxFolderStatistics', $params)
    }

    static [object]GetMailboxStatistics([hashtable]$params) {
        return [ExchangeOnline]::_internalExecuteCommand('Get-MailboxStatistics', $params)
    }

    static [void]RemoveMailboxFolderPermissions([hashtable]$params) {
        [ExchangeOnline]::_internalExecuteCommand('Remove-MailboxFolderPermission', $params)
    }

    static [void]SetMailbox([hashtable]$params) {
        [ExchangeOnline]::_internalExecuteCommand('Set-Mailbox', $params)
    }

    static [void]SetMailboxFolderPermission([hashtable]$params) {
        [ExchangeOnline]::_internalExecuteCommand('Set-MailboxFolderPermission', $params)
    }
}
