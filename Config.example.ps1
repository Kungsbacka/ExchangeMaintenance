$Script:Config = @{
    ExchangeUser = '<Tenant user with Exchange administrative permissions>'
    ExchangePassword = '<Encrypted password>'
    BatchSize = 500
    BatchDelay = ([TimeSpan]::FromMinutes(10))
    ScheduledTaskName = 'ExchangeMaintenance'
    LogPath = $PSScriptRoot
    LogToDatabase = $true
    MetaDirectoryConnectionString = 'Server=<Database server>;Database=<Meta directory database>;Trusted_Connection=true'
    LogDbConnectionString = 'Server=<Database server>;Database=<Log database>;Trusted_Connection=true'
    Simulate = $false # If set to true, all Exchange cmdlets that make changes to a mailbox, are ignored
}