$Script:Config = @{
    AppCertificatePath = '<Path to app registration certificate>'
    AppCertificatePassword = '<Encrypted password for certificate>'
    AppId = '<App registration ID>'
    Organization = 'company.onmicrosoft.com'
    BatchSize = 500
    BatchDelay = ([TimeSpan]::FromMinutes(10))
    ScheduledTaskName = 'ExchangeMaintenance'
    LogPath = $PSScriptRoot
    LogToDatabase = $true
    MetaDirectoryConnectionString = 'Server=<Database server>;Database=<Meta directory database>;Trusted_Connection=true'
    LogDbConnectionString = 'Server=<Database server>;Database=<Log database>;Trusted_Connection=true'
    Simulate = $false # If set to true, all Exchange cmdlets that make changes to a mailbox, are ignored
}