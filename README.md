# Exchange Maintenance

## Introduction

Exchange Maintenance is a collection of mailbox maintenance tasks that are performed on all, or a subset of, mailboxes in an Exchange Online tenant.

Mailboxes are processed in batches and after each batch the remaining mailboxes that have not been processed are written to a CSV file. The script then reschedule itself and continues processing at a later time. When all mailboxes are processed, the script fetches all mailboxes again an starts over.

## Tasks

* CalendarPermissionTask: Restores default permissions and removes more restrictive permissions to make sure all users are at least "Reviewer" on the main calendar.
* AddressBookTask: Sets correct Address Book Policy (ABP) on all mailboxes.
* InventoryTask: Takes inventory of all mailboxes in the tenant.

## Deploy

The script depends on the [Exchange Online PowerShell V2 module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps) and uses a service principal with certificate authentication to connect. Instructions on how to create a service principal and assign permissions can be found [here](https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps).

Rename Config.example.ps1 to Config.ps1 and update configuration settings. AppCertificatePassword must be encrypted using DPAPI with the service account that is going to run the script (preferably a gMSA). This can be accomplished by following the steps below:

1. Start PowerShell as the service account. If the account is a gMSA, you can do this with PsExec (<https://docs.microsoft.com/en-us/sysinternals/downloads/psexec>).

    `psexec -i -u <gMSA> powershell.exe`

2. Run the following one-liner and enter the certificate password in the credentials dialog (username doesn't matter, but cannot be empty).

    `(Get-Credential -UserName '(Not used)' -Message 'Exchange password').Password | ConvertFrom-SecureString | Set-Clipboard`

3. Paste the encrypted password into the configuration file (AppCertificatePassword).

The service account must have write permission to script root to be able to save remaining mailboxes as a CSV file (mailboxes.csv). It must also have appropriate permissions on meta directory and log databases.

## Scheduled task

Create a scheduled task that invokes the dispatcher. Notice that the task is configured to run only once. The dispatcher will automatically reschedule the task when it has completed a batch.

If the dispatcher crashes before it had the opportunity to reschedule, the task must be started manually after the cause of the crash has been investigated.

If you change the task name, the configuration must be updated with the new name (ScheduledTaskName).

```powershell
Unregister-ScheduledTask -TaskName 'ExchangeMaintenance' -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask `
    -TaskName 'ExchangeMaintenance' `
    -TaskPath '\' `
    -Description 'Exchange Online mailbox maintenance' `
    -Principal (New-ScheduledTaskPrincipal -UserId '<gMSA>' -LogonType Password) `
        -Trigger (New-ScheduledTaskTrigger -At (Get-Date).AddHours(1) -Once) `
    -Action (New-ScheduledTaskAction `
        -Execute 'powershell.exe' `
        -Argument '-Command "C:\ExchangeMaintenance\Dispatcher.ps1"' `
        -WorkingDirectory 'C:\ExchangeMaintenance') `
    -Settings (New-ScheduledTaskSettingsSet -StartWhenAvailable)
```

## Database

DDL for log table, stored procedures, and inventory tables.

### Inventory

```sql
CREATE TABLE [dbo].[MailboxInventory_stage](
    [AzureAdGuid] [uniqueidentifier] NULL,
    [ExchangeGuid] [uniqueidentifier] NULL,
    [PrimarySmtpAddress] [nvarchar](256) NULL,
    [IsShared] [tinyint] NULL,
    [IsResource] [tinyint] NULL,
    [IsForwarded] [tinyint] NULL,
    [ResourceType] [nvarchar](50) NULL,
    [ItemCount] [int] NULL,
    [DeletedItemCount] [int] NULL,
    [TotalItemSize] [bigint] NULL,
    [TotalDeletedItemSize] [bigint] NULL
) ON [PRIMARY]

CREATE TABLE [dbo].[MailboxInventory](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [InventoryTime] [datetime] NOT NULL,
    [AzureAdGuid] [uniqueidentifier] NULL,
    [ExchangeGuid] [uniqueidentifier] NULL,
    [PrimarySmtpAddress] [nvarchar](256) NULL,
    [IsShared] [tinyint] NULL,
    [IsResource] [tinyint] NULL,
    [IsForwarded] [tinyint] NULL,
    [ResourceType] [nvarchar](50) NULL,
    [ItemCount] [int] NULL,
    [DeletedItemCount] [int] NULL,
    [TotalItemSize] [bigint] NULL,
    [TotalDeletedItemSize] [bigint] NULL,
    CONSTRAINT [PK_MailboxInventory] PRIMARY KEY CLUSTERED ([Id])
)

CREATE PROCEDURE [dbo].[spMailboxInventoryUpsert]
AS
BEGIN
    SET NOCOUNT ON;

    MERGE
        dbo.MailboxInventory inv
    USING (
        SELECT
            AzureAdGuid,
            ExchangeGuid,
            PrimarySmtpAddress,
            IsShared,
            IsResource,
            IsForwarded,
            ResourceType,
            ItemCount,
            DeletedItemCount,
            TotalItemSize,
            TotalDeletedItemSize
        FROM
            dbo.MailboxInventory_stage
    ) AS stage
    ON
        inv.ExchangeGuid = stage.ExchangeGuid
    WHEN MATCHED THEN
        UPDATE SET
        InventoryTime = GETDATE(),
        PrimarySmtpAddress = stage.PrimarySmtpAddress,
        IsShared = stage.IsShared,
        IsResource = stage.IsResource,
        IsForwarded = stage.IsForwarded,
        ResourceType = stage.ResourceType,
        ItemCount = stage.ItemCount,
        DeletedItemCount = stage.DeletedItemCount,
        TotalItemSize = stage.TotalItemSize,
        TotalDeletedItemSize = stage.TotalDeletedItemSize
    WHEN NOT MATCHED THEN
        INSERT
            (InventoryTime, AzureAdGuid, ExchangeGuid, PrimarySmtpAddress, IsShared, IsResource, IsForwarded, ResourceType, ItemCount, DeletedItemCount, TotalItemSize, TotalDeletedItemSize)
        VALUES
            (GETDATE(),     AzureAdGuid, ExchangeGuid, PrimarySmtpAddress, IsShared, IsResource, IsForwarded, ResourceType, ItemCount, DeletedItemCount, TotalItemSize, TotalDeletedItemSize);

    -- Cleanup stale records
    DELETE FROM dbo.MailboxInventory WHERE InventoryTime < DATEADD(MONTH, -1, GETDATE());
END

CREATE PROCEDURE [dbo].[spMailboxInventoryPrepareStage]
AS
BEGIN
    SET NOCOUNT ON;

    TRUNCATE TABLE dbo.MailboxInventory_stage;
END
```

### Logging

```sql
CREATE TABLE [dbo].[ExchangeMaintenanceLog](
    [id] [int] IDENTITY(1,1) NOT NULL,
    [logTime] [datetime] NOT NULL,
    [task] [nvarchar](50) NOT NULL,
    [mailbox] [nvarchar](100) NULL,
    [simulation] [bit] NOT NULL,
    [result] [nvarchar](1000) NOT NULL,
    CONSTRAINT [PK_ExchangeMaintenanceLog] PRIMARY KEY CLUSTERED ([id])
)

CREATE PROCEDURE [dbo].[spNewExchangeMaintenanceLogEntry]
    @logTime datetime,
    @task nvarchar(50),
    @mailbox nvarchar(100) = NULL,
    @simulation bit = 0,
    @result nvarchar(1000)
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO dbo.ExchangeMaintenanceLog (logTime, task, mailbox, simulation, result)
    VALUES (@logTime, @task, @mailbox, @simulation, @result);
END
```
