class InventoryTask : MailboxTask
{
    hidden [System.Data.DataTable]$_dataTable
    hidden [System.Data.SqlClient.SqlConnection] $_sqlConnection

    hidden [void]_internalInitialize() {
        $this._dataTable = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('MailboxInventory')
        $null = $this._dataTable.Columns.Add('AzureAdGuid', 'guid')
        $null = $this._dataTable.Columns.Add('ExchangeGuid', 'guid')
        $null = $this._dataTable.Columns.Add('PrimarySmtpAddress', 'string')
        $null = $this._dataTable.Columns.Add('IsShared', 'boolean')
        $null = $this._dataTable.Columns.Add('IsResource', 'boolean')
        $null = $this._dataTable.Columns.Add('IsForwarded', 'boolean')
        $null = $this._dataTable.Columns.Add('ResourceType', 'string')
        $null = $this._dataTable.Columns.Add('ItemCount', 'int')
        $null = $this._dataTable.Columns.Add('DeletedItemCount', 'int')
        $null = $this._dataTable.Columns.Add('TotalItemSize', 'long')
        $null = $this._dataTable.Columns.Add('TotalDeletedItemSize', 'long')
    }

    hidden [void]_internalProcessMailbox($mailbox) {
        $params = @{
            Identity = $mailbox.Identity
        }
        $stats = [ExchangeOnline]::GetMailboxStatistics($params)
        $row = $this._dataTable.NewRow()
        $row['AzureAdGuid'] = $mailbox.ExternalDirectoryObjectId
        $row['ExchangeGuid'] = $mailbox.Guid
        $row['PrimarySmtpAddress'] = $mailbox.PrimarySmtpAddress
        $row['IsShared'] = $mailbox.IsShared
        $row['IsResource'] = $mailbox.IsResource
        $row['IsForwarded'] = ($null -ne $mailbox.ForwardingAddress -or $null -ne $mailbox.ForwardingSmtpAddress)
        $row['ResourceType'] = $mailbox.ResourceType
        $row['ItemCount'] = $stats.ItemCount
        $row['DeletedItemCount'] = $stats.DeletedItemCount
        $row['TotalItemSize'] = $this._parseSize($stats.TotalItemSize)
        $row['TotalDeletedItemSize'] = $this._parseSize($stats.TotalDeletedItemSize)
        $this._dataTable.Rows.Add($row)
    }

    hidden [void]_internalCleanup() {
        [MetaDirectoryDb]::ExecuteOnly('dbo.spMailboxInventoryPrepareStage', $null)
        [MetaDirectoryDb]::SaveDataTable('dbo.MailboxInventory_stage', $this._dataTable)
        [MetaDirectoryDb]::ExecuteOnly('dbo.spMailboxInventoryUpsert', $null)
    }

    hidden [long]_parseSize($bqs) {
        $str = $bqs.ToString()
        if ($str -eq 'Unlimited') {
            return 0
        }
        $tmp = $str.Split('(')
        if ($tmp.Count -lt 2) {
            return 0
        }
        $tmp = $tmp[1].Split()
        if ($tmp.Count -lt 1) {
            return 0
        }
        return [long]$tmp[0]
    }
}
