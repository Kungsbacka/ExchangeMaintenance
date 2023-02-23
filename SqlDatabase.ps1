class SqlDatabase
{
    hidden [System.Data.SqlClient.SqlConnection]$_connection

    hidden SqlDatabase([string]$connectionString) {
        $this._connection = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
        $this._connection.ConnectionString = $connectionString
        $this._connection.Open()
    }

    hidden [void]_executeOnly([string]$storedProcedure, [hashtable]$params) {
        $cmd = $this._getCmd($storedProcedure, $params)
        try {
            $null = $cmd.ExecuteNonQuery()
        }
        finally {
            $cmd.Dispose()
        }
    }

    hidden [System.Collections.ArrayList]_getData([string]$query, [hashtable]$params) {
        $result = New-Object -TypeName 'System.Collections.ArrayList'
        $cmd = $null
        $reader = $null
        try {
            $cmd = $this._getCmd($query, $params)
            $reader = $cmd.ExecuteReader()
            $firstRow = $true
            $ht = [ordered]@{}
            $values = [object[]]::new($reader.FieldCount)
            while ($reader.Read()) {
                if ($firstRow) {
                    for($i = 0; $i -lt $reader.FieldCount; $i++) {
                        $value = $reader.GetValue($i)
                        if ($value -is [System.DBNull]) {
                            $value = $null
                        }
                        $ht.Add($reader.GetName($i), $value)
                    }
                    $firstRow = $false
                    $null = $result.Add([pscustomobject]$ht)
                }
                else {
                    $null = $reader.GetValues($values)
                    for($i = 0; $i -lt $reader.FieldCount; $i++) {
                        $value = $values[$i]
                        if ($value -is [System.DBNull]) {
                            $value = $null
                        }
                        $ht.Item($i) = $value
                    }
                    $null = $result.Add([pscustomobject]$ht)
                }
            }
        }
        finally {
            if ($reader) {
                $reader.Dispose()
            }
            if ($cmd) {
                $cmd.Dispose()
            }
        }
        return $result
    }

    hidden [void]_saveDataTable([string]$destinationTable, [System.Data.DataTable]$dataTable) {
        $bulkCopy = New-Object -TypeName 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList @($this._connection)
        $bulkCopy.DestinationTableName = $destinationTable
        try {
            $bulkCopy.WriteToServer($dataTable)
        }
        finally {
            $bulkCopy.Dispose()
        }
    }

    hidden [System.Data.SqlClient.SqlCommand]_getCmd($query, [hashtable]$params) {
        if ($query -like 'dbo.*') {
            $queryType = 'StoredProcedure'
        }
        else {
            $queryType = 'Text'
        }
        $cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
        $cmd.Connection = $this._connection
        $cmd.CommandText = $query
        $cmd.CommandType = $queryType
        if ($null -ne $params) {
            foreach ($item in $params.GetEnumerator()) {
                $null = $cmd.Parameters.AddWithValue($item.Name, $item.Value)
            }
        }
        return $cmd
    }
}

class LogDb
{
    hidden static [SqlDatabase]$_instance

    hidden static [SqlDatabase]_getInstance() {
        if (-not [LogDb]::_instance) {
            [LogDb]::_instance = [SqlDatabase]::new($Script:Config.LogDbConnectionString)
        }
        return [LogDb]::_instance
    }

    static [void]ExecuteOnly([string]$storedProcedure, [hashtable]$params) {
        [LogDb]::_getInstance()._executeOnly($storedProcedure, $params)
    }

    static [void]GetData([string]$storedProcedure, [hashtable]$params) {
        [LogDb]::_getInstance()._getData($storedProcedure, $params)
    }

    static [void]SaveDataTable([string]$destinationTable, [System.Data.DataTable]$dataTable) {
        [LogDb]::_getInstance()._saveDataTable($destinationTable, $dataTable)
    }

    hidden LogDb() {
        throw 'Cannot create instance of static class.'
    }
}

class MetaDirectoryDb
{
    hidden static [SqlDatabase]$_instance

    hidden static [SqlDatabase]_getInstance() {
        if (-not [MetaDirectoryDb]::_instance) {
            [MetaDirectoryDb]::_instance = [SqlDatabase]::new($Script:Config.MetaDirectoryConnectionString)
        }
        return [MetaDirectoryDb]::_instance
    }

    static [void]ExecuteOnly([string]$storedProcedure, [hashtable]$params) {
        [MetaDirectoryDb]::_getInstance()._executeOnly($storedProcedure, $params)
    }

    static [System.Collections.ArrayList]GetData([string]$storedProcedure, [hashtable]$params) {
        return [MetaDirectoryDb]::_getInstance()._getData($storedProcedure, $params)
    }

    static [void]SaveDataTable([string]$destinationTable, [System.Data.DataTable]$dataTable) {
        [MetaDirectoryDb]::_getInstance()._saveDataTable($destinationTable, $dataTable)
    }

    hidden MetaDirectoryDb() {
        throw 'Cannot create instance of static class.'
    }
}
