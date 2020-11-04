class RetentionPolicyTask : MailboxTask
{
    hidden [System.Collections.Generic.HashSet[string]]$_includedUsers

    hidden [void]_internalInitialize() {
        $this._includedUsers = New-Object -TypeName 'System.Collections.Generic.HashSet[string]'
        $query = 'SELECT UserPrincipalName FROM dbo.ExchangeMaintenanceUserView'
        $result = [MetaDirectoryDb]::GetData($query, $null)
        foreach ($item in $result) {
            if (-not $this._includedUsers.Add($item.UserPrincipalName)) {
                $this._addLogItem('RetentionPolicyTask:Initialize', "Duplicate UserPrincipalName '$($item.UserPrincipalName)'")
            }
        }
    }

    hidden [bool]_internalShouldProcess($mailbox) {
        return $this._includedUsers.Contains($mailbox.PrimarySmtpAddress)
    }

    hidden [void]_internalProcessMailbox($mailbox) {
        if ($mailbox.PrimarySmtpAddress -like '*@elev.kungsbacka.se') {
            $policy = 'Elev Retention Policy'
        }
        else {
            $policy = 'Personal Retention Policy'
        }
        if (-not $this._policyEqual($mailbox.RetentionPolicy, $policy)) {
            $this._addLogItem('RetentionPolicyTask', $mailbox, "Change Retention Policy from '$($mailbox.RetentionPolicy)' to '$policy'")
            $params = @{
                Identity = $mailbox.Identity
                RetentionPolicy = $policy
            }
            [ExchangeOnline]::SetMailbox($params)
        }
    }

    hidden [bool]_policyEqual($policy1, $policy2) {
        return ([string]::IsNullOrEmpty($policy1) -and [string]::IsNullOrEmpty($policy2)) -or $policy1 -eq $policy2
    }
}
