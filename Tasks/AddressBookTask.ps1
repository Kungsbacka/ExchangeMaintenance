class AddressBookTask : MailboxTask
{
    hidden [System.Collections.Generic.Dictionary[string,object]]$_includedUsers

    hidden [void]_internalInitialize() {
        $this._includedUsers = New-Object -TypeName 'System.Collections.Generic.Dictionary[string,object]'
        $query = 'SELECT UserPrincipalName,Department,PhysicalDeliveryOfficeName FROM dbo.ExchangeMaintenanceUserView'
        $result = [MetaDirectoryDb]::GetData($query, $null)
        foreach ($item in $result) {
            if ($this._includedUsers.ContainsKey($item.UserPrincipalName)) {
                $this._addLogItem('AddressBookTask:Initialize', "Duplicate UserPrincipalName '$($item.UserPrincipalName)'")
            }
            else {
                $this._includedUsers.Add($item.UserPrincipalName, $item)
            }
        }
    }

    hidden [bool]_internalShouldProcess($mailbox) {
        return $this._includedUsers.ContainsKey($mailbox.PrimarySmtpAddress)
    }

    hidden [void]_internalProcessMailbox($mailbox) {
        $abp = 'ABP-ADM'
        if ($mailbox.PrimarySmtpAddress -like '*@elev.kungsbacka.se') {
            $abp = 'ABP-Skola'
        }
        elseif ($this._isSkolpersonal($mailbox)) {
            $abp = $null
        }
        if (-not $this._abpEqual($mailbox.AddressBookPolicy, $abp)) {
            $this._addLogItem('AddressBookTask', $mailbox, "Change ABP from '$($mailbox.AddressBookPolicy)' to '$abp'")
            $params = @{
                Identity = $mailbox.Identity
                AddressBookPolicy = $abp
            }
            [ExchangeOnline]::SetMailbox($params)
        }
    }

    hidden [bool]_abpEqual($abp1, $abp2) {
        return ([string]::IsNullOrEmpty($abp1) -and [string]::IsNullOrEmpty($abp2)) -or $abp1 -eq $abp2
    }

    hidden [bool]_isSkolpersonal($mailbox) {
        $item = $null
        if ($this._includedUsers.TryGetValue($mailbox.PrimarySmtpAddress, [ref]$item)) {
            return $item.Department -eq 'Förskola & Grundskola' -or $item.Department -eq 'Gymnasium & Arbetsmarknad' -or $item.PhysicalDeliveryOfficeName -eq 'Lärare Kulturskolan'
        }
        return $false
    }
}
