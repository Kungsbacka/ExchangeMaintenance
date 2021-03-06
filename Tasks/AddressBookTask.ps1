﻿class AddressBookTask : MailboxTask
{
    hidden [System.Collections.Generic.Dictionary[string,object]]$_includedUsers
    hidden [System.Collections.Generic.Dictionary[string,string]]$_abpExceptions

    hidden [void]_internalInitialize() {
        $this._includedUsers = New-Object -TypeName 'System.Collections.Generic.Dictionary[string,object]'
        $this._abpExceptions = New-Object -TypeName 'System.Collections.Generic.Dictionary[string,string]'
        $query = 'SELECT UserPrincipalName,Department,PhysicalDeliveryOfficeName,DistinguishedName FROM dbo.ExchangeMaintenanceUserView'
        $result = [MetaDirectoryDb]::GetData($query, $null)
        foreach ($item in $result) {
            if ($this._includedUsers.ContainsKey($item.UserPrincipalName)) {
                $this._addLogItem('AddressBookTask:Initialize', "Duplicate UserPrincipalName '$($item.UserPrincipalName)'")
            }
            else {
                $this._includedUsers.Add($item.UserPrincipalName, $item)
            }
        }
        foreach ($member in (Get-ADGroup 'G-Exchange-undantag-ABP-Skola' -Property 'Members').Members) {
            $this._abpExceptions.Add($member, 'ABP-Skola')
        }
        foreach ($member in (Get-ADGroup 'G-Exchange-undantag-ABP-ADM' -Property 'Members').Members) {
            $this._abpExceptions.Add($member, 'ABP-ADM')
        }
        foreach ($member in (Get-ADGroup 'G-Exchange-undantag-ABP-None' -Property 'Members').Members) {
            $this._abpExceptions.Add($member, $null)
        }
    }

    hidden [bool]_internalShouldProcess($mailbox) {
        return $this._includedUsers.ContainsKey($mailbox.PrimarySmtpAddress)
    }

    hidden [void]_internalProcessMailbox($mailbox) {
        $abp = $this._getException($mailbox)
        if ($abp -eq 'No exception') {
            $abp = 'ABP-ADM'
            if ($mailbox.PrimarySmtpAddress -like '*@elev.kungsbacka.se') {
                $abp = 'ABP-Skola'
            }
            elseif ($this._isSkolpersonal($mailbox)) {
                $abp = $null
            }
        }
        if ($abp -eq '') {
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
            return $item.Department -eq 'Förskola & Grundskola' -or $item.Department -eq 'Gymnasium & Arbetsmarknad' -or $item.PhysicalDeliveryOfficeName -like 'Lärare Kulturskolan*'
        }
        return $false
    }

    hidden [string]_getException($mailbox) {
        $item = $null
        if ($this._includedUsers.TryGetValue($mailbox.PrimarySmtpAddress, [ref]$item)) {
            $abp = $null
            if ($this._abpExceptions.TryGetValue($item.DistinguishedName, [ref]$abp)) {
                return $abp
            }
        }
        return 'No exception'
    }
}
