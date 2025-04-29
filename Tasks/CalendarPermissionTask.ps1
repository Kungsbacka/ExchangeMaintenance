class CalendarPermissionTask : MailboxTask
{

    hidden [System.Collections.Generic.HashSet[string]]$_includedUsers

    hidden [void]_internalInitialize() {
        $exclude = @{}
        Get-ADGroupMember -Identity 'GA-FV-Alla' | ForEach-Object {
            $exclude.Add($_.SamAccountName, 1)
        }
        $query = "SELECT UserPrincipalName,SamAccountName FROM dbo.ExchangeMaintenanceUserView WHERE UserPrincipalName NOT LIKE '%@elev.kungsbacka.se'"
        $result = [MetaDirectoryDb]::GetData($query, $null)
        $this._includedUsers = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($item in $result) {
            if ($exclude.ContainsKey($item.SamAccountName)) {
                continue
            }
            $this._includedUsers.Add($item.UserPrincipalName)
        }
    }

    hidden [bool]_internalShouldProcess($mailbox) {
        return $this._includedUsers.Contains($mailbox.PrimarySmtpAddress)
    }

    hidden [void]_internalProcessMailbox($mailbox) {
        $params = @{
            Identity = $mailbox.PrimarySmtpAddress
            FolderScope = 'Calendar'
        }
        $calendars = [ExchangeOnline]::GetMailboxFolderStatistics($params) | Where-Object {$_.FolderType -eq 'Calendar'}
        foreach ($calendar in $calendars) {
            $calendarId = "$($mailbox.PrimarySmtpAddress):$($calendar.FolderPath.Replace('/', '\'))"
            $this._internalSetCalendarPermission($mailbox, $calendarId)
        }
    }

    hidden [void]_internalSetCalendarPermission($mailbox, $calendarId) {
        $params = @{
            Identity = $calendarId
        }
        $currentPermissions = [ExchangeOnline]::GetMailboxFolderPermission($params)
        $default = $currentPermissions | Where-Object {
            $_.User.DisplayName -eq 'Default'
        }
        $anonymous = $currentPermissions | Where-Object {
            $_.User.DisplayName -eq 'Anonymous'
        }
        $broken = $currentPermissions | Where-Object {
            $_.User.DisplayName -like '*S-1-5-21-*'
        }
        $remaining = $currentPermissions | Where-Object {
            @('Default', 'Anonymous') -notcontains $_.User.DisplayName -and $_.User.DisplayName -notlike '*S-1-5-21-*'
        }
        if ($null -eq $default)
        {
            $this._addLogItem('CalendarPermissionTask', $mailbox, "Missing 'Default' permission")
        }
        elseif ($default.AccessRights -ne 'Reviewer')
        {
            $this._addLogItem('CalendarPermissionTask', $mailbox, "Resetting 'Default' permission from [$($default.AccessRights -join ',')] to [Reviewer]")
            $params = @{
                Identity = $calendarId
                User = 'Default'
                AccessRights = 'Reviewer'
            }
            [ExchangeOnline]::SetMailboxFolderPermission($params)
        }
        if ($null -eq $anonymous)
        {
            $this._addLogItem('CalendarPermissionTask', $mailbox, "Adding missing 'Anonymous' permission as [None]")
            $params = @{
                Identity = $calendarId
                User = 'Anonymous'
                AccessRights = 'None'
            }
            [ExchangeOnline]::AddMailboxFolderPermission($params)
        }
        elseif ($anonymous.AccessRights -ne 'None')
        {
            $this._addLogItem('CalendarPermissionTask', $mailbox, "Resetting 'Anonymous' permission from [$($anonymous.AccessRights -join ',')] to [None]")
            $params = @{
                Identity = $calendarId
                User = 'Anonymous'
                AccessRights = 'None'
            }
            [ExchangeOnline]::SetMailboxFolderPermission($params)
        }
        if ($null -ne $broken)
        {
            foreach ($permission in $broken)
            {
                $this._addLogItem('CalendarPermissionTask', $mailbox, "Removing access right(s) [$($permission.AccessRights -join ',')] for deleted user $($permission.User.DisplayName)")
                $params = @{
                    Identity = $calendarId
                    User = $permission.User.DisplayName
                }
                [ExchangeOnline]::RemoveMailboxFolderPermission($params)
            }
        }
        if ($null -ne $remaining)
        {
            foreach ($permission in $remaining)
            {
                if ($permission.AccessRights -in @('Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author', 'NonEditingAuthor'))
                {
                    continue
                }
                if ($permission.AccessRights -in @('Reviewer', 'Contributor', 'AvailabilityOnly', 'LimitedDetails') -and $permission.User.UserType.Value -eq 'Internal')
                {
                    $this._addLogItem('CalendarPermissionTask', $mailbox, "Removing access right(s) [$($permission.AccessRights -join ',')] for user $($permission.User.DisplayName) <$($permission.User.RecipientPrincipal.PrimarySmtpAddress)>")
                    $params = @{
                        Identity = $calendarId
                        User = $permission.User.RecipientPrincipal.Guid.ToString()
                    }
                    [ExchangeOnline]::RemoveMailboxFolderPermission($params)
                }
                elseif ($permission.User.UserType.Value -eq 'Internal')
                {
                    $accessRights = $permission.AccessRights | ForEach-Object -Process { $_.ToString() }
                    if ($accessRights -notcontains 'ReadItems')
                    {
                        $accessRights += 'ReadItems'
                    }
                    if ($accessRights -notcontains 'FolderVisible')
                    {
                        $accessRights += 'FolderVisible'
                    }
                    if ($permission.AccessRights.Count -ne $accessRights.Count)
                    {
                        $this._addLogItem('CalendarPermissionTask', $mailbox, "Replacing access right(s) for user $($permission.User.DisplayName) <$($permission.User.RecipientPrincipal.PrimarySmtpAddress)> from [$($permission.AccessRights -join ',')] to [$($accessRights -join ',')]")
                        $params = @{
                            Identity = $calendarId
                            User = $permission.User.RecipientPrincipal.Guid.ToString()
                            AccessRights = $accessRights
                        }
                        [ExchangeOnline]::SetMailboxFolderPermission($params)
                    }
                }
            }
        }
    }
}
