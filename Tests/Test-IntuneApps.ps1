[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$checks = 0
function Assert-AppTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
    Write-Host "PASS: $Message"
}
foreach ($path in @('Features\Get-IntuneApps.ps1', 'Features\Invoke-IntuneAppAssignment.ps1', 'Features\Set-IntuneAppAvailableUninstall.ps1', 'Shared\IntuneAppHelpers.ps1')) {
    $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $root $path), [ref]$null, [ref]$errors)
    Assert-AppTest ($errors.Count -eq 0) "App feature syntax: $path"
}
$tenant = '11111111-1111-1111-1111-111111111111'
$appId = '22222222-2222-2222-2222-222222222222'
$groupId = '33333333-3333-3333-3333-333333333333'
$assignmentId = '44444444-4444-4444-4444-444444444444'
$filterId = '55555555-5555-5555-5555-555555555555'
$otherGroupId = '66666666-6666-6666-6666-666666666666'
$mock = @{
    Tenant = $tenant; Scopes = @('DeviceManagementApps.ReadWrite.All', 'Group.Read.All')
    Apps = @(
        @{ id = $appId; displayName = 'Windows [Pilot]'; publisher = 'Publisher'; '@odata.type' = '#microsoft.graph.win32LobApp'; isAssigned = $true; allowAvailableUninstall = $false },
        @{ id = '77777777-7777-7777-7777-777777777777'; displayName = 'Store'; '@odata.type' = '#microsoft.graph.winGetApp' },
        @{ id = '88888888-8888-8888-8888-888888888888'; displayName = 'iPad'; '@odata.type' = '#microsoft.graph.iosVppApp' },
        @{ id = '99999999-9999-9999-9999-999999999999'; displayName = 'Android'; '@odata.type' = '#microsoft.graph.androidManagedStoreApp' },
        @{ id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; displayName = 'Mac'; '@odata.type' = '#microsoft.graph.macOSDmgApp' },
        @{ id = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; displayName = 'Web'; '@odata.type' = '#microsoft.graph.webApp' },
        @{ id = 'cccccccc-cccc-cccc-cccc-cccccccccccc'; displayName = 'Future'; '@odata.type' = '#microsoft.graph.futureApp' }
    )
    Assignments = [System.Collections.Generic.List[object]]::new()
    OtherAssignments = [System.Collections.Generic.List[object]]::new()
    Writes = [System.Collections.Generic.List[object]]::new()
    Reads = [System.Collections.Generic.List[string]]::new()
    FailPage = $false; FailRead = $false; FailWrite = $false; Ambiguous = $false; MissingGroup = $false
    Retry = $false; Sleeps = 0; FailGroup = ''
    IgnoreSettingWrite = $false; FailSettingVerification = $false; SettingWritten = $false
    FailSettingApp = ''; RetrySettingWrite = $false
    WrongSettingIdentity = $false
}
function Get-MgContext { [PSCustomObject]@{ TenantId = $mock.Tenant; Scopes = $mock.Scopes; Account = 'offline@example.test' } }
function Connect-MgGraph { throw 'Unexpected authentication in offline app tests.' }
function Start-Sleep { param($Seconds) $mock.Sleeps++ }
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param($Method, $Uri, $Body, $ContentType, $Headers)
    $uriText = [string]$Uri
    $collection = $mock.Assignments
    if ($uriText -like "*mobileApps/$($mock.Apps[1].id)/*") { $collection = $mock.OtherAssignments }
    if ($Method -in @('POST', 'PATCH', 'DELETE')) {
        $mock.Writes.Add([PSCustomObject]@{ Method = $Method; Uri = $uriText; Body = $Body })
        if ($Method -eq 'PATCH' -and $uriText -match '/mobileApps/([^/?]+)$') {
            if ($mock.FailSettingApp -eq $Matches[1]) { throw '403 Forbidden writing Win32 setting' }
            if ($mock.RetrySettingWrite) { $mock.RetrySettingWrite = $false; throw '503 ServiceUnavailable' }
            $settingApp = @($mock.Apps | Where-Object id -eq $Matches[1])[0]
            if (-not $mock.IgnoreSettingWrite) { $settingApp.allowAvailableUninstall = ($Body | ConvertFrom-Json).allowAvailableUninstall }
            $mock.SettingWritten = $true
            return
        }
        if ($mock.FailWrite) { throw '403 Forbidden writing app assignment' }
        if ($Body -and $mock.FailGroup -and ($Body | ConvertFrom-Json).target.groupId -eq $mock.FailGroup) { throw '400 Incompatible group assignment' }
        if ($Method -eq 'POST') {
            $data = $Body | ConvertFrom-Json
            $data | Add-Member NoteProperty id ([guid]::NewGuid().ToString())
            $data | Add-Member NoteProperty source 'direct'
            $collection.Add($data)
            if ($mock.Ambiguous) { $mock.Ambiguous = $false; throw '503 ServiceUnavailable after create' }
            return $data
        }
        $selectedId = ($uriText -split '/')[-1]
        $current = @($collection | Where-Object id -eq $selectedId)[0]
        if ($Method -eq 'DELETE') { [void]$collection.Remove($current); return }
        $data = $Body | ConvertFrom-Json
        $current.intent = $data.intent; $current.target = $data.target; $current.settings = $data.settings
        return $current
    }
    $mock.Reads.Add($uriText)
    if ($mock.FailRead) { throw '403 Forbidden reading apps' }
    if ($mock.Retry) { $mock.Retry = $false; throw '503 ServiceUnavailable' }
    if ($uriText -like '*/groups/*') {
        if ($mock.MissingGroup) { throw '404 Group not found' }
        return @{ displayName = 'Pilot group' }
    }
    if ($uriText -match '/assignments/([^/?]+)$') {
        $row = @($collection | Where-Object id -eq $Matches[1])
        if (-not $row.Count) { throw '404 Assignment not found' }
        return $row[0]
    }
    if ($uriText -like '*/assignments*') {
        if ($uriText -like '*page=2') {
            if ($mock.FailPage) { throw '403 Forbidden on assignment page two' }
            return @{ value = @($collection | Select-Object -Skip 1) }
        }
        return @{ value = @($collection | Select-Object -First 1); '@odata.nextLink' = "$uriText`?page=2" }
    }
    if ($uriText -like '*mobileApps/*') {
        if ($mock.SettingWritten -and $mock.FailSettingVerification) { throw '403 Forbidden verifying Win32 setting' }
        $selectedId = ($uriText -split '/')[-1]
        if ($mock.WrongSettingIdentity) {
            $wrong = @($mock.Apps | Where-Object id -eq $selectedId)[0].Clone()
            $wrong.id = $otherGroupId
            return $wrong
        }
        return @($mock.Apps | Where-Object id -eq $selectedId)[0]
    }
    if ($uriText -like '*assignmentFilters*' -or $uriText -like '*groups*') {
        if ($uriText -like '*page=2') { return @{ value = @(@{ id = $otherGroupId; displayName = 'Second'; platform = 'windows10' }) } }
        return @{ value = @(@{ id = $groupId; displayName = 'First'; platform = 'windows10' }); '@odata.nextLink' = "$uriText&page=2" }
    }
    if ($uriText -like '*mobileApps*') {
        if ($uriText -like '*page=2') {
            if ($mock.FailPage) { throw '403 Forbidden on app page two' }
            return @{ value = @($mock.Apps | Select-Object -Skip 2) }
        }
        return @{ value = @($mock.Apps | Select-Object -First 2); '@odata.nextLink' = "$uriText`?page=2" }
    }
    throw "Unexpected app test request $Method $Uri"
}
$get = Join-Path $root 'Features\Get-IntuneApps.ps1'
$write = Join-Path $root 'Features\Invoke-IntuneAppAssignment.ps1'
$logs = [System.Collections.Generic.List[string]]::new()
$common = @{ SkipGraphConnect = $true; PassThru = $true; LogAction = { param($message, $level) $logs.Add("$level $message") } }
$inventory = & $get @common
Assert-AppTest ($inventory.Rows.Count -eq 7) 'All app types, including unknown and web apps, load across pages'
foreach ($case in @(@('Windows', 2), @('iOS', 1), @('Android', 1), @('macOS', 1))) {
    $report = & $get @common -PlatformFilter $case[0]
    Assert-AppTest ($report.Rows.Count -eq $case[1]) "App platform filter $($case[0])"
}
$literal = & $get @common -NameFilter '[pilot]'
Assert-AppTest ($literal.Rows.Count -eq 1) 'App name matching is literal and case-insensitive'
$empty = & $get @common -NameFilter 'no-such-app'
Assert-AppTest ($empty.Rows.Count -eq 0 -and $empty.Action -eq 'Inventory') 'Empty inventory is an explicit completed report'
$mock.FailPage = $true
$published = [System.Collections.Generic.List[object]]::new()
$failed = $false
try { & $get @common | ForEach-Object { $published.Add($_) } } catch { $failed = $true }
Assert-AppTest ($failed -and $published.Count -eq 0) 'Late inventory failure publishes no partial inventory'
$mock.FailPage = $false
$mock.Retry = $true
[void](& $get @common)
Assert-AppTest ($mock.Sleeps -eq 1) 'App inventory uses shared transient-read retry'
$app = $inventory.Rows[0]
$settingScript = Join-Path $root 'Features\Set-IntuneAppAvailableUninstall.ps1'
Assert-AppTest ($app.AllowAvailableUninstall -eq $false -and $app.AllowAvailableUninstallDisplay -eq 'No' -and
    $inventory.Rows[1].AllowAvailableUninstallDisplay -eq 'N/A' -and $null -eq $inventory.Rows[1].AllowAvailableUninstall) 'Inventory preserves false and marks non-Win32 values N/A'
$originalApps = $mock.Apps
$mock.Apps = @(
    @{ id = $appId; displayName = 'Yes'; '@odata.type' = '#microsoft.graph.win32LobApp'; allowAvailableUninstall = $true },
    @{ id = $originalApps[1].id; displayName = 'No'; '@odata.type' = '#microsoft.graph.win32LobApp'; allowAvailableUninstall = $false },
    @{ id = $originalApps[2].id; displayName = 'Unknown'; '@odata.type' = '#microsoft.graph.win32LobApp' }
)
$settingsInventory = & $get @common
Assert-AppTest ($settingsInventory.Rows[0].AllowAvailableUninstallDisplay -eq 'Yes' -and
    $settingsInventory.Rows[1].AllowAvailableUninstallDisplay -eq 'No' -and
    $settingsInventory.Rows[2].AllowAvailableUninstallDisplay -eq 'Unknown' -and
    $null -eq $settingsInventory.Rows[2].AllowAvailableUninstall) 'Win32 inventory displays true, false and missing values across pages'
$mock.Apps = $originalApps
$settingParameters = @{ Apps = @($app); AllowAvailableUninstall = $true; Confirm = $false }
$result = & $settingScript @common @settingParameters -WhatIf
Assert-AppTest ($result.Status -eq 'WhatIf' -and $mock.Writes.Count -eq 0) 'Win32 setting WhatIf makes no writes'
$mock.RetrySettingWrite = $true
$sleepsBefore = $mock.Sleeps
$result = & $settingScript @common @settingParameters
$settingPayload = $mock.Writes[1].Body | ConvertFrom-Json
Assert-AppTest ($result.Status -eq 'OK' -and $result.AllowAvailableUninstall -eq $true -and
    $settingPayload.allowAvailableUninstall -is [bool] -and $settingPayload.allowAvailableUninstall -eq $true -and
    $settingPayload.'@odata.type' -eq '#microsoft.graph.win32LobApp' -and $settingPayload.PSObject.Properties.Name.Count -eq 2 -and
    $mock.Writes[1].Uri -like "*/mobileApps/$appId" -and $mock.Sleeps -eq ($sleepsBefore + 1)) 'Win32 setting uses an idempotent retry, minimal Boolean PATCH and verified read-back'
$beforeSettingWrites = $mock.Writes.Count
$result = & $settingScript @common -Apps @($app, $app, $inventory.Rows[1]) -AllowAvailableUninstall $true -Confirm:$false
Assert-AppTest ($result.Count -eq 2 -and $result[0].Status -eq 'NoChange' -and $result[1].Status -eq 'Skipped' -and
    $mock.Writes.Count -eq $beforeSettingWrites) 'Matching settings are no-ops, duplicates are deduplicated and Store apps are skipped'
$result = & $settingScript @common -Apps @($app) -AllowAvailableUninstall $false -Confirm:$false
Assert-AppTest ($result.Status -eq 'OK' -and $result.AllowAvailableUninstallDisplay -eq 'No' -and
    $mock.Apps[0].allowAvailableUninstall -eq $false) 'Win32 available uninstall can be explicitly disabled'
$mock.IgnoreSettingWrite = $true
$result = & $settingScript @common @settingParameters
Assert-AppTest ($result.Status -eq 'Failed' -and $result.AllowAvailableUninstall -eq $false -and
    $result.AllowAvailableUninstallDisplay -eq 'No') 'Read-back mismatch is failure and returns the observed value, not the requested value'
$mock.IgnoreSettingWrite = $false
$mock.SettingWritten = $false; $mock.FailSettingVerification = $true
$result = & $settingScript @common @settingParameters
Assert-AppTest ($result.Status -eq 'Failed' -and $null -eq $result.AllowAvailableUninstall -and
    $result.AllowAvailableUninstallDisplay -eq 'Unknown' -and $result.Message -like '*write was attempted*') 'Read-back failure reports uncertain persistence and Unknown instead of success'
$mock.FailSettingVerification = $false; $mock.SettingWritten = $false
$mock.Apps[0].allowAvailableUninstall = $false
$mock.Apps[0].'@odata.type' = '#microsoft.graph.winGetApp'
$beforeSettingWrites = $mock.Writes.Count
$result = & $settingScript @common @settingParameters
Assert-AppTest ($result.Status -eq 'Failed' -and $mock.Writes.Count -eq $beforeSettingWrites) 'Fresh app type verification blocks stale Win32 inventory writes'
$mock.Apps[0].'@odata.type' = '#microsoft.graph.win32LobApp'
$mock.WrongSettingIdentity = $true
$beforeSettingWrites = $mock.Writes.Count
$result = & $settingScript @common @settingParameters
Assert-AppTest ($result.Status -eq 'Failed' -and $mock.Writes.Count -eq $beforeSettingWrites) 'Fresh identity mismatch blocks Win32 property writes'
$mock.WrongSettingIdentity = $false
$mock.Apps[0].Remove('allowAvailableUninstall')
$result = & $settingScript @common -Apps @($app) -AllowAvailableUninstall $false -Confirm:$false
Assert-AppTest ($result.Status -eq 'OK' -and $mock.Writes.Count -eq ($beforeSettingWrites + 1) -and
    $result.AllowAvailableUninstall -eq $false) 'Missing current setting is not silently treated as No and can be explicitly updated'
$mock.Apps[1].'@odata.type' = '#microsoft.graph.win32LobApp'; $mock.Apps[1].allowAvailableUninstall = $false
$secondWin32 = [PSCustomObject]@{ Id = $mock.Apps[1].id; DisplayName = 'Second Win32'; TenantId = $tenant; AppType = 'win32LobApp' }
$mock.FailSettingApp = $appId
$result = & $settingScript @common -Apps @($app, $secondWin32) -AllowAvailableUninstall $true -Confirm:$false
Assert-AppTest ($result[0].Status -eq 'Failed' -and $result[1].Status -eq 'OK' -and
    $mock.Apps[1].allowAvailableUninstall -eq $true) 'Per-app setting failure does not stop another Win32 app from updating'
$mock.FailSettingApp = ''
$mock.Apps[1].'@odata.type' = '#microsoft.graph.winGetApp'; $mock.Apps[1].Remove('allowAvailableUninstall')
foreach ($invalid in @(
    [PSCustomObject]@{ Id = 'invalid'; TenantId = $tenant; AppType = 'win32LobApp' },
    [PSCustomObject]@{ Id = $appId; TenantId = $otherGroupId; AppType = 'win32LobApp' }
)) {
    $failed = $false; $beforeSettingWrites = $mock.Writes.Count
    try { & $settingScript @common -Apps @($app, $invalid) -AllowAvailableUninstall $true -Confirm:$false | Out-Null } catch { $failed = $true }
    Assert-AppTest ($failed -and $mock.Writes.Count -eq $beforeSettingWrites) 'Invalid/different-tenant selections are rejected before any setting writes'
}
$mock.Tenant = $otherGroupId
$failed = $false
try { & $settingScript @common @settingParameters -TenantId $tenant | Out-Null } catch { $failed = $true }
Assert-AppTest $failed 'Win32 setting rejects a changed Graph tenant'
$mock.Tenant = $tenant
$originalScopes = $mock.Scopes; $mock.Scopes = @('DeviceManagementApps.Read.All')
$failed = $false
try { & $settingScript @common @settingParameters | Out-Null } catch { $failed = $true }
Assert-AppTest $failed 'Win32 setting requires Apps ReadWrite permission'
$mock.Scopes = $originalScopes
$mock.Writes.Clear(); $mock.Apps[0].allowAvailableUninstall = $false
$mock.SettingWritten = $false
$group = [PSCustomObject]@{ Id = $groupId; DisplayName = 'Pilot group' }
$group2 = [PSCustomObject]@{ Id = $otherGroupId; DisplayName = 'Other group' }
$filter = [PSCustomObject]@{ Id = $filterId; DisplayName = 'Windows'; Platform = 'windows10' }
$mutation = @{ Apps = @($app); Action = 'Add'; TargetGroup = @($group); Confirm = $false }
$preview = & $write @common @mutation -WhatIf
Assert-AppTest ($preview.Status -eq 'WhatIf' -and $mock.Writes.Count -eq 0) 'App assignment WhatIf makes no writes'
$mock.Ambiguous = $true
$result = & $write @common @mutation -Intent required -FilterMode include -AssignmentFilter $filter
Assert-AppTest ($result.Status -eq 'OK' -and $mock.Writes.Count -eq 1 -and $mock.Assignments.Count -eq 1) 'Ambiguous create reconciles without replaying POST'
$result = & $write @common @mutation -Intent required -FilterMode include -AssignmentFilter $filter
Assert-AppTest ($result.Status -eq 'NoChange' -and $mock.Writes.Count -eq 1) 'Exact assignment duplicates do not write'
$result = & $write @common @mutation -Intent uninstall
Assert-AppTest ($result.Status -eq 'Failed' -and $mock.Writes.Count -eq 1) 'Conflicting targets require explicit editing'
$current = $mock.Assignments[0]
$current | Add-Member NoteProperty settings ([PSCustomObject]@{ '@odata.type' = '#microsoft.graph.win32LobAppAssignmentSettings'; notifications = 'hideAll'; restartSettings = @{ gracePeriodInMinutes = 90 }; custom = @(@{ retained = $true }) })
$current.target | Add-Member NoteProperty customTargetProperty 'preserve'
$mock.Assignments.Add(@{ id = $assignmentId; source = 'policySets'; intent = 'required'; target = @{ '@odata.type' = '#microsoft.graph.allDevicesAssignmentTarget' } })
$details = & $get @common -Action Assignments -AppId $appId
Assert-AppTest ($details.Rows.Count -eq 2 -and $details.Rows[0].TargetName -eq 'Pilot group' -and -not $details.Rows[1].Editable) 'Assignment inspection pages, resolves names, and protects Policy Set ownership'
$update = @{ Apps = @($app); Action = 'Update'; Assignments = @($details.Rows[0]); TargetGroup = @($group); Confirm = $false }
$result = & $write @common @update -Intent available
$payload = $mock.Writes[$mock.Writes.Count - 1].Body | ConvertFrom-Json
Assert-AppTest ($result.Status -eq 'OK' -and $payload.settings.notifications -eq 'hideAll' -and
    $payload.settings.restartSettings.gracePeriodInMinutes -eq 90 -and $payload.settings.custom[0].retained -and
    $payload.target.customTargetProperty -eq 'preserve' -and $payload.target.deviceAndAppManagementAssignmentFilterType -eq 'none' -and
    $null -eq $payload.target.deviceAndAppManagementAssignmentFilterId) 'Edit preserves nested platform settings and target properties while clearing filters explicitly'
$before = $mock.Writes.Count
$stale = & $write @common @update -Intent uninstall
Assert-AppTest ($stale.Status -eq 'Failed' -and $mock.Writes.Count -eq $before) 'Stale inspected assignments cannot be updated'
$policy = & $write @common -Apps $app -Action Remove -Assignments $details.Rows[1] -Confirm:$false
Assert-AppTest ($policy.Status -eq 'Failed' -and $mock.Writes.Count -eq $before) 'Policy Set-owned assignment deletion is blocked'
$mock.FailPage = $true
$failed = $false; $published.Clear()
try { & $get @common -Action Assignments -AppId $appId | ForEach-Object { $published.Add($_) } } catch { $failed = $true }
Assert-AppTest ($failed -and $published.Count -eq 0) 'Late assignment-read failures do not publish partial details'
$mock.FailPage = $false
$details = & $get @common -Action Assignments -AppId $appId
$removed = & $write @common -Apps $app -Action Remove -Assignments $details.Rows[0] -Confirm:$false
Assert-AppTest ($removed.Status -eq 'OK' -and $mock.Assignments.Count -eq 1 -and $mock.Assignments[0].source -eq 'policySets') 'Removal targets only the selected direct assignment'
$mock.Assignments.Clear()
$multi = @(& $write @common -Apps @($app, $inventory.Rows[1]) -Action Add -TargetGroup @($group, $group2) -Intent uninstall -Confirm:$false)
Assert-AppTest ($multi.Count -eq 4 -and @($multi | Where-Object Status -ne 'OK').Count -eq 0 -and
    $mock.Assignments.Count -eq 2 -and $mock.OtherAssignments.Count -eq 2) 'Bulk add creates an assignment for every app and queued group'
$mock.Assignments.Clear()
$mock.FailGroup = $otherGroupId
$partial = @(& $write @common -Apps $app -Action Add -TargetGroup @($group, $group2) -Confirm:$false)
Assert-AppTest ($partial.Count -eq 2 -and @($partial | Where-Object Status -eq 'OK').Count -eq 1 -and
    @($partial | Where-Object Status -eq 'Failed').Count -eq 1 -and $mock.Assignments.Count -eq 1) 'Bulk failures retain successful targets and expose each partial outcome'
$mock.FailGroup = ''
$mock.Assignments.Clear()
foreach ($intent in @('required', 'available', 'uninstall')) {
    $result = & $write @common -Apps $app -Action Add -TargetType Users -Intent $intent -Confirm:$false
    Assert-AppTest ($result.Status -eq 'OK') "All Users supports requested $intent payload"
    $mock.Assignments.Clear()
}
$result = & $write @common -Apps $app -Action Add -TargetType Devices -Intent available -Confirm:$false
Assert-AppTest ($result.Status -eq 'OK') 'Available Win32 All Devices payload is supported'
$mock.Assignments.Clear()
$savedType = $mock.Apps[0]['@odata.type']
$mock.Apps[0]['@odata.type'] = '#microsoft.graph.iosVppApp'
$before = $mock.Writes.Count
$result = & $write @common -Apps $app -Action Add -TargetType Devices -Intent available -Confirm:$false
Assert-AppTest ($result.Status -eq 'Failed' -and $mock.Writes.Count -eq $before) 'Unsupported Available All Devices combination is rejected before write'
$mock.Apps[0]['@odata.type'] = $savedType
$result = & $write @common @mutation -IncludeExclude exclude -Intent uninstall
$payload = $mock.Writes[$mock.Writes.Count - 1].Body | ConvertFrom-Json
Assert-AppTest ($result.Status -eq 'OK' -and $payload.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget' -and $payload.intent -eq 'uninstall') 'Excluded group preserves intent and uses exclusion target type'
$mock.Assignments.Clear()
$mock.FailRead = $true
$result = & $write @common @mutation
Assert-AppTest ($result.Status -eq 'Failed') 'App assignment read failure is explicit'
$mock.FailRead = $false; $mock.FailWrite = $true
$result = & $write @common @mutation
Assert-AppTest ($result.Status -eq 'Failed') 'App assignment service rejection is explicit'
$mock.FailWrite = $false
$mock.Scopes = @('DeviceManagementApps.Read.All')
$failed = $false
try { & $write @common @mutation | Out-Null } catch { $failed = $_.Exception.Message -like '*ReadWrite*' }
Assert-AppTest $failed 'App writes require Apps ReadWrite permission'
$mock.Scopes = @('DeviceManagementApps.ReadWrite.All')
$mock.Tenant = 'dddddddd-dddd-dddd-dddd-dddddddddddd'
$failed = $false
try { & $write @common @mutation -TenantId $tenant | Out-Null } catch { $failed = $_.Exception.Message -like '*tenant*' }
Assert-AppTest $failed 'App mutation rejects a changed tenant'
$mock.Tenant = $tenant
$targetSearch = Join-Path $root 'Features\Search-IntuneAssignmentTargets.ps1'
$targets = @(& $targetSearch @common -Action Groups -SearchTerm 'Pilot')
$filters = @(& $targetSearch @common -Action Filters)
Assert-AppTest ($targets.Count -eq 2 -and $filters.Count -eq 2) 'Shared group and filter searches follow every page'

function New-MockAssignment {
    param([string]$Id, [string]$Intent, [string]$Type, [string]$Group, [string]$Source = 'direct')
    $target = @{ '@odata.type' = "#microsoft.graph.$Type"; deviceAndAppManagementAssignmentFilterType = 'none'; deviceAndAppManagementAssignmentFilterId = $null }
    if ($Group) { $target.groupId = $Group }
    [PSCustomObject]@{ id = $Id; intent = $Intent; source = $Source; target = $target; settings = $null }
}
function Reset-ReplaceAssignments {
    $mock.Assignments.Clear()
    foreach ($row in @(
        (New-MockAssignment 'a0000000-0000-0000-0000-000000000001' required groupAssignmentTarget $groupId),
        (New-MockAssignment 'a0000000-0000-0000-0000-000000000002' required groupAssignmentTarget $replacedGroupId),
        (New-MockAssignment 'a0000000-0000-0000-0000-000000000003' available groupAssignmentTarget $availableGroupId),
        (New-MockAssignment 'a0000000-0000-0000-0000-000000000004' required exclusionGroupAssignmentTarget $excludedGroupId),
        (New-MockAssignment 'a0000000-0000-0000-0000-000000000005' required allDevicesAssignmentTarget '' 'policySets')
    )) { $mock.Assignments.Add($row) }
}
$replacedGroupId = 'e0000000-0000-0000-0000-000000000002'
$availableGroupId = 'e0000000-0000-0000-0000-000000000003'
$excludedGroupId = 'e0000000-0000-0000-0000-000000000004'
$replace = @{ Apps = @($app); Action = 'Replace'; Intent = 'required'; TargetGroup = @($group, $group2); Confirm = $false }
Reset-ReplaceAssignments
$before = $mock.Writes.Count
$plan = @(& $write @common @replace -PlanOnly)
$operations = (@($plan | Sort-Object Operation | ForEach-Object { "$($_.Operation):$($_.TargetKey.Split('|')[1])" }) -join ',')
Assert-AppTest ($mock.Writes.Count -eq $before -and @($plan | Where-Object Status -ne 'Planned').Count -eq 0 -and
    $operations -eq "Add:$otherGroupId,NoChange:$groupId,Remove:$replacedGroupId") 'Replace preview classifies targets, ignores other intents, exclusions and Policy Sets, and writes nothing'
$result = @(& $write @common @replace -Plan $plan)
$newWrites = @($mock.Writes | Select-Object -Skip $before)
Assert-AppTest ($newWrites.Count -eq 2 -and $newWrites[0].Method -eq 'POST' -and $newWrites[1].Method -eq 'DELETE' -and
    $newWrites[1].Uri -like '*a0000000-0000-0000-0000-000000000002' -and $mock.Assignments.Count -eq 5 -and
    @($result | Where-Object Status -eq 'Failed').Count -eq 0) 'Replace adds queued targets before removing replaced same-intent assignments'
Reset-ReplaceAssignments
$mock.FailGroup = $otherGroupId
$before = $mock.Writes.Count
$result = @(& $write @common @replace)
Assert-AppTest (@($result | Where-Object { $_.Operation -eq 'Remove' -and $_.Status -eq 'Skipped' }).Count -eq 1 -and
    @($mock.Writes | Select-Object -Skip $before | Where-Object Method -eq 'DELETE').Count -eq 0 -and $mock.Assignments.Count -eq 5) 'Failed replacement add keeps the existing assignments'
$mock.FailGroup = ''
Reset-ReplaceAssignments
$before = $mock.Writes.Count
$conflictReplace = $replace.Clone(); $conflictReplace.TargetGroup = @($group, [PSCustomObject]@{ Id = $availableGroupId; DisplayName = 'Available group' })
$result = @(& $write @common @conflictReplace)
Assert-AppTest (@($result | Where-Object Status -eq 'Conflict').Count -eq 1 -and $mock.Writes.Count -eq $before) 'Replace reports other-intent targets as conflicts and then removes nothing'
Reset-ReplaceAssignments
$plan = @(& $write @common @replace -PlanOnly)
$mock.Assignments[1].target.deviceAndAppManagementAssignmentFilterType = 'include'
$mock.Assignments[1].target.deviceAndAppManagementAssignmentFilterId = $filterId
$before = $mock.Writes.Count
$result = @(& $write @common @replace -Plan $plan)
Assert-AppTest (@($result | Where-Object { $_.Operation -eq 'Remove' -and $_.Message -like '*preview*' }).Count -eq 1 -and
    @($mock.Writes | Select-Object -Skip $before | Where-Object Method -eq 'DELETE').Count -eq 0) 'Assignment changed after the preview is not removed'
Reset-ReplaceAssignments
$before = $mock.Writes.Count
$filterReplace = $replace.Clone(); $filterReplace.TargetGroup = @($group)
$result = @(& $write @common @filterReplace -FilterMode include -AssignmentFilter $filter)
$patch = @($mock.Writes | Select-Object -Skip $before | Where-Object Method -eq 'PATCH')
Assert-AppTest ($patch.Count -eq 1 -and ($patch[0].Body | ConvertFrom-Json).target.deviceAndAppManagementAssignmentFilterId -eq $filterId -and
    @($result | Where-Object { $_.Operation -eq 'Update' -and $_.Status -eq 'OK' }).Count -eq 1) 'Replace updates the filter on an existing same-intent target in place'
$mock.Assignments.Clear()

$ui = . (Join-Path $root 'Invoke-IntuneToolbox-GUI.ps1') -ValidateOnly
try {
    Assert-AppTest ($ui.MissingFeatures.Count -eq 0 -and $MainTabs.Items.Count -eq 6 -and $MainTabs.Items[5].Header -eq 'App Management') 'App page is appended without changing existing navigation indexes'
    $ui.Window.FindName('NavAppManagement').IsChecked = $true
    Assert-AppTest ($MainTabs.SelectedIndex -eq 5 -and $TxtPageTitle.Text -eq 'App Management') 'App navigation selects the new page'
    Assert-AppTest ((Get-SelectedTag $AppUI.ComboAppTarget) -eq 'Devices' -and
        $AppUI.PanelAppGroupSearch.Visibility -eq 'Collapsed' -and $AppUI.PanelAppFilter.Visibility -eq 'Collapsed' -and
        $AppUI.PanelAppAssignmentDetails.Visibility -eq 'Collapsed') 'All Devices is the default and optional app panels are collapsed'
    function Invoke-InBackground {
        param($Worker, $Buttons, $BusyLabel, $LogBox, $LogScroll, $OnComplete)
        $script:pendingAppWorker = @{ Worker = $Worker; Complete = $OnComplete }
    }
    $click = [System.Windows.Controls.Button]::ClickEvent
    $AppUI.BtnAppLoad.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    $loadWorker = $pendingAppWorker
    Assert-AppTest ($script:AppState.Busy -and -not $AppUI.PanelAppResultFilters.IsEnabled -and
        $AppUI.TxtAppEmpty.Text -eq 'Loading apps...') 'Load locks result filters and shows a loading state'
    & $loadWorker.Worker.OutputAction $inventory
    & $loadWorker.Complete
    Assert-AppTest ($script:AppState.Inventory.Count -eq 7 -and $AppUI.ListApps.Items.Count -eq 7 -and
        $AppUI.PanelAppResultFilters.IsEnabled -and $AppUI.TxtAppCount.Text -like '7 visible / 7 loaded / 0 selected*') 'Load callback caches all rows and publishes visible/loaded/selected counts'
    $readsBeforeFiltering = $mock.Reads.Count
    $AppUI.ListApps.SelectedIndex = 0
    $script:AppState.InspectedAppId = $AppUI.ListApps.SelectedItem.Id
    $AppUI.PanelAppAssignmentDetails.Visibility = 'Visible'
    Select-AppComboTag $AppUI.ComboAppResultType 'win32LobApp'
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 1 -and $AppUI.ListApps.Items[0].Id -eq $appId -and
        $AppUI.ListApps.SelectedItems.Count -eq 0 -and -not $AppUI.BtnAppAdd.IsEnabled -and
        -not $AppUI.BtnAppInspect.IsEnabled -and -not $script:AppState.InspectedAppId -and
        $AppUI.PanelAppAssignmentDetails.Visibility -eq 'Collapsed') 'Exact Win32 filtering excludes Store apps and clears selection and inspection'
    Select-AppComboTag $AppUI.ComboAppResultPlatform 'iOS'
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 0) 'Platform filter independently excludes nonmatching Win32 rows'
    Select-AppComboTag $AppUI.ComboAppResultPlatform 'Windows'
    Select-AppComboTag $AppUI.ComboAppResultStatus 'Unknown'
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 0) 'Assignment hint filter independently excludes nonmatching Win32 rows'
    Select-AppComboTag $AppUI.ComboAppResultStatus 'Assigned (hint)'
    $AppUI.TxtAppResultName.Text = 'no such name'
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 0) 'App name filter independently excludes nonmatching rows'
    $AppUI.TxtAppResultName.Text = '[pILOT]'
    $AppUI.TxtAppResultPublisher.Text = 'pUbLiShEr'
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 1 -and
        $AppUI.TxtAppCount.Text -like '1 visible / 7 loaded / 0 selected*') 'All five filters combine with literal case-insensitive text matching'
    $AppUI.TxtAppResultPublisher.Text = 'missing publisher'
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 0 -and
        $AppUI.TxtAppEmpty.Text -eq 'No apps match the result filters.' -and
        $AppUI.ComboAppResultType.Items.Count -eq 8) 'Zero matches are explicit and dropdown options still cover the full inventory'
    $AppUI.BtnAppClearResultFilters.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    Select-AppComboTag $AppUI.ComboAppResultType 'futureApp'
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 1 -and $AppUI.ListApps.Items[0].DisplayName -eq 'Future') 'Previously unknown app types remain available as exact result filters'
    $AppUI.BtnAppClearResultFilters.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    $AppUI.TxtAppResultPublisher.Text = 'publisher'
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 1) 'Publisher filtering safely excludes missing publishers'
    $AppUI.BtnAppClearResultFilters.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    [void]$AppUI.ListApps.Items.SortDescriptions.Add(
        [System.ComponentModel.SortDescription]::new('DisplayName', [System.ComponentModel.ListSortDirection]::Descending))
    $expectedOrder = @($AppUI.ListApps.Items | ForEach-Object Id) -join ','
    $AppUI.TxtAppResultName.Text = 'windows'
    $AppUI.BtnAppClearResultFilters.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 7 -and $AppUI.ListApps.SelectedItems.Count -eq 0 -and
        (@($AppUI.ListApps.Items | ForEach-Object Id) -join ',') -eq $expectedOrder -and
        $AppUI.ListApps.Items.SortDescriptions.Count -eq 1) 'Clear filters restores every loaded row and preserves sorting without restoring selection'
    Assert-AppTest ($mock.Reads.Count -eq $readsBeforeFiltering -and $pendingAppWorker -eq $loadWorker -and
        -not $script:AppState.Busy) 'Local result filtering starts no workers or Graph requests'
    $AppUI.TxtAppResultName.Text = 'old filter'
    $AppUI.BtnAppLoad.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    Assert-AppTest ($script:AppState.Inventory.Count -eq 0 -and $AppUI.ListApps.Items.Count -eq 0 -and
        $AppUI.TxtAppResultName.Text -eq '') 'Fresh load clears cached inventory and result filters'
    & $pendingAppWorker.Complete
    Assert-AppTest (-not $script:AppState.InventoryLoaded -and $AppUI.ListApps.Items.Count -eq 0 -and
        $AppUI.TxtAppEmpty.Text -like '*load failed*' -and $AppUI.PanelAppResultFilters.IsEnabled) 'Failed load cannot restore stale inventory and has an explicit failure state'
    $AppUI.BtnAppLoad.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    & $pendingAppWorker.Worker.OutputAction ([PSCustomObject]@{ TenantId = $tenant; Rows = @() })
    & $pendingAppWorker.Complete
    Assert-AppTest ($script:AppState.InventoryLoaded -and
        $AppUI.TxtAppEmpty.Text -eq 'No apps returned by the load criteria.') 'Successful empty inventory is distinct from failure and filtered-out results'
    $AppUI.BtnAppLoad.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    & $pendingAppWorker.Worker.OutputAction ([PSCustomObject]@{ TenantId = $tenant; Rows = @($app) })
    & $pendingAppWorker.Complete
    Assert-AppTest (-not $AppUI.BtnAppAdd.IsEnabled -and -not $AppUI.BtnAppInspect.IsEnabled) 'Inventory does not automatically select apps'
    $AppUI.ListApps.SelectedIndex = 0
    Assert-AppTest ($AppUI.BtnAppInspect.IsEnabled -and $AppUI.BtnAppAdd.IsEnabled -and $AppUI.BtnAppReplace.IsEnabled) 'All Devices is ready when an app is selected without queued groups'
    $script:confirmApp = $false
    function Confirm-AppAssignment { param($Message, $Action) $script:confirmedAppMessage = $Message; return $script:confirmApp }
    $script:pendingAppWorker = $null
    $before = $mock.Writes.Count
    Invoke-AppAssignmentAction Add
    Assert-AppTest ($confirmedAppMessage -like '*All Devices*' -and $confirmedAppMessage -like '*WARNING:*tenant-wide*' -and
        $null -eq $pendingAppWorker -and -not $script:AppState.Busy -and $mock.Writes.Count -eq $before) 'Cancelling the All Devices confirmation starts no worker or write'
    Invoke-AppAssignmentAction Replace
    $defaultPreviewWorker = $pendingAppWorker
    Assert-AppTest ($defaultPreviewWorker.Worker.Parameters.TargetType -eq 'Devices' -and
        $defaultPreviewWorker.Worker.Parameters.PlanOnly -and $defaultPreviewWorker.Worker.Scopes -contains 'DeviceManagementApps.Read.All' -and
        $mock.Writes.Count -eq $before) 'All Devices Replace starts with a read-only preview'
    & $defaultPreviewWorker.Complete
    $script:confirmApp = $true
    Invoke-AppAssignmentAction Add
    $defaultWorker = $pendingAppWorker
    Assert-AppTest ($defaultWorker.Worker.Parameters.TargetType -eq 'Devices' -and
        $defaultWorker.Worker.Parameters.TargetGroup.Count -eq 0 -and $mock.Writes.Count -eq $before) 'Confirmed default assignment snapshots All Devices without queued groups'
    $AppUI.ListApps.SelectedItems.Clear()
    & $defaultWorker.Complete
    $AppUI.ListApps.SelectedIndex = 0
    $script:confirmApp = $false
    Select-AppComboTag $AppUI.ComboAppTarget 'Group'
    Assert-AppTest ($AppUI.PanelAppGroupSearch.Visibility -eq 'Visible' -and $AppUI.BtnAppInspect.IsEnabled -and
        -not $AppUI.BtnAppAdd.IsEnabled) 'Selecting Groups shows group search and requires a queued target'
    Select-AppComboTag $AppUI.ComboAppFilterMode 'include'
    $filterShown = $AppUI.PanelAppFilter.Visibility -eq 'Visible'
    Select-AppComboTag $AppUI.ComboAppFilterMode 'None'
    Assert-AppTest ($filterShown -and $AppUI.PanelAppFilter.Visibility -eq 'Collapsed') 'Filter picker is shown only when a filter mode is chosen'
    [void]$AppUI.ListAppGroups.Items.Add($group)
    $AppUI.ListAppGroups.SelectAll()
    $AppUI.BtnAppQueue.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    $AppUI.ListAppGroups.Items.Clear()
    Assert-AppTest ($AppUI.ListAppQueuedGroups.Items.Count -eq 1 -and $AppUI.BtnAppAdd.IsEnabled) 'App queues persist across result clearing without highlighted groups'
    Select-AppComboTag $AppUI.ComboAppTarget 'Users'
    $usersHidden = $AppUI.PanelAppGroupSearch.Visibility -eq 'Collapsed'
    Select-AppComboTag $AppUI.ComboAppTarget 'Devices'
    $devicesHidden = $AppUI.PanelAppGroupSearch.Visibility -eq 'Collapsed'
    Select-AppComboTag $AppUI.ComboAppTarget 'Group'
    Assert-AppTest ($usersHidden -and $devicesHidden -and $AppUI.PanelAppGroupSearch.Visibility -eq 'Visible' -and
        $AppUI.ListAppQueuedGroups.Items.Count -eq 1 -and $AppUI.BtnAppAdd.IsEnabled) 'Switching targets hides group search without clearing the queue'
    $script:pendingAppWorker = $null
    $before = $mock.Writes.Count
    Invoke-AppAssignmentAction Add
    Assert-AppTest ($mock.Writes.Count -eq $before -and $null -eq $pendingAppWorker -and -not $script:AppState.Busy) 'Cancelled app confirmation starts no worker or write'
    $mock.Assignments.Clear()
    [void](& $write @common @mutation)
    $details = & $get @common -Action Assignments -AppId $appId
    $script:AppState.InspectedAppId = $app.Id
    [void]$AppUI.ListAppAssignments.Items.Add($details.Rows[0])
    $AppUI.ListAppAssignments.SelectedIndex = 0
    Set-AppAssignmentForm
    Assert-AppTest ($AppUI.BtnAppUpdate.IsEnabled -and $AppUI.BtnAppRemove.IsEnabled -and $AppUI.ListAppQueuedGroups.Items.Count -eq 1) 'Direct assignment loads into form and enables selected edit/remove'
    Start-AppInspection
    Assert-AppTest ($script:AppState.Busy -and -not $AppUI.ListApps.IsEnabled -and -not $AppUI.PanelAppActions.IsEnabled) 'App background inspection disables conflicting controls'
    & $pendingAppWorker.Worker.OutputAction $details
    & $pendingAppWorker.Complete
    Assert-AppTest (-not $script:AppState.Busy -and $AppUI.ListApps.IsEnabled -and
        $script:AppState.InspectedAppId -eq $app.Id -and $AppUI.ListAppAssignments.Items.Count -eq 1) 'Real inspection callbacks publish results and restore readiness'
    Assert-AppTest ($AppUI.PanelAppAssignmentDetails.Visibility -eq 'Visible') 'Assignment details appear after Refresh assignments'
    Start-AppInspection
    Clear-AppInspection
    & $pendingAppWorker.Worker.OutputAction $details
    & $pendingAppWorker.Complete
    Assert-AppTest (-not $script:AppState.InspectedAppId -and $AppUI.ListAppAssignments.Items.Count -eq 0 -and
        $AppUI.TxtAppInspection.Text -like '*stale*' -and $AppUI.PanelAppAssignmentDetails.Visibility -eq 'Collapsed') 'Stale background inspection cannot repopulate invalidated details'
    function Confirm-AppAssignment { param($Message, $Action) $script:confirmedAppMessage = $Message; return $true }
    Invoke-AppAssignmentAction Add
    $writeWorker = $pendingAppWorker
    $activityText = [System.Windows.Documents.TextRange]::new($AppUI.TxtAppLog.Document.ContentStart, $AppUI.TxtAppLog.Document.ContentEnd).Text
    Assert-AppTest ($writeWorker.Worker.Parameters.Action -eq 'Add' -and
        $writeWorker.Worker.Parameters.TargetGroup.Count -eq 1 -and $confirmedAppMessage -like '*for 1 app(s)*' -and
        $confirmedAppMessage -notlike "*$appId*" -and $activityText -like "*$appId*") 'Confirmation shows app count only and logs exact app IDs to Activity'
    & $writeWorker.Worker.OutputAction ([PSCustomObject]@{ Status = 'OK' })
    & $writeWorker.Complete
    Assert-AppTest ($script:AppState.Busy -and $pendingAppWorker.Worker.Parameters.Action -eq 'Assignments' -and
        $app.AssignmentStatus -eq 'Refresh required') 'Mutation completion invalidates inventory hint and starts fresh assignment inspection'
    & $pendingAppWorker.Worker.OutputAction $details
    & $pendingAppWorker.Complete
    Assert-AppTest (-not $script:AppState.Busy -and $AppUI.ListAppAssignments.Items.Count -eq 1) 'Post-write refresh restores complete inspected assignments'
    $script:AppState.Busy = $true
    Update-AppState
    Assert-AppTest (-not $AppUI.BtnAppAdd.IsEnabled -and -not $AppUI.BtnAppRemove.IsEnabled) 'Busy app operation disables writes'
    $script:AppState.Busy = $false
    $AppUI.ListApps.SelectedItems.Clear()
    Assert-AppTest ($AppUI.ListAppAssignments.Items.Count -eq 0 -and -not $AppUI.BtnAppRemove.IsEnabled -and
        $AppUI.PanelAppAssignmentDetails.Visibility -eq 'Collapsed') 'Changing app selection invalidates and hides old inspection'
    $AppUI.ListApps.SelectedIndex = 0
    Assert-AppTest ($AppUI.BtnAppReplace.IsEnabled -and $AppUI.ListAppQueuedGroups.Items.Count -eq 1) 'Replace is available for selected apps with a queued target'
    $planRows = @(
        [PSCustomObject]@{ AppId = $appId; AppName = $app.DisplayName; Operation = 'Add'; Status = 'Planned'; Target = 'Pilot group'; Message = ''; AssignmentId = ''; Fingerprint = '' },
        [PSCustomObject]@{ AppId = $appId; AppName = $app.DisplayName; Operation = 'Remove'; Status = 'Planned'; Target = 'Old group'; Message = ''; AssignmentId = $assignmentId; Fingerprint = 'fp' }
    )
    function Confirm-AppAssignment { param($Message, $Action) $script:confirmedAppMessage = $Message; return $script:confirmReplace }
    $script:confirmReplace = $false
    Invoke-AppAssignmentAction Replace
    $previewWorker = $pendingAppWorker
    Assert-AppTest ($previewWorker.Worker.Parameters.PlanOnly -and $previewWorker.Worker.Parameters.Action -eq 'Replace' -and
        $previewWorker.Worker.Scopes -contains 'DeviceManagementApps.Read.All') 'Replace starts with a read-only preview'
    foreach ($row in $planRows) { & $previewWorker.Worker.OutputAction $row }
    & $previewWorker.Complete
    Assert-AppTest ($pendingAppWorker -eq $previewWorker -and -not $script:AppState.Busy -and
        $confirmedAppMessage -like '*- REMOVE: Old group (1 app)*' -and $confirmedAppMessage -like '*Apps to change: 1*' -and
        $confirmedAppMessage -notlike "*$appId*" -and $confirmedAppMessage -like '*1 REMOVE*' -and
        [System.Windows.Documents.TextRange]::new($AppUI.TxtAppLog.Document.ContentStart, $AppUI.TxtAppLog.Document.ContentEnd).Text -like "*Replace preview per app*$appId*- REMOVE: Old group*") 'Cancelled Replace preview counts removals per target without listing apps and starts no write'
    $script:confirmReplace = $true
    Invoke-AppAssignmentAction Replace
    foreach ($row in $planRows) { & $pendingAppWorker.Worker.OutputAction $row }
    & $pendingAppWorker.Complete
    Assert-AppTest ($pendingAppWorker.Worker.Parameters.Action -eq 'Replace' -and -not $pendingAppWorker.Worker.Parameters.PlanOnly -and
        $pendingAppWorker.Worker.Parameters.Plan.Count -eq 2 -and $pendingAppWorker.Worker.Scopes -contains 'DeviceManagementApps.ReadWrite.All') 'Confirmed Replace runs the write pinned to the previewed plan'
    & $pendingAppWorker.Complete
    & $pendingAppWorker.Worker.OutputAction $details
    & $pendingAppWorker.Complete
    $app.AssignmentStatus = 'Assigned (hint)'
    Update-AppResultFilterOptions
    Select-AppComboTag $AppUI.ComboAppResultStatus 'Assigned (hint)'
    $AppUI.ListApps.SelectedIndex = 0
    Invoke-AppAssignmentAction Add
    $statusWorker = $pendingAppWorker
    & $statusWorker.Worker.OutputAction ([PSCustomObject]@{ Status = 'OK' })
    & $statusWorker.Complete
    Assert-AppTest (-not $script:AppState.Busy -and $AppUI.ListApps.Items.Count -eq 0 -and
        $AppUI.ListApps.SelectedItems.Count -eq 0 -and -not $script:AppState.InspectedAppId -and
        (Get-SelectedTag $AppUI.ComboAppResultStatus) -eq 'Assigned (hint)' -and
        @($AppUI.ComboAppResultStatus.Items | Where-Object Tag -eq 'Refresh required').Count -eq 1) 'Mutation updates status options and filters out changed rows without inspecting hidden apps'
    $AppUI.BtnAppClearResultFilters.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 1 -and $AppUI.ListApps.Items[0].AssignmentStatus -eq 'Refresh required' -and
        $AppUI.ListApps.SelectedItems.Count -eq 0) 'Clearing the status filter restores the updated cached row without selecting it'
    Clear-AppManagement
    $AppUI.BtnAppClearResultFilters.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    Assert-AppTest ($AppUI.ListApps.Items.Count -eq 0 -and $AppUI.ListAppQueuedGroups.Items.Count -eq 0 -and
        $script:AppState.Inventory.Count -eq 0 -and -not $script:AppState.InventoryLoaded -and
        $AppUI.ComboAppResultType.Items.Count -eq 1 -and $AppUI.TxtAppEmpty.Text -eq 'No apps loaded.') 'Sign-out cleanup clears cached inventory, result filters, inspection and target queues'

    $uninstallColumn = @($AppUI.ListApps.View.Columns | Where-Object Header -eq 'Allow available uninstall')
    Assert-AppTest ($uninstallColumn.Count -eq 1 -and
        $uninstallColumn[0].DisplayMemberBinding.Path.Path -eq 'AllowAvailableUninstallDisplay' -and
        $AppUI.ComboAppAvailableUninstall.SelectedIndex -eq -1 -and -not $AppUI.BtnAppAvailableUninstall.IsEnabled) 'Uninstall column binds display values and the setting picker has no default'
    $script:AppState.Inventory = @($app, $inventory.Rows[1]); $script:AppState.InventoryLoaded = $true
    Update-AppResultFilterOptions -Reset; Update-AppResultFilter
    $AppUI.ListApps.SelectedItem = $app
    Assert-AppTest (-not $AppUI.BtnAppAvailableUninstall.IsEnabled) 'Win32 selection alone cannot apply a setting without an explicit value'
    $AppUI.ListApps.UnselectAll()
    Select-AppComboTag $AppUI.ComboAppAvailableUninstall 'Yes'
    $AppUI.ListApps.SelectedItem = $inventory.Rows[1]
    Assert-AppTest (-not $AppUI.BtnAppAvailableUninstall.IsEnabled) 'Non-Win32-only selection cannot apply the setting'
    [void]$AppUI.ListApps.SelectedItems.Add($app)
    Assert-AppTest $AppUI.BtnAppAvailableUninstall.IsEnabled 'Mixed selection with a Win32 app can apply an explicit setting'
    $script:confirmUninstall = $false
    function Confirm-AppAvailableUninstall { param($Message) $script:uninstallMessage = $Message; return $script:confirmUninstall }
    $script:pendingAppWorker = $null
    $AppUI.BtnAppAvailableUninstall.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    Assert-AppTest ($null -eq $pendingAppWorker -and $uninstallMessage -like '*to Yes*' -and
        $uninstallMessage -like "*$appId*" -and $uninstallMessage -like '*1 non-Win32*' -and
        $uninstallMessage -like "*$tenant*") 'Setting confirmation lists value, eligible IDs, tenant and skips; cancellation starts no worker'
    $script:confirmUninstall = $true
    $assignmentHint = $app.AssignmentStatus
    $sortCount = $AppUI.ListApps.Items.SortDescriptions.Count
    $AppUI.BtnAppAvailableUninstall.RaiseEvent([System.Windows.RoutedEventArgs]::new($click))
    $settingWorker = $pendingAppWorker
    Assert-AppTest ($script:AppState.Busy -and -not $AppUI.BtnAppAvailableUninstall.IsEnabled -and
        -not $AppUI.PanelAppActions.IsEnabled -and $settingWorker.Worker.Parameters.AllowAvailableUninstall -is [bool] -and
        $settingWorker.Worker.Parameters.AllowAvailableUninstall -eq $true -and $settingWorker.Worker.Parameters.Apps.Count -eq 2 -and
        $settingWorker.Worker.Scopes -contains 'DeviceManagementApps.ReadWrite.All' -and
        $settingWorker.Worker.FeaturePath -like '*Set-IntuneAppAvailableUninstall.ps1') 'Confirmed Win32 setting snapshots a Boolean and mixed selection with Apps Write scope while locking the UI'
    & $settingWorker.Worker.OutputAction ([PSCustomObject]@{ AppId = $appId; Status = 'OK'; AllowAvailableUninstall = $true; AllowAvailableUninstallDisplay = 'Yes' })
    & $settingWorker.Worker.OutputAction ([PSCustomObject]@{ AppId = $inventory.Rows[1].Id; Status = 'Skipped'; AllowAvailableUninstall = $null; AllowAvailableUninstallDisplay = 'N/A' })
    & $settingWorker.Complete
    Assert-AppTest ($app.AllowAvailableUninstallDisplay -eq 'Yes' -and $app.AllowAvailableUninstall -eq $true -and
        $app.AssignmentStatus -eq $assignmentHint -and $AppUI.ListApps.SelectedItems.Count -eq 2 -and
        $AppUI.ListApps.Items.SortDescriptions.Count -eq $sortCount -and -not $script:AppState.Busy) 'Setting completion refreshes verified values and preserves selection/sorting without changing assignment status'
    Select-AppComboTag $AppUI.ComboAppAvailableUninstall 'No'
    Invoke-AppAvailableUninstall
    Assert-AppTest ($pendingAppWorker.Worker.Parameters.AllowAvailableUninstall -eq $false) 'No picker passes Boolean false instead of a truthy string'
    & $pendingAppWorker.Worker.OutputAction ([PSCustomObject]@{ AppId = $appId; Status = 'Failed'; AllowAvailableUninstall = $null; AllowAvailableUninstallDisplay = 'Unknown' })
    & $pendingAppWorker.Complete
    Assert-AppTest ($null -eq $app.AllowAvailableUninstall -and $app.AllowAvailableUninstallDisplay -eq 'Unknown') 'Unverified failure invalidates the old UI value'
    $app.AllowAvailableUninstall = $true; $app.AllowAvailableUninstallDisplay = 'Yes'
    $AppUI.TxtAppResultPublisher.Text = 'publisher'
    $AppUI.ListApps.SelectedItem = $app
    Invoke-AppAvailableUninstall
    & $pendingAppWorker.Complete
    Assert-AppTest ($null -eq $app.AllowAvailableUninstall -and $app.AllowAvailableUninstallDisplay -eq 'Unknown' -and
        $AppUI.TxtAppResultPublisher.Text -eq 'publisher' -and $AppUI.ListApps.Items.Count -eq 1 -and
        $AppUI.ListApps.SelectedItem.Id -eq $appId) 'Missing worker outcomes invalidate cached values while preserving the active filter and selection'
    Clear-AppManagement
    Assert-AppTest ($AppUI.ComboAppAvailableUninstall.SelectedIndex -eq -1 -and -not $AppUI.BtnAppAvailableUninstall.IsEnabled) 'Sign-out cleanup clears the explicit Win32 setting choice'
} finally { $ui.Window.Close() }
Write-Host "Completed $checks app management checks."
