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
foreach ($path in @('Features\Get-IntuneApps.ps1', 'Features\Invoke-IntuneAppAssignment.ps1', 'Shared\IntuneAppHelpers.ps1')) {
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
        @{ id = $appId; displayName = 'Windows [Pilot]'; publisher = 'Publisher'; '@odata.type' = '#microsoft.graph.win32LobApp'; isAssigned = $true },
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
        $selectedId = ($uriText -split '/')[-1]
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
} finally { $ui.Window.Close() }
Write-Host "Completed $checks app management checks."
