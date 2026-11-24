<#
.SYNOPSIS
Adds, replaces, updates or removes assignments on Intune apps.
.DESCRIPTION
Unrelated assignments are never resubmitted. Update preserves platform-specific settings.
New assignments omit settings and request Intune service defaults; service rejections are reported.
Replace makes each selected app's direct assignments with the chosen intent and include/exclude type
match the queued targets. Adds and updates run first; an app's removals run only when every add and
update for that app succeeded. Policy Set-owned assignments and other intents are never changed.
Use -PlanOnly to preview Replace without writing, then pass the preview rows as -Plan so that only
previewed, unchanged assignments are updated or removed.
Bulk operations are not atomic. Pass inspected assignment rows for Update/Remove.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][object[]]$Apps,
    [Parameter(Mandatory)][ValidateSet('Add', 'Update', 'Remove', 'Replace')][string]$Action,
    [ValidateSet('required', 'available', 'uninstall')][string]$Intent = 'required',
    [ValidateSet('Group', 'Users', 'Devices')][string]$TargetType = 'Group',
    [object[]]$TargetGroup,
    [ValidateSet('include', 'exclude')][string]$IncludeExclude = 'include',
    [ValidateSet('None', 'include', 'exclude')][string]$FilterMode = 'None',
    [object]$AssignmentFilter,
    [object[]]$Assignments,
    [switch]$PlanOnly,
    [object[]]$Plan,
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Shared\IntuneAppHelpers.ps1')
if ($PlanOnly -and $Action -ne 'Replace') { throw '-PlanOnly is supported only for Replace.' }
$writeScope = 'DeviceManagementApps.ReadWrite.All'
$scope = if ($PlanOnly) { 'DeviceManagementApps.Read.All' } else { $writeScope }
$tenant = Assert-IntuneAppContext -Scope $scope -Tenant $TenantId -Connect:(-not $SkipGraphConnect)
$base = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps'
if (-not $Apps.Count) { throw 'Select at least one app.' }
if ($Action -in @('Update', 'Remove') -and ($Apps.Count -ne 1 -or -not $Assignments.Count)) { throw 'Select one app and inspected assignments for Update/Remove.' }
if ($Action -eq 'Update' -and $Assignments.Count -ne 1) { throw 'Select exactly one assignment to edit.' }
$targets = @()
if ($Action -ne 'Remove') {
    if ($TargetType -ne 'Group' -and $IncludeExclude -eq 'exclude') { throw 'Only group targets can be excluded.' }
    if ($FilterMode -ne 'None' -and -not $AssignmentFilter.Id) { throw 'Select an assignment filter.' }
    if ($FilterMode -ne 'None') {
        $filterId = [guid]::Empty
        if (-not [guid]::TryParse([string]$AssignmentFilter.Id, [ref]$filterId)) { throw 'The assignment filter must have a valid GUID.' }
        if ($AssignmentFilter.ManagementType -eq 'apps') { throw 'Managed-app filters cannot be used for app deployment assignments. Select a device filter.' }
    }
    if ($IncludeExclude -eq 'exclude' -and $FilterMode -ne 'None') { throw 'Assignment filters cannot be used on excluded groups.' }
    if ($TargetType -eq 'Group') {
        if (-not $TargetGroup.Count) { throw 'Queue at least one group.' }
        $TargetGroup = @($TargetGroup | Group-Object Id | ForEach-Object { $_.Group[0] })
        if ($Action -eq 'Update' -and $TargetGroup.Count -ne 1) { throw 'An edited assignment must have exactly one target group.' }
    }
    $targetCount = if ($TargetType -eq 'Group') { $TargetGroup.Count } else { 1 }
    $targets = @(for ($index = 0; $index -lt $targetCount; $index++) {
        $target = [ordered]@{ '@odata.type' = switch ($TargetType) {
            'Group' { if ($IncludeExclude -eq 'exclude') { '#microsoft.graph.exclusionGroupAssignmentTarget' } else { '#microsoft.graph.groupAssignmentTarget' } }
            'Users' { '#microsoft.graph.allLicensedUsersAssignmentTarget' }
            'Devices' { '#microsoft.graph.allDevicesAssignmentTarget' }
        } }
        if ($TargetType -eq 'Group') {
            $id = [guid]::Empty
            if (-not [guid]::TryParse([string]$TargetGroup[$index].Id, [ref]$id)) { throw 'Every target group must have a valid GUID.' }
            $target.groupId = [string]$TargetGroup[$index].Id
        }
        $target.deviceAndAppManagementAssignmentFilterType = $FilterMode.ToLowerInvariant()
        $target.deviceAndAppManagementAssignmentFilterId = if ($FilterMode -eq 'None') { $null } else { [string]$AssignmentFilter.Id }
        $target
    })
}
$groupNames = @{}
foreach ($queued in @($TargetGroup)) { if ($queued.Id) { $groupNames[[string]$queued.Id] = [string]$queued.DisplayName } }
$results = [System.Collections.Generic.List[object]]::new()

function Assert-IntuneAppTarget {
    param($App)
    $freshApp = Invoke-IntuneGraphWithRetry -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri "$base/$($App.Id)" -ErrorAction Stop
    }
    if ($freshApp.id -ne $App.Id) { throw 'App identity could not be verified. Refresh the inventory.' }
    if ($Action -eq 'Remove') { return }
    if ($Intent -eq 'available' -and $TargetType -eq 'Devices' -and
        [string]$freshApp.'@odata.type' -notmatch '\.(win32LobApp|androidManagedStoreApp)$') {
        throw 'Available device targeting is supported here only for Win32 and managed Google Play apps. Use a user target for other types.'
    }
    if ($FilterMode -ne 'None') {
        $platform = Get-IntuneAppPlatform $freshApp
        $filterPlatform = [string]$AssignmentFilter.Platform
        $expected = switch ($platform) { 'Windows' { 'windows' }; 'Android' { 'android' }; 'iOS' { 'ios' }; 'macOS' { 'macOS' } }
        if ($expected -and $filterPlatform -and $filterPlatform -notmatch "^$expected") { throw 'The assignment filter platform does not match this app.' }
    }
}

function New-IntuneAppAssignment {
    param([string]$Uri, $Target)
    $body = @{ '@odata.type' = '#microsoft.graph.mobileAppAssignment'; intent = $Intent; target = $Target } | ConvertTo-Json -Depth 50
    # POST is not safely replayable after an ambiguous response. Reconcile instead.
    try {
        $created = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $body -ContentType 'application/json' -ErrorAction Stop
        return [string]$created.id
    } catch {
        $failure = $_
        $after = @(Get-IntuneAppPages -Uri $Uri)
        $found = @($after | Where-Object { $_.source -ne 'policySets' -and (Test-IntuneAppAssignmentMatch $_ $Intent $Target) })
        if ($found.Count -ne 1) { throw $failure }
        Write-IntuneAppLog "Create response failed; matching assignment $($found[0].id) was found on reconciliation." 'WARN'
        return [string]$found[0].id
    }
}

function Update-IntuneAppAssignment {
    param([string]$Uri, $Current, [string]$Fingerprint, $NewTarget, [object[]]$Fresh)
    $assignmentKey = [string]$Current.id
    $merged = ConvertTo-IntuneAppDictionary $Current.target
    foreach ($key in @('@odata.type', 'groupId', 'deviceAndAppManagementAssignmentFilterId', 'deviceAndAppManagementAssignmentFilterType')) { $merged.Remove($key) }
    foreach ($key in $NewTarget.Keys) { $merged[$key] = $NewTarget[$key] }
    $conflicts = @($Fresh | Where-Object { $_.id -ne $assignmentKey -and (Get-IntuneAppTargetKey $_.target) -eq (Get-IntuneAppTargetKey $merged) })
    if ($conflicts.Count) { throw 'Another assignment already uses the proposed target.' }
    $body = @{ intent = $Intent; target = $merged; settings = ConvertTo-IntuneAppDictionary $Current.settings } | ConvertTo-Json -Depth 50
    Invoke-IntuneGraphWithRetry -ScriptBlock {
        [void](Assert-IntuneAppContext -Scope $writeScope -Tenant $tenant)
        $check = Invoke-MgGraphRequest -Method GET -Uri "$Uri/$assignmentKey" -ErrorAction Stop
        if ((Get-IntuneAppAssignmentFingerprint $check) -ne $Fingerprint) { throw 'Assignment changed before update. Refresh assignments.' }
        Invoke-MgGraphRequest -Method PATCH -Uri "$Uri/$assignmentKey" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
    }
}

function Remove-IntuneAppAssignment {
    param([string]$Uri, [string]$AssignmentKey, [string]$Fingerprint)
    Invoke-IntuneGraphWithRetry -ScriptBlock {
        [void](Assert-IntuneAppContext -Scope $writeScope -Tenant $tenant)
        $check = Invoke-MgGraphRequest -Method GET -Uri "$Uri/$AssignmentKey" -ErrorAction Stop
        if ((Get-IntuneAppAssignmentFingerprint $check) -ne $Fingerprint) { throw 'Assignment changed before removal. Refresh assignments.' }
        Invoke-MgGraphRequest -Method DELETE -Uri "$Uri/$AssignmentKey" -ErrorAction Stop | Out-Null
    }
}

function Get-IntuneAppTargetLabel {
    param($Target)
    $type = [string]$Target.'@odata.type' -replace '^#?microsoft\.graph\.', ''
    if ($type -eq 'allLicensedUsersAssignmentTarget') { return 'All Users' }
    if ($type -eq 'allDevicesAssignmentTarget') { return 'All Devices' }
    $groupKey = [string]$Target.groupId
    if (-not $groupKey) { return $type }
    if (-not $groupNames.ContainsKey($groupKey)) {
        try {
            $group = Invoke-IntuneGraphWithRetry -ScriptBlock {
                Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupKey`?`$select=id,displayName" -ErrorAction Stop
            }
            $groupNames[$groupKey] = [string]$group.displayName
        } catch { $groupNames[$groupKey] = '' }
    }
    $name = if ($groupNames[$groupKey]) { "$($groupNames[$groupKey]) [$groupKey]" } else { $groupKey }
    if ($type -eq 'exclusionGroupAssignmentTarget') { return "$name (exclude)" }
    return $name
}

function Add-IntuneAppResult {
    param($App, [string]$Operation, [string]$Label, [string]$Status, [string]$Message,
        [string]$AssignmentKey = '', [string]$Fingerprint = '', [string]$TargetKey = '')
    $results.Add([PSCustomObject]@{
        AppId = [string]$App.Id; AppName = [string]$App.DisplayName; AssignmentId = $AssignmentKey; Target = $Label
        Action = $Action; Operation = $Operation; Status = $Status; Message = $Message; TargetKey = $TargetKey; Fingerprint = $Fingerprint
    })
    $prefix = if ($Action -eq 'Replace') { "Replace/$Operation" } else { $Action }
    $level = switch ($Status) { 'Failed' { 'ERR' }; 'OK' { 'OK' }; 'Planned' { 'INFO' }; default { 'WARN' } }
    Write-IntuneAppLog "[$($App.DisplayName)] $prefix $Label - ${Status}: $Message" $level
}

function Invoke-IntuneAppReplace {
    param($App, [string]$Uri)
    try {
        [void](Assert-IntuneAppContext -Scope $scope -Tenant $tenant)
        Assert-IntuneAppTarget $App
        $fresh = @(Get-IntuneAppPages -Uri $Uri)
        $planned = @(Get-IntuneAppReplacePlan -Assignments $fresh -Intent $Intent -IncludeExclude $IncludeExclude -Targets $targets)
    } catch {
        Add-IntuneAppResult -App $App -Operation 'Plan' -Label '' -Status 'Failed' -Message $_.Exception.Message
        return
    }
    # Removals run last so a failed add never leaves the app without its replacement targets.
    $ordered = @($planned | Where-Object Operation -ne 'Remove') + @($planned | Where-Object Operation -eq 'Remove')
    $blocked = $false
    foreach ($step in $ordered) {
        $label = Get-IntuneAppTargetLabel $step.Target
        $assignmentKey = [string]$step.Assignment.id
        $fingerprint = if ($step.Assignment) { Get-IntuneAppAssignmentFingerprint $step.Assignment } else { '' }
        $status = 'Skipped'; $message = $step.Reason
        if ($PlanOnly) {
            $status = 'Planned'
        } elseif ($step.Operation -eq 'NoChange') {
            $status = 'NoChange'
        } elseif ($step.Operation -eq 'Conflict') {
            $status = 'Conflict'; $blocked = $true
        } elseif ($step.Operation -eq 'Remove' -and $blocked) {
            $message = 'Kept because an add or update for this app did not succeed.'
        } elseif ($Plan -and $step.Operation -ne 'Add' -and -not @($Plan | Where-Object {
                $_.AppId -eq $App.Id -and $_.Operation -eq $step.Operation -and $_.AssignmentId -eq $assignmentKey -and $_.Fingerprint -eq $fingerprint
            }).Count) {
            $message = 'Changed since the preview. Preview Replace again.'
            if ($step.Operation -ne 'Remove') { $blocked = $true }
        } elseif ($PSCmdlet.ShouldProcess("$($App.DisplayName) [$($App.Id)] / $label", "Replace: $($step.Operation) $Intent assignment")) {
            try {
                [void](Assert-IntuneAppContext -Scope $writeScope -Tenant $tenant)
                switch ($step.Operation) {
                    'Add' {
                        $assignmentKey = New-IntuneAppAssignment -Uri $Uri -Target $step.Target
                        $message = 'Assignment created (service defaults requested).'
                    }
                    'Update' {
                        Update-IntuneAppAssignment -Uri $Uri -Current $step.Assignment -Fingerprint $fingerprint -NewTarget $step.Target -Fresh $fresh
                        $message = 'Assignment filter updated; existing settings preserved.'
                    }
                    'Remove' {
                        Remove-IntuneAppAssignment -Uri $Uri -AssignmentKey $assignmentKey -Fingerprint $fingerprint
                        $message = 'Replaced assignment removed.'
                    }
                }
                $status = 'OK'
            } catch {
                $status = 'Failed'; $message = $_.Exception.Message
                if ($step.Operation -ne 'Remove') { $blocked = $true }
            }
        } elseif ($WhatIfPreference) {
            $status = 'WhatIf'; $message = "$($step.Operation) $Intent / $label"
        } elseif ($step.Operation -ne 'Remove') {
            $blocked = $true
        }
        Add-IntuneAppResult -App $App -Operation $step.Operation -Label $label -Status $status -Message $message `
            -AssignmentKey $assignmentKey -Fingerprint $fingerprint -TargetKey $step.TargetKey
    }
}

foreach ($app in $Apps) {
    $id = [guid]::Empty
    if (-not [guid]::TryParse([string]$app.Id, [ref]$id) -or $app.TenantId -ne $tenant) { throw 'Invalid or different-tenant app. Refresh inventory.' }
    $uri = "$base/$($app.Id)/assignments"
    if ($Action -eq 'Replace') { Invoke-IntuneAppReplace -App $app -Uri $uri; continue }
    $operations = if ($Action -eq 'Add') { $targets } else { $Assignments }
    foreach ($operation in $operations) {
        $status = 'Skipped'; $message = ''; $assignmentId = ''
        $label = if ($Action -eq 'Add') { Get-IntuneAppTargetKey $operation } else { [string]$operation.Id }
        if ($PSCmdlet.ShouldProcess("$($app.DisplayName) [$($app.Id)] / $label", "$Action $Intent assignment")) {
            try {
                [void](Assert-IntuneAppContext -Scope $writeScope -Tenant $tenant)
                Assert-IntuneAppTarget $app
                $fresh = @(Get-IntuneAppPages -Uri $uri)
                if ($Action -eq 'Add') {
                    $sameTarget = @($fresh | Where-Object { (Get-IntuneAppTargetKey $_.target) -eq (Get-IntuneAppTargetKey $operation) })
                    if ($sameTarget.Count) {
                        if (@($sameTarget | Where-Object { Test-IntuneAppAssignmentMatch $_ $Intent $operation }).Count -eq $sameTarget.Count) {
                            $status = 'NoChange'; $message = 'Matching assignment already exists.'
                        } else { throw 'This target already has a conflicting assignment. Inspect and explicitly edit it, or use Replace.' }
                    } else {
                        $assignmentId = New-IntuneAppAssignment -Uri $uri -Target $operation
                        $status = 'OK'; $message = 'Assignment created (service defaults requested).'
                    }
                } else {
                    $assignmentId = [string]$operation.Id
                    $assignmentGuid = [guid]::Empty
                    if (-not [guid]::TryParse($assignmentId, [ref]$assignmentGuid) -or $operation.AppId -ne $app.Id -or $operation.TenantId -ne $tenant) {
                        throw 'Assignment snapshot does not belong to this app and tenant.'
                    }
                    $current = @($fresh | Where-Object { $_.id -eq $assignmentId })
                    if ($current.Count -ne 1 -or ($current[0].source -and $current[0].source -ne 'direct')) { throw 'Assignment is missing or not directly owned. Refresh assignments.' }
                    if ((Get-IntuneAppAssignmentFingerprint $current[0]) -ne $operation.Fingerprint) { throw 'Assignment changed since inspection. Refresh before editing/removing.' }
                    if ($Action -eq 'Remove') {
                        Remove-IntuneAppAssignment -Uri $uri -AssignmentKey $assignmentId -Fingerprint $operation.Fingerprint
                        $message = 'Selected assignment removed.'
                    } else {
                        Update-IntuneAppAssignment -Uri $uri -Current $current[0] -Fingerprint $operation.Fingerprint -NewTarget $targets[0] -Fresh $fresh
                        $message = 'Assignment updated; existing settings preserved.'
                    }
                    $status = 'OK'
                }
            } catch { $status = 'Failed'; $message = $_.Exception.Message }
        } elseif ($WhatIfPreference) { $status = 'WhatIf'; $message = "$Action $Intent / $label" }
        Add-IntuneAppResult -App $app -Operation $Action -Label $label -Status $status -Message $message -AssignmentKey $assignmentId
    }
}
if ($PassThru) { $results.ToArray() } else { $results | Format-Table AppName, Operation, Target, Status, Message }
