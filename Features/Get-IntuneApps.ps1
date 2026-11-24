<#
.SYNOPSIS
Loads complete Intune app inventory or assignments for one app.
.DESCRIPTION
All app types are included by default. Assignment status in inventory is a Graph hint.
Assignment inspection retains original settings and tenant-bound snapshots for safe editing.
#>
[CmdletBinding()]
param(
    [ValidateSet('Inventory', 'Assignments')][string]$Action = 'Inventory',
    [string]$AppId,
    [string]$NameFilter = '',
    [ValidateSet('All', 'Windows', 'Android', 'iOS', 'macOS')][string]$PlatformFilter = 'All',
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Shared\IntuneAppHelpers.ps1')
$tenant = Assert-IntuneAppContext -Scope 'DeviceManagementApps.Read.All' -Tenant $TenantId -Connect:(-not $SkipGraphConnect)
$base = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps'
$rows = [System.Collections.Generic.List[object]]::new()
if ($Action -eq 'Inventory') {
    foreach ($app in @(Get-IntuneAppPages -Uri $base)) {
        $platform = Get-IntuneAppPlatform $app
        if ($PlatformFilter -ne 'All' -and $platform -ne $PlatformFilter) { continue }
        if ($NameFilter -and ([string]$app.displayName).IndexOf($NameFilter, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        $rows.Add([PSCustomObject]@{
            Id = [string]$app.id; DisplayName = [string]$app.displayName; Publisher = [string]$app.publisher
            AppType = ([string]$app.'@odata.type' -replace '^#?microsoft\.graph\.', '')
            Platform = $platform; AssignmentStatus = if ($null -eq $app.isAssigned) { 'Unknown' } elseif ($app.isAssigned) { 'Assigned (hint)' } else { 'Unassigned (hint)' }
            TenantId = $tenant
        })
    }
} else {
    $parsedId = [guid]::Empty
    if (-not [guid]::TryParse($AppId, [ref]$parsedId)) { throw 'Assignments requires a valid app GUID.' }
    $groups = @{}
    foreach ($assignment in @(Get-IntuneAppPages -Uri "$base/$AppId/assignments")) {
        $target = $assignment.target
        $groupId = [string]$target.groupId
        $targetName = switch (([string]$target.'@odata.type' -replace '^#', '')) {
            'microsoft.graph.allLicensedUsersAssignmentTarget' { 'All Users' }
            'microsoft.graph.allDevicesAssignmentTarget' { 'All Devices' }
            default { [string]$target.'@odata.type' }
        }
        if ($groupId) {
            if (-not $groups.ContainsKey($groupId)) {
                try {
                    $group = Invoke-IntuneGraphWithRetry -ScriptBlock {
                        Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId`?`$select=id,displayName" -ErrorAction Stop
                    }
                    $groups[$groupId] = [string]$group.displayName
                } catch {
                    Write-IntuneAppLog "Group name unavailable for ${groupId}: $($_.Exception.Message)" 'WARN'
                    $groups[$groupId] = "Unavailable ($groupId)"
                }
            }
            $targetName = $groups[$groupId]
        }
        $source = if ($assignment.source) { [string]$assignment.source } else { 'direct' }
        $rows.Add([PSCustomObject]@{
            Id = [string]$assignment.id; AppId = $AppId; TenantId = $tenant; Intent = [string]$assignment.intent
            TargetName = $targetName; GroupId = $groupId
            IncludeExclude = if ([string]$target.'@odata.type' -match 'exclusionGroup') { 'exclude' } else { 'include' }
            FilterMode = if ($target.deviceAndAppManagementAssignmentFilterType) { [string]$target.deviceAndAppManagementAssignmentFilterType } else { 'none' }
            FilterId = [string]$target.deviceAndAppManagementAssignmentFilterId; Source = $source
            Editable = $source -eq 'direct'; Assignment = $assignment
            Fingerprint = Get-IntuneAppAssignmentFingerprint $assignment
        })
    }
}
Write-IntuneAppLog "Loaded $($rows.Count) $($Action.ToLowerInvariant()) record(s)." 'OK'
$report = [PSCustomObject]@{ Action = $Action; AppId = $AppId; TenantId = $tenant; Rows = $rows.ToArray() }
if ($PassThru) { $report } else { $report.Rows | Format-Table }
