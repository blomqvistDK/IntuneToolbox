[CmdletBinding()]
param(
    [string]$NameFilter = '',
    [ValidateSet('SettingsCatalog', 'DeviceConfiguration', 'AdministrativeTemplates')]
    [string[]]$ProfileType = @('SettingsCatalog', 'DeviceConfiguration', 'AdministrativeTemplates'),
    [ValidateSet('Any', 'None', 'Group', 'AllDevices', 'AllUsers')]
    [string]$AssignmentFilter = 'Any',
    [ValidateSet('All', 'Windows', 'iOS', 'macOS')]
    [string]$PlatformFilter = 'All',
    [switch]$IncludeSucceeded,
    [string]$Match = '',
    [string]$Replace = '',
    [string]$Prefix = '',
    [string]$Suffix = '',
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)

$requiredScopes = @('DeviceManagementConfiguration.Read.All')
if ($IncludeSucceeded) { $requiredScopes += 'DeviceManagementManagedDevices.Read.All' }

if (-not $SkipGraphConnect) {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $context = Get-MgContext
    if (-not $context -or @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count) {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ContextScope Process -ErrorAction Stop | Out-Null
    }
}

function Write-FeatureLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ($LogAction) { & $LogAction $Message $Level }
    else { Write-Information $Message -InformationAction Continue }
}

function Get-IntuneFeaturePages {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    do {
        $response = Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop
        foreach ($item in @($response.value)) { $items.Add($item) }
        $Uri = [string]$response.'@odata.nextLink'
    } while ($Uri)
    return $items.ToArray()
}

function Get-IntuneFeatureAssignments {
    param([string]$ApiType, [string]$Id)
    $uri = switch ($ApiType) {
        'configurationPolicies' { "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$Id')/assignments" }
        'deviceConfigurations' { "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$Id/assignments" }
        'groupPolicyConfigurations' { "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations/$Id/assignments" }
        default { throw "Unsupported profile API type: '$ApiType'." }
    }
    try { return @(Get-IntuneFeaturePages -Uri $uri) }
    catch { Write-FeatureLog "Could not read assignments for profile '$Id': $($_.Exception.Message)" 'WARN'; return @() }
}

function Test-IntuneAssignmentFilter {
    param([object[]]$Assignments, [string]$Filter)
    switch ($Filter) {
        'None' { return $Assignments.Count -eq 0 }
        'Group' { return @($Assignments | Where-Object { $_.target.'@odata.type' -in @('#microsoft.graph.groupAssignmentTarget', '#microsoft.graph.exclusionGroupAssignmentTarget') }).Count -gt 0 }
        'AllDevices' { return @($Assignments | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget' }).Count -gt 0 }
        'AllUsers' { return @($Assignments | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget' }).Count -gt 0 }
        default { return $true }
    }
}

function Get-IntuneAssignmentSummary {
    param([object[]]$Assignments)
    if (-not $Assignments.Count) { return 'None' }
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($Assignments | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget' }) { $parts.Add('All Devices') }
    if ($Assignments | Where-Object { $_.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget' }) { $parts.Add('All Users') }
    $groupCount = @($Assignments | Where-Object { $_.target.'@odata.type' -in @('#microsoft.graph.groupAssignmentTarget', '#microsoft.graph.exclusionGroupAssignmentTarget') }).Count
    if ($groupCount) { $parts.Add("Group ($groupCount)") }
    return $parts -join ', '
}

function Get-IntuneSucceededCount {
    param([string]$ApiType, [string]$Id)
    if ($ApiType -eq 'groupPolicyConfigurations') { return 'N/A' }
    try {
        if ($ApiType -eq 'deviceConfigurations') {
            $overview = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$Id/deviceStatusOverview" -ErrorAction Stop
            return [int]$overview.successCount
        }
        $body = @{ filter = "(PolicyId eq '$Id')"; select = @('PolicyStatus'); skip = 0; top = 10000 } | ConvertTo-Json -Depth 5
        $report = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/beta/deviceManagement/reports/getConfigurationPolicyDevicesReport' -Body $body -ContentType 'application/json' -ErrorAction Stop
        $statusIndex = [array]::IndexOf($report.Schema.Column, 'PolicyStatus')
        if ($statusIndex -lt 0) { return 0 }
        return [int]@($report.Values | Where-Object { $_[$statusIndex] -eq 'Success' }).Count
    } catch { return '—' }
}

$results = [System.Collections.Generic.List[object]]::new()
$endpoints = @()
if ('SettingsCatalog' -in $ProfileType) { $endpoints += @{ ApiType = 'configurationPolicies'; Name = 'Settings Catalog'; Uri = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' } }
if ('DeviceConfiguration' -in $ProfileType) { $endpoints += @{ ApiType = 'deviceConfigurations'; Name = 'Device Config'; Uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations' } }
if ('AdministrativeTemplates' -in $ProfileType) { $endpoints += @{ ApiType = 'groupPolicyConfigurations'; Name = 'Administrative Templates'; Uri = 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations' } }

foreach ($endpoint in $endpoints) {
    Write-FeatureLog "Fetching $($endpoint.Name) profiles..."
    $countBefore = $results.Count
    try {
        foreach ($profileItem in Get-IntuneFeaturePages -Uri $endpoint.Uri) {
            $currentName = if ($endpoint.ApiType -eq 'configurationPolicies') { [string]$profileItem.name } else { [string]$profileItem.displayName }
            if ($NameFilter -and $currentName -notlike "*$NameFilter*") { continue }
            $platform = switch ($endpoint.ApiType) {
                'configurationPolicies' { [string]$profileItem.platforms }
                'groupPolicyConfigurations' { 'Windows' }
                default { [string]$profileItem.'@odata.type' }
            }
            if ($PlatformFilter -ne 'All' -and ($platform -notmatch [regex]::Escape($PlatformFilter))) { continue }

            $assignments = @()
            $assignmentSummary = '—'
            if ($AssignmentFilter -ne 'Any') {
                $assignments = @(Get-IntuneFeatureAssignments -ApiType $endpoint.ApiType -Id $profileItem.id)
                if (-not (Test-IntuneAssignmentFilter -Assignments $assignments -Filter $AssignmentFilter)) { continue }
                $assignmentSummary = Get-IntuneAssignmentSummary -Assignments $assignments
            }

            $baseName = if ($Match -and $Replace) { $currentName.Replace($Match, $Replace) } else { $currentName }
            $newName = $Prefix + $baseName + $Suffix
            $succeeded = if ($IncludeSucceeded) { Get-IntuneSucceededCount -ApiType $endpoint.ApiType -Id $profileItem.id } else { '—' }
            $results.Add([PSCustomObject]@{
                CurrentName = $currentName
                NewName = $newName
                ProfileType = $endpoint.Name
                Assignments = $assignmentSummary
                Succeeded = $succeeded
                Compliant = '—'
                Id = [string]$profileItem.id
                ApiType = $endpoint.ApiType
            })
        }
        Write-FeatureLog "$($endpoint.Name): loaded $($results.Count - $countBefore) matching profile(s)." 'OK'
    } catch {
        Write-FeatureLog "Failed to fetch $($endpoint.Name) profiles: $($_.Exception.Message)" 'ERR'
    }
}

Write-FeatureLog "Loaded $($results.Count) profile(s)." $(if ($results.Count) { 'OK' } else { 'WARN' })
if ($PassThru) { $results.ToArray() }
else { $results | Format-Table CurrentName, ProfileType, Assignments, Succeeded }
