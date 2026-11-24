[CmdletBinding()]
param([switch]$RunScaleBenchmark)
$ErrorActionPreference = 'Stop'
$script:CheckCount = 0

function Assert-IntuneTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:CheckCount++
    Write-Host "PASS: $Message"
}

$toolboxRoot = Split-Path $PSScriptRoot -Parent
$sourcePaths = @(
    'Invoke-IntuneToolbox-GUI.ps1',
    'Shared\Invoke-IntuneToolboxOperation.ps1',
    'Shared\Invoke-IntuneGraphWithRetry.ps1',
    'Features\Get-IntuneProfiles.ps1',
    'Features\Invoke-IntuneProfileAssignment.ps1',
    'Features\Rename-IntuneProfiles.ps1',
    'Features\Remove-IntuneProfiles.ps1',
    'Features\Invoke-AutopilotGroupTagOperation.ps1',
    'Features\Search-IntuneAssignmentTargets.ps1',
    'Features\Search-IntuneSettings.ps1',
    'Features\Search-IntuneAuditEvents.ps1',
    'Shared\Get-IntuneDuplicateDeviceAnalysis.ps1',
    'Features\Get-IntuneDuplicateDevices.ps1',
    'Features\Remove-IntuneDuplicateDevices.ps1',
    'Tests\Test-IntuneDuplicateDevices.ps1',
    'Tests\Invoke-IntuneDuplicateScanDiagnostic.ps1'
)
foreach ($relativePath in $sourcePaths) {
    $tokens = $null; $parseErrors = $null
    [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $toolboxRoot $relativePath), [ref]$tokens, [ref]$parseErrors)
    Assert-IntuneTest ($parseErrors.Count -eq 0) "PowerShell syntax: $relativePath"
}

$guiAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $toolboxRoot 'Invoke-IntuneToolbox-GUI.ps1'), [ref]$null, [ref]$null)

$profileOutputAssignment = $guiAst.Find({
    param($node)
    $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$profileOutputAction'
}, $true)
$profileOutput = $profileOutputAssignment.Right.Find({
    param($node)
    $node -is [Management.Automation.Language.ScriptBlockExpressionAst]
}, $true).ScriptBlock.GetScriptBlock()
$previewBag = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
$sortProfiles = @(foreach ($countValue in @(10, 2, 0, 1, [string][char]0x2014, '-', $null, '')) {
    $sortProfile = [PSCustomObject]@{ CurrentName = "Profile $($previewBag.Count)"; Succeeded = $countValue }
    & $profileOutput $sortProfile
    $expectedSortValue = if ($countValue -is [int]) { $countValue } else { -1 }
    Assert-IntuneTest ($sortProfile.SucceededSortValue -is [int] -and
        $sortProfile.SucceededSortValue -eq $expectedSortValue -and
        [object]::Equals($sortProfile.Succeeded, $countValue)) 'Succeeded sort key is numeric and preserves the displayed value'
    $sortProfile
})
Assert-IntuneTest ($previewBag.Count -eq $sortProfiles.Count) 'Profile output collects all normalized profiles'

$global:IntuneMockState = @{
    Writes = 0
    LastBody = $null
    Requests = [System.Collections.Generic.List[object]]::new()
    MultiGroup = $false
    FailNextWrite = ''
    FailAssignmentRead = $false
}
function global:Get-MgContext {
    [PSCustomObject]@{ Account = 'tester@example.test'; TenantId = '11111111-1111-1111-1111-111111111111'; Scopes = @('DeviceManagementConfiguration.Read.All') }
}
function global:Connect-MgGraph { throw 'Unexpected sign-in in an offline test.' }
function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param($Method, $Uri, $Body, $ContentType, $Headers)
    $global:IntuneMockState.Requests.Add([PSCustomObject]@{ Method = $Method; Uri = $Uri; Body = $Body })
    if ($Method -in @('POST', 'PATCH', 'DELETE')) {
        $global:IntuneMockState.Writes++
        $global:IntuneMockState.LastBody = $Body
        if ($global:IntuneMockState.FailNextWrite) {
            $failure = $global:IntuneMockState.FailNextWrite
            $global:IntuneMockState.FailNextWrite = ''
            throw $failure
        }
        return [PSCustomObject]@{}
    }
    if ($Uri -match '/assignments(\?|$)') {
        if ($global:IntuneMockState.FailAssignmentRead) { throw '403 Forbidden reading assignments' }
        if ($global:IntuneMockState.MultiGroup) {
            if ($Uri -like '*page=2') {
                return @{ value = @(
                    @{ target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-keep'; deviceAndAppManagementAssignmentFilterId = 'filter-keep'; deviceAndAppManagementAssignmentFilterType = 'exclude' } },
                    @{ target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-2' } },
                    @{ target = @{ '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget'; groupId = 'group-1' } }
                ) }
            }
            return @{ value = @(
                @{ source = 'direct'; target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-1'; deviceAndAppManagementAssignmentFilterId = 'filter-old'; deviceAndAppManagementAssignmentFilterType = 'include' } },
                @{ source = 'direct'; target = @{ '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget'; groupId = 'group-2' } },
                @{ source = 'policySets'; target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-policyset' } }
            ); '@odata.nextLink' = "${Uri}?page=2" }
        }
    }
    if ($Uri -like '*configurationPolicies*assignments*') {
        return [PSCustomObject]@{ value = @(
            [PSCustomObject]@{ source = 'direct'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-direct' } },
            [PSCustomObject]@{ source = 'policySets'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-policyset' } }
        ) }
    }
    if ($Uri -like '*configurationPolicies*settings*') {
        return [PSCustomObject]@{ value = @([PSCustomObject]@{
            settingDefinitions = @([PSCustomObject]@{ id = 'definition-company'; displayName = 'Company Name'; options = @() })
            settingInstance = [PSCustomObject]@{ settingDefinitionId = 'definition-company'; simpleSettingValue = [PSCustomObject]@{ value = 'Example Corp' } }
        }); '@odata.nextLink' = $null }
    }
    if ($Uri -like '*configurationPolicies*') {
        return [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'profile-1'; name = 'Company Policy'; platforms = 'windows10' }); '@odata.nextLink' = $null }
    }
    if ($Uri -like '*deviceConfigurations*') {
        return [PSCustomObject]@{ value = @(); '@odata.nextLink' = $null }
    }
    if ($Uri -like '*groupPolicyConfigurations*assignments*') {
        return @{ value = @(
            @{ target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-template' } },
            @{ target = @{ '@odata.type' = '#microsoft.graph.allDevicesAssignmentTarget' } },
            @{ target = @{ '@odata.type' = '#microsoft.graph.allLicensedUsersAssignmentTarget' } }
        ) }
    }
    if ($Uri -like '*groupPolicyConfigurations*') {
        if ($Uri -like '*page=2') {
            return @{ value = @(@{ id = 'template-2'; displayName = 'Template BuiltIn'; '@odata.type' = '#microsoft.graph.groupPolicyConfiguration'; policyConfigurationIngestionType = 'builtIn' }) }
        }
        return @{ value = @(@{ id = 'template-1'; displayName = 'Template Firefox'; '@odata.type' = '#microsoft.graph.groupPolicyConfiguration'; policyConfigurationIngestionType = 'custom' })
            '@odata.nextLink' = "${Uri}?page=2" }
    }
    if ($Uri -like '*windowsAutopilotDeviceIdentities*') {
        return [PSCustomObject]@{ value = @(
            [PSCustomObject]@{ id = 'device-1'; serialNumber = 'SERIAL-1'; model = 'Model'; manufacturer = 'Vendor'; displayName = 'PC-1'; groupTag = 'Pilot' }
        ); '@odata.nextLink' = $null }
    }
    if ($Uri -like '*assignmentFilters*') {
        return [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'filter-1'; displayName = 'Windows Filter'; platform = 'windows10'; rule = '(device.deviceOSType -eq "Windows")' }) }
    }
    if ($Uri -like '*groups*') {
        return [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'group-1'; displayName = 'Pilot Group' }) }
    }
    throw "Unexpected mocked Graph request: $Method $Uri"
}

$profilePath = Join-Path $toolboxRoot 'Features\Get-IntuneProfiles.ps1'
$profiles = @(& $profilePath -NameFilter 'Company' -ProfileType SettingsCatalog -AssignmentFilter Group -Match 'Company' -Replace 'Pilot' -Prefix 'New ' -SkipGraphConnect -PassThru)
Assert-IntuneTest ($profiles.Count -eq 1 -and $profiles[0].CurrentName -eq 'Company Policy' -and $profiles[0].NewName -eq 'New Pilot Policy' -and $profiles[0].Assignments -eq 'Group (2)') 'Profile filters, assignment summary, and preview naming'

$assignmentPath = Join-Path $toolboxRoot 'Features\Invoke-IntuneProfileAssignment.ps1'
$global:IntuneMockState.Writes = 0
    $assignmentPreview = @(& $assignmentPath -Profiles $profiles[0] -Action Add -TargetType Users -SkipGraphConnect -PassThru -WhatIf)
Assert-IntuneTest ($assignmentPreview.Count -eq 1 -and $assignmentPreview[0].Status -eq 'WhatIf' -and $global:IntuneMockState.Writes -eq 0) 'Assignment WhatIf makes no Graph writes'
$assignmentResult = @(& $assignmentPath -Profiles $profiles[0] -Action Add -TargetType Users -SkipGraphConnect -PassThru -Confirm:$false)
$assignmentBody = $global:IntuneMockState.LastBody | ConvertFrom-Json
Assert-IntuneTest ($assignmentResult[0].Status -eq 'OK' -and $assignmentBody.assignments.Count -eq 2 -and $assignmentBody.assignments.target.groupId -notcontains 'group-policyset') 'Assignment add preserves direct targets and excludes policy-set targets'

$renamePath = Join-Path $toolboxRoot 'Features\Rename-IntuneProfiles.ps1'
$renamePreview = @(& $renamePath -Profiles $profiles[0] -Prefix 'Preview ' -SkipGraphConnect -PassThru -WhatIf)
Assert-IntuneTest ($renamePreview[0].Status -eq 'WhatIf' -and $renamePreview[0].NewName -eq 'Preview Company Policy') 'Rename WhatIf returns proposed names without writing'
$deletePath = Join-Path $toolboxRoot 'Features\Remove-IntuneProfiles.ps1'
$deletePreview = @(& $deletePath -Profiles $profiles[0] -SkipGraphConnect -PassThru -WhatIf)
Assert-IntuneTest ($deletePreview[0].Status -eq 'WhatIf' -and $global:IntuneMockState.Writes -eq 1) 'Delete WhatIf makes no Graph writes'

$templates = @(& $profilePath -ProfileType AdministrativeTemplates -PlatformFilter Windows -IncludeSucceeded -Match Template -Replace Pilot -Prefix 'New ' -SkipGraphConnect -PassThru)
Assert-IntuneTest ($templates.Count -eq 2 -and $templates[0].CurrentName -eq 'Template Firefox' -and
    $templates[0].NewName -eq 'New Pilot Firefox' -and $templates[0].ApiType -eq 'groupPolicyConfigurations' -and
    $templates[0].ProfileType -eq 'Administrative Templates' -and @($templates | Where-Object Succeeded -ne 'N/A').Count -eq 0) 'Administrative Templates inventory follows pages, normalizes names, and reports N/A'
Assert-IntuneTest (@($global:IntuneMockState.Requests | Where-Object Uri -match 'groupPolicyConfigurations.*(deviceStatusOverview|Report)').Count -eq 0) 'Administrative Templates never use unrelated succeeded reporting endpoints'
$defaultProfiles = @(& $profilePath -SkipGraphConnect -PassThru)
Assert-IntuneTest ($defaultProfiles.Count -eq 3 -and @($defaultProfiles | Where-Object ApiType -eq 'groupPolicyConfigurations').Count -eq 2) 'Default profile types include Administrative Templates'
foreach ($platform in @('All', 'Windows', 'iOS', 'macOS')) {
    $filteredTemplates = @(& $profilePath -ProfileType AdministrativeTemplates -PlatformFilter $platform -NameFilter Firefox -SkipGraphConnect -PassThru)
    $expectedCount = if ($platform -in @('All', 'Windows')) { 1 } else { 0 }
    Assert-IntuneTest ($filteredTemplates.Count -eq $expectedCount) "Administrative Templates name and platform filter: $platform"
}
foreach ($assignmentFilter in @('Any', 'Group', 'AllDevices', 'AllUsers', 'None')) {
    $filteredTemplates = @(& $profilePath -ProfileType AdministrativeTemplates -AssignmentFilter $assignmentFilter -SkipGraphConnect -PassThru)
    $expectedCount = if ($assignmentFilter -eq 'None') { 0 } else { 2 }
    Assert-IntuneTest ($filteredTemplates.Count -eq $expectedCount) "Administrative Templates assignment filter: $assignmentFilter"
    if ($assignmentFilter -eq 'Group') {
        Assert-IntuneTest ($filteredTemplates[0].Assignments -eq 'All Devices, All Users, Group (1)') 'Administrative Templates assignment summary includes all target types'
    }
}
$global:IntuneMockState.Requests.Clear()
$renamedTemplate = @(& $renamePath -Profiles $templates[0] -Prefix 'Pilot ' -SkipGraphConnect -PassThru -Confirm:$false)
$renameRequest = $global:IntuneMockState.Requests[0]
Assert-IntuneTest ($renamedTemplate[0].Status -eq 'OK' -and $renameRequest.Method -eq 'PATCH' -and
    $renameRequest.Uri -eq 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations/template-1' -and
    ($renameRequest.Body | ConvertFrom-Json).displayName -eq 'Pilot Template Firefox' -and
    @(($renameRequest.Body | ConvertFrom-Json).PSObject.Properties).Count -eq 1) 'Administrative Templates rename uses its own endpoint and displayName-only payload'
$global:IntuneMockState.Requests.Clear()
$global:IntuneMockState.FailNextWrite = '400 Bad Request'
$failedRename = @(& $renamePath -Profiles $templates[0] -Prefix 'Pilot ' -SkipGraphConnect -PassThru -Confirm:$false)
Assert-IntuneTest ($failedRename[0].Status -eq 'Failed' -and $global:IntuneMockState.Requests.Count -eq 1) 'Administrative Templates rename failure never enters Device Configuration fallback'
$global:IntuneMockState.Requests.Clear()
$deletedTemplate = @(& $deletePath -Profiles $templates[0] -SkipGraphConnect -PassThru -Confirm:$false)
Assert-IntuneTest ($deletedTemplate[0].Status -eq 'OK' -and $global:IntuneMockState.Requests[0].Method -eq 'DELETE' -and
    $global:IntuneMockState.Requests[0].Uri -eq 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations/template-1') 'Administrative Templates delete uses its own endpoint'

$global:IntuneMockState.MultiGroup = $true
$multiGroups = @(
    [PSCustomObject]@{ Id = 'group-1'; DisplayName = 'First group' },
    [PSCustomObject]@{ Id = 'group-2'; DisplayName = 'Second group' },
    [PSCustomObject]@{ Id = 'GROUP-1'; DisplayName = 'Duplicate ID' }
)
foreach ($apiType in @('configurationPolicies', 'deviceConfigurations', 'groupPolicyConfigurations')) {
    $testProfile = [PSCustomObject]@{ Id = 'test-profile'; CurrentName = 'Test profile'; ProfileType = $apiType; ApiType = $apiType }
    $expectedBase = if ($apiType -eq 'configurationPolicies') { "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('test-profile')" }
        else { "https://graph.microsoft.com/beta/deviceManagement/$apiType/test-profile" }
    foreach ($action in @('Add', 'Replace')) {
        foreach ($direction in @('include', 'exclude')) {
            $global:IntuneMockState.Requests.Clear()
            $multiResult = @(& $assignmentPath -Profiles $testProfile -Action $action -TargetType Group -TargetGroup $multiGroups `
                -IncludeExclude $direction -AssignmentFilter @{ Id = 'filter-new' } -FilterMode include -SkipGraphConnect -PassThru -Confirm:$false)
            $posts = @($global:IntuneMockState.Requests | Where-Object Method -eq POST)
            $payload = $posts[0].Body | ConvertFrom-Json
            $targetOdata = if ($direction -eq 'include') { '#microsoft.graph.groupAssignmentTarget' } else { '#microsoft.graph.exclusionGroupAssignmentTarget' }
            $newTargets = @($payload.assignments.target | Where-Object { $_.'@odata.type' -eq $targetOdata -and $_.groupId -in $multiGroups.Id })
            $expectedCount = if ($action -eq 'Add') { 5 } else { 2 }
            Assert-IntuneTest ($multiResult[0].Status -eq 'OK' -and $posts.Count -eq 1 -and $posts[0].Uri -eq "$expectedBase/assign" -and
                $payload.assignments.Count -eq $expectedCount -and $newTargets.Count -eq 2 -and
                @($newTargets | Where-Object deviceAndAppManagementAssignmentFilterId -ne 'filter-new').Count -eq 0 -and
                $payload.assignments.target.groupId -notcontains 'group-policyset') "$apiType $action $direction deduplicates all groups into one complete filtered assignment payload"
            $getCount = @($global:IntuneMockState.Requests | Where-Object Method -eq GET).Count
            Assert-IntuneTest ($getCount -eq $(if ($action -eq 'Add') { 2 } else { 0 })) "$apiType $action reads every assignment page only when needed"
            if ($action -eq 'Add') {
                $kept = @($payload.assignments.target | Where-Object groupId -eq 'group-keep')
                Assert-IntuneTest ($kept.Count -eq 1 -and $kept[0].deviceAndAppManagementAssignmentFilterId -eq 'filter-keep' -and
                    $kept[0].deviceAndAppManagementAssignmentFilterType -eq 'exclude') "$apiType Add preserves unrelated direct assignments and filters from later pages"
            }
        }
    }
    foreach ($removeType in @('include', 'exclude', 'both')) {
        $global:IntuneMockState.Requests.Clear()
        $multiResult = @(& $assignmentPath -Profiles $testProfile -Action Remove -RemoveScope Group -RemoveType $removeType -RemoveGroup $multiGroups -SkipGraphConnect -PassThru -Confirm:$false)
        Assert-IntuneTest ($multiResult[0].Status -eq 'OK') "$apiType Remove $removeType succeeds: $($multiResult[0].Error)"
        $posts = @($global:IntuneMockState.Requests | Where-Object Method -eq POST)
        $payload = $posts[0].Body | ConvertFrom-Json
        $expectedCount = if ($removeType -eq 'both') { 1 } else { 3 }
        Assert-IntuneTest ($multiResult[0].Status -eq 'OK' -and $posts.Count -eq 1 -and $posts[0].Uri -eq "$expectedBase/assign" -and
            $payload.assignments.Count -eq $expectedCount -and $payload.assignments.target.groupId -contains 'group-keep' -and
            $payload.assignments.target.groupId -notcontains 'group-policyset') "$apiType Remove $removeType removes all requested groups and preserves unrelated direct targets"
    }
    foreach ($targetType in @('Users', 'Devices')) {
        $global:IntuneMockState.Requests.Clear()
        $allResult = @(& $assignmentPath -Profiles $testProfile -Action Replace -TargetType $targetType -SkipGraphConnect -PassThru -Confirm:$false)
        $payload = $global:IntuneMockState.LastBody | ConvertFrom-Json
        $expectedType = if ($targetType -eq 'Users') { '#microsoft.graph.allLicensedUsersAssignmentTarget' } else { '#microsoft.graph.allDevicesAssignmentTarget' }
        Assert-IntuneTest ($allResult[0].Status -eq 'OK' -and $payload.assignments.Count -eq 1 -and
            $payload.assignments[0].target.'@odata.type' -eq $expectedType) "$apiType All $targetType retains single-target behavior"
    }
    $removedAll = @(& $assignmentPath -Profiles $testProfile -Action Remove -RemoveScope All -SkipGraphConnect -PassThru -Confirm:$false)
    Assert-IntuneTest ($removedAll[0].Status -eq 'OK' -and ($global:IntuneMockState.LastBody | ConvertFrom-Json).assignments.Count -eq 0) "$apiType Remove All submits an empty assignment array"
    $global:IntuneMockState.Requests.Clear()
    $preview = @(& $assignmentPath -Profiles $testProfile -Action Replace -TargetType Group -TargetGroup $multiGroups -WhatIf -SkipGraphConnect -PassThru)
    $renamePreview = @(& $renamePath -Profiles $testProfile -Prefix 'Preview ' -WhatIf -SkipGraphConnect -PassThru)
    $deletePreview = @(& $deletePath -Profiles $testProfile -WhatIf -SkipGraphConnect -PassThru)
    Assert-IntuneTest ($preview[0].Status -eq 'WhatIf' -and $renamePreview[0].Status -eq 'WhatIf' -and
        $deletePreview[0].Status -eq 'WhatIf' -and $global:IntuneMockState.Requests.Count -eq 0) "$apiType WhatIf performs no Graph calls for assignment, rename, or delete"
}
$global:IntuneMockState.Requests.Clear()
$singleGroupResult = @(& $assignmentPath -Profiles $testProfile -Action Replace -TargetType Group -TargetGroup $multiGroups[0] -SkipGraphConnect -PassThru -Confirm:$false)
Assert-IntuneTest ($singleGroupResult[0].Status -eq 'OK' -and ($global:IntuneMockState.LastBody | ConvertFrom-Json).assignments.Count -eq 1) 'Existing single TargetGroup argument remains supported'
$singleRemove = @(& $assignmentPath -Profiles $testProfile -Action Remove -RemoveScope Group -RemoveGroup $multiGroups[0] -SkipGraphConnect -PassThru -Confirm:$false)
Assert-IntuneTest ($singleRemove[0].Status -eq 'OK' -and ($global:IntuneMockState.LastBody | ConvertFrom-Json).assignments.Count -eq 3) 'Existing single RemoveGroup argument remains supported'
$global:IntuneMockState.Requests.Clear()
$noMatch = @(& $assignmentPath -Profiles $testProfile -Action Remove -RemoveScope Group -RemoveGroup @{ Id = 'absent' } -SkipGraphConnect -PassThru -Confirm:$false)
Assert-IntuneTest ($noMatch[0].Status -eq 'OK' -and @($global:IntuneMockState.Requests | Where-Object Method -eq POST).Count -eq 0) 'Unmatched group removal performs no writes'
foreach ($badParameters in @(
    @{ Action = 'Add'; TargetType = 'Group'; TargetGroup = @() },
    @{ Action = 'Replace'; TargetType = 'Group'; TargetGroup = @($multiGroups[0], @{ Id = '' }) },
    @{ Action = 'Remove'; RemoveScope = 'Group'; RemoveGroup = @() },
    @{ Action = 'Remove'; RemoveScope = 'Group'; RemoveGroup = @{ DisplayName = 'Missing ID' } },
    @{ Action = 'Add' },
    @{ Action = 'Add'; TargetType = 'Users'; FilterMode = 'include' }
)) {
    $global:IntuneMockState.Requests.Clear()
    $rejected = $false
    try { & $assignmentPath -Profiles $testProfile @badParameters -SkipGraphConnect -PassThru -Confirm:$false | Out-Null }
    catch { $rejected = $true }
    Assert-IntuneTest ($rejected -and $global:IntuneMockState.Requests.Count -eq 0) 'Invalid assignment inputs fail explicitly before Graph calls'
}
$global:IntuneMockState.Requests.Clear()
$global:IntuneMockState.FailAssignmentRead = $true
$readFailure = @(& $assignmentPath -Profiles $testProfile -Action Add -TargetType Group -TargetGroup $multiGroups -SkipGraphConnect -PassThru -Confirm:$false)
$global:IntuneMockState.FailAssignmentRead = $false
Assert-IntuneTest ($readFailure[0].Status -eq 'Failed' -and $readFailure[0].Error -like '*Forbidden*' -and
    @($global:IntuneMockState.Requests | Where-Object Method -eq POST).Count -eq 0) 'Failed assignment reads prevent destructive partial writes'
$global:IntuneMockState.Requests.Clear()
$global:IntuneMockState.FailNextWrite = '503 ServiceUnavailable'
$retried = @(& $assignmentPath -Profiles $testProfile -Action Replace -TargetType Group -TargetGroup $multiGroups -SkipGraphConnect -PassThru -Confirm:$false)
$retryPosts = @($global:IntuneMockState.Requests | Where-Object Method -eq POST)
Assert-IntuneTest ($retried[0].Status -eq 'OK' -and $retryPosts.Count -eq 2 -and $retryPosts[0].Body -eq $retryPosts[1].Body) 'Retried group assignment retains the complete target set'
$testProfile.ApiType = 'unsupported'
foreach ($operationPath in @($assignmentPath, $renamePath, $deletePath)) {
    $global:IntuneMockState.Requests.Clear()
    $operationParameters = if ($operationPath -eq $assignmentPath) { @{ Action = 'Replace'; TargetType = 'Users' } } else { @{} }
    $unknownResult = @(& $operationPath -Profiles $testProfile @operationParameters -SkipGraphConnect -PassThru -Confirm:$false)
    Assert-IntuneTest ($unknownResult[0].Status -eq 'Failed' -and $unknownResult[0].Error -like '*Unsupported profile API type*' -and
        $global:IntuneMockState.Requests.Count -eq 0) 'Unknown API types fail instead of falling through to Device Configuration'
}
$paginatedTemplates = @(& $profilePath -ProfileType AdministrativeTemplates -AssignmentFilter Group -SkipGraphConnect -PassThru)
Assert-IntuneTest ($paginatedTemplates[0].Assignments -eq 'Group (6)') 'Inventory summaries read every assignment page'
$global:IntuneMockState.MultiGroup = $false

$targetsPath = Join-Path $toolboxRoot 'Features\Search-IntuneAssignmentTargets.ps1'
$groups = @(& $targetsPath -Action Groups -SearchTerm Pilot -SkipGraphConnect -PassThru)
$filters = @(& $targetsPath -Action Filters -SkipGraphConnect -PassThru)
Assert-IntuneTest ($groups[0].Id -eq 'group-1' -and $filters[0].Id -eq 'filter-1') 'Assignment group and filter lookup outputs'

$settingsPath = Join-Path $toolboxRoot 'Features\Search-IntuneSettings.ps1'
$settings = @(& $settingsPath -SearchTerm 'company' -Profiles $profiles -SkipGraphConnect -PassThru)
Assert-IntuneTest ($settings.Count -eq 1 -and $settings[0].SettingName -eq 'Company Name' -and $settings[0].Value -eq 'Example Corp') 'Settings Search extracts definition names and values'

$state = [hashtable]::Synchronized(@{ RequestedTenant = ''; ExpectedTenant = '22222222-2222-2222-2222-222222222222'; Connected = $false })
$tenantRejected = $false
try {
    & (Join-Path $toolboxRoot 'Shared\Invoke-IntuneToolboxOperation.ps1') -Path $profilePath `
        -Parameters @{ ProfileType = @('SettingsCatalog'); SkipGraphConnect = $true; PassThru = $true } `
        -Modules @() -Scopes @() -State $state
} catch { $tenantRejected = $_.Exception.Message -like '*tenant*' }
Assert-IntuneTest $tenantRejected 'Shared worker rejects a context from a different tenant'

& (Join-Path $PSScriptRoot 'Test-IntuneDuplicateDevices.ps1') -RunScaleBenchmark:$RunScaleBenchmark

$auditTestRows = @(& {
    $auditPath = Join-Path $toolboxRoot 'Features\Search-IntuneAuditEvents.ps1'
    $start = [datetime]'2026-11-12'
    $end = [datetime]'2026-11-13'
    $startUtc = [TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($start, [DateTimeKind]::Unspecified))
    $endUtc = [TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($end.AddDays(1), [DateTimeKind]::Unspecified))
    $userActor = @{ userPrincipalName = 'Admin@contoso.test'; userId = 'admin-id'; ipAddress = '192.0.2.1'; applicationDisplayName = 'Intune portal' }
    $target = @{ displayName = "Pilot [*]'Policy"; resourceId = 'target-1'; type = 'Policy'; modifiedProperties = @(
        @{ displayName = 'Description'; oldValue = 'Before'; newValue = 'After' },
        @{ displayName = 'Optional'; oldValue = $null; newValue = '' }
    ) }
    $fixtures = @(
        @{ id = 'event-1'; activityDateTime = $startUtc.ToString('o'); activity = 'Update policy'; activityType = 'Update'; activityOperationType = 'PATCH'; activityResult = 'Success'; category = 'Configuration'; correlationId = 'correlation-1'; actor = $userActor; resources = @(@{ displayName = 'Other'; resourceId = 'other-id' }, $target) },
        @{ id = 'event-2'; activityDateTime = $startUtc.AddHours(1).ToString('o'); displayName = 'App operation'; actor = @{ applicationDisplayName = 'Automation'; applicationId = 'app-id' }; resources = @($target) },
        @{ id = 'event-3'; activityDateTime = $startUtc.AddHours(2).ToString('o'); actor = @{ userPrincipalName = 'Someone@contoso.test' }; resources = @($target) },
        @{ id = 'event-4'; activityDateTime = $endUtc.AddTicks(-1).ToString('o'); resources = @(@{ displayName = "Pilot [*]'Policy" }) },
        @{ id = 'event-5'; activityDateTime = $startUtc.AddHours(3).ToString('o'); actor = $userActor; resources = @() },
        @{ id = 'outside-end'; activityDateTime = $endUtc.ToString('o'); actor = $userActor; resources = @($target) },
        @{ id = 'outside-start'; activityDateTime = $startUtc.AddTicks(-1).ToString('o'); actor = $userActor; resources = @($target) }
    )
    $mock = @{ Calls = [System.Collections.Generic.List[object]]::new(); Failure = ''; Retry = $false; Sleeps = 0; Scopes = @('DeviceManagementApps.ReadWrite.All') }
    function Get-MgContext { [PSCustomObject]@{ TenantId = 'audit-tenant'; Scopes = $mock.Scopes } }
    function Connect-MgGraph { throw 'Offline audit tests must not sign in.' }
    function Start-Sleep { param($Seconds) $mock.Sleeps++ }
    function Invoke-MgGraphRequest {
        [CmdletBinding()]
        param($Uri, $Method, $OutputType)
        $mock.Calls.Add(@{ Uri = $Uri; Method = $Method })
        if ($mock.Retry) { $mock.Retry = $false; throw '503 ServiceUnavailable' }
        if ($Uri -like '*skiptoken*') {
            if ($mock.Failure -eq 'late') { throw '403 Forbidden on second audit page' }
            if ($mock.Failure -eq 'invalid') { return @{ unexpected = 'not a collection' } }
            return @{ value = @($fixtures | Select-Object -Skip 3); '@odata.nextLink' = if ($mock.Failure -eq 'repeat') { $Uri } else { $null } }
        }
        return [PSCustomObject]@{ value = @($fixtures | Select-Object -First 3); '@odata.nextLink' = 'https://graph.microsoft.com/beta/deviceManagement/auditEvents?$skiptoken=second' }
    }
    $logs = [System.Collections.Generic.List[string]]::new()
    $log = { param($Message, $Level) $logs.Add("$Level $Message") }
    $common = @{ StartDate = $start; EndDate = $end; SkipGraphConnect = $true; PassThru = $true; LogAction = $log }
    $report = & $auditPath @common -Target 'pIlOt' -AsReport
    Assert-IntuneTest ($report.Completed -and $report.Rows.Count -eq 4 -and $report.PageCount -eq 2 -and $report.ScannedCount -eq 7) 'Audit target search retrieves all pages without duplicating multi-target events'
    Assert-IntuneTest ($report.Rows[0].Id -eq 'event-4' -and $report.Rows[-1].Id -eq 'event-1') 'Audit results sort newest first with inclusive local end-date boundaries'
    $initialUri = [uri]::UnescapeDataString($mock.Calls[0].Uri)
    $expectedFilter = "activityDateTime ge $($startUtc.ToString('o').Substring(0, 19))Z and activityDateTime lt $($endUtc.ToString('o').Substring(0, 19))Z"
    Assert-IntuneTest ($initialUri -eq "https://graph.microsoft.com/beta/deviceManagement/auditEvents?`$filter=$expectedFilter" -and
        @($mock.Calls | Where-Object { $_.Method -ne 'GET' }).Count -eq 0) 'Audit request uses exact UTC date bounds and read-only GET requests'
    Assert-IntuneTest ($report.Rows[2].InitiatedBy -eq 'Automation' -and $report.Rows[0].InitiatedBy -eq 'Not supplied') 'Audit target search includes application-only and missing-actor events'
    Assert-IntuneTest ($report.Rows[3].OperationKind -eq 'Update' -and $report.Rows[2].OperationKind -eq 'Other' -and
        $report.Rows[0].OperationKind -eq 'Unknown') 'Audit rows classify PATCH as Update, unrecognized names as Other, and missing operations as Unknown'
    $kindFunction = [Management.Automation.Language.Parser]::ParseFile($auditPath, [ref]$null, [ref]$null).Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-IntuneAuditOperationKind'
    }, $true)
    . ([scriptblock]::Create($kindFunction.Extent.Text))
    $kindCases = @(
        @{ Args = @{ OperationType = 'Create' }; Expected = 'Create' },
        @{ Args = @{ OperationType = 'POST' }; Expected = 'Create' },
        @{ Args = @{ OperationType = ' delete ' }; Expected = 'Delete' },
        @{ Args = @{ OperationType = 'Patch' }; Expected = 'Update' },
        @{ Args = @{ OperationType = 'PUT' }; Expected = 'Update' },
        @{ Args = @{ OperationType = 'Action'; Activity = 'Delete device' }; Expected = 'Other' },
        @{ Args = @{ ActivityType = 'Delete DeviceConfiguration' }; Expected = 'Delete' },
        @{ Args = @{ ActivityType = ' '; Activity = 'create policy' }; Expected = 'Create' },
        @{ Args = @{ Activity = 'Wipe ManagedDevice' }; Expected = 'Other' },
        @{ Args = @{}; Expected = 'Unknown' }
    )
    $wrongKinds = @($kindCases | Where-Object { $p = $_.Args; (Get-IntuneAuditOperationKind @p) -ne $_.Expected })
    Assert-IntuneTest ($wrongKinds.Count -eq 0) 'Audit operation kind uses the operation type case-insensitively before falling back to the activity verb'
    $users = @(& $auditPath @common -User 'ADMIN')
    Assert-IntuneTest ($users.Count -eq 2 -and 'event-5' -in $users.Id) 'Audit user-only matching is case-insensitive and permits events without targets'
    $both = @(& $auditPath @common -User ' ADMIN-ID ' -Target " [*]' ")
    Assert-IntuneTest ($both.Count -eq 1 -and $both[0].Id -eq 'event-1' -and $both[0].TargetNames -like 'Other; Pilot*') 'Audit combines user ID and literal punctuation target matching with AND'
    $noMatch = & $auditPath @common -User 'Automation' -AsReport
    Assert-IntuneTest ($noMatch.Completed -and $null -ne $noMatch.Rows -and $noMatch.Rows.Count -eq 0) 'Application names do not satisfy the User filter; successful empty searches carry a complete report'
    $noMatch = & $auditPath @common -Target 'Pilot*' -AsReport
    Assert-IntuneTest ($noMatch.Rows.Count -eq 0) 'Audit partial matching treats wildcard characters literally'
    $preserved = $both[0].ResourcesJson | ConvertFrom-Json
    Assert-IntuneTest ($preserved.Count -eq 2 -and $preserved[1].modifiedProperties[0].oldValue -eq 'Before' -and
        $preserved[1].modifiedProperties[0].newValue -eq 'After' -and
        $null -eq $preserved[1].modifiedProperties[1].oldValue -and $preserved[1].modifiedProperties[1].newValue -eq '') 'Audit results retain every target and distinguish null/empty changed values'
    Assert-IntuneTest ($both[0].TimestampUtc -match '\+00:00$' -and
        [DateTimeOffset]::Parse($both[0].TimestampUtc).UtcDateTime -eq $startUtc -and
        $both[0].Event.actor.ipAddress -eq '192.0.2.1') 'Audit retains unambiguous UTC timestamps and complete actor details'
    $originalDate = $fixtures[0].activityDateTime
    foreach ($typedDate in @([DateTimeOffset]$startUtc, $startUtc, $startUtc.ToLocalTime(),
        [datetime]::SpecifyKind($startUtc, [DateTimeKind]::Unspecified))) {
        $fixtures[0].activityDateTime = $typedDate
        $typedRows = @(& $auditPath @common -User 'admin-id' -Target Pilot)
        Assert-IntuneTest ($typedRows.Count -eq 1 -and $typedRows[0].Timestamp.UtcDateTime -eq $startUtc) 'Audit handles SDK timestamps as DateTimeOffset, UTC/local DateTime, and unspecified UTC without culture conversion'
    }
    $fixtures[0].activityDateTime = $originalDate
    foreach ($invalid in @(
        @{ User = ' '; Target = "`t"; StartDate = $start; EndDate = $end },
        @{ Target = 'Pilot'; StartDate = $end; EndDate = $start },
        @{ Target = 'Pilot'; StartDate = 'not-a-date'; EndDate = $end },
        @{ Target = 'Pilot'; StartDate = ''; EndDate = $end },
        @{ Target = 'Pilot'; StartDate = [datetime]::MinValue; EndDate = $end }
    )) {
        $before = $mock.Calls.Count
        $rejected = $false
        try { & $auditPath @invalid -SkipGraphConnect -PassThru -LogAction $log | Out-Null } catch { $rejected = $true }
        Assert-IntuneTest ($rejected -and $mock.Calls.Count -eq $before) 'Invalid audit criteria/date rejected before network activity'
    }
    $mock.Scopes = @('User.Read')
    $rejected = $false
    try { & $auditPath @common -Target Pilot | Out-Null } catch { $rejected = $_.Exception.Message -like '*DeviceManagementApps.Read.All*' }
    Assert-IntuneTest $rejected 'Audit requires the Apps read permission'
    $mock.Scopes = @('DeviceManagementApps.Read.All')
    $rejected = $false
    try { & $auditPath @common -Target Pilot -TenantId 'another-tenant' | Out-Null } catch { $rejected = $_.Exception.Message -like '*tenant*' }
    Assert-IntuneTest $rejected 'Standalone audit search rejects a different signed-in tenant'
    $mock.Retry = $true
    $retried = & $auditPath @common -Target Pilot -AsReport
    Assert-IntuneTest ($retried.Completed -and $retried.Rows.Count -eq 4 -and $mock.Sleeps -eq 1 -and
        @($logs | Where-Object { $_ -like '*retry 1*' }).Count -eq 1) 'Audit reuses transient retry handling and reports retries'
    foreach ($failure in @('late', 'invalid', 'repeat')) {
        $mock.Failure = $failure
        $published = [System.Collections.Generic.List[object]]::new()
        $failed = $false
        try { & $auditPath @common -Target Pilot -AsReport | ForEach-Object { $published.Add($_) } } catch { $failed = $true }
        Assert-IntuneTest ($failed -and $published.Count -eq 0) "Audit $failure page failure publishes no partial or successful-empty report"
    }
    $auditAst = [Management.Automation.Language.Parser]::ParseFile($auditPath, [ref]$null, [ref]$null)
    $dateFunction = $auditAst.Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-IntuneAuditDateWindow'
    }, $true)
    . ([scriptblock]::Create($dateFunction.Extent.Text))
    $dstZone = [TimeZoneInfo]::FindSystemTimeZoneById('W. Europe Standard Time')
    foreach ($dstCase in @(
        @{ Date = [datetime]'2026-03-29'; Hours = 23; Start = '2026-03-28T23:00:00Z'; End = '2026-03-29T22:00:00Z' },
        @{ Date = [datetime]'2026-10-25'; Hours = 25; Start = '2026-10-24T22:00:00Z'; End = '2026-10-25T23:00:00Z' }
    )) {
        $bounds = Get-IntuneAuditDateWindow -StartDate $dstCase.Date -EndDate $dstCase.Date -TimeZone $dstZone
        Assert-IntuneTest (($bounds.EndUtc - $bounds.StartUtc).TotalHours -eq $dstCase.Hours -and
            ($bounds.StartUtc.ToString('o').Substring(0, 19) + 'Z') -eq $dstCase.Start -and
            ($bounds.EndUtc.ToString('o').Substring(0, 19) + 'Z') -eq $dstCase.End) 'Local audit calendar day converts each boundary independently across daylight saving'
    }
    $report.Rows
})

$ui = . (Join-Path $toolboxRoot 'Invoke-IntuneToolbox-GUI.ps1') -ValidateOnly
Assert-IntuneTest (-not $ui.Connected -and $ui.MissingFeatures.Count -eq 0) 'ValidateOnly builds disconnected UI and resolves all feature paths'

function New-IntuneAPTestRow {
    param([string]$Serial, [string]$Tag = 'Pilot', [string]$ManagedDeviceId = 'managed-1', [string]$UpdateStatus = '', [bool]$Eligible = $true)
    [PSCustomObject]@{
        RecordType = 'Device'; RequestedSerial = $Serial; SerialNumber = $Serial; DeviceName = "PC-$Serial"; GroupTag = $Tag
        CsvTag = ''; CsvIssue = ''; LookupStatus = $(if ($Eligible) { 'Found' } else { 'Not found' }); Eligible = $Eligible
        AutopilotId = $(if ($Eligible) { "ap-$Serial" } else { '' }); ManagedDeviceId = $ManagedDeviceId
        UpdateStatus = $UpdateStatus; FreshStartStatus = ''; Details = ''
    }
}
Assert-IntuneTest ($ListDevices.View.Columns.Count -eq 11 -and $script:APActionButtons.Count -eq 11 -and
    $BtnAPFreshStart -and $ChkAPSync -and $BtnAPImportCsv -and -not $BtnAPLookup.IsEnabled -and
    $AutopilotEmptyState.Visibility -eq 'Visible') 'Autopilot tab exposes the group tag tool controls with lookup disabled until input'
$TxtAPSerials.Text = "abc1, ABC1;`r`nxyz2"
Assert-IntuneTest ((@(Get-APSerialInput $TxtAPSerials.Text) -join '|') -eq 'abc1|xyz2' -and $BtnAPLookup.IsEnabled) 'Autopilot serial input splits, de-duplicates case-insensitively, and enables lookup'
$TxtAPSerials.Text = ''

$script:APOperationContext = @{ ReplaceAll = $true }
$script:APOperationResults = [System.Collections.Generic.List[object]]::new()
Complete-APDeviceLookup
Assert-IntuneTest ($script:APRows.Count -eq 0) 'Autopilot empty lookup result leaves the device list unchanged'
$script:APOperationResults.Add((New-IntuneAPTestRow 'SER-1'))
$script:APOperationResults.Add((New-IntuneAPTestRow 'SER-2' -ManagedDeviceId '' -Tag 'Production'))
$script:APOperationResults.Add((New-IntuneAPTestRow 'SER-3' -Eligible $false))
Complete-APDeviceLookup
Assert-IntuneTest ($script:APRows.Count -eq 3 -and $ListDevices.Items.Count -eq 3 -and $TxtCountFound.Text -eq '2' -and
    $AutopilotEmptyState.Visibility -eq 'Collapsed' -and $BtnAPExport.IsEnabled) 'Autopilot lookup completion binds rows and counters'
$script:APOperationContext = @{}
$script:APOperationResults = [System.Collections.Generic.List[object]]::new()
$script:APOperationResults.Add((New-IntuneAPTestRow 'ser-1' -Tag 'Updated'))
Complete-APDeviceLookup
Assert-IntuneTest ($script:APRows.Count -eq 3 -and $script:APRows[0].GroupTag -eq 'Updated') 'Autopilot lookup merges refreshed rows by normalized serial'

$TxtAPDeviceFilter.Text = 'production'
Assert-IntuneTest ($ListDevices.Items.Count -eq 1 -and $ListDevices.Items[0].RequestedSerial -eq 'SER-2') 'Autopilot device filter matches tags case-insensitively'
$TxtAPDeviceFilter.Text = ''
foreach ($item in @($ListDevices.Items)) { [void]$ListDevices.SelectedItems.Add($item) }
$TxtNewTag.Text = 'Updated'
$apPlan = Get-APUpdatePlan -Rows @(Get-APSelectedEligibleRows) -Mode ApplyTag -Tag $TxtNewTag.Text
Assert-IntuneTest (@(Get-APSelectedEligibleRows).Count -eq 2 -and $apPlan.Changes.Count -eq 1 -and $apPlan.Changes[0].Serial -eq 'SER-2' -and
    $apPlan.Skipped.Count -eq 1 -and $BtnAPApplyTag.IsEnabled -and $BtnAPApplyTag.Content -eq 'Apply to 2 selected') 'Autopilot tag plan ignores ineligible rows and skips unchanged tags'
Assert-IntuneTest (@(Get-APSelectedFreshStartRows).Count -eq 1 -and @(Get-APSelectedFreshStartRows)[0].RequestedSerial -eq 'ser-1' -and
    $BtnAPFreshStart.IsEnabled) 'Autopilot Fresh Start only targets found rows with a managed device'
$TxtNewTag.Text = 'x' * 129
Assert-IntuneTest (-not $BtnAPApplyTag.IsEnabled -and $TxtAPTagHint.Text -like '*exceeds 128*') 'Autopilot rejects group tags longer than 128 characters'
$TxtNewTag.Text = ''
$apExportPath = Join-Path ([IO.Path]::GetTempPath()) ("IntuneToolboxAP-" + [guid]::NewGuid().ToString('N') + '.csv')
try {
    Export-APDeviceRows -Path $apExportPath
    $apExported = @(Import-Csv -LiteralPath $apExportPath)
    Assert-IntuneTest ($apExported.Count -eq 3 -and $apExported[1].GroupTag -eq 'Production') 'Autopilot export writes every loaded row'
} finally { Remove-Item -LiteralPath $apExportPath -ErrorAction SilentlyContinue }
$script:APRows.Clear()
Update-APDeviceView
function Find-IntuneColumnHeader {
    param($Visual, $Column)
    if ($Visual -is [System.Windows.Controls.GridViewColumnHeader] -and $Visual.Column -eq $Column) { return $Visual }
    for ($childIndex = 0; $childIndex -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Visual); $childIndex++) {
        $foundHeader = Find-IntuneColumnHeader -Visual ([System.Windows.Media.VisualTreeHelper]::GetChild($Visual, $childIndex)) -Column $Column
        if ($foundHeader) { return $foundHeader }
    }
}
function Test-IntuneLoadingTheme {
    $savedTab = $MainTabs.SelectedIndex
    $savedTheme = $script:IsDarkMode
    try {
        foreach ($case in @(
            @{ Name = 'ListApps'; Tab = 5 },
            @{ Name = 'ListProfiles'; Tab = 0 },
            @{ Name = 'ListGroups'; Tab = 0 }
        )) {
            $control = $window.FindName($case.Name)
            $parent = [System.Windows.LogicalTreeHelper]::GetParent($control)
            $savedItems = @($control.Items)
            $savedSelection = @($control.SelectedItems)
            $savedEnabled = $control.IsEnabled
            $savedParentEnabled = $parent.IsEnabled
            try {
                $MainTabs.SelectedIndex = $case.Tab
                $control.Items.Clear()
                foreach ($populated in @($false, $true)) {
                    if ($populated) {
                        [void]$control.Items.Add([PSCustomObject]@{ CurrentName = 'Theme B'; DisplayName = 'Theme B' })
                        [void]$control.Items.Add([PSCustomObject]@{ CurrentName = 'Theme A'; DisplayName = 'Theme A' })
                    }
                    for ($themePass = 0; $themePass -lt 2; $themePass++) {
                        foreach ($state in @('Enabled', 'Disabled', 'InheritedDisabled', 'Restored')) {
                            $parent.IsEnabled = $state -ne 'InheritedDisabled'
                            $control.IsEnabled = $state -ne 'Disabled'
                            [void]$control.ApplyTemplate()
                            $window.UpdateLayout()
                            $surface = [System.Windows.Media.VisualTreeHelper]::GetChild($control, 0)
                            Assert-IntuneTest ($surface.Background.Color -eq $window.Resources['InputBgBrush'].Color -and
                                $surface.BorderBrush.Color -eq $window.Resources['BorderBrush2'].Color) "$($case.Name) rendered surface retains theme colors: dark=$script:IsDarkMode populated=$populated state=$state"
                            if ($control -is [System.Windows.Controls.ListView]) {
                                $header = Find-IntuneColumnHeader $control $control.View.Columns[0]
                                Assert-IntuneTest ($null -ne $header) "$($case.Name) retains GridView headers"
                                $headerSurface = $header.Template.FindName('HeaderBorder', $header)
                                Assert-IntuneTest ($headerSurface.Background.Color -eq $window.Resources['SurfaceBrush'].Color) "$($case.Name) header retains the theme surface in $state"
                            }
                        }
                        $control.IsEnabled = $false
                        $BtnThemeToggle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        $window.UpdateLayout()
                        Assert-IntuneTest ($surface.Background.Color -eq $window.Resources['InputBgBrush'].Color) "$($case.Name) disabled surface updates when the theme changes"
                        $control.IsEnabled = $true
                    }
                }
                [void]$control.SelectedItems.Add($control.Items[0])
                [void]$control.SelectedItems.Add($control.Items[1])
                Assert-IntuneTest ($control.SelectedItems.Count -eq 2) "$($case.Name) retains multi-selection"
                if ($control -is [System.Windows.Controls.ListView]) {
                    $header = Find-IntuneColumnHeader $control $control.View.Columns[0]
                    $gripper = $header.Template.FindName('PART_HeaderGripper', $header)
                    Assert-IntuneTest ($null -ne $gripper) "$($case.Name) retains its column resize grip"
                    $columnWidth = $header.Column.ActualWidth
                    $drag = [System.Windows.Controls.Primitives.DragDeltaEventArgs]::new(12, 0)
                    $drag.RoutedEvent = [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent
                    $gripper.RaiseEvent($drag)
                    Assert-IntuneTest ($header.Column.Width -eq ($columnWidth + 12)) "$($case.Name) column resize grip changes the column width"
                    $header.Column.Width = $columnWidth
                }
            } finally {
                $parent.IsEnabled = $savedParentEnabled
                $control.IsEnabled = $savedEnabled
                $control.Items.Clear()
                foreach ($item in $savedItems) { [void]$control.Items.Add($item) }
                foreach ($item in $savedSelection) { [void]$control.SelectedItems.Add($item) }
            }
        }
    } finally {
        if ($script:IsDarkMode -ne $savedTheme) {
            $BtnThemeToggle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        }
        $MainTabs.SelectedIndex = $savedTab
    }
}
function Wait-IntuneBackground {
    param([scriptblock]$Until, [int]$TimeoutSeconds = 20)
    $waitTimer = [Diagnostics.Stopwatch]::StartNew()
    while (-not (& $Until) -and $waitTimer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $frame = [System.Windows.Threading.DispatcherFrame]::new()
        [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
            [System.Windows.Threading.DispatcherPriority]::Background, [Action]{ $frame.Continue = $false })
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
        Start-Sleep -Milliseconds 50
    }
}
function Get-IntuneLogText {
    param($LogBox)
    [System.Windows.Documents.TextRange]::new($LogBox.Document.ContentStart, $LogBox.Document.ContentEnd).Text
}
function New-IntuneTestCompletion {
    param([switch]$WorkerFailed, [hashtable]$CancelHandle)
    $backgroundFunction = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-InBackground'
    }, $true)
    $captureStatements = @($backgroundFunction.Body.EndBlock.Statements | Where-Object {
        $_.Extent.Text -like '$captured*' -or $_.Extent.Text -like '$timer.Add_Tick(*'
    })
    $Buttons = @()
    $savedContent = @{}
    $savedEnabled = @{}
    $OnComplete = $null
    $OverlayMessage = ''
    $handle = [PSCustomObject]@{ IsCompleted = $true }
    $ps = [PSCustomObject]@{ WorkerFailed = [bool]$WorkerFailed }
    $ps | Add-Member -MemberType ScriptMethod -Name EndInvoke -Value { param($handle) if ($this.WorkerFailed) { throw 'Mock worker failure' } }
    $ps | Add-Member -MemberType ScriptMethod -Name Dispose -Value {}
    $pool = [PSCustomObject]@{}
    $pool | Add-Member -MemberType ScriptMethod -Name Close -Value {}
    $pool | Add-Member -MemberType ScriptMethod -Name Dispose -Value {}
    $timer = [PSCustomObject]@{ Callback = $null }
    $timer | Add-Member -MemberType ScriptMethod -Name Stop -Value {}
    $timer | Add-Member -MemberType ScriptMethod -Name Add_Tick -Value { param($callback) $this.Callback = $callback }
    . ([scriptblock]::Create(($captureStatements.Extent.Text -join "`n")))
    return $timer.Callback
}
try {
    $availabilityButton = [System.Windows.Controls.Button]::new()
    $availabilityButton.Name = 'TestAvailability'
    foreach ($availabilityCase in @(
        @{ Available = $true; BusyCount = 0; Expected = $true; Name = 'idle available button' },
        @{ Available = $false; BusyCount = 0; Expected = $false; Name = 'idle unavailable button' },
        @{ Available = $true; BusyCount = 1; Expected = $false; Name = 'busy available button' },
        @{ Available = $true; BusyCount = 2; Expected = $false; Name = 'button shared by busy operations' },
        @{ Available = $false; BusyCount = 1; Expected = $false; Name = 'busy unavailable button' },
        @{ Available = $true; BusyCount = 0; Expected = $true; Name = 'available button after operations finish' }
    )) {
        if ($availabilityCase.BusyCount -gt 0) {
            $script:BusyButtonCounts[$availabilityButton.Name] = $availabilityCase.BusyCount
        } else {
            [void]$script:BusyButtonCounts.Remove($availabilityButton.Name)
        }
        Set-ButtonAvailability -Button $availabilityButton -Available $availabilityCase.Available
        Assert-IntuneTest ($availabilityButton.IsEnabled -eq $availabilityCase.Expected) "Button availability: $($availabilityCase.Name)"
    }

    $ui.Window.ShowActivated = $false
    $ui.Window.ShowInTaskbar = $false
    $ui.Window.WindowStartupLocation = 'Manual'
    $ui.Window.Left = -10000
    $ui.Window.Top = -10000
    $ui.Window.Show()
    $ui.Window.UpdateLayout()
    Test-IntuneLoadingTheme
    $ui.Window.FindName('NavAuditLookup').IsChecked = $true
    Assert-IntuneTest ($MainTabs.SelectedIndex -eq 4 -and $TxtPageTitle.Text -eq 'Audit Lookup' -and
        $DateAuditStart.SelectedDate -eq [datetime]::Today.AddDays(-6) -and $DateAuditEnd.SelectedDate -eq [datetime]::Today -and
        -not $BtnAuditExport.IsEnabled) 'Audit navigation builds with the last seven local calendar dates and no initial export'
    for ($themePass = 0; $themePass -lt 2; $themePass++) {
        $ui.Window.UpdateLayout()
        $unreadableInputs = @(@($DateAuditStart, $DateAuditEnd) | Where-Object {
            $inputBox = $_.Template.FindName('PART_TextBox', $_)
            $null -eq $inputBox -or $inputBox.Background.Color -ne $window.Resources['InputBgBrush'].Color -or
                $inputBox.Foreground.Color -ne $window.Resources['TextBrush'].Color
        })
        Assert-IntuneTest ($unreadableInputs.Count -eq 0) 'Audit date text and background follow the active light/dark theme'
        $BtnThemeToggle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        $ui.Window.UpdateLayout()
    }
    $auditMocks = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
    $global:IntuneMockState.AuditOutput = [PSCustomObject]@{ Completed = $true; Rows = $auditTestRows }
    $global:IntuneMockState.AuditWorkerCalls = 0
    $auditMocks.Add('New-IntuneToolboxFeatureWorker', {
        param($RelativePath, $Parameters, $Scopes, $OutputAction)
        $global:IntuneMockState.AuditWorkerCalls++
        $global:IntuneMockState.AuditParameters = $Parameters
        if ($global:IntuneMockState.AuditOutput) { & $OutputAction $global:IntuneMockState.AuditOutput }
        return {}
    })
    $auditMocks.Add('Invoke-InBackground', {
        param($Worker, $Buttons, $BusyLabel, $LogBox, $LogScroll, $OnComplete)
        & $OnComplete
    })
    $auditSearchFunction = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Start-IntuneAuditSearch'
    }, $true).Body.GetScriptBlock()
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    Assert-IntuneTest ($global:IntuneMockState.AuditWorkerCalls -eq 0 -and
        (Get-IntuneLogText $TxtAuditLog) -like '*Enter a User*') 'Audit UI rejects blank criteria with a visible warning'
    $TxtAuditTarget.Text = 'Pilot'
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    Assert-IntuneTest ($ListAuditEvents.Items.Count -eq 4 -and $BtnAuditExport.IsEnabled -and
        $global:IntuneMockState.AuditParameters.AsReport -and -not $script:AuditState.Busy -and
        $global:IntuneMockState.AuditParameters.StartDate -eq [datetime]::Today.AddDays(-6) -and
        $global:IntuneMockState.AuditParameters.EndDate -eq [datetime]::Today) 'Audit UI publishes only a completed report and preserves local calendar inputs without culture swaps'
    $ui.Window.UpdateLayout()
    $auditTimeHeader = Find-IntuneColumnHeader -Visual $ListAuditEvents -Column $ListAuditEvents.View.Columns[0]
    $auditTimeHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    Assert-IntuneTest ($ListAuditEvents.Items.SortDescriptions[0].PropertyName -eq 'Timestamp' -and
        $ListAuditEvents.Items[0].Id -eq 'event-1') 'Audit timestamp header sorts typed instants rather than localized display strings'
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    Assert-IntuneTest ($ListAuditEvents.Items.SortDescriptions.Count -eq 0 -and $ListAuditEvents.Items[0].Id -eq 'event-4' -and
        $ListAuditEvents.View.Columns[0].Header -eq 'Timestamp (local)') 'New audit search resets prior column sorting and restores newest-first order'
    $ListAuditEvents.SelectedIndex = 3
    Assert-IntuneTest ($TxtAuditDetails.Text -like '*correlation-1*' -and $TxtAuditDetails.Text -like '*192.0.2.1*' -and
        $TxtAuditDetails.Text -like '*Other*' -and $TxtAuditDetails.Text -like '*Old value: Before*New value: After*' -and
        $TxtAuditDetails.Text -like '*Old value: Not supplied*New value: (empty)*') 'Selected audit event shows actor, every target, and old/new property values'
    $ListAuditEvents.SelectedIndex = 0
    Assert-IntuneTest ($TxtAuditDetails.Text -like '*Changed properties: Not supplied*') 'Audit details explicitly identify missing changed properties'
    $auditCsv = [IO.Path]::GetTempFileName()
    try {
        $dialog = [PSCustomObject]@{ FileName = $auditCsv }
        $auditExportPipeline = $guiAst.Find({
            param($node)
            $node -is [Management.Automation.Language.PipelineAst] -and $node.Extent.Text -like '$script:AuditState.Report.Rows*Export-Csv*'
        }, $true)
        & ([scriptblock]::Create($auditExportPipeline.Extent.Text))
        $csvRows = @(Import-Csv -LiteralPath $auditCsv)
        $csvTargets = $csvRows[-1].ResourcesJson | ConvertFrom-Json
        Assert-IntuneTest ($csvRows.Count -eq 4 -and $csvRows[0].TimestampUtc -match '\+00:00$' -and
            $csvTargets.Count -eq 2 -and $csvTargets[1].modifiedProperties[0].oldValue -eq 'Before' -and
            ($csvRows[-1].ActorJson | ConvertFrom-Json).ipAddress -eq '192.0.2.1') 'Audit CSV exports all results with UTC timestamps and intact actor/target/change JSON'
        Assert-IntuneTest ($csvRows[-1].OperationKind -eq 'Update' -and $csvRows[0].OperationKind -eq 'Unknown') 'Audit CSV includes the operation type classification'
    } finally { Remove-Item -LiteralPath $auditCsv -Force }
    $colorRows = @('Create', 'Delete', 'Update', 'Other', 'Unknown') | ForEach-Object -Begin { $offset = 0 } -Process {
        $offset++
        [PSCustomObject]@{ Id = "color-$_"; OperationKind = $_; Operation = "$_ op"; Timestamp = [DateTimeOffset]::UtcNow.AddMinutes(-$offset)
            TimestampLocal = ''; TimestampUtc = ''; Event = [PSCustomObject]@{ actor = $null; resources = @() } }
    }
    $global:IntuneMockState.AuditOutput = [PSCustomObject]@{ Completed = $true; Rows = $colorRows }
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    $ui.Window.UpdateLayout()
    function Find-IntuneVisualText {
        param($Visual, [string]$Text)
        if ($Visual -is [System.Windows.Controls.TextBlock] -and $Visual.Text -eq $Text) { return $Visual }
        for ($childIndex = 0; $childIndex -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Visual); $childIndex++) {
            $found = Find-IntuneVisualText -Visual ([System.Windows.Media.VisualTreeHelper]::GetChild($Visual, $childIndex)) -Text $Text
            if ($found) { return $found }
        }
    }
    function Get-IntuneAuditColorMismatches {
        $expectedKeys = @{ Create = 'AuditCreateBrush'; Delete = 'AuditDeleteBrush'; Update = 'AuditUpdateBrush'; Other = 'AuditOtherBrush' }
        @(for ($rowIndex = 0; $rowIndex -lt $ListAuditEvents.Items.Count; $rowIndex++) {
            $row = $ListAuditEvents.Items[$rowIndex]
            $container = $ListAuditEvents.ItemContainerGenerator.ContainerFromIndex($rowIndex)
            $operationText = Find-IntuneVisualText -Visual $container -Text $row.Operation
            $typeText = Find-IntuneVisualText -Visual $container -Text $row.OperationKind
            if ($expectedKeys.ContainsKey($row.OperationKind)) {
                $expected = $window.Resources[$expectedKeys[$row.OperationKind]].Color
                if ($container.BorderBrush.Color -ne $expected -or $container.BorderThickness.Left -ne 4 -or
                    $operationText.Foreground.Color -ne $expected -or $typeText.Foreground.Color -ne $expected) { $row.OperationKind }
            } elseif ($operationText.Foreground.Color -ne $window.Resources['TextBrush'].Color -and
                $operationText.Foreground.Color -ne $window.Resources['SelectionForeground'].Color) { $row.OperationKind }
        })
    }
    Assert-IntuneTest ((Get-IntuneAuditColorMismatches).Count -eq 0) 'Audit rows show green/red/blue/amber stripes and Operation/Type text, leaving Unknown uncolored'
    $ListAuditEvents.SelectedIndex = 1
    $ui.Window.UpdateLayout()
    Assert-IntuneTest ((Get-IntuneAuditColorMismatches).Count -eq 0 -and
        $ListAuditEvents.ItemContainerGenerator.ContainerFromIndex(1).Background.Color -eq $window.Resources['SelectionBackground'].Color -and
        $TxtAuditDetails.Text -like '*Change type: Delete*') 'Selected audit row keeps its type stripe and color while showing selection and change type'
    $BtnThemeToggle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    $ui.Window.UpdateLayout()
    Assert-IntuneTest ((Get-IntuneAuditColorMismatches).Count -eq 0 -and
        $window.Resources['AuditDeleteBrush'].Color.ToString() -eq '#FFB91C1C') 'Audit type colors switch to the light-theme palette'
    $BtnThemeToggle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    $ui.Window.UpdateLayout()
    Assert-IntuneTest ((Get-IntuneAuditColorMismatches).Count -eq 0 -and
        $window.Resources['AuditDeleteBrush'].Color.ToString() -eq '#FFF87171') 'Audit type colors return to the dark-theme palette'
    $legendLabels = @($window.FindName('AuditLegend').Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | ForEach-Object Text)
    Assert-IntuneTest (($legendLabels -join ',') -eq 'Create,Delete,Update/Patch,Other' -and
        $ListAuditEvents.View.Columns[1].Header -eq 'Type') 'Audit legend and Type column are shown'
    $auditTypeHeader = Find-IntuneColumnHeader -Visual $ListAuditEvents -Column $ListAuditEvents.View.Columns[1]
    $auditTypeHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    $typeSorted = $ListAuditEvents.Items.SortDescriptions[0].PropertyName -eq 'OperationKind' -and
        $ListAuditEvents.Items[0].OperationKind -eq 'Create' -and $ListAuditEvents.Items[4].OperationKind -eq 'Update' -and
        $ListAuditEvents.View.Columns[1].Header -eq "Type $([char]0x25B2)"
    $auditOperationHeader = Find-IntuneColumnHeader -Visual $ListAuditEvents -Column $ListAuditEvents.View.Columns[2]
    $auditOperationHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    Assert-IntuneTest ($typeSorted -and $ListAuditEvents.Items.SortDescriptions[0].PropertyName -eq 'Operation' -and
        $ListAuditEvents.View.Columns[1].Header -eq 'Type') 'Audit Type and Operation columns sort by their own properties'
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    Assert-IntuneTest ($ListAuditEvents.Items.SortDescriptions.Count -eq 0 -and $ListAuditEvents.View.Columns[2].Header -eq 'Operation') 'New audit search resets Type/Operation sorting'
    $global:IntuneMockState.AuditOutput = [PSCustomObject]@{ Completed = $true; Rows = @() }
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    Assert-IntuneTest ($ListAuditEvents.Items.Count -eq 0 -and -not $BtnAuditExport.IsEnabled -and
        $TxtAuditCount.Text -like 'No matching events*' -and $TxtAuditDetails.Text -like 'Select an event*') 'Successful empty audit search clears stale results and details without reporting failure'
    $global:IntuneMockState.AuditOutput = $null
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    Assert-IntuneTest ($null -eq $script:AuditState.Report -and -not $BtnAuditExport.IsEnabled -and
        $TxtAuditCount.Text -like 'Search failed*') 'Failed audit worker cannot appear as a successful no-match search'
    $beforeAuditCalls = $global:IntuneMockState.AuditWorkerCalls
    $DateAuditStart.Text = ''
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    Assert-IntuneTest ($global:IntuneMockState.AuditWorkerCalls -eq $beforeAuditCalls -and
        (Get-IntuneLogText $TxtAuditLog) -like '*valid start/end dates*') 'Audit UI rejects blank dates before starting a worker'
    $DateAuditStart.Text = 'not-a-date'
    [void]$auditSearchFunction.InvokeWithContext($auditMocks, $null, @())
    Assert-IntuneTest ($global:IntuneMockState.AuditWorkerCalls -eq $beforeAuditCalls) 'Malformed calendar input cannot silently start a search with a restored date'
    $DateAuditStart.SelectedDate = [datetime]::Today.AddDays(-6)
    $DateAuditStart.Text = [datetime]::Today.AddDays(-6).ToShortDateString()
    $script:AuditState.Busy = $true
    $BtnSignInOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    Assert-IntuneTest ((Get-IntuneLogText $TxtAuditLog) -like '*Wait for the audit search*') 'Audit search blocks sign-in changes while pages are being retrieved'
    $script:AuditState.Busy = $false
    $ui.Window.FindName('NavProfileManagement').IsChecked = $true
    Assert-IntuneTest ([bool]$ChkAdministrativeTemplates.IsChecked) 'Administrative Templates checkbox starts checked'
    $loadHandler = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and
            $node.Expression.Extent.Text -eq '$BtnPreview' -and $node.Member.Value -eq 'Add_Click'
    }, $true)
    $loadMocks = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
    $loadMocks.Add('New-IntuneToolboxFeatureWorker', {
        param($RelativePath, $Parameters, $Scopes, $OutputAction)
        $global:IntuneMockState.LoadParameters = $Parameters
        return {}
    })
    $loadMocks.Add('Invoke-InBackground', {})
    [void]$loadHandler.Arguments[0].ScriptBlock.GetScriptBlock().InvokeWithContext($loadMocks, $null, @())
    Assert-IntuneTest (($global:IntuneMockState.LoadParameters.ProfileType -join ',') -eq 'SettingsCatalog,DeviceConfiguration,AdministrativeTemplates') 'GUI inventory includes all checked profile types'
    $ChkSettingsCatalog.IsChecked = $false
    $ChkDeviceConfig.IsChecked = $false
    [void]$loadHandler.Arguments[0].ScriptBlock.GetScriptBlock().InvokeWithContext($loadMocks, $null, @())
    Assert-IntuneTest (($global:IntuneMockState.LoadParameters.ProfileType -join ',') -eq 'AdministrativeTemplates') 'GUI inventory accepts Administrative Templates as the only selected type'
    $ChkAdministrativeTemplates.IsChecked = $false
    $global:IntuneMockState.LoadParameters = $null
    [void]$loadHandler.Arguments[0].ScriptBlock.GetScriptBlock().InvokeWithContext($loadMocks, $null, @())
    Assert-IntuneTest ($null -eq $global:IntuneMockState.LoadParameters) 'GUI rejects inventory when no profile types are checked'
    $ChkSettingsCatalog.IsChecked = $true
    $ChkDeviceConfig.IsChecked = $true
    $ChkAdministrativeTemplates.IsChecked = $true

    $ComboAction.SelectedIndex = 0
    $ComboTargetType.SelectedIndex = 2
    [void]$ListProfiles.Items.Add($templates[0])
    $ListProfiles.SelectAll()
    foreach ($group in $multiGroups[0..1]) { [void]$ListGroups.Items.Add($group) }
    $ListGroups.SelectAll()
    Assert-IntuneTest ($BtnQueueGroups.IsEnabled -and -not $BtnAssignExecute.IsEnabled) 'Highlighted search results alone do not enable Apply before groups are queued'
    $buttonClick = [System.Windows.Controls.Button]::ClickEvent
    $BtnQueueGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    $ListGroups.SelectedItems.Clear()
    $ListQueuedGroups.SelectedItems.Clear()
    Assert-IntuneTest ($ListQueuedGroups.Items.Count -eq 2 -and $BtnAssignExecute.IsEnabled -and
        $ListQueuedGroups.SelectedItems.Count -eq 0 -and $ListGroups.SelectedItems.Count -eq 0) 'Apply is enabled for every queued group without highlighting any group rows'
    $ListQueuedGroups.SelectedItems.Add($ListQueuedGroups.Items[0]) | Out-Null
    $snapshot = Get-AssignmentParameters
    Assert-IntuneTest ($snapshot.TargetGroup.Count -eq 2 -and $snapshot.Profiles.Count -eq 1 -and
        $snapshot.Profiles[0].ApiType -eq 'groupPolicyConfigurations') 'Assignment snapshot includes all queued groups, not just highlighted rows, and selected Administrative Templates'
    $ListQueuedGroups.SelectedItems.Clear()
    $script:BusyButtonCounts[$BtnAssignExecute.Name] = 1
    Update-ProfileCounters
    Assert-IntuneTest (-not $BtnAssignExecute.IsEnabled) 'Queue updates respect the busy Apply button guard'
    $script:BusyButtonCounts.Remove($BtnAssignExecute.Name)
    $ListProfiles.SelectedItems.Clear()
    Assert-IntuneTest (-not $BtnAssignExecute.IsEnabled) 'Queued groups never bypass required profile selection'
    $ListProfiles.SelectAll()

    foreach ($searchButtonName in @('$BtnSearchGroup', '$BtnSearchRemoveGroup')) {
        $searchHandler = $guiAst.Find({
            param($node)
            $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and
                $node.Expression.Extent.Text -eq $searchButtonName -and $node.Member.Value -eq 'Add_Click'
        }, $true)
        $searchMocks = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
        $searchMocks.Add('New-IntuneToolboxFeatureWorker', {
            param($RelativePath, $Parameters, $Scopes, $OutputAction)
            & $OutputAction ([PSCustomObject]@{ Id = 'group-3'; DisplayName = 'Third group' })
            & $OutputAction ([PSCustomObject]@{ Id = 'GROUP-1'; DisplayName = 'Duplicate ID' })
            return {}
        })
        $searchMocks.Add('Invoke-InBackground', {
            param($Worker, $Buttons, $BusyLabel, $LogBox, $LogScroll, $DarkMode, $OnComplete)
            $global:IntuneMockState.SearchCompletion = $OnComplete
        })
        $TxtGroupSearch.Text = 'group'
        $TxtRemoveGroupSearch.Text = 'group'
        [void]$searchHandler.Arguments[0].ScriptBlock.GetScriptBlock().InvokeWithContext($searchMocks, $null, @())
        $_groupList = if ($searchButtonName -eq '$BtnSearchGroup') { $ListGroups } else { $ListRemoveGroups }
        & $global:IntuneMockState.SearchCompletion
        Assert-IntuneTest ($_groupList.Items.Count -eq 2 -and $ListQueuedGroups.Items.Count -eq 2) 'Actual search completion replaces results without clearing queued assignments'
    }
    $ListGroups.SelectAll()
    $BtnQueueGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    Assert-IntuneTest ($ListQueuedGroups.Items.Count -eq 3) 'Groups accumulate across searches and duplicate IDs are ignored case-insensitively'
    Assert-IntuneTest ($snapshot.TargetGroup.Count -eq 2) 'Previously captured assignment snapshot is independent of later queue edits'
    $ListQueuedGroups.SelectedItem = $ListQueuedGroups.Items[2]
    $BtnUnqueueGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    Assert-IntuneTest ($ListQueuedGroups.Items.Count -eq 2 -and $BtnAssignExecute.IsEnabled) 'Remove selected edits the assignment queue without requiring other rows to be highlighted'
    $ComboAction.SelectedIndex = 1
    Assert-IntuneTest ((Get-AssignmentParameters).TargetGroup.Count -eq 2 -and $BtnAssignExecute.IsEnabled) 'Replace uses the entire persistent assignment queue'
    $ComboAction.SelectedIndex = 2
    Assert-IntuneTest (-not $BtnAssignExecute.IsEnabled) 'Group Remove requires its own independent queue'
    $ListRemoveGroups.SelectAll()
    $BtnQueueRemoveGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    $ListRemoveGroups.SelectedItems.Clear()
    $ListQueuedRemoveGroups.SelectedItems.Clear()
    $removalSnapshot = Get-AssignmentParameters
    Assert-IntuneTest ($BtnAssignExecute.IsEnabled -and $removalSnapshot.RemoveGroup.Count -eq 2 -and
        -not $removalSnapshot.ContainsKey('TargetType') -and $ListQueuedGroups.Items.Count -eq 2) 'Remove snapshot uses every removal-queue group without unrelated Add parameters or highlighting'
    $ListQueuedRemoveGroups.SelectAll()
    $BtnUnqueueRemoveGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    Assert-IntuneTest ($ListQueuedRemoveGroups.Items.Count -eq 0 -and -not $BtnAssignExecute.IsEnabled) 'Removing all queued removal groups disables Apply'
    $ComboRemoveScope.SelectedIndex = 1
    Assert-IntuneTest ($BtnAssignExecute.IsEnabled -and -not (Get-AssignmentParameters).ContainsKey('RemoveGroup')) 'Remove All needs no queued groups'
    $ComboRemoveScope.SelectedIndex = 0
    $ListRemoveGroups.SelectAll()
    $BtnQueueRemoveGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    $BtnClearQueuedRemoveGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    Assert-IntuneTest ($ListQueuedRemoveGroups.Items.Count -eq 0 -and $ListQueuedGroups.Items.Count -eq 2) 'Clear removal queue leaves assignment queue unchanged'
    $ComboAction.SelectedIndex = 0
    $BtnClearQueuedGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    Assert-IntuneTest (-not $BtnAssignExecute.IsEnabled -and $ListQueuedGroups.Items.Count -eq 0) 'Clear assignment queue disables group Apply'
    $ListGroups.SelectAll()
    $BtnQueueGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    $script:AllFilters = @([PSCustomObject]@{ Id = 'filter-ui'; DisplayName = 'UI test filter' })
    $ComboFilterMode.SelectedIndex = 1
    Assert-IntuneTest (-not $BtnAssignExecute.IsEnabled) 'Enabled assignment filter still requires a selected filter'
    [void]$ListFilters.Items.Add($script:AllFilters[0])
    $ListFilters.SelectedIndex = 0
    Assert-IntuneTest ($BtnAssignExecute.IsEnabled -and (Get-AssignmentParameters).AssignmentFilter.Id -eq 'filter-ui') 'Selected optional filter is captured for all queued groups'
    $ComboFilterMode.SelectedIndex = 0
    $ListRemoveGroups.SelectAll()
    $BtnQueueRemoveGroups.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    $BtnRenameReset.RaiseEvent([System.Windows.RoutedEventArgs]::new($buttonClick))
    Assert-IntuneTest ($ListQueuedGroups.Items.Count -eq 0 -and $ListQueuedRemoveGroups.Items.Count -eq 0 -and
        $ListGroups.Items.Count -eq 0 -and $ListRemoveGroups.Items.Count -eq 0 -and -not $BtnAssignExecute.IsEnabled) 'Form reset clears both queues and search results safely'

    $exportHandler = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and
            $node.Expression.Extent.Text -eq '$BtnExportCsv' -and $node.Member.Value -eq 'Add_Click'
    }, $true)
    $exportPipeline = $exportHandler.Arguments[0].ScriptBlock.Find({
        param($node)
        $node -is [Management.Automation.Language.PipelineAst] -and $node.Extent.Text -like '*Select-Object CurrentName*Export-Csv*'
    }, $true)
    $script:AllProfiles.Add($templates[0])
    $global:IntuneMockState.ExportedProfiles = [System.Collections.Generic.List[object]]::new()
    $exportMocks = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
    $exportMocks.Add('Export-Csv', {
        param([Parameter(ValueFromPipeline)]$InputObject, $Path, [switch]$NoTypeInformation, $Encoding)
        process { $global:IntuneMockState.ExportedProfiles.Add($InputObject) }
    })
    $dlg = [PSCustomObject]@{ FileName = 'unused-offline-test.csv' }
    [void]([scriptblock]::Create($exportPipeline.Extent.Text)).InvokeWithContext($exportMocks, $null, @())
    Assert-IntuneTest ($global:IntuneMockState.ExportedProfiles.Count -eq 1 -and
        $global:IntuneMockState.ExportedProfiles[0].ProfileType -eq 'Administrative Templates' -and
        ($global:IntuneMockState.ExportedProfiles[0].PSObject.Properties.Name -join ',') -eq 'CurrentName,NewName,ProfileType,Assignments,Succeeded,Compliant,Id') 'Existing CSV export pipeline includes Administrative Templates with the unchanged column schema'
    $script:AllProfiles.Clear()

    $ui.Window.FindName('NavDuplicateDevices').IsChecked = $true
    Assert-IntuneTest ($MainTabs.SelectedIndex -eq 3 -and $TxtPageTitle.Text -eq 'Duplicate Devices') 'Duplicate Devices sidebar opens the correct page'
    Assert-IntuneTest (-not $BtnDuplicateDelete.IsEnabled -and -not $BtnDuplicateExport.IsEnabled) 'Duplicate actions start safely disabled'
    $testTenant = '11111111-1111-1111-1111-111111111111'
    $script:GraphState.ExpectedTenant = $testTenant
    $eligibleRow = [PSCustomObject]@{ GroupId = 'Hardware-1'; MatchType = 'Hardware'; Id = 'device-1'; DeviceName = 'Old device'; CanDelete = $true; ProtectionReason = ''; SerialNumber = 'SERIAL'; UserPrincipalName = '' }
    $protectedRow = [PSCustomObject]@{ GroupId = 'Hardware-1'; MatchType = 'Hardware'; Id = 'device-2'; DeviceName = 'Keeper'; CanDelete = $false; ProtectionReason = 'Suggested keeper'; SerialNumber = 'SERIAL'; UserPrincipalName = '' }
    $script:DuplicateState.Report = [PSCustomObject]@{ TenantId = $testTenant; Rows = @($eligibleRow, $protectedRow) }
    $script:DuplicateState.Valid = $true
    Update-DuplicateDeviceFilter
    Assert-IntuneTest ($ListDuplicateDevices.SelectedItems.Count -eq 0 -and -not $BtnDuplicateDelete.IsEnabled -and $BtnDuplicateExport.IsEnabled) 'Loaded duplicate report never auto-selects records'
    [void]$ListDuplicateDevices.SelectedItems.Add($eligibleRow)
    Assert-IntuneTest $BtnDuplicateDelete.IsEnabled 'Eligible duplicate selection enables deletion'
    [void]$ListDuplicateDevices.SelectedItems.Add($protectedRow)
    Assert-IntuneTest (-not $BtnDuplicateDelete.IsEnabled) 'Mixed keeper/candidate selection disables deletion'
    $TxtDuplicateFilter.Text = 'keeper'
    Assert-IntuneTest ($ListDuplicateDevices.Items.Count -eq 1 -and $script:DuplicateState.Report.Rows.Count -eq 2) 'Duplicate text filter does not truncate the export report'
    $TxtDuplicateFilter.Text = ''
    $ComboDuplicateMatch.SelectedIndex = 3
    Assert-IntuneTest ($ListDuplicateDevices.Items.Count -eq 0) 'Duplicate match-type filter isolates name-only candidates'
    $ComboDuplicateMatch.SelectedIndex = 0
    $columns = @($script:DuplicateState.Report.Rows | Select-Object -Property $script:DuplicateExportColumns | ConvertTo-Csv -NoTypeInformation)
    Assert-IntuneTest ($columns.Count -eq 3 -and $columns[0] -like '*"ProtectionReason"*' -and $columns[0] -like '*"TenantId"*' -and $columns[0] -notlike '*Fingerprint*') 'Duplicate CSV schema includes context and protection without internal fingerprints'
    $script:GraphState.ExpectedTenant = '22222222-2222-2222-2222-222222222222'
    Update-DuplicateDeviceState
    Assert-IntuneTest (-not $script:DuplicateState.Valid -and $ListDuplicateDevices.Items.Count -eq 0 -and -not $BtnDuplicateDelete.IsEnabled) 'Tenant change invalidates duplicate report and selection'
    $script:GraphState.ExpectedTenant = ''
    $scanFunction = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Start-IntuneDuplicateScan'
    }, $true)
    $scanCompletionAssignment = $scanFunction.Body.Find({
        param($node)
        $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$onComplete'
    }, $true)
    $scanCompletion = $scanCompletionAssignment.Right.Find({
        param($node)
        $node -is [Management.Automation.Language.ScriptBlockExpressionAst]
    }, $true).ScriptBlock.GetScriptBlock()
    $state = $script:DuplicateState
    $graphState = $script:GraphState
    $graphState.ExpectedTenant = $testTenant
    $bag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $bag.Add([PSCustomObject]@{ TenantId = $testTenant; Rows = @() })
    $state.Busy = $true
    & $scanCompletion
    Assert-IntuneTest ($state.Valid -and -not $state.Busy -and $state.Report.Rows.Count -eq 0 -and
        $state.Message -like '*no duplicate*' -and -not $BtnDuplicateDelete.IsEnabled) 'Successful zero-result duplicate scan is distinguishable from failure'
    $bag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    & $scanCompletion
    Assert-IntuneTest (-not $state.Valid -and $null -eq $state.Report -and -not $BtnDuplicateExport.IsEnabled -and
        $state.Message -like '*failed*') 'Failed duplicate scan leaves no old report actionable'
    $deleteHandler = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and
            $node.Expression.Extent.Text -eq '$BtnDuplicateDelete' -and $node.Member.Value -eq 'Add_Click'
    }, $true)
    $deleteCompletionAssignment = $deleteHandler.Arguments[0].ScriptBlock.Find({
        param($node)
        $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$onComplete'
    }, $true)
    $deleteCompletion = $deleteCompletionAssignment.Right.Find({
        param($node)
        $node -is [Management.Automation.Language.ScriptBlockExpressionAst]
    }, $true).ScriptBlock.GetScriptBlock()
    $global:IntuneMockState.Refreshes = 0
    $refreshMock = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
    $refreshMock.Add('Start-IntuneDuplicateScan', { $global:IntuneMockState.Refreshes++ })
    $bag.Add([PSCustomObject]@{ Id = 'device-1'; Status = 'Failed'; Message = 'Mock failure' })
    $state.Busy = $true
    [void]$deleteCompletion.InvokeWithContext($refreshMock, $null, @())
    Assert-IntuneTest ($global:IntuneMockState.Refreshes -eq 1 -and $state.Outcomes.Count -eq 1 -and
        -not $state.Valid -and -not $BtnDuplicateDelete.IsEnabled) 'Delete completion retains outcomes and requests a fresh scan without enabling stale deletion'
    $global:IntuneMockState.DuplicateReport = [PSCustomObject]@{ TenantId = $testTenant; Rows = @($eligibleRow, $protectedRow) }
    $scanMocks = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
    $scanMocks.Add('New-IntuneToolboxFeatureWorker', {
        param($RelativePath, $Parameters, $Scopes, $OutputAction)
        & $OutputAction $global:IntuneMockState.DuplicateReport
        return {}
    })
    $scanMocks.Add('Invoke-InBackground', {
        param($Worker, $Buttons, $BusyLabel, $LogBox, $LogScroll, $DarkMode, $OverlayMessage, $OnComplete)
        & $OnComplete
    })
    [void]$scanFunction.Body.GetScriptBlock().InvokeWithContext($scanMocks, $null, @())
    Assert-IntuneTest ($state.Valid -and -not $state.Busy -and $ListDuplicateDevices.Items.Count -eq 2 -and
        $ListDuplicateDevices.SelectedItems.Count -eq 0) 'Actual scan completion closures publish report rows without auto-selection'
    Clear-DuplicateReport 'Test complete'
    $script:GraphState.ExpectedTenant = ''
    $ui.Window.FindName('NavProfileManagement').IsChecked = $true
    $sortList = $ui.Window.FindName('ListProfiles')
    foreach ($sortProfile in $sortProfiles) { [void]$sortList.Items.Add($sortProfile) }
    $succeededColumn = $sortList.View.Columns[4]
    $succeededHeader = Find-IntuneColumnHeader -Visual $sortList -Column $succeededColumn
    $nameHeader = Find-IntuneColumnHeader -Visual $sortList -Column $sortList.View.Columns[0]
    Assert-IntuneTest ($null -ne $succeededHeader -and $null -ne $nameHeader) 'WPF generates sortable profile column headers'
    $clickEvent = [System.Windows.Controls.GridViewColumnHeader]::ClickEvent
    $succeededHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new($clickEvent))
    Assert-IntuneTest (($sortList.Items.SucceededSortValue -join ',') -eq '-1,-1,-1,-1,0,1,2,10' -and
        $succeededColumn.Header -eq "Succeeded $([char]0x25B2)") 'Succeeded header sorts mixed counts numerically ascending'
    $succeededHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new($clickEvent))
    Assert-IntuneTest (($sortList.Items.SucceededSortValue -join ',') -eq '10,2,1,0,-1,-1,-1,-1' -and
        $succeededColumn.Header -eq "Succeeded $([char]0x25BC)") 'Succeeded header toggles descending with unavailable counts last'
    $nameHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new($clickEvent))
    Assert-IntuneTest (($sortList.Items.CurrentName -join ',') -eq (($sortProfiles.CurrentName | Sort-Object) -join ',')) 'Other profile columns retain their existing sorting'
    $succeededHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new($clickEvent))
    Assert-IntuneTest (($sortList.Items.SucceededSortValue -join ',') -eq '-1,-1,-1,-1,0,1,2,10') 'Switching back to Succeeded resets to ascending'
    $sortList.Items.Clear()
    foreach ($sortProfile in $sortProfiles | Where-Object SucceededSortValue -eq -1) { [void]$sortList.Items.Add($sortProfile) }
    $succeededHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new($clickEvent))
    Assert-IntuneTest ($sortList.Items.Count -eq 4 -and @($sortList.Items | Where-Object SucceededSortValue -ne -1).Count -eq 0) 'Succeeded sorting handles all-unavailable counts'
    $sortList.Items.Clear()
    $succeededHeader.RaiseEvent([System.Windows.RoutedEventArgs]::new($clickEvent))
    Assert-IntuneTest ($sortList.Items.Count -eq 0) 'Succeeded sorting handles an empty profile list'

    $completion = New-IntuneTestCompletion
    & $completion
    Assert-IntuneTest (-not $script:IsConnected -and $TxtGraphConnectionStatus.Text -eq 'Not connected' -and
        $BtnSignInOut.Content -eq 'Sign in') 'Disconnected background completion does not report a connection'
    $script:GraphState.Connected = $true
    $script:GraphState.Account = 'background@example.test'
    $script:GraphState.Tenant = '11111111-1111-1111-1111-111111111111'
    & $completion
    Assert-IntuneTest ($script:IsConnected -and $script:GraphUser -eq $script:GraphState.Account -and
        $script:GraphTenantId -eq $script:GraphState.Tenant -and $TxtGraphConnectionStatus.Text -eq 'Connected' -and
        $BtnSignInOut.Content -eq 'Sign out' -and $TxtTenantName.Text -eq $script:GraphState.Account -and
        $TxtTenantName.ToolTip -eq "User: $($script:GraphState.Account)`nTenant ID: $($script:GraphState.Tenant)" -and
        $EllipseStatus.Fill.Color.ToString() -eq '#FF16A34A') 'Background completion synchronizes original connection state and sidebar'
    $script:GraphState.Account = ''
    & $completion
    Assert-IntuneTest ($TxtTenantName.Text -eq $script:GraphState.Tenant -and $script:GraphUser -eq '') 'Connection sidebar falls back to tenant when account is empty'
    $failedCompletion = New-IntuneTestCompletion -WorkerFailed
    & $failedCompletion
    Assert-IntuneTest ($script:IsConnected -and $TxtGraphConnectionStatus.Text -eq 'Connected') 'Worker failure does not hide an authenticated Graph connection'
    $signInOutHandler = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and
            $node.Expression.Extent.Text -eq '$BtnSignInOut' -and $node.Member.Value -eq 'Add_Click'
    }, $true)
    $global:IntuneMockState.Disconnects = 0
    $graphMocks = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
    $graphMocks.Add('Import-Module', {})
    $graphMocks.Add('Disconnect-MgGraph', { $global:IntuneMockState.Disconnects++ })
    [void]$signInOutHandler.Arguments[0].ScriptBlock.GetScriptBlock().InvokeWithContext($graphMocks, $null, @())
    Assert-IntuneTest ($global:IntuneMockState.Disconnects -eq 1 -and -not $script:IsConnected -and
        -not $script:GraphState.Connected -and $script:GraphState.ExpectedTenant -eq '' -and
        $script:GraphUser -eq '' -and $script:GraphTenantId -eq '' -and $TxtTenantName.Text -eq '' -and
        $TxtGraphConnectionStatus.Text -eq 'Not connected' -and $BtnSignInOut.Content -eq 'Sign in') 'Sign out uses synchronized connection flag and clears state without live Graph calls'
    & $completion
    Assert-IntuneTest (-not $script:IsConnected -and $TxtGraphConnectionStatus.Text -eq 'Not connected') 'Completion reads current shared state after sign out'

    $testTenant = '11111111-1111-1111-1111-111111111111'
    $script:GraphState.ExpectedTenant = $testTenant
    $global:IntuneMockState.GrantedScopes = @('User.Read')
    $global:IntuneMockState.Connects = [System.Collections.Generic.List[object]]::new()
    $global:IntuneMockState.ConnectBehavior = 'Grant'
    $consentMocks = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
    $consentMocks.Add('Import-Module', {})
    $consentMocks.Add('Update-TenantInfo', {})
    $consentMocks.Add('Get-MgContext', {
        [PSCustomObject]@{ Account = 'tester@example.test'; TenantId = '11111111-1111-1111-1111-111111111111'; Scopes = $global:IntuneMockState.GrantedScopes }
    })
    $consentMocks.Add('Connect-MgGraph', {
        param($Scopes, $NoWelcome, $ContextScope, $TenantId, $ErrorAction)
        $global:IntuneMockState.Connects.Add([PSCustomObject]@{ Scopes = $Scopes; TenantId = $TenantId })
        switch ($global:IntuneMockState.ConnectBehavior) {
            'Cancel' { throw 'User canceled authentication.' }
            'NoConsent' { }
            default { $global:IntuneMockState.GrantedScopes = $Scopes }
        }
    })
    $requestScan = { Request-IntuneToolboxScopes -Scopes @('DeviceManagementManagedDevices.Read.All') }
    $TxtGraphConnectionStatus.Text = 'Connected'
    $note = $requestScan.InvokeWithContext($consentMocks, $null, @())
    $connect = $global:IntuneMockState.Connects[0]
    Assert-IntuneTest ($global:IntuneMockState.Connects.Count -eq 1 -and $connect.TenantId -eq $testTenant -and
        'User.Read' -in $connect.Scopes -and 'DeviceManagementManagedDevices.Read.All' -in $connect.Scopes -and
        "$note" -like '*Granted*DeviceManagementManagedDevices.Read.All*') 'Missing permissions are requested on the UI thread for the current tenant'
    $note = $requestScan.InvokeWithContext($consentMocks, $null, @())
    Assert-IntuneTest ($global:IntuneMockState.Connects.Count -eq 1 -and "$note" -eq '') 'Granted permissions do not prompt again'
    $global:IntuneMockState.GrantedScopes = @('User.Read', 'DeviceManagementManagedDevices.ReadWrite.All')
    [void]$requestScan.InvokeWithContext($consentMocks, $null, @())
    Assert-IntuneTest ($global:IntuneMockState.Connects.Count -eq 1) 'Existing write permission satisfies the read-only scan'
    $global:IntuneMockState.GrantedScopes = @('User.Read')
    $global:IntuneMockState.ConnectBehavior = 'Cancel'
    $TxtGraphConnectionStatus.Text = 'Connected'
    $failure = ''
    try { [void]$requestScan.InvokeWithContext($consentMocks, $null, @()) } catch { $failure = $_.Exception.Message }
    Assert-IntuneTest ($failure -like '*canceled*' -and $TxtGraphConnectionStatus.Text -eq 'Connected') 'Cancelled sign-in fails clearly and restores the status text'
    $global:IntuneMockState.ConnectBehavior = 'NoConsent'
    $failure = ''
    try { [void]$requestScan.InvokeWithContext($consentMocks, $null, @()) } catch { $failure = $_.Exception.Message }
    Assert-IntuneTest ($failure -like '*did not grant*administrator*') 'Sign-in without the required consent fails clearly'
    $global:IntuneMockState.ConnectBehavior = 'Cancel'
    $global:IntuneMockState.WorkerOutputs = 0
    $workerFactory = {
        New-IntuneToolboxFeatureWorker -RelativePath 'Get-IntuneDuplicateDevices.ps1' -Parameters @{ AsReport = $true } `
            -Scopes @('DeviceManagementManagedDevices.Read.All') -OutputAction { $global:IntuneMockState.WorkerOutputs++ }
    }
    $deniedWorker = @($workerFactory.InvokeWithContext($consentMocks, $null, @()))[0]
    Assert-IntuneTest ($deniedWorker -isnot [scriptblock] -and $deniedWorker.PermissionError -like '*canceled*' -and
        $deniedWorker.FeaturePath -like '*Features*Get-IntuneDuplicateDevices.ps1') 'Feature workers are plain data that carry the permission failure'
    Invoke-InBackground -Worker $deniedWorker -Buttons @($BtnDuplicateScan) -BusyLabel 'Testing...' -LogBox $TxtDuplicateLog `
        -LogScroll $DuplicateLogScroll
    Wait-IntuneBackground { $BtnDuplicateScan.Content -ne 'Testing...' }
    Assert-IntuneTest ((Get-IntuneLogText $TxtDuplicateLog) -like '*Microsoft Graph permission request failed:*' -and
        $global:IntuneMockState.WorkerOutputs -eq 0) 'Failed permission request stops the worker before it runs'
    $workerAst = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-IntuneToolboxFeatureWorker'
    }, $true)
    $backgroundAst = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-InBackground'
    }, $true)
    Assert-IntuneTest ($backgroundAst.Extent.Text -match '\$job\.RunnerPath -Path \$job\.FeaturePath -Parameters \$featureParameters -NonInteractive') 'GUI workers never sign in from a background thread'
    $workerScriptText = [regex]::Match($backgroundAst.Extent.Text, "(?s)\`$workerScript = @'(.*?)'@").Groups[1].Value
    Assert-IntuneTest ($workerScriptText.Length -gt 0 -and $workerAst.Extent.Text -notmatch 'GetNewClosure' -and
        $workerScriptText -notmatch 'Dispatcher|\$OutputAction') 'Background workers never run UI-runspace script blocks or block on the UI dispatcher'
    $script:GraphState.ExpectedTenant = ''
    $script:GraphState.Connected = $false
    $TxtGraphConnectionStatus.Text = 'Not connected'

    $ui.Window.FindName('NavDuplicateDevices').IsChecked = $true
    $state = $script:DuplicateState
    $script:GraphState.ExpectedTenant = $testTenant
    $global:IntuneMockState.Stops = 0
    $global:IntuneMockState.ScanComplete = $null
    $global:IntuneMockState.ScanOverlay = ''
    $cancelMocks = [System.Collections.Generic.Dictionary[string,scriptblock]]::new()
    $cancelMocks.Add('New-IntuneToolboxFeatureWorker', {
        param($RelativePath, $Parameters, $Scopes, $OutputAction)
        & $OutputAction ([PSCustomObject]@{ TenantId = '11111111-1111-1111-1111-111111111111'; Rows = @($eligibleRow) })
        return {}
    })
    $cancelMocks.Add('Invoke-InBackground', {
        param($Worker, $Buttons, $BusyLabel, $LogBox, $LogScroll, $DarkMode, $OverlayMessage, $OnComplete, $CancelHandle)
        $CancelHandle.Stop = { $global:IntuneMockState.Stops++ }
        $global:IntuneMockState.ScanComplete = $OnComplete
        $global:IntuneMockState.ScanOverlay = $OverlayMessage
    })
    [void]$scanFunction.Body.GetScriptBlock().InvokeWithContext($cancelMocks, $null, @())
    Assert-IntuneTest ($state.Busy -and $BtnDuplicateCancel.IsEnabled -and -not $BtnDuplicateDelete.IsEnabled -and
        -not $global:IntuneMockState.ScanOverlay) 'Scan leaves progress and cancellation accessible without a blocking overlay'
    $BtnDuplicateCancel.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    Assert-IntuneTest ($global:IntuneMockState.Stops -eq 1 -and -not $BtnDuplicateCancel.IsEnabled) 'Cancel stops the scan worker once'
    & $global:IntuneMockState.ScanComplete
    Update-DuplicateDeviceState
    Assert-IntuneTest (-not $state.Valid -and $null -eq $state.Report -and -not $state.Busy -and $state.Message -like '*cancelled*' -and
        -not $BtnDuplicateDelete.IsEnabled -and -not $BtnDuplicateExport.IsEnabled -and -not $BtnDuplicateCancel.IsEnabled) 'Cancelled scan publishes no report even if results arrived'
    $state.Busy = $true
    Update-DuplicateDeviceState
    Assert-IntuneTest (-not $BtnDuplicateCancel.IsEnabled) 'Deletion runs cannot be cancelled'
    $state.Busy = $false
    $completionHandle = @{ Stop = { 'pending' } }
    $completion = New-IntuneTestCompletion -CancelHandle $completionHandle
    & $completion
    Assert-IntuneTest ($null -eq $completionHandle.Stop) 'Background completion clears the cancel hook'

    $global:IntuneMockState.BackgroundDone = $false
    $liveHandle = @{ Stop = $null }
    $state.Busy = $true
    $state.Cancel = $liveHandle
    $state.ScanTimer = [Diagnostics.Stopwatch]::StartNew()
    $sleepWorker = [scriptblock]::Create('param($WriteLog) & $WriteLog "Requesting inventory page 1 (waiting for Graph)..." "INFO"; Start-Sleep -Seconds 60')
    Invoke-InBackground -Worker $sleepWorker -Buttons @($BtnDuplicateScan) -BusyLabel 'Testing...' -LogBox $TxtDuplicateLog `
        -LogScroll $DuplicateLogScroll -OnComplete {
            $script:DuplicateState.Busy = $false
            $script:DuplicateState.Cancel = $null
            $script:DuplicateState.ScanTimer.Stop()
            $global:IntuneMockState.BackgroundDone = $true
        } -CancelHandle $liveHandle
    Wait-IntuneBackground { $TxtDuplicateCount.Text -like '*waiting for Graph*Elapsed:*' }
    Assert-IntuneTest ($state.Busy -and $TxtDuplicateCount.Text -like '*waiting for Graph*Elapsed:*') 'UI timer shows last scan phase and elapsed time while a background request waits'
    $stopTimer = [Diagnostics.Stopwatch]::StartNew()
    & $liveHandle.Stop
    while (-not $global:IntuneMockState.BackgroundDone -and $stopTimer.Elapsed.TotalSeconds -lt 15) {
        $frame = [System.Windows.Threading.DispatcherFrame]::new()
        [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
            [System.Windows.Threading.DispatcherPriority]::Background, [Action]{ $frame.Continue = $false })
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
        Start-Sleep -Milliseconds 50
    }
    Assert-IntuneTest ($global:IntuneMockState.BackgroundDone -and $stopTimer.Elapsed.TotalSeconds -lt 15 -and
        $BtnDuplicateScan.IsEnabled -and $null -eq $liveHandle.Stop) 'Cancel hook stops a running background worker and restores the buttons'
    $state.Busy = $false
    $state.Cancel = $null
    $state.ScanTimer.Stop()

    # End-to-end: real runner + worker script + queues, with a fake Graph module supplying a signed-in context.
    $e2eRoot = Join-Path ([IO.Path]::GetTempPath()) ("IntuneToolboxE2E-" + [guid]::NewGuid().ToString('N'))
    $fakeModuleDir = Join-Path $e2eRoot 'Modules\Microsoft.Graph.Authentication'
    [void](New-Item -ItemType Directory -Path $fakeModuleDir -Force)
    Set-Content -LiteralPath (Join-Path $fakeModuleDir 'Microsoft.Graph.Authentication.psm1') -Value @'
function Get-MgContext {
    [PSCustomObject]@{ TenantId = 'e2e-tenant'; Account = 'e2e@contoso.test'; Scopes = @('DeviceManagementConfiguration.ReadWrite.All', 'DeviceManagementApps.Read.All') }
}
function Connect-MgGraph { throw 'Connect-MgGraph must not be called from a GUI worker.' }
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param($Uri, $Method, $OutputType)
    if ($Uri -like '*skiptoken*') { return @{ value = @(); '@odata.nextLink' = $null } }
    return @{ value = @(@{
        id = 'e2e-audit'; activityDateTime = [DateTimeOffset]::Now.ToString('o'); activity = 'Update policy'
        actor = @{ userPrincipalName = 'e2e@contoso.test' }; resources = @(@{ displayName = 'E2E policy' })
    }); '@odata.nextLink' = 'https://graph.microsoft.com/beta/deviceManagement/auditEvents?$skiptoken=e2e' }
}
Export-ModuleMember -Function Get-MgContext, Connect-MgGraph, Invoke-MgGraphRequest
'@
    $stubFeature = Join-Path $e2eRoot 'Stub-Feature.ps1'
    Set-Content -LiteralPath $stubFeature -Value @'
param([string]$Name, [switch]$SkipGraphConnect, [switch]$PassThru, [scriptblock]$LogAction)
& $LogAction "Stub feature ran for $Name" 'OK'
foreach ($index in 1..3) { [PSCustomObject]@{ Name = "$Name-$index"; SkipGraphConnect = [bool]$SkipGraphConnect } }
'@
    $savedModulePath = $env:PSModulePath
    $env:PSModulePath = (Join-Path $e2eRoot 'Modules') + [IO.Path]::PathSeparator + $savedModulePath
    try {
        $e2eItems = [System.Collections.Generic.List[object]]::new()
        $e2eState = @{ ExpectedTenant = 'e2e-tenant' }
        $global:IntuneMockState.BackgroundDone = $false
        $e2eWorker = [PSCustomObject]@{
            FeaturePath = $stubFeature
            RunnerPath = Join-Path $toolboxRoot 'Shared\Invoke-IntuneToolboxOperation.ps1'
            Parameters = @{ Name = 'Profile' }
            Scopes = @('DeviceManagementConfiguration.Read.All')
            State = $e2eState
            PermissionNote = $null
            PermissionError = $null
            OutputAction = { param($Item) $e2eItems.Add($Item) }.GetNewClosure()
        }
        Invoke-InBackground -Worker $e2eWorker -Buttons @($BtnDuplicateScan) -BusyLabel 'Testing...' -LogBox $TxtDuplicateLog `
            -LogScroll $DuplicateLogScroll -OnComplete { $global:IntuneMockState.BackgroundDone = $true }
        Wait-IntuneBackground { $global:IntuneMockState.BackgroundDone }
        $e2eLog = Get-IntuneLogText $TxtDuplicateLog
        Assert-IntuneTest ($global:IntuneMockState.BackgroundDone -and $BtnDuplicateScan.IsEnabled -and $BtnDuplicateScan.Content -ne 'Testing...') 'Signed-in background worker completes and restores the buttons'
        Assert-IntuneTest ($e2eLog -like '*Stub feature ran for Profile*' -and $e2eLog -notlike '*Feature failed*') 'Signed-in background worker streams feature logs to the UI'
        Assert-IntuneTest ($e2eItems.Count -eq 3 -and $e2eItems[2].Name -eq 'Profile-3' -and $e2eItems[0].SkipGraphConnect) 'Signed-in background worker delivers every result to the UI-thread output action'
        Assert-IntuneTest ($e2eState.Connected -and $e2eState.Account -eq 'e2e@contoso.test') 'Read scope is satisfied by the matching ReadWrite consent'
        $e2eItems.Clear()
        $global:IntuneMockState.BackgroundDone = $false
        $e2eWorker.FeaturePath = Join-Path $toolboxRoot 'Features\Search-IntuneAuditEvents.ps1'
        $e2eWorker.Parameters = @{ Target = 'E2E'; AsReport = $true }
        $e2eWorker.Scopes = @('DeviceManagementApps.Read.All')
        Invoke-InBackground -Worker $e2eWorker -Buttons @($BtnAuditSearch) -BusyLabel 'Testing...' -LogBox $TxtAuditLog `
            -LogScroll $AuditLogScroll -OnComplete { $global:IntuneMockState.BackgroundDone = $true }
        Wait-IntuneBackground { $global:IntuneMockState.BackgroundDone }
        Assert-IntuneTest ($global:IntuneMockState.BackgroundDone -and $e2eItems.Count -eq 1 -and
            $e2eItems[0].Completed -and $e2eItems[0].Rows.Count -eq 1 -and $e2eItems[0].PageCount -eq 2 -and
            $e2eItems[0].Rows[0].Id -eq 'e2e-audit' -and (Get-IntuneLogText $TxtAuditLog) -like '*Search complete: 1 matching*') 'Real audit feature runs through signed-in background runner, pages, logs, and complete-report queue'
    } finally {
        $env:PSModulePath = $savedModulePath
        Remove-Item -LiteralPath $e2eRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    $signInScopesAst = $guiAst.Find({
        param($node)
        $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$script:ToolboxGraphScopes'
    }, $true)
    $signInScopes = @([regex]::Matches($signInScopesAst.Right.Extent.Text, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
    $requestedScopes = @([regex]::Matches($guiAst.Extent.Text, "-Scopes @\(([^)]*)\)") | ForEach-Object {
        [regex]::Matches($_.Groups[1].Value, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value }
    } | Sort-Object -Unique)
    $uncoveredScopes = @($requestedScopes | Where-Object {
        $_ -notin $signInScopes -and -not ($_ -like '*.Read.All' -and ($_ -replace '\.Read\.All$', '.ReadWrite.All') -in $signInScopes)
    })
    Assert-IntuneTest ($requestedScopes.Count -ge 5 -and $uncoveredScopes.Count -eq 0) 'Sign-in requests every Graph permission the toolbox uses'
    $script:GraphState.ExpectedTenant = $testTenant
    $state = $script:DuplicateState
    $largeUiRows = @(foreach ($number in 1..7000) {
        $pair = [int][math]::Ceiling($number / 2)
        foreach ($type in @('Hardware', 'Entra')) {
            [PSCustomObject]@{
                GroupId = "$type-$pair"; MatchType = $type
                Id = ('00000000-0000-0000-0001-{0:D12}' -f $number); DeviceName = "PC-$number"
                SerialNumber = "SERIAL-$pair"; UserPrincipalName = 'user@example.test'
                ProtectionReason = $(if ($number % 2) { '' } else { 'Suggested keeper' }); CanDelete = ($number % 2 -eq 1)
            }
        }
    })
    $largeUiReport = [PSCustomObject]@{ TenantId = $testTenant; Rows = $largeUiRows }
    $global:IntuneMockState.DuplicateReport = $largeUiReport
    [void]$scanFunction.Body.GetScriptBlock().InvokeWithContext($scanMocks, $null, @())
    Assert-IntuneTest ($state.Valid -and $ListDuplicateDevices.Items.Count -eq 14000 -and
        $state.UniqueDeviceCount -eq 7000 -and $state.GroupCount -eq 7000 -and
        [object]::ReferenceEquals($state.CountsFor, $largeUiReport) -and $ListDuplicateDevices.SelectedItems.Count -eq 0) 'Actual scan completion publishes 14000 rows with cached exact device/group counts and no selection'
    [void]$ListDuplicateDevices.SelectedItems.Add($largeUiRows[0])
    [void]$ListDuplicateDevices.SelectedItems.Add($largeUiRows[1])
    Assert-IntuneTest ($BtnDuplicateDelete.IsEnabled -and $TxtDuplicateCount.Text -like '*1 selected.*' -and
        [object]::ReferenceEquals($state.CountsFor, $largeUiReport)) 'Selecting two memberships counts one eligible device without replacing cached report counts'
    [void]$ListDuplicateDevices.Items.SortDescriptions.Add([ComponentModel.SortDescription]::new('DeviceName', [ComponentModel.ListSortDirection]::Descending))
    $TxtDuplicateFilter.Text = ' pc-7000 '
    Assert-IntuneTest ($ListDuplicateDevices.Items.Count -eq 2 -and $ListDuplicateDevices.SelectedItems.Count -eq 0 -and
        -not $BtnDuplicateDelete.IsEnabled -and $state.Report.Rows.Count -eq 14000 -and
        $ListDuplicateDevices.Items.SortDescriptions.Count -eq 1 -and -not $state.Refreshing) 'Large text filter clears selection, preserves sorting/full export data, and restores refresh state'
    $ComboDuplicateMatch.SelectedIndex = 2
    Assert-IntuneTest ($ListDuplicateDevices.Items.Count -eq 1 -and $ListDuplicateDevices.Items[0].MatchType -eq 'Entra' -and
        $TxtDuplicateCount.Text -like '7000 unique devices / 7000 groups / 14000 membership rows (1 visible)*') 'Large match-type filter uses full-report cached totals and exact visible membership count'
    $ListDuplicateDevices.Items.SortDescriptions.Clear()
    $ComboDuplicateMatch.SelectedIndex = 0
    $TxtDuplicateFilter.Text = ''
    $state.Report = [PSCustomObject]@{ TenantId = $testTenant; Rows = @($eligibleRow, $protectedRow) }
    Update-DuplicateDeviceFilter
    Assert-IntuneTest ($state.UniqueDeviceCount -eq 2 -and $state.GroupCount -eq 1 -and
        [object]::ReferenceEquals($state.CountsFor, $state.Report)) 'Replacing a report recomputes cached counts rather than keeping stale large-report totals'
    $state.Report = $largeUiReport
    Update-DuplicateDeviceFilter
    $script:GraphState.ExpectedTenant = '22222222-2222-2222-2222-222222222222'
    Update-DuplicateDeviceState
    Assert-IntuneTest (-not $state.Valid -and $null -eq $state.CountsFor -and $state.UniqueDeviceCount -eq 0 -and
        $ListDuplicateDevices.Items.Count -eq 0 -and -not $BtnDuplicateExport.IsEnabled -and -not $BtnDuplicateDelete.IsEnabled) 'Tenant invalidation clears large report, count cache, selection, and actionable buttons'
    $script:GraphState.ExpectedTenant = $testTenant

    if ($RunScaleBenchmark) {
        foreach ($count in @(1000, 7000, 14000)) {
            $state.Report = [PSCustomObject]@{
                TenantId = $testTenant
                Rows = @(foreach ($number in 1..$count) {
                    [PSCustomObject]@{ Id = "device-$number"; GroupId = "Hardware-$([math]::Ceiling($number / 2))"; MatchType = 'Hardware'; DeviceName = "PC-$number"; SerialNumber = "SERIAL-$number"; UserPrincipalName = ''; ProtectionReason = ''; CanDelete = ($number % 2 -eq 1) }
                })
            }
            $state.Valid = $true
            Update-DuplicateDeviceFilter
            foreach ($run in 1..3) {
                $timer = [Diagnostics.Stopwatch]::StartNew()
                Update-DuplicateDeviceFilter
                $populateSeconds = $timer.Elapsed.TotalSeconds
                $timer.Restart()
                $ListDuplicateDevices.SelectedIndex = 0
                $selectionSeconds = $timer.Elapsed.TotalSeconds
                $timer.Restart()
                $TxtDuplicateFilter.Text = "PC-$count"
                $filterSeconds = $timer.Elapsed.TotalSeconds
                Assert-IntuneTest ($ListDuplicateDevices.Items.Count -eq 1 -and $state.Report.Rows.Count -eq $count -and
                    $ListDuplicateDevices.SelectedItems.Count -eq 0) "WPF benchmark $count rows run $run preserves selection/filter behavior"
                Write-Host ("WPF BENCHMARK rows={0} run={1} populateSeconds={2:N3} selectSeconds={3:N3} filterSeconds={4:N3}" -f $count, $run, $populateSeconds, $selectionSeconds, $filterSeconds)
                $TxtDuplicateFilter.Text = ''
            }
        }
    }
    Clear-DuplicateReport 'Test complete'
    $script:GraphState.ExpectedTenant = ''
} finally { $ui.Window.Close() }

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('Intune Toolbox Tests ' + [guid]::NewGuid().ToString('N'))
$portableRoot = Join-Path $temporaryRoot 'Copied Toolbox'
$portableUi = $null
try {
    Copy-Item -LiteralPath $toolboxRoot -Destination $portableRoot -Recurse
    Push-Location $temporaryRoot
    try {
        $portableUi = & (Join-Path $portableRoot 'Invoke-IntuneToolbox-GUI.ps1') -ValidateOnly
        Assert-IntuneTest ($portableUi.MissingFeatures.Count -eq 0) 'Copied toolbox resolves features from another working directory and a path with spaces'
    } finally { Pop-Location }
} finally {
    if ($portableUi) { $portableUi.Window.Close() }
    if (Test-Path -LiteralPath $temporaryRoot) { Remove-Item -LiteralPath $temporaryRoot -Recurse -Force }
}

& (Join-Path $PSScriptRoot 'Test-IntuneApps.ps1')
Write-Host "Completed $script:CheckCount Intune Toolbox checks."
