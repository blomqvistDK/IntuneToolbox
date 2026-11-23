. (Join-Path $PSScriptRoot 'Invoke-IntuneGraphWithRetry.ps1')

function Write-IntuneAppLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ($LogAction) { & $LogAction $Message $Level }
    else { Write-Information $Message -InformationAction Continue }
}

function Assert-IntuneAppContext {
    param([string]$Scope, [string]$Tenant, [switch]$Connect)
    if ($Connect) { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
    $context = Get-MgContext
    if ($Tenant -and $context -and $context.TenantId -ne $Tenant) { throw 'The Graph tenant changed. Sign out and refresh the app inventory.' }
    $hasScope = $context -and ($Scope -in $context.Scopes -or
        ($Scope -like '*.Read.All' -and ($Scope -replace '\.Read\.All$', '.ReadWrite.All') -in $context.Scopes))
    if (-not $hasScope -and $Connect) {
        $parameters = @{ Scopes = @(@($context.Scopes) + $Scope | Where-Object { $_ } | Select-Object -Unique); NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
        if ($Tenant) { $parameters.TenantId = $Tenant }
        elseif ($context.TenantId) { $parameters.TenantId = $context.TenantId }
        Connect-MgGraph @parameters | Out-Null
        return Assert-IntuneAppContext -Scope $Scope -Tenant $Tenant
    }
    if (-not $hasScope) { throw "Missing Microsoft Graph permission: $Scope." }
    return [string]$context.TenantId
}

function Get-IntuneAppPages {
    param([string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $page = 0
    do {
        $response = Invoke-IntuneGraphWithRetry -ScriptBlock {
            Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        } -OnRetry { param($attempt, $delay, $message) Write-IntuneAppLog "Read retry $attempt in ${delay}s: $message" 'WARN' }
        foreach ($item in @($response.value)) { $items.Add($item) }
        $page++
        Write-IntuneAppLog "Read page $page ($($items.Count) records)."
        $Uri = [string]$response.'@odata.nextLink'
    } while ($Uri)
    return $items.ToArray()
}

function Get-IntuneAppPlatform {
    param($App)
    $type = [string]$App.'@odata.type'
    switch -Regex ($type) {
        '\.(windows|win32|winGet|officeSuite|microsoftStoreForBusinessApp)' { return 'Windows' }
        '\.(managedAndroid|android)' { return 'Android' }
        '\.(managedIOS|ios)' { return 'iOS' }
        '\.(macOS|mac)' { return 'macOS' }
        default { return 'Other / platform-neutral' }
    }
}

function ConvertTo-IntuneAppDictionary {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in @($Value.Keys | Sort-Object)) { $result[$key] = ConvertTo-IntuneAppDictionary $Value[$key] }
        return $result
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return ,@(foreach ($item in $Value) { ConvertTo-IntuneAppDictionary $item })
    }
    if ($Value -is [PSCustomObject]) {
        $result = [ordered]@{}
        foreach ($property in @($Value.PSObject.Properties | Sort-Object Name)) {
            $result[$property.Name] = ConvertTo-IntuneAppDictionary $property.Value
        }
        return $result
    }
    return $Value
}

function Get-IntuneAppAssignmentFingerprint {
    param($Assignment)
    $data = [ordered]@{
        id = [string]$Assignment.id
        intent = [string]$Assignment.intent
        source = [string]$Assignment.source
        sourceId = [string]$Assignment.sourceId
        target = ConvertTo-IntuneAppDictionary $Assignment.target
        settings = ConvertTo-IntuneAppDictionary $Assignment.settings
    }
    return ($data | ConvertTo-Json -Depth 50 -Compress)
}

function Get-IntuneAppTargetKey {
    param($Target)
    return (([string]$Target.'@odata.type' -replace '^#', '') + '|' + [string]$Target.groupId).ToLowerInvariant()
}

function Get-IntuneAppTargetScope {
    param($Target)
    if ([string]$Target.'@odata.type' -match 'exclusionGroupAssignmentTarget$') { return 'exclude' }
    return 'include'
}

function Get-IntuneAppReplacePlan {
    <#
    Compares an app's current assignments with the desired targets for one intent and include/exclude scope.
    Policy Set-owned assignments, other intents and the other scope are never planned for removal.
    #>
    param([object[]]$Assignments, [string]$Intent, [string]$IncludeExclude, [object[]]$Targets)
    $plan = [System.Collections.Generic.List[object]]::new()
    $desiredKeys = @{}
    foreach ($target in $Targets) {
        $key = Get-IntuneAppTargetKey $target
        $desiredKeys[$key] = $true
        $existing = @($Assignments | Where-Object { (Get-IntuneAppTargetKey $_.target) -eq $key })
        $operation = 'Add'; $reason = 'Target has no assignment.'; $assignment = $null
        if ($existing.Count -gt 1) {
            $operation = 'Conflict'; $reason = 'Multiple assignments already use this target.'
        } elseif ($existing.Count -eq 1) {
            $assignment = $existing[0]
            if ($assignment.source -and $assignment.source -ne 'direct') {
                $operation = 'Conflict'; $reason = 'Target is assigned by a Policy Set.'
            } elseif ($assignment.intent -ne $Intent) {
                $operation = 'Conflict'; $reason = "Target already has a '$($assignment.intent)' assignment. Edit it explicitly."
            } elseif (Test-IntuneAppAssignmentMatch $assignment $Intent $target) {
                $operation = 'NoChange'; $reason = 'Matching assignment already exists.'
            } else {
                $operation = 'Update'; $reason = 'Assignment filter differs.'
            }
        }
        $plan.Add([PSCustomObject]@{ Operation = $operation; TargetKey = $key; Target = $target; Assignment = $assignment; Reason = $reason })
    }
    foreach ($assignment in $Assignments) {
        $key = Get-IntuneAppTargetKey $assignment.target
        if ($desiredKeys.ContainsKey($key) -or ($assignment.source -and $assignment.source -ne 'direct') -or
            $assignment.intent -ne $Intent -or (Get-IntuneAppTargetScope $assignment.target) -ne $IncludeExclude) { continue }
        $plan.Add([PSCustomObject]@{ Operation = 'Remove'; TargetKey = $key; Target = $assignment.target; Assignment = $assignment; Reason = 'Not in the queued targets.' })
    }
    return $plan.ToArray()
}

function Test-IntuneAppAssignmentMatch {
    param($Assignment, [string]$Intent, $Target)
    $existingMode = if ($Assignment.target.deviceAndAppManagementAssignmentFilterType) { [string]$Assignment.target.deviceAndAppManagementAssignmentFilterType } else { 'none' }
    $newMode = if ($Target.deviceAndAppManagementAssignmentFilterType) { [string]$Target.deviceAndAppManagementAssignmentFilterType } else { 'none' }
    return ($Assignment.intent -eq $Intent -and
        (Get-IntuneAppTargetKey $Assignment.target) -eq (Get-IntuneAppTargetKey $Target) -and
        $existingMode -eq $newMode -and ($newMode -eq 'none' -or
        $Assignment.target.deviceAndAppManagementAssignmentFilterId -eq $Target.deviceAndAppManagementAssignmentFilterId))
}
