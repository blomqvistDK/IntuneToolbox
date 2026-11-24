[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][object[]]$Profiles,
    [Parameter(Mandatory)][ValidateSet('Add', 'Replace', 'Remove')][string]$Action,
    [ValidateSet('Users', 'Devices', 'Group')][string]$TargetType,
    [ValidateSet('include', 'exclude')][string]$IncludeExclude = 'include',
    [object[]]$TargetGroup,
    [object]$AssignmentFilter,
    [ValidateSet('None', 'include', 'exclude')][string]$FilterMode = 'None',
    [ValidateSet('All', 'Group')][string]$RemoveScope = 'All',
    [ValidateSet('include', 'exclude', 'both')][string]$RemoveType = 'both',
    [object[]]$RemoveGroup,
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)

begin {
    function Get-UniqueAssignmentGroups {
        param([object[]]$Groups)
        if (-not $Groups.Count) { throw 'At least one group is required.' }
        $seen = @{}
        foreach ($group in $Groups) {
            if ([string]::IsNullOrWhiteSpace([string]$group.Id)) { throw 'Every group must have a non-empty Id.' }
            if (-not $seen.ContainsKey([string]$group.Id)) {
                $seen[[string]$group.Id] = $true
                $group
            }
        }
    }

    function Get-DirectAssignments {
        param([string]$Uri)
        do {
            $response = Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop
            foreach ($assignment in @($response.value)) {
                if (-not $assignment.source -or $assignment.source -eq 'direct') { $assignment }
            }
            $Uri = [string]$response.'@odata.nextLink'
        } while ($Uri)
    }

    function ConvertTo-AssignmentPayload {
        param([object]$Assignment)
        $target = [ordered]@{ '@odata.type' = $Assignment.target.'@odata.type' }
        if ($Assignment.target.groupId) { $target.groupId = $Assignment.target.groupId }
        if ($Assignment.target.deviceAndAppManagementAssignmentFilterId -and
            $Assignment.target.deviceAndAppManagementAssignmentFilterType -and
            $Assignment.target.deviceAndAppManagementAssignmentFilterType -ne 'none') {
            $target.deviceAndAppManagementAssignmentFilterId = $Assignment.target.deviceAndAppManagementAssignmentFilterId
            $target.deviceAndAppManagementAssignmentFilterType = $Assignment.target.deviceAndAppManagementAssignmentFilterType
        }
        @{ target = $target }
    }

    if ($Action -in @('Add', 'Replace')) {
        if (-not $TargetType) { throw 'An assignment target type is required.' }
        if ($FilterMode -ne 'None' -and [string]::IsNullOrWhiteSpace([string]$AssignmentFilter.Id)) {
            throw 'An assignment filter with a non-empty Id is required when FilterMode is enabled.'
        }
        if ($TargetType -eq 'Group') { $TargetGroup = @(Get-UniqueAssignmentGroups -Groups $TargetGroup) }
    } elseif ($RemoveScope -eq 'Group') {
        $RemoveGroup = @(Get-UniqueAssignmentGroups -Groups $RemoveGroup)
    }
    $targetName = if ($Action -eq 'Remove') {
        if ($RemoveScope -eq 'All') { 'all direct assignments' }
        else { "$RemoveType assignments for $($RemoveGroup.Count) group(s): $($RemoveGroup.DisplayName -join ', ')" }
    } elseif ($TargetType -eq 'Group') {
        "$($TargetGroup.Count) group(s): $($TargetGroup.DisplayName -join ', ')"
    } else { "All $TargetType" }

    $requiredScopes = @('DeviceManagementConfiguration.ReadWrite.All')
    if (-not $SkipGraphConnect -and -not $WhatIfPreference) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $context = Get-MgContext
        if (-not $context -or @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count) {
            $connectParameters = @{ Scopes = $requiredScopes; NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
            if ($TenantId) { $connectParameters.TenantId = $TenantId }
            Connect-MgGraph @connectParameters | Out-Null
        }
    }
    . (Join-Path $PSScriptRoot '..\Shared\Invoke-IntuneGraphWithRetry.ps1')
    $results = [System.Collections.Generic.List[object]]::new()
}
process {
    foreach ($item in $Profiles) {
        $status = 'Skipped'
        $message = ''
        $errorMessage = ''
        $operation = "$Action assignments for '$($item.CurrentName)'"
        if ($PSCmdlet.ShouldProcess($operation, "$Action $targetName")) {
            try {
                $result = Invoke-IntuneGraphWithRetry -ScriptBlock {
                    $base = 'https://graph.microsoft.com/beta/deviceManagement'
                    $profileUri = switch ($item.ApiType) {
                        'configurationPolicies' { "$base/configurationPolicies('$($item.Id)')" }
                        'deviceConfigurations' { "$base/deviceConfigurations/$($item.Id)" }
                        'groupPolicyConfigurations' { "$base/groupPolicyConfigurations/$($item.Id)" }
                        default { throw "Unsupported profile API type: '$($item.ApiType)'." }
                    }
                    $assignmentsUri = "$profileUri/assignments"
                    $assignUri = "$profileUri/assign"

                    if ($Action -in @('Add', 'Replace')) {
                        $odataType = switch ($TargetType) {
                            'Users' { '#microsoft.graph.allLicensedUsersAssignmentTarget' }
                            'Devices' { '#microsoft.graph.allDevicesAssignmentTarget' }
                            'Group' { if ($IncludeExclude -eq 'include') { '#microsoft.graph.groupAssignmentTarget' } else { '#microsoft.graph.exclusionGroupAssignmentTarget' } }
                        }
                        $targetCount = if ($TargetType -eq 'Group') { $TargetGroup.Count } else { 1 }
                        $newAssignments = @(for ($index = 0; $index -lt $targetCount; $index++) {
                            $target = [ordered]@{ '@odata.type' = $odataType }
                            if ($TargetType -eq 'Group') { $target.groupId = [string]$TargetGroup[$index].Id }
                            if ($FilterMode -ne 'None') {
                                $target.deviceAndAppManagementAssignmentFilterId = [string]$AssignmentFilter.Id
                                $target.deviceAndAppManagementAssignmentFilterType = $FilterMode.ToLowerInvariant()
                            }
                            @{ target = $target }
                        })
                        $assignments = @()
                        if ($Action -eq 'Add') {
                            $existing = @(Get-DirectAssignments -Uri $assignmentsUri)
                            $assignments = @($existing | Where-Object {
                                -not ($_.target.'@odata.type' -eq $odataType -and ($TargetType -ne 'Group' -or $_.target.groupId -in $TargetGroup.Id))
                            } | ForEach-Object {
                                ConvertTo-AssignmentPayload -Assignment $_
                            })
                        }
                        $assignments += $newAssignments
                        $body = @{ assignments = $assignments } | ConvertTo-Json -Depth 10
                        Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                        return "Assigned $targetName"
                    }

                    $direct = @(Get-DirectAssignments -Uri $assignmentsUri)
                    if ($RemoveScope -eq 'All') {
                        if (-not $direct.Count) { return 'No direct assignments found' }
                        Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body (@{ assignments = @() } | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop | Out-Null
                        return 'All direct assignments removed'
                    }
                    $types = switch ($RemoveType) {
                        'include' { @('#microsoft.graph.groupAssignmentTarget') }
                        'exclude' { @('#microsoft.graph.exclusionGroupAssignmentTarget') }
                        'both' { @('#microsoft.graph.groupAssignmentTarget', '#microsoft.graph.exclusionGroupAssignmentTarget') }
                    }
                    $updated = @($direct | Where-Object {
                        -not ($_.target.groupId -in $RemoveGroup.Id -and $_.target.'@odata.type' -in $types)
                    } | ForEach-Object {
                        ConvertTo-AssignmentPayload -Assignment $_
                    })
                    if ($updated.Count -ge $direct.Count) { return "No matching $targetName" }
                    Invoke-MgGraphRequest -Method POST -Uri $assignUri -Body (@{ assignments = $updated } | ConvertTo-Json -Depth 10) -ContentType 'application/json' -ErrorAction Stop | Out-Null
                    return "Removed $targetName"
                } -OnRetry {
                    param($attempt, $delaySeconds, $retryMessage)
                    if ($LogAction) { & $LogAction "[$($item.CurrentName)] Throttled; retry $attempt after ${delaySeconds}s." 'WARN' }
                }
                $status = 'OK'
                $message = [string]$result
            } catch {
                $status = 'Failed'
                $errorMessage = $_.Exception.Message
            }
        } elseif ($WhatIfPreference) { $status = 'WhatIf'; $message = "$Action $targetName" }

        $row = [PSCustomObject]@{ Id = [string]$item.Id; ProfileName = [string]$item.CurrentName; Action = $Action; Status = $status; Message = $message; Error = $errorMessage }
        $results.Add($row)
        $level = if ($status -eq 'OK') { 'OK' } elseif ($status -eq 'Failed') { 'ERR' } elseif ($status -eq 'WhatIf') { 'WARN' } else { 'INFO' }
        $detail = if ([string]::IsNullOrWhiteSpace($errorMessage)) { '' } else { ": $errorMessage" }
        if ($LogAction) { & $LogAction "[$($item.CurrentName)] $message$detail" $level }
    }
}
end {
    if ($PassThru) { $results.ToArray() }
    else { $results | Format-Table ProfileName, Action, Status, Message, Error }
}
