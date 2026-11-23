function Connect-IntuneDuplicateGraph {
    [CmdletBinding()]
    param([string]$TenantId, [switch]$WriteAccess, [switch]$SkipGraphConnect)

    $scope = if ($WriteAccess) { 'DeviceManagementManagedDevices.ReadWrite.All' } else { 'DeviceManagementManagedDevices.Read.All' }
    if (-not $SkipGraphConnect) { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
    $context = Get-MgContext
    if ($TenantId -and $context -and $context.TenantId -ne $TenantId) { throw 'The Graph session belongs to a different tenant. Sign out first.' }
    $hasScope = $context -and ($scope -in $context.Scopes -or 'DeviceManagementManagedDevices.ReadWrite.All' -in $context.Scopes)
    if (-not $context -or -not $hasScope) {
        if ($SkipGraphConnect) { throw "An existing Graph session with $scope is required." }
        $parameters = @{ Scopes = @($scope); NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
        if ($TenantId) { $parameters.TenantId = $TenantId }
        elseif ($context) { $parameters.TenantId = $context.TenantId }
        Connect-MgGraph @parameters | Out-Null
        $context = Get-MgContext
    }
    if (-not $context -or ($scope -notin $context.Scopes -and 'DeviceManagementManagedDevices.ReadWrite.All' -notin $context.Scopes)) {
        throw "Sign-in did not grant $scope."
    }
    if ($TenantId -and $context.TenantId -ne $TenantId) { throw 'Tenant verification failed.' }
    return [string]$context.TenantId
}

function Assert-IntuneDuplicateTenant {
    param([Parameter(Mandatory)][string]$TenantId)
    $context = Get-MgContext
    if (-not $context -or $context.TenantId -ne $TenantId) { throw 'The Graph tenant changed. Scan again after reconnecting.' }
}

function Get-IntuneDuplicateDeviceFields {
    return 'id,deviceName,serialNumber,manufacturer,model,operatingSystem,osVersion,azureADDeviceId,enrolledDateTime,lastSyncDateTime,userPrincipalName,managementAgent,managementState,complianceState'
}

function Get-IntuneDuplicateInventory {
    param([Parameter(Mandatory)][string]$TenantId, [scriptblock]$LogAction)
    $baseUri = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=' + (Get-IntuneDuplicateDeviceFields)
    $filteredUri = $baseUri + '&$filter=' + [uri]::EscapeDataString("operatingSystem eq 'Windows'")
    $progress = @{ Pages = 0 }
    try {
        return (Get-IntuneDuplicateInventoryPages -Uri $filteredUri -TenantId $TenantId -LogAction $LogAction -Progress $progress)
    } catch {
        # Fall back only when Graph rejects the server-side filter on the first page.
        if ($progress.Pages -gt 0 -or $_.Exception.Message -notmatch '400|BadRequest|filter|not supported') { throw }
        if ($LogAction) { & $LogAction "Graph rejected the Windows filter; downloading all managed devices and filtering locally. ($($_.Exception.Message))" 'WARN' }
    }
    return (Get-IntuneDuplicateInventoryPages -Uri $baseUri -TenantId $TenantId -LogAction $LogAction -Progress @{ Pages = 0 })
}

function Get-IntuneDuplicateInventoryPages {
    param(
        [Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$TenantId,
        [scriptblock]$LogAction, [Parameter(Mandatory)][hashtable]$Progress
    )
    $uri = $Uri
    $devices = [System.Collections.Generic.List[object]]::new()
    $visited = [System.Collections.Generic.HashSet[string]]::new()
    $downloaded = 0
    $downloadTimer = [Diagnostics.Stopwatch]::StartNew()
    do {
        if (-not $visited.Add($uri)) { throw 'Graph repeated a managed-device continuation URL. Scan aborted.' }
        Assert-IntuneDuplicateTenant $TenantId
        $requestState = @{ Attempt = 0 }
        $page = $Progress.Pages + 1
        $pageTimer = [Diagnostics.Stopwatch]::StartNew()
        try {
            $response = Invoke-IntuneGraphWithRetry -ScriptBlock {
                Assert-IntuneDuplicateTenant $TenantId
                $requestState.Attempt++
                if ($LogAction) { & $LogAction "Requesting inventory page ${page}, attempt $($requestState.Attempt) (waiting for Graph; SDK retries may occur)..." 'INFO' }
                Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            } -OnRetry {
                param($attempt, $delay, $message)
                if ($LogAction) { & $LogAction "Inventory retry $attempt after ${delay}s: $message" 'WARN' }
            }
        } catch {
            if ($LogAction) { & $LogAction "Inventory page $page failed after $([math]::Round($pageTimer.Elapsed.TotalSeconds, 1)) s ($($requestState.Attempt) attempts): $($_.Exception.Message)" 'ERR' }
            throw
        }
        $Progress.Pages++
        $page = $Progress.Pages
        $hasValue = if ($response -is [System.Collections.IDictionary]) { $response.Contains('value') } else { $null -ne $response -and $null -ne $response.PSObject.Properties['value'] }
        if (-not $hasValue -or $null -eq $response.value -or $response.value -is [string] -or
            $response.value -isnot [System.Collections.IEnumerable]) {
            throw 'Graph returned an invalid managed-device collection.'
        }
        foreach ($device in @($response.value)) {
            $downloaded++
            if ([string]$device.operatingSystem -eq 'Windows') { $devices.Add($device) }
        }
        if ($LogAction) { & $LogAction "Downloaded page ${page}: $downloaded records so far ($($devices.Count) Windows)." 'INFO' }
        if ($LogAction) { & $LogAction "Inventory page $page finished in $([math]::Round($pageTimer.Elapsed.TotalSeconds, 1)) s; download elapsed $([math]::Round($downloadTimer.Elapsed.TotalSeconds, 1)) s." 'INFO' }
        $uri = [string]$response.'@odata.nextLink'
        if ($uri -and $uri -notlike 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?*') {
            throw 'Graph returned an unexpected managed-device continuation URL.'
        }
    } while ($uri)
    Assert-IntuneDuplicateTenant $TenantId
    return $devices.ToArray()
}

function ConvertTo-IntuneDuplicateKey {
    param($Value)
    return ([string]$Value).Trim().ToUpperInvariant()
}

function ConvertTo-IntuneDuplicateGuid {
    param($Value)
    $guid = [guid]::Empty
    if ([guid]::TryParse(([string]$Value).Trim(), [ref]$guid) -and $guid -ne [guid]::Empty) { return $guid.ToString() }
    return ''
}

function ConvertTo-IntuneDuplicateSerial {
    param($Value)
    $key = ConvertTo-IntuneDuplicateKey $Value
    if ($key -in @('', 'UNKNOWN', 'NONE', 'N/A', 'NA', 'NOT AVAILABLE', 'NOT SPECIFIED', 'DEFAULT STRING',
            'SYSTEM SERIAL NUMBER', 'TO BE FILLED BY O.E.M.', 'TO BE FILLED BY OEM') -or $key -match '^0+$') { return '' }
    return $key
}

function ConvertTo-IntuneDuplicateDate {
    param($Value, [DateTimeOffset]$NowUtc)
    $date = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$date) -and
        $date.Year -gt 1970 -and $date -le $NowUtc) { return $date.ToUniversalTime() }
    return $null
}

function Get-IntuneDuplicateFingerprint {
    param([Parameter(Mandatory)]$Device)
    $values = [ordered]@{}
    foreach ($field in ((Get-IntuneDuplicateDeviceFields) -split ',')) {
        $values[$field] = [string]$Device.$field
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($values | ConvertTo-Json -Compress)))).Replace('-', '')
    } finally { $sha.Dispose() }
}

function Get-IntuneDuplicateDeviceAnalysis {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][object[]]$Devices = @(),
        [Parameter(Mandatory)][string]$TenantId,
        [DateTimeOffset]$NowUtc = [DateTimeOffset]::UtcNow,
        [scriptblock]$LogAction
    )
    $phaseTimer = [Diagnostics.Stopwatch]::StartNew()
    if ($LogAction) { & $LogAction "Analysis phase: indexing $($Devices.Count) inventory records..." 'INFO' }
    $indexes = @{ Hardware = @{}; Entra = @{}; Name = @{} }
    $records = @{}
    $processed = 0
    foreach ($device in $Devices) {
        $processed++
        if ($LogAction -and $processed % 1000 -eq 0) { & $LogAction "Analysis indexing: $processed / $($Devices.Count) records." 'INFO' }
        if ([string]$device.operatingSystem -ne 'Windows') { continue }
        $id = ConvertTo-IntuneDuplicateGuid $device.id
        if (-not $id) { throw "Invalid managedDevice ID in inventory: '$($device.id)'." }
        $fingerprint = Get-IntuneDuplicateFingerprint $device
        if ($records.ContainsKey($id)) {
            if ($records[$id].Fingerprint -ne $fingerprint) { throw "Inventory changed across pages for managedDevice '$id'. Scan again." }
            continue
        }
        $serial = ConvertTo-IntuneDuplicateSerial $device.serialNumber
        $manufacturer = ConvertTo-IntuneDuplicateSerial $device.manufacturer
        $hardware = if ($serial -and $manufacturer) { ConvertTo-Json -InputObject @($serial, $manufacturer) -Compress } else { '' }
        $keys = @{ Hardware = $hardware; Entra = (ConvertTo-IntuneDuplicateGuid $device.azureADDeviceId); Name = (ConvertTo-IntuneDuplicateKey $device.deviceName) }
        $record = [PSCustomObject]@{
            Id = $id; Device = $device; Serial = $serial; Manufacturer = $manufacturer
            Keys = $keys; Sync = (ConvertTo-IntuneDuplicateDate $device.lastSyncDateTime $NowUtc)
            Fingerprint = $fingerprint; Groups = [System.Collections.Generic.List[object]]::new()
        }
        $records[$id] = $record
        foreach ($type in @('Hardware', 'Entra', 'Name')) {
            $key = $keys[$type]
            if (-not $key) { continue }
            if (-not $indexes[$type].ContainsKey($key)) { $indexes[$type][$key] = [System.Collections.Generic.List[object]]::new() }
            $indexes[$type][$key].Add($record)
        }
    }
    if ($LogAction) { & $LogAction "Analysis indexing finished in $([math]::Round($phaseTimer.Elapsed.TotalSeconds, 1)) s ($($records.Count) unique Windows records)." 'INFO' }
    $phaseTimer.Restart()
    if ($LogAction) { & $LogAction 'Analysis phase: building evidence groups...' 'INFO' }
    $groups = [System.Collections.Generic.List[object]]::new()
    $processed = 0
    foreach ($type in @('Hardware', 'Entra', 'Name')) {
        foreach ($key in @($indexes[$type].Keys | Sort-Object)) {
            $processed++
            if ($LogAction -and $processed % 1000 -eq 0) { & $LogAction "Analysis grouping: $processed keys processed ($($groups.Count) evidence groups)." 'INFO' }
            $members = @($indexes[$type][$key] | Sort-Object Id)
            if ($members.Count -lt 2) { continue }
            if ($type -eq 'Name') {
                $hardwareKeys = @($members | ForEach-Object { $_.Keys.Hardware } | Select-Object -Unique)
                $entraKeys = @($members | ForEach-Object { $_.Keys.Entra } | Select-Object -Unique)
                if (($hardwareKeys.Count -eq 1 -and $hardwareKeys[0]) -or ($entraKeys.Count -eq 1 -and $entraKeys[0])) { continue }
            }
            $reason = ''
            if ($type -eq 'Name') { $reason = 'Name-only candidate; not hardware evidence.' }
            elseif (@($members | Where-Object { $null -eq $_.Sync }).Count) { $reason = 'Missing, invalid, default, or future sync date; review required.' }
            elseif (@($members | Where-Object { -not $_.Serial -or -not $_.Manufacturer }).Count) { $reason = 'Incomplete or placeholder hardware identifiers; review required.' }
            elseif (@($members.Serial | Select-Object -Unique).Count -gt 1 -or
                @($members.Manufacturer | Select-Object -Unique).Count -gt 1 -or
                @($members | ForEach-Object { ConvertTo-IntuneDuplicateKey $_.Device.model } | Where-Object { $_ } | Select-Object -Unique).Count -gt 1) {
                $reason = 'Conflicting hardware identifiers; review required.'
            } elseif (@($members | Where-Object { [string]$_.Device.model -match 'Virtual|VMware|KVM|HVM|Parallels' }).Count) {
                $reason = 'Virtual-machine identifiers may be cloned; review required.'
            }
            $keeper = $null
            if (-not $reason) {
                $ordered = @($members | Sort-Object Sync -Descending)
                if ($ordered[0].Sync -eq $ordered[1].Sync) { $reason = 'Newest sync is tied; review required.' }
                else { $keeper = $ordered[0] }
            }
            $group = [PSCustomObject]@{
                Id = "$type-$($groups.Count + 1)"; Type = $type; Key = $key; Members = $members
                MemberIds = ($members.Id -join ','); Reason = $reason; Keeper = $keeper
            }
            $groups.Add($group)
            foreach ($member in $members) { $member.Groups.Add($group) }
        }
    }
    if ($LogAction) { & $LogAction "Analysis grouping finished in $([math]::Round($phaseTimer.Elapsed.TotalSeconds, 1)) s ($($groups.Count) evidence groups)." 'INFO' }
    $phaseTimer.Restart()
    if ($LogAction) { & $LogAction 'Analysis phase: checking overlapping-group protections...' 'INFO' }
    $processed = 0
    foreach ($record in $records.Values) {
        $processed++
        if ($LogAction -and $processed % 1000 -eq 0) { & $LogAction "Analysis protections: $processed / $($records.Count) records." 'INFO' }
        $strong = @($record.Groups | Where-Object Type -ne 'Name')
        if (@($strong.MemberIds | Select-Object -Unique).Count -gt 1) {
            foreach ($group in $strong) { $group.Reason = 'Overlapping strong-match groups disagree; review required.'; $group.Keeper = $null }
        }
    }
    if ($LogAction) { & $LogAction "Analysis protections finished in $([math]::Round($phaseTimer.Elapsed.TotalSeconds, 1)) s." 'INFO' }
    $phaseTimer.Restart()
    if ($LogAction) { & $LogAction 'Analysis phase: constructing report rows...' 'INFO' }
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $groups) {
        foreach ($record in $group.Members) {
            $device = $record.Device
            $protections = [System.Collections.Generic.List[string]]::new()
            foreach ($membership in $record.Groups) {
                if ($membership.Reason) { $protections.Add($membership.Reason) }
                elseif ($membership.Keeper.Id -eq $record.Id) { $protections.Add('Suggested keeper; protected.') }
            }
            $days = if ($null -ne $record.Sync) { [math]::Floor([decimal]($NowUtc - $record.Sync).Ticks / [TimeSpan]::TicksPerDay) } else { $null }
            if ($null -eq $days) { $protections.Add('Unknown last sync; review required.') }
            elseif ($days -lt 60) { $protections.Add('Last sync is less than 60 days ago.') }
            $canDelete = $protections.Count -eq 0
            $recommendation = if ($canDelete) { 'Stale duplicate candidate' } elseif ($group.Keeper -and $group.Keeper.Id -eq $record.Id) { 'Suggested keeper' } else { 'Protected / review' }
            $rows.Add([PSCustomObject][ordered]@{
                GroupId = $group.Id; MatchType = $group.Type
                MatchReason = switch ($group.Type) {
                    Hardware { 'Same valid serial number and manufacturer' }
                    Entra { 'Same Entra device ID (shared registration)' }
                    Name { 'Same device name only' }
                }
                Confidence = if ($group.Type -eq 'Name') { 'Low' } elseif ($group.Reason) { 'Review' } else { 'Strong evidence' }
                GroupSize = $group.Members.Count; Id = $record.Id; DeviceName = [string]$device.deviceName
                SerialNumber = [string]$device.serialNumber; Manufacturer = [string]$device.manufacturer
                Model = [string]$device.model; AzureADDeviceId = [string]$device.azureADDeviceId
                UserPrincipalName = [string]$device.userPrincipalName; OperatingSystem = [string]$device.operatingSystem
                OSVersion = [string]$device.osVersion; EnrolledDateTime = [string]$device.enrolledDateTime
                LastSyncDateTime = [string]$device.lastSyncDateTime; DaysSinceLastSync = $days
                ComplianceState = [string]$device.complianceState; ManagementAgent = [string]$device.managementAgent
                ManagementState = [string]$device.managementState; KeeperId = if ($group.Keeper) { $group.Keeper.Id } else { '' }
                Recommendation = $recommendation; CanDelete = $canDelete
                ProtectionReason = (@($protections | Select-Object -Unique) -join ' ')
                TenantId = $TenantId; ScanTimeUtc = $NowUtc.ToUniversalTime().ToString('o')
                Fingerprint = $record.Fingerprint; KeeperFingerprint = if ($group.Keeper) { $group.Keeper.Fingerprint } else { '' }
            })
            if ($LogAction -and $rows.Count % 1000 -eq 0) { & $LogAction "Analysis report: $($rows.Count) membership rows constructed." 'INFO' }
        }
    }
    if ($LogAction) { & $LogAction "Analysis report finished in $([math]::Round($phaseTimer.Elapsed.TotalSeconds, 1)) s ($($rows.Count) membership rows)." 'INFO' }
    return $rows.ToArray()
}
