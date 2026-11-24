[CmdletBinding()]
param([switch]$RunScaleBenchmark)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Shared\Get-IntuneDuplicateDeviceAnalysis.ps1')
$script:DuplicateChecks = 0
function Assert-DuplicateTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:DuplicateChecks++
    Write-Host "PASS: $Message"
}
$tenant = '11111111-1111-1111-1111-111111111111'
$now = [DateTimeOffset]::UtcNow
function New-DuplicateTestDevice {
    param([int]$Number, [string]$Serial = 'SERIAL-1', [int]$Age = 90, [string]$Name = 'PC-1')
    [PSCustomObject]@{
        id = ('00000000-0000-0000-0001-{0:D12}' -f $Number); deviceName = $Name
        serialNumber = $Serial; manufacturer = 'Contoso'; model = 'Laptop'
        operatingSystem = 'Windows'; osVersion = '10.0'; azureADDeviceId = ''
        enrolledDateTime = $now.AddDays(-120).ToString('o'); lastSyncDateTime = $now.AddDays(-$Age).ToString('o')
        userPrincipalName = 'user@example.test'; managementAgent = 'mdm'; managementState = 'managed'; complianceState = 'compliant'
    }
}
function Get-TestAnalysis {
    param([object[]]$Devices)
    @(Get-IntuneDuplicateDeviceAnalysis -Devices $Devices -TenantId $tenant -NowUtc $now)
}
$old = New-DuplicateTestDevice 1
$keeper = New-DuplicateTestDevice 2 -Age 1
$rows = @(Get-TestAnalysis @($old, $keeper))
Assert-DuplicateTest ($rows.Count -eq 2 -and @($rows | Where-Object CanDelete).Count -eq 1 -and $rows[0].KeeperId -eq $keeper.id) 'Hardware group protects newest-sync keeper'
Assert-DuplicateTest ($rows[0].MatchType -eq 'Hardware' -and $rows[0].TenantId -eq $tenant -and $rows[0].Fingerprint.Length -eq 64) 'Report includes evidence, tenant, and revalidation fingerprints'
$exact = New-DuplicateTestDevice 1 -Age 60
$rows = @(Get-TestAnalysis @($exact, $keeper))
Assert-DuplicateTest (($rows | Where-Object Id -eq $exact.id).CanDelete) 'Exactly 60 full days is eligible'
$exact.lastSyncDateTime = $now.AddDays(-60).AddTicks(1).ToString('o')
$rows = @(Get-TestAnalysis @($exact, $keeper))
Assert-DuplicateTest (-not ($rows | Where-Object Id -eq $exact.id).CanDelete) 'Less than 60 days by one tick is protected'
foreach ($invalid in @('', 'not-a-date', '0001-01-01T00:00:00Z', '1970-01-01T00:00:00Z', $now.AddDays(1).ToString('o'))) {
    $item = New-DuplicateTestDevice 1
    $item.lastSyncDateTime = $invalid
    $rows = @(Get-TestAnalysis @($item, $keeper))
    Assert-DuplicateTest (@($rows | Where-Object CanDelete).Count -eq 0) "Invalid/default/future sync protects entire group: '$invalid'"
}
$tied = New-DuplicateTestDevice 2
Assert-DuplicateTest (@((Get-TestAnalysis @($old, $tied)) | Where-Object CanDelete).Count -eq 0) 'Tied newest sync requires review'
$trimmed = New-DuplicateTestDevice 1 -Serial ' serial-1 '
$trimmed.manufacturer = ' contoso '
Assert-DuplicateTest (@(Get-TestAnalysis @($trimmed, $keeper)).Count -eq 2) 'Identifiers normalize casing and outer whitespace'
$different = New-DuplicateTestDevice 3 -Serial 'SERIAL1' -Name 'PC-3'
Assert-DuplicateTest (@(Get-TestAnalysis @($old, $different)).Count -eq 0) 'Meaningful serial punctuation is preserved'
foreach ($serial in @('', 'To Be Filled By O.E.M.', 'Default string', '000000', 'UNKNOWN')) {
    $a = New-DuplicateTestDevice 1 -Serial $serial
    $b = New-DuplicateTestDevice 2 -Serial $serial -Age 1
    $rows = @(Get-TestAnalysis @($a, $b))
    Assert-DuplicateTest ($rows.Count -eq 2 -and @($rows | Where-Object MatchType -ne 'Name').Count -eq 0 -and
        @($rows | Where-Object CanDelete).Count -eq 0) "Placeholder serial never creates hardware evidence: '$serial'"
}
$a = New-DuplicateTestDevice 1 -Serial 'A'
$b = New-DuplicateTestDevice 2 -Serial 'B' -Age 1
$rows = @(Get-TestAnalysis @($a, $b))
Assert-DuplicateTest ($rows.Count -eq 2 -and $rows[0].MatchType -eq 'Name' -and @($rows | Where-Object CanDelete).Count -eq 0) 'Same name with different hardware is report-only'
$a.azureADDeviceId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
$b.azureADDeviceId = $a.azureADDeviceId
$rows = @(Get-TestAnalysis @($a, $b))
Assert-DuplicateTest ($rows[0].MatchType -eq 'Entra' -and @($rows | Where-Object CanDelete).Count -eq 0) 'Shared Entra ID with conflicting serials requires review'
$b.serialNumber = 'A'
$rows = @(Get-TestAnalysis @($a, $b))
Assert-DuplicateTest ($rows.Count -eq 4 -and @($rows | Where-Object CanDelete).Count -eq 2) 'Identical hardware and Entra memberships retain both evidence groups'
$c = New-DuplicateTestDevice 3 -Serial 'C' -Name 'PC-3'
$c.azureADDeviceId = $a.azureADDeviceId
$b.azureADDeviceId = ''
$rows = @(Get-TestAnalysis @($a, $b, $c))
Assert-DuplicateTest (@($rows | Where-Object CanDelete).Count -eq 0 -and @($rows | Where-Object ProtectionReason -like '*Overlapping*').Count -gt 0) 'Disagreeing overlapping evidence groups are protected'
foreach ($guid in @('', 'not-a-guid', '00000000-0000-0000-0000-000000000000')) {
    Assert-DuplicateTest ((ConvertTo-IntuneDuplicateGuid $guid) -eq '') "Invalid or empty Entra GUID is ignored: '$guid'"
}
$vm = New-DuplicateTestDevice 1
$vm.model = 'Virtual Machine'
$vmKeeper = New-DuplicateTestDevice 2 -Age 1
$vmKeeper.model = 'Virtual Machine'
Assert-DuplicateTest (@((Get-TestAnalysis @($vm, $vmKeeper)) | Where-Object CanDelete).Count -eq 0) 'Virtual-machine serial matches require review'
$conflict = New-DuplicateTestDevice 2 -Age 1
$conflict.model = 'Different model'
Assert-DuplicateTest (@((Get-TestAnalysis @($old, $conflict)) | Where-Object CanDelete).Count -eq 0) 'Conflicting models protect hardware matches'
$android = New-DuplicateTestDevice 3
$android.operatingSystem = 'Android'
Assert-DuplicateTest (@(Get-TestAnalysis @($old, $android)).Count -eq 0) 'Non-Windows inventory is excluded'
$android.operatingSystem = 'WindowsPhone'
Assert-DuplicateTest (@(Get-TestAnalysis @($old, $android)).Count -eq 0) 'WindowsPhone is not treated as a Windows desktop'
$android.operatingSystem = 'Android'
Assert-DuplicateTest (@(Get-TestAnalysis @($old, $old, $keeper)).Count -eq 2) 'Repeated identical managedDevice IDs are deduplicated'
$badDuplicate = New-DuplicateTestDevice 1 -Serial 'CHANGED'
$rejected = $false
try { Get-TestAnalysis @($old, $badDuplicate) | Out-Null } catch { $rejected = $true }
Assert-DuplicateTest $rejected 'Conflicting copies of one ID fail the scan'

$script:Mock = @{
    Devices = @($old, $keeper, $android); Writes = [System.Collections.Generic.List[string]]::new()
    ContextTenant = $tenant; Scopes = @('DeviceManagementManagedDevices.ReadWrite.All')
    FailPage = $false; RetryPage = $false; Sleeps = 0; ChangedLive = ''
    FailDelete = ''; MissingLive = ''; PageCalls = 0; InvalidPage = $false; RepeatedPage = $false
    FilteredCalls = 0; UnfilteredCalls = 0; RejectFilter = ''
    PageSize = 0; FailPageNumber = 0; TransientFailures = 0; RequestSettingsFail = $false
    ReadOnly = $false
}
$mockState = $script:Mock
function Get-MgContext { [PSCustomObject]@{ TenantId = $mockState.ContextTenant; Scopes = $mockState.Scopes } }
function Connect-MgGraph { throw 'Unexpected authentication in offline duplicate tests.' }
function Start-Sleep { param($Seconds) $mockState.Sleeps++ }
function Get-MgRequestContext {
    [CmdletBinding()]
    param()
    if ($mockState.RequestSettingsFail) { throw 'Mock request context unavailable.' }
    [PSCustomObject]@{ ClientTimeout = [TimeSpan]::FromSeconds(300); RetryDelay = 3; MaxRetry = 3; RetriesTimeLimit = [TimeSpan]::Zero }
}
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param($Method, $Uri)
    if ($Method -eq 'DELETE') {
        if ($mockState.ReadOnly) { throw 'Unexpected write in read-only scale test.' }
        $id = ($Uri -split '/')[-1]
        if ($id -eq $mockState.FailDelete) { throw 'Forbidden delete (403).' }
        $mockState.Writes.Add($id)
        return
    }
    if ($Uri -match '/managedDevices/([^?]+)') {
        $id = $Matches[1]
        if ($id -eq $mockState.MissingLive) { throw 'Not Found (404).' }
        $device = $mockState.Devices | Where-Object id -eq $id | Select-Object -First 1
        if ($id -eq $mockState.ChangedLive) {
            $copy = $device | Select-Object *
            $copy.lastSyncDateTime = $now.ToString('o')
            return $copy
        }
        return $device
    }
    $mockState.PageCalls++
    if ($mockState.PageSize -gt 0) {
        if ($mockState.TransientFailures -gt 0) { $mockState.TransientFailures--; throw '503 ServiceUnavailable' }
        $pageNumber = if ($Uri -match '\$skiptoken=scale-(\d+)') { [int]$Matches[1] } else { 1 }
        if ($pageNumber -eq $mockState.FailPageNumber) { throw 'Forbidden scale page (403).' }
        $start = ($pageNumber - 1) * $mockState.PageSize
        $end = [math]::Min($start + $mockState.PageSize, $mockState.Devices.Count) - 1
        $values = if ($start -le $end) { @($mockState.Devices[$start..$end]) } else { @() }
        return @{
            value = @($values)
            '@odata.nextLink' = if ($end + 1 -lt $mockState.Devices.Count) { "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$skiptoken=scale-$($pageNumber + 1)" } else { $null }
        }
    }
    if ($Uri -like '*$filter=*') {
        $mockState.FilteredCalls++
        if ($mockState.RejectFilter) { throw $mockState.RejectFilter }
    } else { $mockState.UnfilteredCalls++ }
    if ($Uri -like '*skiptoken*') {
        if ($mockState.FailPage) { throw 'Forbidden page (403).' }
        return [PSCustomObject]@{ value = @($mockState.Devices | Select-Object -Skip 1); '@odata.nextLink' = if ($mockState.RepeatedPage) { $Uri } else { $null } }
    }
    if ($mockState.RetryPage) { $mockState.RetryPage = $false; throw '503 ServiceUnavailable' }
    if ($mockState.InvalidPage) { return @{ value = $null } }
    return @{ value = @($mockState.Devices | Select-Object -First 1); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$skiptoken=page2' }
}
$getPath = Join-Path $root 'Features\Get-IntuneDuplicateDevices.ps1'
$removePath = Join-Path $root 'Features\Remove-IntuneDuplicateDevices.ps1'
$scan = & $getPath -AsReport -PassThru -SkipGraphConnect
Assert-DuplicateTest ($scan.Rows.Count -eq 2 -and $scan.InventoryCount -eq 2 -and $script:Mock.PageCalls -eq 2) 'Scan follows all pages and applies Windows filtering'
$runnerReport = & (Join-Path $root 'Shared\Invoke-IntuneToolboxOperation.ps1') -Path $getPath -Modules @() `
    -Scopes @('DeviceManagementManagedDevices.Read.All') -ExpectedTenant $tenant `
    -Parameters @{ AsReport = $true; PassThru = $true; SkipGraphConnect = $true }
Assert-DuplicateTest ($runnerReport.Rows.Count -eq 2) 'GUI operation runner accepts existing managed-device write scope for read-only scan'
$script:Mock.RetryPage = $true
$null = & $getPath -AsReport -PassThru -SkipGraphConnect
Assert-DuplicateTest ($script:Mock.Sleeps -eq 1) 'Inventory retries transient Graph failures'
$script:Mock.FailPage = $true
$rejected = $false
$partial = @()
try { $partial = @(& $getPath -AsReport -PassThru -SkipGraphConnect) } catch { $rejected = $true }
Assert-DuplicateTest ($rejected -and $partial.Count -eq 0) 'Failed continuation page never returns a partial success report'
$script:Mock.FailPage = $false
foreach ($flag in @('InvalidPage', 'RepeatedPage')) {
    $script:Mock[$flag] = $true
    $rejected = $false
    try { & $getPath -AsReport -PassThru -SkipGraphConnect | Out-Null } catch { $rejected = $true }
    Assert-DuplicateTest $rejected "Malformed/repeating Graph page fails explicitly: $flag"
    $script:Mock[$flag] = $false
}
$result = @(& $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -WhatIf)
Assert-DuplicateTest ($result[0].Status -eq 'WhatIf' -and $script:Mock.Writes.Count -eq 0) 'Duplicate deletion WhatIf revalidates without writes'
$script:Mock.Scopes = @('DeviceManagementManagedDevices.Read.All')
$result = @(& $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -WhatIf)
Assert-DuplicateTest ($result[0].Status -eq 'WhatIf') 'WhatIf needs no write consent'
$script:Mock.Scopes = @('DeviceManagementManagedDevices.ReadWrite.All')
$result = @(& $removePath -Id $keeper.id -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'Protected' -and $script:Mock.Writes.Count -eq 0) 'Keeper cannot be deleted'
$script:Mock.ChangedLive = $old.id
$result = @(& $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'ChangedSinceScan' -and $script:Mock.Writes.Count -eq 0) 'Fresh target changes block deletion'
$script:Mock.ChangedLive = $keeper.id
$result = @(& $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'ChangedSinceScan' -and $script:Mock.Writes.Count -eq 0) 'Fresh keeper changes block deletion'
$script:Mock.ChangedLive = ''
$script:Mock.MissingLive = $keeper.id
$result = @(& $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'Failed' -and $script:Mock.Writes.Count -eq 0) 'Missing keeper during final validation fails closed'
$script:Mock.MissingLive = ''
$script:Mock.FailDelete = $old.id
$result = @(& $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'Failed' -and $result[0].Message -like '*403*') 'Delete failures return explicit error outcomes'
$script:Mock.FailDelete = ''
$result = @(& $removePath -Id @($old.id, $old.id) -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result.Count -eq 1 -and $result[0].Status -eq 'Deleted' -and $script:Mock.Writes.Count -eq 1) 'Repeated selected IDs send exactly one DELETE'
$script:Mock.Devices = @($keeper)
$result = @(& $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'AlreadyAbsent' -and $script:Mock.Writes.Count -eq 1) 'Absent targets do not send writes'
$script:Mock.Devices = @($old, $keeper)
$script:Mock.ContextTenant = '22222222-2222-2222-2222-222222222222'
$rejected = $false
try { & $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false | Out-Null } catch { $rejected = $_.Exception.Message -like '*tenant*' }
Assert-DuplicateTest ($rejected -and $script:Mock.Writes.Count -eq 1) 'Tenant mismatch blocks deletion'
$script:Mock.ContextTenant = $tenant
$rejected = $false
try { & $removePath -Id 'not-a-guid' -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false | Out-Null } catch { $rejected = $_.Exception.Message -like '*Invalid managedDevice*' }
Assert-DuplicateTest $rejected 'Invalid selected device IDs are rejected explicitly'
$recent = New-DuplicateTestDevice 1 -Age 59
$script:Mock.Devices = @($recent, $keeper)
$recentScan = & $getPath -AsReport -PassThru -SkipGraphConnect
foreach ($row in $recentScan.Rows) { $row.CanDelete = $true }
$result = @(& $removePath -Id $recent.id -Scan $recentScan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'Protected' -and $script:Mock.Writes.Count -eq 1) 'Fresh analysis protects recent records even when caller flags are altered'
$nameOnly = New-DuplicateTestDevice 1 -Serial 'OTHER-SERIAL'
$script:Mock.Devices = @($nameOnly, $keeper)
$nameScan = & $getPath -AsReport -PassThru -SkipGraphConnect
foreach ($row in $nameScan.Rows) { $row.CanDelete = $true }
$result = @(& $removePath -Id $nameOnly.id -Scan $nameScan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'Protected' -and $script:Mock.Writes.Count -eq 1) 'Fresh analysis protects name-only records even when caller flags are altered'
$older = New-DuplicateTestDevice 3 -Age 100
$script:Mock.Devices = @($old, $older, $keeper)
$multiScan = & $getPath -AsReport -PassThru -SkipGraphConnect
$script:Mock.FailDelete = $old.id
$result = @(& $removePath -Id @($old.id, $older.id) -Scan $multiScan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result.Count -eq 2 -and $result[0].Status -eq 'Failed' -and $result[1].Status -eq 'Deleted') 'Partial delete failure preserves separate outcomes and continues eligible records'
$script:Mock.FailDelete = ''
$script:Mock.Devices = @($old, $keeper)
$snapshotChanged = New-DuplicateTestDevice 1
$snapshotChanged.enrolledDateTime = $now.AddDays(-115).ToString('o')
$script:Mock.Devices = @($snapshotChanged, $keeper)
$before = $script:Mock.Writes.Count
$result = @(& $removePath -Id $old.id -Scan $scan -SkipGraphConnect -PassThru -Confirm:$false)
Assert-DuplicateTest ($result[0].Status -eq 'ChangedSinceScan' -and $script:Mock.Writes.Count -eq $before) 'Inventory snapshot changes block deletion before final reads'
$script:Mock.Devices = @()
$empty = & $getPath -AsReport -PassThru -SkipGraphConnect
Assert-DuplicateTest ($empty.Rows.Count -eq 0 -and $empty.TenantId -eq $tenant) 'Empty scan returns explicit successful report context'

$script:Mock.Devices = @($old, $keeper, $android)
$script:Mock.FilteredCalls = 0
$script:Mock.UnfilteredCalls = 0
$logs = [System.Collections.Generic.List[string]]::new()
$logAction = { param($Message, $Level) $logs.Add("$Level|$Message") }.GetNewClosure()
$scan = & $getPath -AsReport -PassThru -SkipGraphConnect -LogAction $logAction
Assert-DuplicateTest ($script:Mock.FilteredCalls -eq 1 -and $script:Mock.UnfilteredCalls -eq 1 -and $scan.Rows.Count -eq 2) 'Inventory asks Graph for Windows devices and keeps the local check'
Assert-DuplicateTest (@($logs | Where-Object { $_ -like 'INFO|Downloaded page 1:*' }).Count -eq 1 -and
    @($logs | Where-Object { $_ -like 'INFO|Downloaded page 2: 3 records so far (2 Windows).' }).Count -eq 1) 'Scan logs download progress per page'
Assert-DuplicateTest (@($logs | Where-Object { $_ -like 'INFO|Download complete. Analysing 2 Windows records*' }).Count -eq 1 -and
    @($logs | Where-Object { $_ -like 'INFO|Analysis finished in*' }).Count -eq 1) 'Scan logs analysis start and finish'
$script:Mock.RejectFilter = 'BadRequest (400): Invalid filter clause.'
$script:Mock.FilteredCalls = 0
$script:Mock.UnfilteredCalls = 0
$logs.Clear()
$scan = & $getPath -AsReport -PassThru -SkipGraphConnect -LogAction $logAction
Assert-DuplicateTest ($script:Mock.FilteredCalls -eq 1 -and $script:Mock.UnfilteredCalls -eq 2 -and $scan.Rows.Count -eq 2 -and
    @($logs | Where-Object { $_ -like 'WARN|Graph rejected the Windows filter*' }).Count -eq 1) 'Rejected Windows filter falls back to the unfiltered inventory'
$script:Mock.RejectFilter = 'Forbidden (403).'
$script:Mock.UnfilteredCalls = 0
$rejected = $false
try { & $getPath -AsReport -PassThru -SkipGraphConnect | Out-Null } catch { $rejected = $_.Exception.Message -like '*403*' }
Assert-DuplicateTest ($rejected -and $script:Mock.UnfilteredCalls -eq 0) 'Permission errors are not hidden by the filter fallback'
$script:Mock.RejectFilter = ''

$script:Mock.Scopes = @('User.Read')
$rejected = $false
try {
    & (Join-Path $root 'Shared\Invoke-IntuneToolboxOperation.ps1') -Path $getPath -Modules @() -NonInteractive `
        -Scopes @('DeviceManagementManagedDevices.Read.All') -ExpectedTenant $tenant `
        -Parameters @{ AsReport = $true; PassThru = $true; SkipGraphConnect = $true } | Out-Null
} catch { $rejected = $_.Exception.Message -like 'Missing Microsoft Graph permission(s): DeviceManagementManagedDevices.Read.All.*' }
Assert-DuplicateTest $rejected 'Non-interactive runner reports missing permissions instead of opening sign-in'
$script:Mock.Scopes = @('DeviceManagementManagedDevices.ReadWrite.All')

$script:Mock.ReadOnly = $true
$script:Mock.PageSize = 250
$largeInventory = @(foreach ($number in 1..7000) {
    $pair = [int][math]::Ceiling($number / 2)
    $device = New-DuplicateTestDevice $number -Serial "SERIAL-$pair" -Name "PC-$pair" -Age $(if ($number % 2) { 90 } else { 1 })
    $device.azureADDeviceId = ('10000000-0000-0000-0001-{0:D12}' -f $pair)
    $device
})
$script:Mock.Devices = $largeInventory
$script:Mock.PageCalls = 0
$writesBefore = $script:Mock.Writes.Count
$logs.Clear()
$largeScan = & $getPath -AsReport -PassThru -SkipGraphConnect -LogAction $logAction
Assert-DuplicateTest ($largeScan.InventoryCount -eq 7000 -and $largeScan.Rows.Count -eq 14000 -and
    $script:Mock.PageCalls -eq 28 -and $script:Mock.Writes.Count -eq $writesBefore) '7000-device scan uses 28 page GETs, not per-duplicate requests, and sends no writes'
Assert-DuplicateTest (@($largeScan.Rows | Where-Object CanDelete).Count -eq 7000 -and
    @($largeScan.Rows | Where-Object { $_.Id -eq $largeInventory[0].id -and $_.KeeperId -eq $largeInventory[1].id }).Count -eq 2) 'Large hardware/Entra report keeps exact membership eligibility and keeper evidence'
Assert-DuplicateTest (($largeScan.PSObject.Properties.Name -join ',') -eq 'TenantId,ScanTimeUtc,MinimumDays,InventoryCount,Rows' -and
    ($largeScan.Rows[0].PSObject.Properties.Name -join ',') -eq ($scan.Rows[0].PSObject.Properties.Name -join ',')) 'Scale diagnostics preserve report and row schemas'
$firstRequest = $logs.FindIndex([Predicate[string]]{ param($line) $line -like 'INFO|Requesting inventory page 1,*' })
$firstCompletion = $logs.FindIndex([Predicate[string]]{ param($line) $line -like 'INFO|Downloaded page 1:*' })
Assert-DuplicateTest ($firstRequest -ge 0 -and $firstCompletion -gt $firstRequest -and
    @($logs | Where-Object { $_ -like 'INFO|Inventory page * finished in*' }).Count -eq 28) 'Diagnostics log request start before completion and duration for every page'
Assert-DuplicateTest (@($logs | Where-Object { $_ -like 'INFO|Analysis indexing:*' }).Count -eq 7 -and
    @($logs | Where-Object { $_ -like 'INFO|Analysis * finished in*' }).Count -eq 4 -and
    @($logs | Where-Object { $_ -like 'INFO|SDK request settings:*settings are unchanged.*' }).Count -eq 1 -and
    @($logs | Where-Object { $_ -like 'INFO|Report summary finished in*' }).Count -eq 1) 'Scale scan exposes SDK settings, coarse analysis progress, and phase timings'
Assert-DuplicateTest (@($logs | Where-Object { $_ -match 'skiptoken|SERIAL-|user@example|PC-' }).Count -eq 0) 'New scale progress messages do not expose continuation tokens or device/user payloads'

$script:Mock.Devices = @($largeInventory[0], $largeInventory[1], $largeInventory[0])
$script:Mock.PageSize = 1
$smallScan = & $getPath -AsReport -PassThru -SkipGraphConnect
Assert-DuplicateTest ($smallScan.InventoryCount -eq 3 -and $smallScan.Rows.Count -eq 4) 'Identical device IDs repeated across pages retain one set of memberships'
$changed = $largeInventory[0] | Select-Object *
$changed.serialNumber = 'CHANGED'
$script:Mock.Devices = @($largeInventory[0], $changed)
$logs.Clear()
$rejected = $false
try { & $getPath -AsReport -PassThru -SkipGraphConnect -LogAction $logAction | Out-Null } catch { $rejected = $true }
Assert-DuplicateTest ($rejected -and @($logs | Where-Object { $_ -like 'ERR|Scan failed during duplicate analysis*' }).Count -eq 1) 'Conflicting IDs across pages fail with analysis-phase diagnostics'

$script:Mock.Devices = $largeInventory
$script:Mock.PageSize = 250
$script:Mock.FailPageNumber = 20
$logs.Clear()
$partial = @()
$rejected = $false
try { $partial = @(& $getPath -AsReport -PassThru -SkipGraphConnect -LogAction $logAction) } catch { $rejected = $true }
Assert-DuplicateTest ($rejected -and $partial.Count -eq 0 -and
    @($logs | Where-Object { $_ -like 'ERR|Inventory page 20 failed after*' }).Count -eq 1 -and
    @($logs | Where-Object { $_ -like 'ERR|Scan failed during inventory download*' }).Count -eq 1) 'Late page failure reports phase/page context without publishing partial scale results'
$script:Mock.FailPageNumber = 0
$script:Mock.TransientFailures = 6
$script:Mock.PageCalls = 0
$sleepsBefore = $script:Mock.Sleeps
$logs.Clear()
$rejected = $false
try { & $getPath -AsReport -PassThru -SkipGraphConnect -LogAction $logAction | Out-Null } catch { $rejected = $true }
Assert-DuplicateTest ($rejected -and $script:Mock.PageCalls -eq 6 -and $script:Mock.Sleeps - $sleepsBefore -eq 5 -and
    @($logs | Where-Object { $_ -like 'WARN|Inventory retry*' }).Count -eq 5) 'Exhausted transient retries keep existing attempt limits and log waits before failing'
$script:Mock.Devices = @()
$script:Mock.RequestSettingsFail = $true
$logs.Clear()
$empty = & $getPath -AsReport -PassThru -SkipGraphConnect -LogAction $logAction
Assert-DuplicateTest ($empty.Rows.Count -eq 0 -and @($logs | Where-Object { $_ -like 'WARN|SDK request settings unavailable: Mock*' }).Count -eq 1) 'Unavailable optional SDK diagnostics are explicitly warned without losing a valid scan'
$script:Mock.RequestSettingsFail = $false

$diagnosticPath = Join-Path $PSScriptRoot 'Invoke-IntuneDuplicateScanDiagnostic.ps1'
$diagnosticDir = Join-Path ([IO.Path]::GetTempPath()) ("IntuneDuplicateDiag_" + [guid]::NewGuid().ToString('N'))
try {
    $script:Mock.Devices = $largeInventory
    $script:Mock.PageSize = 250
    $writesBefore = $script:Mock.Writes.Count
    $summary = & $diagnosticPath -SkipGraphConnect -OutputDirectory $diagnosticDir -PassThru 6>$null
    $diagnosticFiles = @(Get-ChildItem -LiteralPath $diagnosticDir -File)
    $diagnosticText = ($diagnosticFiles | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
    Assert-DuplicateTest ($summary.Status -eq 'Completed' -and $summary.InventoryCount -eq 7000 -and $summary.ReportRows -eq 14000 -and
        $summary.UniqueDevices -eq 7000 -and $summary.EvidenceGroups -eq 7000 -and $summary.DeletionCandidateRows -eq 7000 -and
        $summary.PagesCompleted -eq 28 -and $script:Mock.Writes.Count -eq $writesBefore -and $diagnosticFiles.Count -eq 2) 'Live diagnostic script summarizes a read-only 7000-device scan into log and summary files'
    Assert-DuplicateTest ($diagnosticText -like '*Inventory page 28 finished in*' -and $diagnosticText -notmatch [regex]::Escape($tenant) -and
        $diagnosticText -notmatch 'skiptoken|https://|SERIAL-|user@example|PC-') 'Live diagnostic files mask tenant IDs and URLs and omit device payloads'
    Remove-Item -LiteralPath $diagnosticDir -Recurse -Force
    $script:Mock.FailPageNumber = 3
    $failure = & $diagnosticPath -SkipGraphConnect -OutputDirectory $diagnosticDir -PassThru 6>$null
    $failureSummary = Get-Content -LiteralPath (Get-ChildItem -LiteralPath $diagnosticDir -Filter '*_summary.txt').FullName -Raw
    Assert-DuplicateTest ($LASTEXITCODE -eq 1 -and $failure.Status -eq 'Failed' -and $failure.PagesCompleted -eq 2 -and $failure.ReportRows -eq 0 -and
        $failureSummary -like '*Scan failed during inventory download*') 'Live diagnostic script records the failing phase and still writes a summary'
} finally {
    $script:Mock.FailPageNumber = 0
    $global:LASTEXITCODE = 0
    if (Test-Path -LiteralPath $diagnosticDir) { Remove-Item -LiteralPath $diagnosticDir -Recurse -Force }
}

$uniqueInventory = @($largeInventory | ForEach-Object {
    $copy = $_ | Select-Object *
    $copy.serialNumber = "UNIQUE-$($_.id)"
    $copy.deviceName = "PC-$($_.id)"
    $copy.azureADDeviceId = ''
    $copy
})
$uniqueRows = @(Get-TestAnalysis $uniqueInventory)
Assert-DuplicateTest ($uniqueRows.Count -eq 0) '7000 distinct devices produce no false duplicate rows'
foreach ($device in $uniqueInventory) { $device.deviceName = 'SHARED-NAME' }
$nameRows = @(Get-TestAnalysis $uniqueInventory)
Assert-DuplicateTest ($nameRows.Count -eq 7000 -and @($nameRows | Where-Object CanDelete).Count -eq 0 -and
    @($nameRows | Where-Object MatchType -ne 'Name').Count -eq 0) 'One 7000-member name-only group remains entirely protected'

if ($RunScaleBenchmark) {
    foreach ($count in @(1000, 7000, 14000)) {
        $script:Mock.Devices = @(foreach ($number in 1..$count) {
            $pair = [int][math]::Ceiling($number / 2)
            New-DuplicateTestDevice $number -Serial "SERIAL-$pair" -Name "PC-$pair" -Age $(if ($number % 2) { 90 } else { 1 })
        })
        $null = & $getPath -AsReport -PassThru -SkipGraphConnect
        foreach ($run in 1..3) {
            $script:Mock.PageCalls = 0
            $timer = [Diagnostics.Stopwatch]::StartNew()
            $benchmark = & $getPath -AsReport -PassThru -SkipGraphConnect -LogAction {
                param($Message, $Level)
                if ($Message -match 'finished in|total scan') { Write-Host "$Level $Message" }
            }
            Assert-DuplicateTest ($benchmark.InventoryCount -eq $count -and $benchmark.Rows.Count -eq $count -and
                $script:Mock.PageCalls -eq [math]::Ceiling($count / 250)) "Benchmark $count devices run $run preserves exact counts and page-only requests"
            Write-Host ("BENCHMARK devices={0} rows={1} pages={2} run={3} totalSeconds={4:N3}" -f $count, $benchmark.Rows.Count, $script:Mock.PageCalls, $run, $timer.Elapsed.TotalSeconds)
        }
    }
}
Write-Host "Completed $script:DuplicateChecks duplicate-device checks."
