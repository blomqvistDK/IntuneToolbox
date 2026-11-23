<#
.SYNOPSIS
    Looks up Windows Autopilot devices by serial number and changes their group tags.

.DESCRIPTION
    Actions:
      ImportCsv    - Reads a headered CSV with a serial-number column and an optional GroupTag/NewGroupTag column.
      Lookup       - Resolves serial numbers to Autopilot devices and returns the current group tag,
                     assigned deployment profile, profile assignment date, and Intune device name.
      ApplyTag     - Sets one entered tag (-NewTag) on the supplied, uniquely matched devices.
      ApplyCsvTags - Sets each device's CSV tag (CsvTag). Blank CSV tags mean no change.
      ClearTag     - Clears the group tag on the supplied devices.
    FreshStart   - Requests Intune Fresh Start on selected Windows devices, removing apps and user data.
      Sync         - Requests a tenant-wide Autopilot sync.

    Profile details and assignment dates come from the Microsoft Graph beta endpoint; tag updates,
    sync, Fresh Start, and Intune device names use v1.0. Wait for the intended deployment profile
    assignment before Fresh Start; a tag update does not guarantee profile readiness.

.EXAMPLE
    & .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action Lookup -SerialNumber 'PF12345','PF67890' -PassThru

.EXAMPLE
    $rows = & .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action ImportCsv -Path .\devices.csv -PassThru
    $devices = & .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action Lookup -Device $rows -PassThru
    & .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action ApplyCsvTags -Device $devices -WhatIf -PassThru
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][ValidateSet('ImportCsv', 'Lookup', 'ApplyTag', 'ApplyCsvTags', 'ClearTag', 'FreshStart', 'Sync')][string]$Action,
    [string[]]$SerialNumber = @(),
    [object[]]$Device = @(),
    [string]$Path,
    [string]$NewTag = '',
    [switch]$SyncAfterUpdate,
    [int]$BulkLookupThreshold = 20,
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)

$ErrorActionPreference = 'Stop'
$script:MaxGroupTagLength = 128
$script:GraphV1 = 'https://graph.microsoft.com/v1.0/deviceManagement'
$script:GraphBeta = 'https://graph.microsoft.com/beta/deviceManagement'
$script:EmptyGuid = '00000000-0000-0000-0000-000000000000'
$writeActions = @('ApplyTag', 'ApplyCsvTags', 'ClearTag')

$requiredScopes = switch ($Action) {
    'Lookup' { @('DeviceManagementServiceConfig.Read.All', 'DeviceManagementManagedDevices.Read.All') }
    'FreshStart' { @('DeviceManagementServiceConfig.Read.All', 'DeviceManagementManagedDevices.Read.All', 'DeviceManagementManagedDevices.PrivilegedOperations.All') }
    'ImportCsv' { @() }
    default { @('DeviceManagementServiceConfig.ReadWrite.All') }
}
$needsConnect = $requiredScopes.Count -gt 0 -and -not $SkipGraphConnect -and
    -not ($WhatIfPreference -and ($Action -in $writeActions -or $Action -in @('Sync', 'FreshStart')))
if ($needsConnect) {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $context = Get-MgContext
    $missingScopes = @($requiredScopes | Where-Object { -not $context -or $_ -notin $context.Scopes })
    if (-not $context -or $missingScopes.Count) {
        $connectParameters = @{ Scopes = $requiredScopes; NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
        if ($TenantId) { $connectParameters.TenantId = $TenantId }
        Connect-MgGraph @connectParameters | Out-Null
    }
}
if ($Action -ne 'ImportCsv') { . (Join-Path $PSScriptRoot '..\Shared\Invoke-IntuneGraphWithRetry.ps1') }

#region Helpers
function Write-FeatureLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ($LogAction) { & $LogAction $Message $Level }
    else { Write-Information "[$Level] $Message" -InformationAction Continue }
}

function Get-NormalizedSerial {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToUpperInvariant()
}

function Get-StringValue {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return [string]$Value
}

function Test-GroupTagValue {
    param([AllowNull()][string]$Tag)
    if ($null -eq $Tag) { return '' }
    if ($Tag.Length -gt $script:MaxGroupTagLength) { return "Group tag exceeds $script:MaxGroupTagLength characters." }
    if ($Tag -match '[\x00-\x1F\x7F]') { return 'Group tag contains control characters.' }
    return ''
}

function Get-GraphErrorMessage {
    param($ErrorRecord)
    $message = $ErrorRecord.Exception.Message
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        try { $message = ($ErrorRecord.ErrorDetails.Message | ConvertFrom-Json).error.message }
        catch { $message = $ErrorRecord.ErrorDetails.Message }
    }
    $text = (($message -replace '\s+', ' ').Trim())
    if ("$($ErrorRecord.Exception.Message) $text" -match '\b403\b|Forbidden') { return "Permission denied: $text" }
    return $text
}

function Test-GraphNotFound {
    param($ErrorRecord)
    return ("$($ErrorRecord.Exception.Message) $($ErrorRecord.ErrorDetails.Message)" -match '\b404\b|NotFound|ResourceNotFound')
}

function Invoke-GraphRead {
    param([Parameter(Mandatory)][string]$Uri)
    $requestUri = $Uri
    Invoke-IntuneGraphWithRetry -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri $requestUri -ErrorAction Stop
    } -OnRetry {
        param($attempt, $delaySeconds, $retryMessage)
        Write-FeatureLog "Graph request throttled or unavailable; retry $attempt after ${delaySeconds}s." 'WARN'
    }
}

function Get-GraphPages {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    while ($nextUri) {
        $response = Invoke-GraphRead -Uri $nextUri
        foreach ($item in @($response.value)) { if ($null -ne $item) { $items.Add($item) } }
        $nextUri = [string]$response.'@odata.nextLink'
    }
    return $items.ToArray()
}

function ConvertTo-GraphDate {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $result = $null
    if ($Value -is [DateTimeOffset]) { $result = $Value }
    elseif ($Value -is [datetime]) {
        $dateValue = if ($Value.Kind -eq [DateTimeKind]::Unspecified) { [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) } else { $Value }
        $result = [DateTimeOffset]::new($dateValue.ToUniversalTime())
    } else {
        $parsed = [DateTimeOffset]::MinValue
        $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal
        if (-not [DateTimeOffset]::TryParse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) { return $null }
        $result = $parsed
    }
    if ($result.UtcDateTime.Year -le 1601) { return $null }
    return $result
}

function Format-GraphDate {
    param([AllowNull()][object]$Date)
    if ($null -eq $Date) { return '' }
    $local = $Date.ToLocalTime()
    return "$($local.ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)) (UTC$($local.ToString('zzz', [System.Globalization.CultureInfo]::InvariantCulture)))"
}

function New-DeviceRow {
    param([string]$RequestedSerial, [string]$CsvTag = '', [string]$CsvIssue = '', [string]$Source = 'Manual')
    [PSCustomObject]@{
        RecordType = 'Device'
        RequestedSerial = $RequestedSerial
        SerialNumber = ''
        DeviceName = ''
        GroupTag = ''
        CsvTag = $CsvTag
        CsvIssue = $CsvIssue
        ProfileName = ''
        ProfileId = ''
        ProfileAssignmentStatus = ''
        ProfileAssignedDate = ''
        ProfileAssignedDateTimeUtc = ''
        LookupStatus = ''
        ManagedDeviceStatus = ''
        Eligible = $false
        AutopilotId = ''
        ManagedDeviceId = ''
        AutopilotDisplayName = ''
        Model = ''
        Manufacturer = ''
        UpdateStatus = ''
        FreshStartStatus = ''
        PreviousTag = ''
        RequestedTag = ''
        LastAction = ''
        Details = ''
        Source = $Source
    }
}

function Add-RowDetail {
    param($Row, [string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $Row.Details = if ($Row.Details) { "$($Row.Details); $Text" } else { $Text }
}

function Get-CsvDelimiter {
    param([string]$HeaderLine)
    $commaCount = 0; $semicolonCount = 0; $inQuotes = $false
    foreach ($character in $HeaderLine.ToCharArray()) {
        if ($character -eq '"') { $inQuotes = -not $inQuotes }
        elseif (-not $inQuotes -and $character -eq ',') { $commaCount++ }
        elseif (-not $inQuotes -and $character -eq ';') { $semicolonCount++ }
    }
    if ($semicolonCount -gt $commaCount) { return ';' }
    return ','
}
#endregion Helpers

#region CSV import
function Import-SerialCsv {
    param([Parameter(Mandatory)][string]$CsvPath)
    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { throw "CSV file not found: $CsvPath" }
    $text = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $CsvPath).ProviderPath)
    $lines = @($text -split "\r?\n")
    $headerLine = $lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
    if (-not $headerLine) { throw 'The CSV file is empty.' }
    $delimiter = Get-CsvDelimiter -HeaderLine $headerLine

    try { $records = @($text | ConvertFrom-Csv -Delimiter $delimiter) }
    catch { throw "The CSV file could not be parsed: $($_.Exception.Message)" }
    if ($records.Count -eq 0) { throw 'The CSV file has a header but no data rows.' }

    $headers = @($records[0].PSObject.Properties.Name)
    $serialHeaders = @($headers | Where-Object { $_.Trim() -match '^(device\s*)?serial\s*(number|no\.?|nr\.?|#)?$' })
    $tagHeaders = @($headers | Where-Object { $_.Trim() -match '^(new\s*)?group\s*tag$' })
    if ($serialHeaders.Count -eq 0) {
        throw "The CSV file has no serial number column. Expected a header such as 'SerialNumber' (found: $($headers -join ', '))."
    }
    if ($serialHeaders.Count -gt 1) { throw "The CSV file has more than one serial number column: $($serialHeaders -join ', ')." }
    if ($tagHeaders.Count -gt 1) { throw "The CSV file has conflicting group tag columns: $($tagHeaders -join ', '). Keep only one." }
    $serialHeader = $serialHeaders[0]
    $tagHeader = if ($tagHeaders.Count) { $tagHeaders[0] } else { $null }

    $entries = [ordered]@{}
    $rowNumber = 1
    foreach ($record in $records) {
        $rowNumber++
        $serial = (Get-StringValue $record.$serialHeader).Trim()
        $tag = if ($tagHeader) { (Get-StringValue $record.$tagHeader).Trim() } else { '' }
        if (-not $serial) {
            if ($tag -or @($record.PSObject.Properties | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Value) }).Count) {
                $entries["#row$rowNumber"] = @{ Serial = ''; Tags = [System.Collections.Generic.List[string]]::new(); Rows = @($rowNumber); Issue = "Blank serial number on CSV row $rowNumber." }
            }
            continue
        }
        $key = Get-NormalizedSerial $serial
        if (-not $entries.Contains($key)) {
            $entries[$key] = @{ Serial = $serial; Tags = [System.Collections.Generic.List[string]]::new(); Rows = @(); Issue = '' }
        }
        $entries[$key].Tags.Add($tag)
        $entries[$key].Rows += $rowNumber
    }

    foreach ($key in $entries.Keys) {
        $entry = $entries[$key]
        $tagSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        $distinctTags = @($entry.Tags | Where-Object { $tagSet.Add($_) })
        $csvTag = ''
        $issue = $entry.Issue
        if ($distinctTags.Count -gt 1) {
            $shownTags = ($distinctTags | ForEach-Object { if ($_) { "'$_'" } else { '(blank)' } }) -join ', '
            $issue = "Conflicting CSV tags on rows $($entry.Rows -join ', '): $shownTags."
        } elseif ($distinctTags.Count -eq 1) {
            $csvTag = $distinctTags[0]
            $tagError = Test-GroupTagValue $csvTag
            if ($tagError) { $issue = "Invalid CSV tag: $tagError"; $csvTag = '' }
        }
        $row = New-DeviceRow -RequestedSerial $entry.Serial -CsvTag $csvTag -CsvIssue $issue -Source 'CSV'
        if (-not $entry.Serial) { $row.LookupStatus = 'Invalid'; $row.Details = $issue }
        $row
    }
}
#endregion CSV import

#region Lookup
function Get-ManagedDeviceInfo {
    param([string]$ManagedDeviceId, [string]$Serial, [hashtable]$ManagedIndex)
    $info = @{ Name = ''; Id = ''; Status = 'Not found in Intune'; Error = '' }
    if ($ManagedIndex) {
        $candidates = @($ManagedIndex[(Get-NormalizedSerial $Serial)])
        $linked = @($candidates | Where-Object { $_ -and [string]$_.id -eq $ManagedDeviceId })
        if ($linked.Count -eq 1) { $candidates = $linked }
    } else {
        $candidates = @()
        if ($ManagedDeviceId -and $ManagedDeviceId -ne $script:EmptyGuid) {
            try {
                $managed = Invoke-GraphRead -Uri "$script:GraphV1/managedDevices/$([uri]::EscapeDataString($ManagedDeviceId))?`$select=id,deviceName,serialNumber"
                if ($managed) { $candidates = @($managed) }
            } catch {
                if (-not (Test-GraphNotFound $_)) { $info.Error = "Intune device lookup failed: $(Get-GraphErrorMessage $_)"; $info.Status = 'Error'; return $info }
            }
        }
        if ($candidates.Count -eq 0 -and $Serial) {
            try {
                $filter = [uri]::EscapeDataString("serialNumber eq '$($Serial.Replace("'", "''"))'")
                $candidates = @(Get-GraphPages -Uri "$script:GraphV1/managedDevices?`$filter=$filter&`$select=id,deviceName,serialNumber" |
                    Where-Object { (Get-NormalizedSerial $_.serialNumber) -eq (Get-NormalizedSerial $Serial) })
            } catch {
                $info.Error = "Intune device lookup failed: $(Get-GraphErrorMessage $_)"; $info.Status = 'Error'; return $info
            }
        }
    }
    $candidates = @($candidates | Where-Object { $_ })
    if ($candidates.Count -eq 1) {
        $info.Name = Get-StringValue $candidates[0].deviceName
        $info.Id = Get-StringValue $candidates[0].id
        $info.Status = 'Found'
    } elseif ($candidates.Count -gt 1) {
        $info.Name = (@($candidates | ForEach-Object { Get-StringValue $_.deviceName }) | Sort-Object -Unique) -join ' | '
        $info.Status = "Multiple Intune devices ($($candidates.Count))"
    }
    return $info
}

function Get-ProfileInfo {
    param($AutopilotDevice, [hashtable]$ProfileMap)
    $info = @{ Name = ''; Id = ''; Error = '' }
    $status = Get-StringValue $AutopilotDevice.deploymentProfileAssignmentStatus
    $id = Get-StringValue $AutopilotDevice.id
    if ($ProfileMap -and $ProfileMap.ContainsKey($id)) {
        $info.Name = $ProfileMap[$id].Name; $info.Id = $ProfileMap[$id].Id
        return $info
    }
    if ($status -eq 'notAssigned') { $info.Name = '(not assigned)'; return $info }
    if ($ProfileMap) { return $info }
    try {
        $detail = Invoke-GraphRead -Uri "$script:GraphBeta/windowsAutopilotDeviceIdentities/$([uri]::EscapeDataString($id))?`$expand=deploymentProfile"
        if ($detail.deploymentProfile) {
            $info.Name = Get-StringValue $detail.deploymentProfile.displayName
            $info.Id = Get-StringValue $detail.deploymentProfile.id
        } elseif ($status -in @('', 'unknown', 'notAssigned')) {
            $info.Name = '(not assigned)'
        }
    } catch {
        $info.Error = "Profile lookup failed: $(Get-GraphErrorMessage $_)"
    }
    return $info
}

function Get-ProfileMap {
    $map = @{}
    $profiles = Get-GraphPages -Uri "$script:GraphBeta/windowsAutopilotDeploymentProfiles"
    foreach ($profileItem in $profiles) {
        $profileId = Get-StringValue $profileItem.id
        $assigned = Get-GraphPages -Uri "$script:GraphBeta/windowsAutopilotDeploymentProfiles/$([uri]::EscapeDataString($profileId))/assignedDevices"
        foreach ($assignedDevice in $assigned) {
            $map[[string]$assignedDevice.id] = @{ Name = Get-StringValue $profileItem.displayName; Id = $profileId }
        }
    }
    return $map
}

function Invoke-DeviceLookup {
    param([object[]]$Requests)
    $rows = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($request in $Requests) {
        if ($request -is [string]) { $request = New-DeviceRow -RequestedSerial $request.Trim() }
        $requested = Get-StringValue $request.RequestedSerial
        if (-not $requested) { $requested = Get-StringValue $request.SerialNumber }
        $row = New-DeviceRow -RequestedSerial $requested.Trim() -CsvTag (Get-StringValue $request.CsvTag) `
            -CsvIssue (Get-StringValue $request.CsvIssue) -Source $(if ($request.Source) { [string]$request.Source } else { 'Manual' })
        $key = Get-NormalizedSerial $row.RequestedSerial
        if (-not $key) {
            $row.LookupStatus = 'Invalid'
            $row.Details = if ($row.CsvIssue) { $row.CsvIssue } else { 'Blank serial number.' }
            $rows.Add($row); continue
        }
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $rows.Add($row)
    }
    $lookupRows = @($rows | Where-Object { $_.LookupStatus -ne 'Invalid' })
    if ($lookupRows.Count -eq 0) { return $rows.ToArray() }

    $bulk = $lookupRows.Count -gt $BulkLookupThreshold
    $autopilotIndex = $null
    $managedIndex = $null
    $profileMap = $null
    $perSerialMatches = @{}

    if (-not $bulk) {
        try {
            foreach ($row in $lookupRows) {
                $filter = [uri]::EscapeDataString("contains(serialNumber,'$($row.RequestedSerial.Replace("'", "''"))')")
                $perSerialMatches[(Get-NormalizedSerial $row.RequestedSerial)] = @(Get-GraphPages -Uri "$script:GraphBeta/windowsAutopilotDeviceIdentities?`$filter=$filter")
            }
        } catch {
            Write-FeatureLog "Filtered Autopilot lookup failed ($(Get-GraphErrorMessage $_)); loading the full Autopilot inventory instead." 'WARN'
            $bulk = $true
        }
    }

    if ($bulk) {
        Write-FeatureLog 'Loading Autopilot inventory...'
        $inventory = Get-GraphPages -Uri "$script:GraphBeta/windowsAutopilotDeviceIdentities"
        $autopilotIndex = @{}
        foreach ($item in $inventory) {
            $itemKey = Get-NormalizedSerial $item.serialNumber
            if (-not $itemKey) { continue }
            if (-not $autopilotIndex.ContainsKey($itemKey)) { $autopilotIndex[$itemKey] = [System.Collections.Generic.List[object]]::new() }
            $autopilotIndex[$itemKey].Add($item)
        }
        Write-FeatureLog "Loaded $($inventory.Count) Autopilot device(s)."
        try {
            Write-FeatureLog 'Loading Intune managed devices...'
            $managedDevices = Get-GraphPages -Uri "$script:GraphV1/managedDevices?`$select=id,deviceName,serialNumber"
            $managedIndex = @{}
            foreach ($item in $managedDevices) {
                $itemKey = Get-NormalizedSerial $item.serialNumber
                if (-not $itemKey) { continue }
                if (-not $managedIndex.ContainsKey($itemKey)) { $managedIndex[$itemKey] = [System.Collections.Generic.List[object]]::new() }
                $managedIndex[$itemKey].Add($item)
            }
        } catch {
            Write-FeatureLog "Intune managed-device inventory failed: $(Get-GraphErrorMessage $_). Falling back to per-device lookups." 'WARN'
            $managedIndex = $null
        }
        try {
            Write-FeatureLog 'Loading Autopilot deployment profile assignments...'
            $profileMap = Get-ProfileMap
        } catch {
            Write-FeatureLog "Deployment profile map failed: $(Get-GraphErrorMessage $_). Falling back to per-device lookups." 'WARN'
            $profileMap = $null
        }
    }

    $index = 0
    foreach ($row in $lookupRows) {
        $index++
        $key = Get-NormalizedSerial $row.RequestedSerial
        $candidates = if ($bulk) { @($autopilotIndex[$key]) } else { @($perSerialMatches[$key] | Where-Object { (Get-NormalizedSerial $_.serialNumber) -eq $key }) }
        $candidates = @($candidates | Where-Object { $_ })
        if ($candidates.Count -gt 1) {
            $row.LookupStatus = 'Ambiguous'
            $row.SerialNumber = Get-StringValue $candidates[0].serialNumber
            $row.GroupTag = (@($candidates | ForEach-Object { Get-StringValue $_.groupTag }) | Sort-Object -Unique) -join ' | '
            Add-RowDetail $row "$($candidates.Count) Autopilot records share this serial number; resolve the duplicates in Intune before updating."
        } elseif ($candidates.Count -eq 1) {
            $apDevice = $candidates[0]
            $row.SerialNumber = Get-StringValue $apDevice.serialNumber
            $row.AutopilotId = Get-StringValue $apDevice.id
            $row.GroupTag = Get-StringValue $apDevice.groupTag
            $row.ManagedDeviceId = Get-StringValue $apDevice.managedDeviceId
            if ($row.ManagedDeviceId -eq $script:EmptyGuid) { $row.ManagedDeviceId = '' }
            $row.AutopilotDisplayName = Get-StringValue $apDevice.displayName
            $row.Model = Get-StringValue $apDevice.model
            $row.Manufacturer = Get-StringValue $apDevice.manufacturer
            $row.ProfileAssignmentStatus = Get-StringValue $apDevice.deploymentProfileAssignmentStatus
            $assignedDate = ConvertTo-GraphDate $apDevice.deploymentProfileAssignedDateTime
            if ($assignedDate) {
                $row.ProfileAssignedDate = Format-GraphDate $assignedDate
                $row.ProfileAssignedDateTimeUtc = $assignedDate.UtcDateTime.ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
            }
            $profileInfo = Get-ProfileInfo -AutopilotDevice $apDevice -ProfileMap $profileMap
            $row.ProfileName = $profileInfo.Name
            $row.ProfileId = $profileInfo.Id
            Add-RowDetail $row $profileInfo.Error
            $row.LookupStatus = 'Found'
            $row.Eligible = $true
        } else {
            $row.LookupStatus = 'Not found'
            Add-RowDetail $row 'No Autopilot record with this serial number.'
        }

        if ($row.LookupStatus -ne 'Ambiguous') {
            $managedInfo = Get-ManagedDeviceInfo -ManagedDeviceId $row.ManagedDeviceId -Serial $row.RequestedSerial -ManagedIndex $managedIndex
            $row.DeviceName = $managedInfo.Name
            if ($managedInfo.Id) { $row.ManagedDeviceId = $managedInfo.Id }
            $row.ManagedDeviceStatus = $managedInfo.Status
            Add-RowDetail $row $managedInfo.Error
            if ($row.LookupStatus -eq 'Not found' -and $managedInfo.Status -ne 'Not found in Intune' -and $managedInfo.Status -ne 'Error') {
                $row.LookupStatus = 'Intune only'
                $row.Details = 'Enrolled in Intune but not registered in Autopilot; group tags cannot be set.'
            }
        }
        if ($row.CsvIssue) { Add-RowDetail $row $row.CsvIssue }
        if ($index % 25 -eq 0) { Write-FeatureLog "Looked up $index of $($lookupRows.Count) serial number(s)..." }
    }

    $found = @($rows | Where-Object { $_.LookupStatus -eq 'Found' }).Count
    Write-FeatureLog "Lookup complete: $found of $($rows.Count) serial number(s) matched a single Autopilot device." $(if ($found -eq $rows.Count) { 'OK' } else { 'WARN' })
    return $rows.ToArray()
}
#endregion Lookup

#region Updates
function Get-AutopilotIdentityV1 {
    param([string]$Id)
    Invoke-GraphRead -Uri "$script:GraphV1/windowsAutopilotDeviceIdentities/$([uri]::EscapeDataString($Id))"
}

function Invoke-TagUpdate {
    param([object[]]$Rows, [ValidateSet('ApplyTag', 'ApplyCsvTags', 'ClearTag')][string]$Mode, [string]$EnteredTag)

    if ($Mode -eq 'ApplyTag') {
        $EnteredTag = ([string]$EnteredTag).Trim()
        if ([string]::IsNullOrEmpty($EnteredTag)) { throw 'NewTag must not be blank. Use the ClearTag action to remove group tags.' }
        $tagError = Test-GroupTagValue $EnteredTag
        if ($tagError) { throw $tagError }
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $writes = [System.Collections.Generic.List[object]]::new()
    $seenIds = @{}
    foreach ($source in $Rows) {
        $row = New-DeviceRow -RequestedSerial ''
        foreach ($property in @($source.PSObject.Properties)) {
            if ($row.PSObject.Properties[$property.Name]) { $row.($property.Name) = $property.Value }
        }
        $row.RecordType = 'Device'
        $row.Details = ''
        $row.LastAction = $Mode
        $row.PreviousTag = Get-StringValue $source.GroupTag
        $row.RequestedTag = ''
        $row.UpdateStatus = ''
        $results.Add($row)

        if ($row.LookupStatus -ne 'Found' -or -not $row.Eligible -or -not $row.AutopilotId) {
            $row.UpdateStatus = "Skipped: $(if ($row.LookupStatus) { $row.LookupStatus } else { 'not looked up' })"; continue
        }
        if ($seenIds.ContainsKey($row.AutopilotId)) { $row.UpdateStatus = 'Skipped: duplicate selection'; continue }
        $seenIds[$row.AutopilotId] = $true

        $requested = switch ($Mode) {
            'ApplyTag' { $EnteredTag }
            'ClearTag' { '' }
            'ApplyCsvTags' { (Get-StringValue $row.CsvTag).Trim() }
        }
        if ($Mode -eq 'ApplyCsvTags') {
            if ($row.CsvIssue) { $row.UpdateStatus = "Skipped: $($row.CsvIssue)"; continue }
            if (-not $requested) { $row.UpdateStatus = 'Skipped: no CSV tag'; continue }
            $tagError = Test-GroupTagValue $requested
            if ($tagError) { $row.UpdateStatus = "Skipped: $tagError"; continue }
        }
        $row.RequestedTag = $requested
        if ($row.PreviousTag -ceq $requested) { $row.UpdateStatus = 'Skipped: already set'; continue }
        $writes.Add($row)
    }

    $accepted = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $writes) {
        $description = if ($row.RequestedTag) { "Set group tag '$($row.PreviousTag)' -> '$($row.RequestedTag)'" } else { "Clear group tag '$($row.PreviousTag)'" }
        if (-not $PSCmdlet.ShouldProcess("$($row.SerialNumber) ($($row.AutopilotId))", $description)) {
            $row.UpdateStatus = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped: not confirmed' }
            continue
        }
        try {
            $current = Get-AutopilotIdentityV1 -Id $row.AutopilotId
            $currentTag = Get-StringValue $current.groupTag
            if ((Get-NormalizedSerial $current.serialNumber) -ne (Get-NormalizedSerial $row.SerialNumber)) {
                $row.UpdateStatus = 'Stale: the Autopilot record no longer matches this serial number; refresh and retry.'
                continue
            }
            if ($currentTag -cne $row.PreviousTag) {
                $row.GroupTag = $currentTag
                $row.UpdateStatus = "Stale: tag changed to '$currentTag' since lookup; refresh and confirm again."
                continue
            }
        } catch {
            $row.UpdateStatus = if (Test-GraphNotFound $_) { 'Failed: Autopilot record no longer exists.' } else { "Failed: $(Get-GraphErrorMessage $_)" }
            continue
        }

        try {
            $updateUri = "$script:GraphV1/windowsAutopilotDeviceIdentities/$([uri]::EscapeDataString($row.AutopilotId))/updateDeviceProperties"
            $updateBody = @{ groupTag = $row.RequestedTag } | ConvertTo-Json -Compress
            $serial = $row.SerialNumber
            Invoke-IntuneGraphWithRetry -ScriptBlock {
                Invoke-MgGraphRequest -Method POST -Uri $updateUri -Body $updateBody -ContentType 'application/json' -ErrorAction Stop | Out-Null
            } -OnRetry {
                param($attempt, $delaySeconds, $retryMessage)
                Write-FeatureLog "[$serial] Throttled; retry $attempt after ${delaySeconds}s." 'WARN'
            }
            $row.UpdateStatus = 'Accepted'
            $accepted.Add($row)
        } catch {
            $row.UpdateStatus = "Failed: $(Get-GraphErrorMessage $_)"
        }
        $level = if ($row.UpdateStatus -eq 'Accepted') { 'OK' } else { 'ERR' }
        Write-FeatureLog "[$($row.SerialNumber)] '$($row.PreviousTag)' -> '$($row.RequestedTag)': $($row.UpdateStatus)" $level
    }

    foreach ($row in $accepted) {
        try {
            $readBack = Get-AutopilotIdentityV1 -Id $row.AutopilotId
            $readTag = Get-StringValue $readBack.groupTag
            if ($readTag -ceq $row.RequestedTag) {
                $row.GroupTag = $readTag
                $row.UpdateStatus = 'Updated'
            } else {
                $row.GroupTag = $readTag
                $row.UpdateStatus = 'Pending verification'
                Add-RowDetail $row 'Graph accepted the change but still reports the previous tag; refresh later to confirm.'
            }
        } catch {
            $row.UpdateStatus = 'Pending verification'
            Add-RowDetail $row "Read-back failed: $(Get-GraphErrorMessage $_)"
        }
    }

    foreach ($row in $results) {
        if ($row.UpdateStatus -like 'Skipped*' -or $row.UpdateStatus -like 'Stale*' -or $row.UpdateStatus -eq 'WhatIf') {
            $level = if ($row.UpdateStatus -like 'Stale*') { 'WARN' } else { 'INFO' }
            Write-FeatureLog "[$(if ($row.SerialNumber) { $row.SerialNumber } else { $row.RequestedSerial })] $($row.UpdateStatus)" $level
        }
    }

    $changed = @($results | Where-Object { $_.UpdateStatus -in @('Updated', 'Pending verification') }).Count
    $failed = @($results | Where-Object { $_.UpdateStatus -like 'Failed*' -or $_.UpdateStatus -like 'Stale*' }).Count
    Write-FeatureLog "Done: $changed changed, $failed failed or stale, $($results.Count - $changed - $failed) skipped." $(if ($failed) { 'WARN' } elseif ($changed) { 'OK' } else { 'INFO' })

    $output = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $results) { $output.Add($row) }
    if ($SyncAfterUpdate -and $changed -gt 0 -and -not $WhatIfPreference) { $output.Add((Invoke-AutopilotSync)) }
    return $output.ToArray()
}

function Invoke-DeviceFreshStart {
    param([object[]]$Rows)
    $seenIds = @{}
    foreach ($source in $Rows) {
        $row = New-DeviceRow -RequestedSerial ''
        foreach ($property in @($source.PSObject.Properties)) {
            if ($row.PSObject.Properties[$property.Name]) { $row.($property.Name) = $property.Value }
        }
        $row.LastAction = 'FreshStart'
        $row.FreshStartStatus = ''
        try {
            if (-not $row.Eligible -or $row.LookupStatus -ne 'Found' -or -not $row.AutopilotId -or
                -not $row.ManagedDeviceId -or $row.ManagedDeviceId -eq $script:EmptyGuid) {
                $row.FreshStartStatus = 'Skipped: requires a unique Autopilot and Intune device'
            } elseif ($row.FreshStartStatus -eq 'Requested' -or $source.FreshStartStatus -eq 'Requested') {
                $row.FreshStartStatus = 'Skipped: Fresh Start already requested; check Intune action status'
            } elseif ($row.UpdateStatus -eq 'Pending verification' -or $row.UpdateStatus -like 'Failed*' -or $row.UpdateStatus -like 'Stale*') {
                $row.FreshStartStatus = 'Skipped: resolve the group tag update and refresh first'
            } elseif ($seenIds.ContainsKey($row.ManagedDeviceId)) {
                $row.FreshStartStatus = 'Skipped: duplicate selection'
            } else {
                $seenIds[$row.ManagedDeviceId] = $true
                if ($PSCmdlet.ShouldProcess("$($row.SerialNumber) ($($row.ManagedDeviceId))", 'Fresh Start: remove apps and user data, reinstall Windows')) {
                    $current = Get-AutopilotIdentityV1 -Id $row.AutopilotId
                    if ((Get-NormalizedSerial $current.serialNumber) -ne (Get-NormalizedSerial $row.SerialNumber) -or
                        [string]$current.groupTag -cne [string]$row.GroupTag -or
                        ($current.managedDeviceId -and $current.managedDeviceId -ne $script:EmptyGuid -and $current.managedDeviceId -ne $row.ManagedDeviceId)) {
                        $row.FreshStartStatus = 'Stale: Autopilot identity, tag or Intune link changed; refresh first'
                    } else {
                        $managedUri = "$script:GraphV1/managedDevices/$([uri]::EscapeDataString($row.ManagedDeviceId))"
                        $managed = Invoke-GraphRead -Uri "$managedUri`?`$select=id,serialNumber,operatingSystem"
                        if ([string]$managed.id -ne $row.ManagedDeviceId -or
                            (Get-NormalizedSerial $managed.serialNumber) -ne (Get-NormalizedSerial $row.SerialNumber) -or
                            $managed.operatingSystem -ne 'Windows') {
                            $row.FreshStartStatus = 'Stale: Intune device is not the expected Windows device; refresh first'
                        } else {
                            $body = @{ keepUserData = $false } | ConvertTo-Json -Compress
                            Invoke-MgGraphRequest -Method POST -Uri "$managedUri/cleanWindowsDevice" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                            $row.FreshStartStatus = 'Requested'
                        }
                    }
                } else {
                    $row.FreshStartStatus = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped: not confirmed' }
                }
            }
        } catch {
            $row.FreshStartStatus = "Failed: $(Get-GraphErrorMessage $_)"
            Add-RowDetail $row 'Check Intune device action status before retrying; the request may have reached the service.'
        }
        Write-FeatureLog "[$($row.RequestedSerial)] Fresh Start: $($row.FreshStartStatus)" $(if ($row.FreshStartStatus -eq 'Requested') { 'OK' } else { 'WARN' })
        $row
    }
}

function Invoke-AutopilotSync {
    $status = 'Requested'; $errorMessage = ''
    if ($PSCmdlet.ShouldProcess('Windows Autopilot (tenant-wide)', 'Request Autopilot sync')) {
        try {
            Invoke-IntuneGraphWithRetry -ScriptBlock {
                Invoke-MgGraphRequest -Method POST -Uri "$script:GraphV1/windowsAutopilotSettings/sync" -ErrorAction Stop | Out-Null
            } -MaxRetries 2
            Write-FeatureLog 'Tenant-wide Autopilot sync requested.' 'OK'
        } catch {
            $status = 'Failed'; $errorMessage = Get-GraphErrorMessage $_
            Write-FeatureLog "Autopilot sync request failed: $errorMessage" 'ERR'
        }
    } else { $status = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' } }
    [PSCustomObject]@{ RecordType = 'Sync'; UpdateStatus = $status; Details = $errorMessage }
}
#endregion Updates

$results = switch ($Action) {
    'ImportCsv' {
        if (-not $Path) { throw 'Path is required for ImportCsv.' }
        $imported = @(Import-SerialCsv -CsvPath $Path)
        $issues = @($imported | Where-Object { $_.CsvIssue }).Count
        Write-FeatureLog "Imported $($imported.Count) unique serial number(s) from CSV$(if ($issues) { "; $issues row(s) need attention" })." $(if ($issues) { 'WARN' } else { 'OK' })
        $imported
    }
    'Lookup' {
        $requests = @($Device) + @($SerialNumber | Where-Object { $null -ne $_ } | ForEach-Object { New-DeviceRow -RequestedSerial ([string]$_).Trim() })
        if ($requests.Count -eq 0) { throw 'Provide at least one serial number or device row to look up.' }
        Invoke-DeviceLookup -Requests $requests
    }
    'Sync' { Invoke-AutopilotSync }
    'FreshStart' { Invoke-DeviceFreshStart -Rows $Device }
    default { Invoke-TagUpdate -Rows $Device -Mode $Action -EnteredTag $NewTag }
}

if ($PassThru) { $results }
else {
    $results | Where-Object { $_.RecordType -eq 'Device' } |
        Format-Table RequestedSerial, DeviceName, GroupTag, CsvTag, ProfileName, ProfileAssignedDate, LookupStatus, UpdateStatus, FreshStartStatus -AutoSize
}
