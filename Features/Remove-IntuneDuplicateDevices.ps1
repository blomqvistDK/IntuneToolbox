<#
.SYNOPSIS
Deletes explicitly selected stale strong-match Intune records after revalidation.
.DESCRIPTION
Requires the report from Get-IntuneDuplicateDevices.ps1 -AsReport -PassThru.
Protects keepers, name-only matches, ambiguous groups, and records synced within 60 days.
Windows deletion triggers Retire and can affect BitLocker protection. Back up recovery
keys and local administrator credentials before intentionally removing managed devices.
Entra and Autopilot registrations are not deleted. Supports -WhatIf.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][string[]]$Id,
    [Parameter(Mandatory)]$Scan,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Shared\Invoke-IntuneGraphWithRetry.ps1')
. (Join-Path $PSScriptRoot '..\Shared\Get-IntuneDuplicateDeviceAnalysis.ps1')
$tenant = ConvertTo-IntuneDuplicateGuid $Scan.TenantId
if (-not $tenant -or $Scan.MinimumDays -ne 60 -or -not $Scan.ScanTimeUtc) { throw 'A complete duplicate-device scan report is required.' }
$ids = @($Id | ForEach-Object {
    $parsed = ConvertTo-IntuneDuplicateGuid $_
    if (-not $parsed) { throw "Invalid managedDevice ID: '$_'." }
    $parsed
} | Select-Object -Unique)
if ($ids.Count -eq 0) { throw 'Select at least one managedDevice ID.' }
# WhatIf still revalidates using read permission, but never requests write consent.
$null = Connect-IntuneDuplicateGraph -TenantId $tenant -WriteAccess:(-not $WhatIfPreference) -SkipGraphConnect:$SkipGraphConnect
if ($LogAction) { & $LogAction "Revalidating: downloading a fresh Windows inventory for tenant $tenant..." 'INFO' }
$inventory = @(Get-IntuneDuplicateInventory -TenantId $tenant -LogAction $LogAction)
$rows = @(Get-IntuneDuplicateDeviceAnalysis -Devices $inventory -TenantId $tenant)
$currentById = @{}
foreach ($device in $inventory) { $currentById[(ConvertTo-IntuneDuplicateGuid $device.id)] = $device }
$results = foreach ($deviceId in $ids) {
    $status = 'Protected'
    $message = ''
    try {
        Assert-IntuneDuplicateTenant $tenant
        $old = @($Scan.Rows | Where-Object Id -eq $deviceId)
        $current = @($rows | Where-Object Id -eq $deviceId)
        if ($old.Count -eq 0) { $message = 'Selected ID was not in the original scan.' }
        elseif (-not $currentById.ContainsKey($deviceId)) { $status = 'AlreadyAbsent'; $message = 'Not present in the fresh Windows inventory; no delete sent.' }
        elseif ($current.Count -eq 0 -or @($current | Where-Object { -not $_.CanDelete }).Count) { $message = 'No longer an eligible stale strong match. Scan again.' }
        elseif (@($old | Where-Object { -not $_.CanDelete -or $_.TenantId -ne $tenant }).Count) { $message = 'Original scan protects this record.' }
        else {
            $changed = $false
            $oldEvidence = @($old | ForEach-Object { "$($_.MatchType)|$($_.KeeperId)|$($_.Fingerprint)|$($_.KeeperFingerprint)" } | Sort-Object)
            $newEvidence = @($current | ForEach-Object { "$($_.MatchType)|$($_.KeeperId)|$($_.Fingerprint)|$($_.KeeperFingerprint)" } | Sort-Object)
            if (($oldEvidence -join ';') -ne ($newEvidence -join ';')) { $changed = $true }
            $supportIds = @(@($current.KeeperId | Select-Object -Unique) + @($deviceId) | Select-Object -Unique)
            if (@($current.KeeperId | Where-Object { $_ -in $ids }).Count) { throw 'The selection includes a supporting keeper.' }
            foreach ($supportId in $supportIds) {
                if ($changed) { break }
                Assert-IntuneDuplicateTenant $tenant
                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${supportId}?`$select=$(Get-IntuneDuplicateDeviceFields)"
                $live = Invoke-IntuneGraphWithRetry -ScriptBlock {
                    Assert-IntuneDuplicateTenant $tenant
                    Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                }
                if (-not $live -or (ConvertTo-IntuneDuplicateGuid $live.id) -ne $supportId -or
                    (Get-IntuneDuplicateFingerprint $live) -ne (Get-IntuneDuplicateFingerprint $currentById[$supportId])) { $changed = $true }
            }
            if ($changed) { $status = 'ChangedSinceScan'; $message = 'Target or supporting keeper changed; scan again.' }
            elseif ($PSCmdlet.ShouldProcess("$($current[0].DeviceName) ($deviceId), tenant $tenant", 'Delete Intune device (triggers Windows Retire; can affect BitLocker)')) {
                Invoke-IntuneGraphWithRetry -ScriptBlock {
                    Assert-IntuneDuplicateTenant $tenant
                    # Recheck the target and keepers on each retry, not only before the first attempt.
                    foreach ($supportId in $supportIds) {
                        $live = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${supportId}?`$select=$(Get-IntuneDuplicateDeviceFields)" -ErrorAction Stop
                        if (-not $live -or (Get-IntuneDuplicateFingerprint $live) -ne (Get-IntuneDuplicateFingerprint $currentById[$supportId])) {
                            throw 'Target or keeper changed immediately before deletion. Scan again.'
                        }
                    }
                    Assert-IntuneDuplicateTenant $tenant
                    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$deviceId" -ErrorAction Stop | Out-Null
                } -OnRetry {
                    param($attempt, $delay, $errorMessage)
                    if ($LogAction) { & $LogAction "Delete retry $attempt after ${delay}s for ${deviceId}: $errorMessage" 'WARN' }
                }
                $status = 'Deleted'; $message = 'Server-side deletion completed; endpoint retirement is not confirmed.'
            } elseif ($WhatIfPreference) { $status = 'WhatIf'; $message = 'Eligible after fresh validation; no write sent.' }
            else { $status = 'Protected'; $message = 'Deletion was not confirmed.' }
        }
    } catch { $status = 'Failed'; $message = $_.Exception.Message }
    $result = [PSCustomObject]@{ Id = $deviceId; TenantId = $tenant; Status = $status; Message = $message }
    $level = if ($status -eq 'Deleted') { 'OK' } elseif ($status -eq 'Failed') { 'ERR' } else { 'WARN' }
    if ($LogAction) { & $LogAction "$deviceId [$status] $message" $level }
    $result
}
if ($PassThru) { $results } else { $results | Format-Table Id, Status, Message -Wrap }
