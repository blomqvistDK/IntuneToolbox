<#
.SYNOPSIS
Reports likely duplicate Windows managed devices without changing Intune.
.DESCRIPTION
Strong serial/manufacturer and Entra registration matches are separate from name-only candidates.
Only unambiguous non-keepers last synced at least 60 days ago can be deletion candidates.
Use -AsReport -PassThru to retain the complete scan context for guarded deletion.
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [switch]$AsReport,
    [scriptblock]$LogAction
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Shared\Invoke-IntuneGraphWithRetry.ps1')
. (Join-Path $PSScriptRoot '..\Shared\Get-IntuneDuplicateDeviceAnalysis.ps1')
$tenant = Connect-IntuneDuplicateGraph -TenantId $TenantId -SkipGraphConnect:$SkipGraphConnect
if ($LogAction) { & $LogAction "Scanning Windows managed devices in tenant $tenant..." 'INFO' }
if ($LogAction) {
    $graphModule = Get-Module Microsoft.Graph.Authentication | Select-Object -First 1
    $graphVersion = if ($graphModule) { [string]$graphModule.Version } else { 'not loaded (mock or caller-provided commands)' }
    & $LogAction "Scan runtime: PowerShell $($PSVersionTable.PSVersion); Graph Authentication $graphVersion." 'INFO'
    if (Get-Command Get-MgRequestContext -ErrorAction SilentlyContinue) {
        $settingsError = ''
        try {
            $requestContext = Get-MgRequestContext -ErrorAction Stop
            if (-not $requestContext) { throw 'No SDK request context was returned.' }
        } catch { $settingsError = $_.Exception.Message }
        if ($settingsError) { & $LogAction "SDK request settings unavailable: $settingsError" 'WARN' }
        else { & $LogAction "SDK request settings: ClientTimeout=$($requestContext.ClientTimeout); RetryDelay=$($requestContext.RetryDelay); MaxRetry=$($requestContext.MaxRetry); RetriesTimeLimit=$($requestContext.RetriesTimeLimit). Outer retries are additional; settings are unchanged." 'INFO' }
    } else { & $LogAction 'SDK request settings unavailable: Get-MgRequestContext is not available.' 'WARN' }
}
$scanTimer = [Diagnostics.Stopwatch]::StartNew()
$phase = 'inventory download'
try {
    $devices = @(Get-IntuneDuplicateInventory -TenantId $tenant -LogAction $LogAction)
    $now = [DateTimeOffset]::UtcNow
    if ($LogAction) { & $LogAction "Inventory download finished in $([math]::Round($scanTimer.Elapsed.TotalSeconds, 1)) s." 'INFO' }
    if ($LogAction) { & $LogAction "Download complete. Analysing $($devices.Count) Windows records for duplicates..." 'INFO' }
    $phase = 'duplicate analysis'
    $analysisTimer = [Diagnostics.Stopwatch]::StartNew()
    $rows = @(Get-IntuneDuplicateDeviceAnalysis -Devices $devices -TenantId $tenant -NowUtc $now -LogAction $LogAction)
    if ($LogAction) { & $LogAction "Analysis finished in $([math]::Round($analysisTimer.Elapsed.TotalSeconds, 1)) s." 'INFO' }
    $phase = 'report summary'
    $summaryTimer = [Diagnostics.Stopwatch]::StartNew()
    $deviceIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $rows) {
        [void]$deviceIds.Add($row.Id)
        [void]$groupIds.Add($row.GroupId)
    }
    if ($LogAction) { & $LogAction "Report summary finished in $([math]::Round($summaryTimer.Elapsed.TotalSeconds, 1)) s; total scan $([math]::Round($scanTimer.Elapsed.TotalSeconds, 1)) s." 'INFO' }
} catch {
    if ($LogAction) { & $LogAction "Scan failed during $phase after $([math]::Round($scanTimer.Elapsed.TotalSeconds, 1)) s: $($_.Exception.Message)" 'ERR' }
    throw
}
$report = [PSCustomObject]@{
    TenantId = $tenant; ScanTimeUtc = $now.ToString('o'); MinimumDays = 60
    InventoryCount = $devices.Count; Rows = $rows
}
if ($LogAction) { & $LogAction "Scan complete: $($devices.Count) Windows records, $($deviceIds.Count) matching devices in $($groupIds.Count) evidence groups." 'OK' }
if ($PassThru) {
    if ($AsReport) { $report } else { $rows }
} else { $rows | Format-Table GroupId, MatchType, DeviceName, LastSyncDateTime, Recommendation, ProtectionReason }
