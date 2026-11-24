<#
.SYNOPSIS
Runs a read-only duplicate-device scan against a live tenant and records where the time is spent.
.DESCRIPTION
Calls Features\Get-IntuneDuplicateDevices.ps1 with diagnostic logging. Only read access
(DeviceManagementManagedDevices.Read.All) is requested and the removal feature is never called.

Two files are written to -OutputDirectory:
  DuplicateScanDiagnostic_<timestamp>.log          Timestamped activity, with the gap since the previous line.
  DuplicateScanDiagnostic_<timestamp>_summary.txt  Status, counts, phase timings, slowest page, longest wait.

GUIDs, e-mail addresses/UPNs and URLs (including continuation tokens) are masked by default so the
files can be shared. Device records/report rows are never written. The summary is also written when the
scan fails or is stopped with Ctrl+C, so a hanging scan can be interrupted and still diagnosed.
.EXAMPLE
.\Tests\Invoke-IntuneDuplicateScanDiagnostic.ps1 -TenantId '00000000-0000-0000-0000-000000000000'
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputDirectory = (Get-Location).Path,
    # Keep tenant/device IDs, UPNs and URLs in the files. Do not share such files.
    [switch]$IncludeIdentifiers,
    # Reuse the current Graph session instead of signing in.
    [switch]$SkipGraphConnect,
    [switch]$PassThru
)
$ErrorActionPreference = 'Stop'
$featurePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Features\Get-IntuneDuplicateDevices.ps1'
if (-not (Test-Path -LiteralPath $featurePath)) { throw "Feature script not found: $featurePath" }
if (-not (Test-Path -LiteralPath $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $OutputDirectory "DuplicateScanDiagnostic_$stamp.log"
$summaryPath = Join-Path $OutputDirectory "DuplicateScanDiagnostic_${stamp}_summary.txt"

$mask = {
    param([string]$Text)
    if ($IncludeIdentifiers) { return $Text }
    $Text = [regex]::Replace($Text, 'https?://\S+', '<url>')
    $Text = [regex]::Replace($Text, '[^\s@''"<>()]+@[^\s@''"<>()]+\.[A-Za-z]{2,}', '<upn>')
    [regex]::Replace($Text, '[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}', '<guid>')
}.GetNewClosure()

$state = @{
    Timer = [Diagnostics.Stopwatch]::StartNew(); LastSeconds = 0.0
    LongestGap = 0.0; LongestGapAfter = ''; LastMessage = '(start)'
    Pages = 0; Retries = 0; SlowestPage = ''; SlowestPageSeconds = -1.0; Phases = [System.Collections.Generic.List[string]]::new()
}
$logAction = {
    param($Message, $Level)
    $elapsed = $state.Timer.Elapsed.TotalSeconds
    $gap = $elapsed - $state.LastSeconds
    if ($gap -gt $state.LongestGap) { $state.LongestGap = $gap; $state.LongestGapAfter = $state.LastMessage }
    $state.LastSeconds = $elapsed
    $text = & $mask ([string]$Message)
    $state.LastMessage = $text
    if ($text -match '^Inventory page (\d+) finished in ([\d.]+) s') {
        $state.Pages++
        $seconds = [double]::Parse($Matches[2], [Globalization.CultureInfo]::InvariantCulture)
        if ($seconds -gt $state.SlowestPageSeconds) { $state.SlowestPageSeconds = $seconds; $state.SlowestPage = $Matches[1] }
    } elseif ($text -like 'Inventory retry*') {
        $state.Retries++
        $state.Phases.Add($text)
    }
    elseif ($text -match 'finished in|^Scan (failed|complete)|SDK request settings|Scan runtime|Graph rejected|failed after') { $state.Phases.Add($text) }
    $line = [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:yyyy-MM-ddTHH:mm:ss.fffzzz} +{1,8:F1}s [{2,-4}] {3}',
        [DateTimeOffset]::Now, $gap, $Level, $text)
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    $color = switch ($Level) { 'ERR' { 'Red' } 'WARN' { 'Yellow' } 'OK' { 'Green' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
}.GetNewClosure()

& $logAction "Diagnostic started: PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)) on $([Environment]::OSVersion.VersionString); read-only scan." 'INFO'
Write-Host "Log: $logPath" -ForegroundColor Cyan
Write-Host 'If no new line appears for several minutes, the last line shows what it is waiting for. Ctrl+C still writes the summary.' -ForegroundColor Cyan

$status = 'Interrupted (Ctrl+C or host stopped)'
$errorText = ''
$scan = $null
try {
    $parameters = @{ AsReport = $true; PassThru = $true; LogAction = $logAction; SkipGraphConnect = $SkipGraphConnect }
    if ($TenantId) { $parameters.TenantId = $TenantId }
    $scan = & $featurePath @parameters
    $status = 'Completed'
} catch {
    $status = 'Failed'
    $errorText = & $mask $_.Exception.Message
    & $logAction "Diagnostic caught failure: $errorText" 'ERR'
} finally {
    $rows = if ($scan) { @($scan.Rows) } else { @() }
    $devices = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $byType = @{}
    foreach ($row in $rows) {
        [void]$devices.Add($row.Id); [void]$groups.Add($row.GroupId)
        $byType[$row.MatchType] = 1 + [int]$byType[$row.MatchType]
    }
    $summary = [PSCustomObject][ordered]@{
        Status = $status
        TotalSeconds = [math]::Round($state.Timer.Elapsed.TotalSeconds, 1)
        InventoryCount = if ($scan) { $scan.InventoryCount } else { $null }
        ReportRows = $rows.Count
        UniqueDevices = $devices.Count
        EvidenceGroups = $groups.Count
        RowsByMatchType = (@($byType.Keys | Sort-Object | ForEach-Object { "$_=$($byType[$_])" }) -join ', ')
        DeletionCandidateRows = @($rows | Where-Object CanDelete).Count
        PagesCompleted = $state.Pages
        OuterRetries = $state.Retries
        SlowestPage = if ($state.Pages) { "page $($state.SlowestPage): $($state.SlowestPageSeconds) s" } else { '' }
        LongestWaitSeconds = [math]::Round($state.LongestGap, 1)
        LongestWaitAfter = $state.LongestGapAfter
        LastActivity = $state.LastMessage
        Error = $errorText
    }
    $content = @('Intune duplicate-device scan diagnostic (read-only)', '')
    $content += ($summary | Format-List | Out-String -Width 400).Trim()
    $content += @('', 'Phase timings and settings:')
    $content += @($state.Phases | ForEach-Object { "  $_" })
    Set-Content -LiteralPath $summaryPath -Value $content -Encoding UTF8
    Write-Host ''
    Write-Host ($content -join [Environment]::NewLine) -ForegroundColor Cyan
    Write-Host ''
    Write-Host "Summary: $summaryPath" -ForegroundColor Cyan
    if (-not $IncludeIdentifiers) { Write-Host 'IDs, UPNs and URLs are masked. Review both files before sharing.' -ForegroundColor Cyan }
}
if ($PassThru) { $summary }
if ($status -ne 'Completed') { exit 1 }
