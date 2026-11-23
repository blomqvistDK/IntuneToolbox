<#
.SYNOPSIS
Searches Intune audit events by initiating user and/or historical target name.
.DESCRIPTION
Uses the beta auditEvents endpoint across all Intune categories. User (UPN or ID)
and Target use case-insensitive literal partial matching; both must match when supplied.
StartDate and EndDate are inclusive local calendar dates. No partial results are returned.
.EXAMPLE
.\Search-IntuneAuditEvents.ps1 -User 'admin@contoso.com' -Target 'Firewall' -PassThru
.EXAMPLE
.\Search-IntuneAuditEvents.ps1 -Target 'Pilot' -StartDate '2026-11-01' -EndDate '2026-11-13' -PassThru
#>
[CmdletBinding()]
param(
    [string]$User,
    [string]$Target,
    [datetime]$StartDate = [datetime]::Today.AddDays(-6),
    [datetime]$EndDate = [datetime]::Today,
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [switch]$AsReport,
    [scriptblock]$LogAction
)
$ErrorActionPreference = 'Stop'

function Get-IntuneAuditDateWindow {
    param([datetime]$StartDate, [datetime]$EndDate, [TimeZoneInfo]$TimeZone = [TimeZoneInfo]::Local)
    if ($StartDate -eq [datetime]::MinValue -or $EndDate -eq [datetime]::MinValue -or
        $EndDate.Date -eq [datetime]::MaxValue.Date) {
        throw 'Enter valid start and end dates.'
    }
    if ($StartDate.Date -gt $EndDate.Date) { throw 'Start date must be on or before end date.' }
    # Convert the two local midnights separately: a daylight-saving day need not be 24 hours.
    $start = [datetime]::SpecifyKind($StartDate.Date, [DateTimeKind]::Unspecified)
    $end = [datetime]::SpecifyKind($EndDate.Date.AddDays(1), [DateTimeKind]::Unspecified)
    [PSCustomObject]@{
        StartUtc = [TimeZoneInfo]::ConvertTimeToUtc($start, $TimeZone)
        EndUtc = [TimeZoneInfo]::ConvertTimeToUtc($end, $TimeZone)
    }
}

function Get-IntuneAuditOperationKind {
    param([string]$OperationType, [string]$ActivityType, [string]$Activity)
    $kinds = @{
        create = 'Create'; post = 'Create'
        delete = 'Delete'
        patch = 'Update'; update = 'Update'; put = 'Update'
    }
    $operation = if ($OperationType) { $OperationType.Trim() } else { '' }
    if ($operation) {
        if ($kinds.ContainsKey($operation.ToLowerInvariant())) { return $kinds[$operation.ToLowerInvariant()] }
        return 'Other'
    }
    # Without an operation type, infer only from the leading verb of the activity, e.g. "Delete DeviceConfiguration".
    foreach ($name in @($ActivityType, $Activity)) {
        if (-not $name -or -not $name.Trim()) { continue }
        $verb = ($name.Trim() -split '\s+')[0].ToLowerInvariant()
        if ($kinds.ContainsKey($verb)) { return $kinds[$verb] }
        return 'Other'
    }
    'Unknown'
}

function Write-AuditFeatureLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ($LogAction) { & $LogAction $Message $Level }
    else { Write-Information $Message -InformationAction Continue }
}

function Test-IntuneAuditText {
    param([string]$Value, [string]$Term)
    $Value.IndexOf($Term, [StringComparison]::OrdinalIgnoreCase) -ge 0
}

$User = $User.Trim()
$Target = $Target.Trim()
if (-not $User -and -not $Target) { throw 'Enter a User (UPN or ID), a Target name, or both.' }
$window = Get-IntuneAuditDateWindow -StartDate $StartDate -EndDate $EndDate

. (Join-Path $PSScriptRoot '..\Shared\Invoke-IntuneGraphWithRetry.ps1')
if (-not $SkipGraphConnect) { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
$scope = 'DeviceManagementApps.Read.All'
$context = Get-MgContext
if ($TenantId -and $context -and $context.TenantId -ne $TenantId) {
    throw 'The current Graph session belongs to a different tenant. Sign out first.'
}
$expectedTenant = if ($TenantId) { $TenantId } elseif ($context) { [string]$context.TenantId } else { '' }
$hasScope = $context -and ($scope -in $context.Scopes -or 'DeviceManagementApps.ReadWrite.All' -in $context.Scopes)
if (-not $hasScope -and -not $SkipGraphConnect) {
    $connectParameters = @{ Scopes = @($scope); NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
    if ($expectedTenant) { $connectParameters.TenantId = $expectedTenant }
    Connect-MgGraph @connectParameters | Out-Null
    $context = Get-MgContext
}
if (-not $context -or ($scope -notin $context.Scopes -and 'DeviceManagementApps.ReadWrite.All' -notin $context.Scopes)) {
    throw "Missing Microsoft Graph permission: $scope. Sign in with audit-read access before searching."
}
if ($expectedTenant -and $context.TenantId -ne $expectedTenant) { throw 'Tenant verification failed; audit search was not started.' }
$tenant = [string]$context.TenantId
$culture = [Globalization.CultureInfo]::InvariantCulture
$filter = "activityDateTime ge $($window.StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ', $culture)) and activityDateTime lt $($window.EndUtc.ToString('yyyy-MM-ddTHH:mm:ssZ', $culture))"
$uri = 'https://graph.microsoft.com/beta/deviceManagement/auditEvents?$filter=' + [uri]::EscapeDataString($filter)
$rows = [System.Collections.Generic.List[object]]::new()
$visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$page = 0
$scanned = 0
Write-AuditFeatureLog "Searching Intune audit events from $($window.StartUtc.ToString('o')) to $($window.EndUtc.ToString('o')) (exclusive), tenant $tenant."
do {
    if (-not $visited.Add($uri)) { throw 'Graph returned a repeated audit page link; no partial report was published.' }
    $page++
    Write-AuditFeatureLog "Requesting audit page $page..."
    $response = Invoke-IntuneGraphWithRetry -ScriptBlock {
        Invoke-MgGraphRequest -Uri $uri -Method GET -OutputType PSObject -ErrorAction Stop
    } -OnRetry { param($Attempt, $Delay, $Message) Write-AuditFeatureLog "Audit request retry $Attempt in $Delay seconds: $Message" 'WARN' }
    if ($null -eq $response -or $null -eq $response.value) { throw 'Graph returned an invalid audit collection; no partial report was published.' }
    foreach ($event in $response.value) {
        $scanned++
        $dateValue = $event.activityDateTime
        if ($dateValue -is [DateTimeOffset]) { $timestamp = $dateValue }
        elseif ($dateValue -is [datetime]) {
            if ($dateValue.Kind -eq [DateTimeKind]::Unspecified) { $dateValue = [datetime]::SpecifyKind($dateValue, [DateTimeKind]::Utc) }
            $timestamp = [DateTimeOffset]$dateValue
        } else {
            $timestamp = [DateTimeOffset]::Parse([string]$dateValue, $culture, [Globalization.DateTimeStyles]::AssumeUniversal)
        }
        if ($timestamp.UtcDateTime -lt $window.StartUtc -or $timestamp.UtcDateTime -ge $window.EndUtc) { continue }
        $actor = $event.actor
        if ($User -and -not ((Test-IntuneAuditText ([string]$actor.userPrincipalName) $User) -or
            (Test-IntuneAuditText ([string]$actor.userId) $User))) { continue }
        $resources = @($event.resources | Where-Object { $null -ne $_ })
        if ($Target -and -not @($resources | Where-Object { Test-IntuneAuditText ([string]$_.displayName) $Target }).Count) { continue }
        $initiatedBy = if ($actor.userPrincipalName) { [string]$actor.userPrincipalName }
            elseif ($actor.userId) { [string]$actor.userId }
            elseif ($actor.applicationDisplayName) { [string]$actor.applicationDisplayName }
            elseif ($actor.servicePrincipalName) { [string]$actor.servicePrincipalName }
            elseif ($actor.applicationId) { [string]$actor.applicationId }
            else { 'Not supplied' }
        $operation = if ($event.activity) { [string]$event.activity } else { [string]$event.displayName }
        $rows.Add([PSCustomObject]@{
            Id = [string]$event.id
            Timestamp = $timestamp
            TimestampLocal = $timestamp.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss zzz', $culture)
            TimestampUtc = $timestamp.ToUniversalTime().ToString('o')
            Operation = $operation
            OperationKind = Get-IntuneAuditOperationKind -OperationType ([string]$event.activityOperationType) `
                -ActivityType ([string]$event.activityType) -Activity $operation
            ActivityType = [string]$event.activityType
            HttpOperation = [string]$event.activityOperationType
            Result = [string]$event.activityResult
            InitiatedBy = $initiatedBy
            Application = [string]$actor.applicationDisplayName
            UserPrincipalName = [string]$actor.userPrincipalName
            UserId = [string]$actor.userId
            ApplicationId = [string]$actor.applicationId
            TargetNames = (@($resources | ForEach-Object { $_.displayName }) -join '; ')
            Category = [string]$event.category
            Component = [string]$event.componentName
            CorrelationId = [string]$event.correlationId
            TenantId = $tenant
            ActorJson = ConvertTo-Json -InputObject $actor -Depth 100 -Compress
            ResourcesJson = ConvertTo-Json -InputObject $resources -Depth 100 -Compress
            Event = $event
        })
    }
    Write-AuditFeatureLog "Read audit page $page; $scanned events checked, $($rows.Count) matching."
    $uri = [string]$response.'@odata.nextLink'
} while ($uri)
$sorted = @($rows | Sort-Object Timestamp -Descending)
Write-AuditFeatureLog "Search complete: $($sorted.Count) matching event(s) from $page page(s)." 'OK'
if ($PassThru) {
    if ($AsReport) {
        [PSCustomObject]@{ Completed = $true; TenantId = $tenant; Rows = $sorted; ScannedCount = $scanned; PageCount = $page }
    } else { $sorted }
} else { $sorted | Format-Table TimestampLocal, OperationKind, Operation, Result, InitiatedBy, TargetNames, Category }
