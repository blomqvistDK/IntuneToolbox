[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('Groups', 'Filters')][string]$Action,
    [string]$SearchTerm = '',
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)

$requiredScope = if ($Action -eq 'Groups') { 'Group.Read.All' } else { 'DeviceManagementConfiguration.Read.All' }
if (-not $SkipGraphConnect) {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $context = Get-MgContext
    if (-not $context -or $requiredScope -notin $context.Scopes) {
        $connectParameters = @{ Scopes = @($requiredScope); NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
        if ($TenantId) { $connectParameters.TenantId = $TenantId }
        Connect-MgGraph @connectParameters | Out-Null
    }
}

function Get-AssignmentTargetPages {
    param([string]$Uri, [hashtable]$Headers = @{})
    $items = [System.Collections.Generic.List[object]]::new()
    do {
        $response = Invoke-MgGraphRequest -Uri $Uri -Method GET -Headers $Headers -ErrorAction Stop
        foreach ($item in @($response.value)) { $items.Add($item) }
        $Uri = [string]$response.'@odata.nextLink'
    } while ($Uri)
    $items.ToArray()
}

if ($Action -eq 'Groups') {
    if ([string]::IsNullOrWhiteSpace($SearchTerm)) { throw 'SearchTerm must not be empty.' }
    $search = '$search=' + [uri]::EscapeDataString('"displayName:' + $SearchTerm.Replace('\', '\\').Replace('"', '\"') + '"')
    $uri = "https://graph.microsoft.com/beta/groups?$search"
    $results = @(Get-AssignmentTargetPages -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' } | ForEach-Object { [PSCustomObject]@{ DisplayName = [string]$_.displayName; Id = [string]$_.id } })
    if ($LogAction) { & $LogAction "Found $($results.Count) group(s)." 'OK' }
} else {
    $results = @(Get-AssignmentTargetPages -Uri 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters' | ForEach-Object { [PSCustomObject]@{ DisplayName = [string]$_.displayName; Id = [string]$_.id; Platform = [string]$_.platform; Rule = [string]$_.rule; ManagementType = [string]$_.assignmentFilterManagementType } })
    if ($LogAction) { & $LogAction "Found $($results.Count) assignment filter(s)." 'OK' }
}

if ($PassThru) { $results }
else { $results | Format-Table DisplayName, Id }
