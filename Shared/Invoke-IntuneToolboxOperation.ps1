[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [hashtable]$Parameters = @{},
    [string[]]$Modules = @('Microsoft.Graph.Authentication'),
    [string[]]$Scopes = @(),
    [string]$ExpectedTenant,
    [string]$RequestedTenant,
    [scriptblock]$LogAction,
    [hashtable]$State,
    # Set by the GUI: never open interactive sign-in from a background thread.
    [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'
$ConfirmPreference = 'None'
if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Feature script not found: $Path" }
foreach ($module in $Modules) { Import-Module $module -ErrorAction Stop }

if ($State) {
    if (-not $ExpectedTenant) { $ExpectedTenant = [string]$State.ExpectedTenant }
    if (-not $RequestedTenant) { $RequestedTenant = [string]$State.RequestedTenant }
}

$context = Get-MgContext
if ($ExpectedTenant -and $context -and $context.TenantId -ne $ExpectedTenant) {
    throw 'The Graph tenant changed. Sign out and reconnect before continuing.'
}
if ($RequestedTenant -and $context -and $context.TenantId -ne $RequestedTenant) {
    throw 'The current Graph session belongs to a different tenant. Sign out first.'
}

$missingScopes = @($Scopes | Where-Object {
    -not $context -or ($_ -notin $context.Scopes -and
        -not ($_ -like '*.Read.All' -and ($_ -replace '\.Read\.All$', '.ReadWrite.All') -in $context.Scopes))
})
if (-not $context -or $missingScopes.Count) {
    if ($NonInteractive) {
        $missingText = if ($missingScopes.Count) { $missingScopes -join ', ' } else { $Scopes -join ', ' }
        throw "Missing Microsoft Graph permission(s): $missingText. Run the action again and complete the Microsoft sign-in window."
    }
    $connectParameters = @{ Scopes = $Scopes; NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
    if ($RequestedTenant) { $connectParameters.TenantId = $RequestedTenant }
    elseif ($context -and $context.TenantId) { $connectParameters.TenantId = $context.TenantId }
    Connect-MgGraph @connectParameters | Out-Null
    $context = Get-MgContext
}

$missingScopes = @($Scopes | Where-Object {
    -not $context -or ($_ -notin $context.Scopes -and
        -not ($_ -like '*.Read.All' -and ($_ -replace '\.Read\.All$', '.ReadWrite.All') -in $context.Scopes))
})
if (-not $context -or $missingScopes.Count) { throw 'Sign-in did not grant all required Microsoft Graph scopes.' }
if (($RequestedTenant -and $context.TenantId -ne $RequestedTenant) -or ($ExpectedTenant -and $context.TenantId -ne $ExpectedTenant)) {
    throw 'Tenant verification failed; the operation was not started.'
}

if ($State) {
    $State.ExpectedTenant = [string]$context.TenantId
    $State.Tenant = [string]$context.TenantId
    $State.Account = [string]$context.Account
    $State.Connected = $true
}

if ($Parameters.ContainsKey('SkipGraphConnect')) { $Parameters.SkipGraphConnect = $true }
if ($LogAction -and $Parameters.ContainsKey('LogAction')) { $Parameters.LogAction = $LogAction }
& $Path @Parameters
