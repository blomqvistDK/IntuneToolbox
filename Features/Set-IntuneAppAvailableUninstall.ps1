<#
.SYNOPSIS
Sets Allow available uninstall on selected Intune Win32 apps.
.DESCRIPTION
Changes only the app-level property, not assignments or uninstall commands.
Non-Win32 apps are skipped. Successful updates are verified by reading the app back.
Bulk operations are not atomic; failures are reported per app. Supports -WhatIf.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][object[]]$Apps,
    [Parameter(Mandatory)][bool]$AllowAvailableUninstall,
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Shared\IntuneAppHelpers.ps1')
$scope = 'DeviceManagementApps.ReadWrite.All'
$tenant = Assert-IntuneAppContext -Scope $scope -Tenant $TenantId -Connect:(-not $SkipGraphConnect)
if (-not $Apps.Count) { throw 'Select at least one app.' }
foreach ($app in $Apps) {
    $id = [guid]::Empty
    if (-not [guid]::TryParse([string]$app.Id, [ref]$id) -or $app.TenantId -ne $tenant) {
        throw 'Invalid or different-tenant app. Refresh inventory.'
    }
}
$results = [System.Collections.Generic.List[object]]::new()
$requested = Format-IntuneAppAvailableUninstall -AppType 'win32LobApp' -Value $AllowAvailableUninstall
foreach ($app in @($Apps | Group-Object Id | ForEach-Object { $_.Group[0] })) {
    $status = 'Skipped'; $message = 'Not a Win32 app.'
    $observed = $null
    if ($app.AppType -eq 'win32LobApp') {
        if ($PSCmdlet.ShouldProcess("$($app.DisplayName) [$($app.Id)]", "Set Allow available uninstall to $requested")) {
            $writeAttempted = $false
            try {
                $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.Id)"
                $fresh = Invoke-IntuneGraphWithRetry -ScriptBlock {
                    [void](Assert-IntuneAppContext -Scope $scope -Tenant $tenant)
                    Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                }
                if ($fresh.id -ne $app.Id -or ([string]$fresh.'@odata.type' -replace '^#', '') -cne 'microsoft.graph.win32LobApp') {
                    throw 'App identity or Win32 type could not be verified. Refresh inventory.'
                }
                if ($fresh.allowAvailableUninstall -is [bool] -and $fresh.allowAvailableUninstall -eq $AllowAvailableUninstall) {
                    $observed = $fresh.allowAvailableUninstall
                    $status = 'NoChange'; $message = 'The stored setting already matches.'
                } else {
                    $body = @{ '@odata.type' = '#microsoft.graph.win32LobApp'; allowAvailableUninstall = $AllowAvailableUninstall } | ConvertTo-Json
                    $writeAttempted = $true
                    Invoke-IntuneGraphWithRetry -ScriptBlock {
                        [void](Assert-IntuneAppContext -Scope $scope -Tenant $tenant)
                        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                    } -OnRetry { param($attempt, $delay, $reason) Write-IntuneAppLog "Update retry $attempt in ${delay}s: $reason" 'WARN' }
                    $check = Invoke-IntuneGraphWithRetry -ScriptBlock {
                        [void](Assert-IntuneAppContext -Scope $scope -Tenant $tenant)
                        Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                    }
                    if ($check.id -ne $app.Id -or ([string]$check.'@odata.type' -replace '^#', '') -cne 'microsoft.graph.win32LobApp' -or
                        $check.allowAvailableUninstall -isnot [bool]) { throw 'Read-back did not return a verified Win32 Boolean value.' }
                    $observed = $check.allowAvailableUninstall
                    if ($observed -ne $AllowAvailableUninstall) { throw 'Read-back value does not match the requested setting.' }
                    $status = 'OK'; $message = 'Setting updated and verified.'
                }
            } catch {
                $status = 'Failed'; $message = $_.Exception.Message
                if ($writeAttempted) { $message += ' A write was attempted; refresh inventory before retrying.' }
            }
        } elseif ($WhatIfPreference) {
            $status = 'WhatIf'; $message = "Would set Allow available uninstall to $requested."
        } else { $message = 'Change cancelled.' }
    }
    $results.Add([PSCustomObject]@{
        AppId = [string]$app.Id; AppName = [string]$app.DisplayName
        RequestedValue = $AllowAvailableUninstall; AllowAvailableUninstall = $observed
        AllowAvailableUninstallDisplay = Format-IntuneAppAvailableUninstall -AppType $app.AppType -Value $observed
        Status = $status; Message = $message
    })
    $level = switch ($status) { 'Failed' { 'ERR' }; 'OK' { 'OK' }; 'NoChange' { 'INFO' }; default { 'WARN' } }
    Write-IntuneAppLog "[$($app.DisplayName)] Allow available uninstall = $requested - ${status}: $message" $level
}
if ($PassThru) { $results.ToArray() } else { $results | Format-Table AppName, RequestedValue, AllowAvailableUninstallDisplay, Status, Message }
