[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory, ValueFromPipeline)][object[]]$Profiles,
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)

begin {
    $requiredScopes = @('DeviceManagementConfiguration.ReadWrite.All')
    if (-not $SkipGraphConnect -and -not $WhatIfPreference) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $context = Get-MgContext
        if (-not $context -or @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count) {
            $connectParameters = @{ Scopes = $requiredScopes; NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
            if ($TenantId) { $connectParameters.TenantId = $TenantId }
            Connect-MgGraph @connectParameters | Out-Null
        }
    }
    . (Join-Path $PSScriptRoot '..\Shared\Invoke-IntuneGraphWithRetry.ps1')
    $results = [System.Collections.Generic.List[object]]::new()
}
process {
    foreach ($item in $Profiles) {
        $status = 'Skipped'
        $errorMessage = ''
        if ($PSCmdlet.ShouldProcess("$($item.CurrentName) ($($item.Id))", 'Delete Intune profile')) {
            try {
                Invoke-IntuneGraphWithRetry -ScriptBlock {
                    if ($item.ApiType -eq 'configurationPolicies') {
                        Invoke-MgGraphRequest -Method DELETE `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($item.Id)')" -ErrorAction Stop
                    } elseif ($item.ApiType -in @('deviceConfigurations', 'groupPolicyConfigurations')) {
                        Invoke-MgGraphRequest -Method DELETE `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/$($item.ApiType)/$($item.Id)" -ErrorAction Stop
                    } else { throw "Unsupported profile API type: '$($item.ApiType)'." }
                }.GetNewClosure() -OnRetry {
                    param($attempt, $delaySeconds, $message)
                    if ($LogAction) { & $LogAction "[$($item.ProfileType)] Throttled; retry $attempt after ${delaySeconds}s." 'WARN' }
                }.GetNewClosure() | Out-Null
                $status = 'OK'
            } catch {
                $status = 'Failed'
                $errorMessage = $_.Exception.Message
            }
        } elseif ($WhatIfPreference) { $status = 'WhatIf' }

        $result = [PSCustomObject]@{ Id = [string]$item.Id; ApiType = [string]$item.ApiType; CurrentName = [string]$item.CurrentName; Status = $status; Error = $errorMessage }
        $results.Add($result)
        $level = if ($status -eq 'OK') { 'OK' } elseif ($status -eq 'Failed') { 'ERR' } elseif ($status -eq 'WhatIf') { 'WARN' } else { 'INFO' }
        if ($LogAction) { & $LogAction "[$($item.ProfileType)] '$($item.CurrentName)' ($status)" $level }
    }
}
end {
    if ($PassThru) { $results.ToArray() }
    else { $results | Format-Table CurrentName, Status, Error }
}
