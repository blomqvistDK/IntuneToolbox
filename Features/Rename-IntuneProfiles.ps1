[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory, ValueFromPipeline)][object[]]$Profiles,
    [string]$Match = '',
    [string]$Replace = '',
    [string]$Prefix = '',
    [string]$Suffix = '',
    [switch]$RemoveWhitespace,
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
        $baseName = if ($Match) { ([string]$item.CurrentName).Replace($Match, $Replace) } else { [string]$item.CurrentName }
        $newName = $Prefix + $baseName + $Suffix
        if ($RemoveWhitespace) { $newName = $newName -replace '\s', '' }
        $status = 'Skipped'
        $errorMessage = ''

        if ($PSCmdlet.ShouldProcess("$($item.CurrentName) ($($item.Id))", "Rename to '$newName'")) {
            try {
                Invoke-IntuneGraphWithRetry -ScriptBlock {
                    if ($item.ApiType -eq 'configurationPolicies') {
                        Invoke-MgGraphRequest -Method PATCH `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($item.Id)')" `
                            -Body (@{ name = $newName } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop
                    } elseif ($item.ApiType -eq 'groupPolicyConfigurations') {
                        Invoke-MgGraphRequest -Method PATCH `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations/$($item.Id)" `
                            -Body (@{ displayName = $newName } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop
                    } elseif ($item.ApiType -eq 'deviceConfigurations') {
                        $deviceConfigUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($item.Id)"
                        try {
                            Invoke-MgGraphRequest -Method PATCH -Uri $deviceConfigUri `
                                -Body (@{ displayName = $newName } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop
                        } catch {
                            $statusCode = $null
                            try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch { }
                            if (($statusCode -ne 400) -and ($_.ToString() -notmatch '400|Bad Request')) { throw }
                            $full = Invoke-MgGraphRequest -Method GET -Uri $deviceConfigUri -ErrorAction Stop
                            if (-not $full) { throw 'The full device configuration could not be read for the fallback rename.' }
                            $payload = $full | ConvertTo-Json -Depth 100 | ConvertFrom-Json
                            $payload.displayName = $newName
                            foreach ($property in @('createdDateTime', 'lastModifiedDateTime', 'version', 'supportsScopeTags', 'settingCount')) {
                                if ($payload.PSObject.Properties[$property]) { [void]$payload.PSObject.Properties.Remove($property) }
                            }
                            foreach ($property in @($payload.PSObject.Properties.Name | Where-Object { $_ -like '*@odata.context' })) {
                                [void]$payload.PSObject.Properties.Remove($property)
                            }
                            Invoke-MgGraphRequest -Method PATCH -Uri $deviceConfigUri `
                                -Body ($payload | ConvertTo-Json -Depth 100) -ContentType 'application/json' -ErrorAction Stop
                        }
                    } else { throw "Unsupported profile API type: '$($item.ApiType)'." }
                } -OnRetry {
                    param($attempt, $delaySeconds, $message)
                    if ($LogAction) { & $LogAction "[$($item.ProfileType)] Throttled; retry $attempt after ${delaySeconds}s." 'WARN' }
                }.GetNewClosure() | Out-Null
                $status = 'OK'
            } catch {
                $status = 'Failed'
                $errorMessage = $_.Exception.Message
            }
        } elseif ($WhatIfPreference) { $status = 'WhatIf' }

        $result = [PSCustomObject]@{ Id = [string]$item.Id; ApiType = [string]$item.ApiType; CurrentName = [string]$item.CurrentName; NewName = $newName; Status = $status; Error = $errorMessage }
        $results.Add($result)
        $level = if ($status -eq 'OK') { 'OK' } elseif ($status -eq 'Failed') { 'ERR' } elseif ($status -eq 'WhatIf') { 'WARN' } else { 'INFO' }
        if ($LogAction) { & $LogAction "[$($item.ProfileType)] '$($item.CurrentName)' -> '$newName' ($status)" $level }
    }
}
end {
    if ($PassThru) { $results.ToArray() }
    else { $results | Format-Table CurrentName, NewName, Status, Error }
}
