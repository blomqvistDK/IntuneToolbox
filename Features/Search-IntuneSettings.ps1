[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SearchTerm,
    [Parameter(Mandatory)][object[]]$Profiles,
    [string]$TenantId,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [scriptblock]$LogAction
)

$requiredScope = 'DeviceManagementConfiguration.Read.All'
if (-not $SkipGraphConnect) {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $context = Get-MgContext
    if (-not $context -or $requiredScope -notin $context.Scopes) {
        $connectParameters = @{ Scopes = @($requiredScope); NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
        if ($TenantId) { $connectParameters.TenantId = $TenantId }
        Connect-MgGraph @connectParameters | Out-Null
    }
}

function Write-FeatureLog {
    param([string]$Message, [string]$Level = 'INFO')
    if ($LogAction) { & $LogAction $Message $Level }
    else { Write-Information $Message -InformationAction Continue }
}

function Get-IntuneSettingPages {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    do {
        $response = Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop
        foreach ($item in @($response.value)) { $items.Add($item) }
        $Uri = [string]$response.'@odata.nextLink'
    } while ($Uri)
    return $items.ToArray()
}

function Get-IntuneSettingProperty {
    param([object]$Object, [string]$Name)
    if (-not $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) { if ([string]$key -ieq $Name) { return $Object[$key] } }
        return $null
    }
    foreach ($property in $Object.PSObject.Properties) { if ($property.Name -ieq $Name) { return $property.Value } }
    return $null
}

function Get-IntuneSettingLeaves {
    param([object]$Instance, [int]$Depth = 0)
    if (-not $Instance) { return }
    $groupCollection = Get-IntuneSettingProperty $Instance 'groupSettingCollectionValue'
    if ($groupCollection -and $groupCollection -is [System.Collections.IEnumerable] -and -not ($groupCollection -is [string])) {
        foreach ($group in $groupCollection) {
            $children = Get-IntuneSettingProperty $group 'children'
            foreach ($child in $children) { Get-IntuneSettingLeaves -Instance $child -Depth $Depth }
        }
        return
    }
    $groupValue = Get-IntuneSettingProperty $Instance 'groupSettingValue'
    if ($groupValue) {
        $children = Get-IntuneSettingProperty $groupValue 'children'
        foreach ($child in $children) { Get-IntuneSettingLeaves -Instance $child -Depth $Depth }
        return
    }
    $choiceValue = Get-IntuneSettingProperty $Instance 'choiceSettingValue'
    if ($choiceValue) {
        [PSCustomObject]@{ Instance = $Instance; Depth = $Depth }
        $children = Get-IntuneSettingProperty $choiceValue 'children'
        if ($children -and $children -is [System.Collections.IEnumerable] -and -not ($children -is [string])) {
            foreach ($child in $children) { Get-IntuneSettingLeaves -Instance $child -Depth ($Depth + 1) }
        }
        return
    }
    $simpleValue = Get-IntuneSettingProperty $Instance 'simpleSettingValue'
    if ($simpleValue) {
        [PSCustomObject]@{ Instance = $Instance; Depth = $Depth }
        $children = Get-IntuneSettingProperty $simpleValue 'children'
        if ($children -and $children -is [System.Collections.IEnumerable] -and -not ($children -is [string])) {
            foreach ($child in $children) { Get-IntuneSettingLeaves -Instance $child -Depth ($Depth + 1) }
        }
        return
    }
    [PSCustomObject]@{ Instance = $Instance; Depth = $Depth }
}

$results = [System.Collections.Generic.List[object]]::new()
$searchLower = $SearchTerm.ToLowerInvariant()
Write-FeatureLog "Searching $($Profiles.Count) Settings Catalog profile(s) for '$SearchTerm'..."
$current = 0
foreach ($profileItem in $Profiles) {
    $current++
    Write-FeatureLog "[$current/$($Profiles.Count)] $($profileItem.CurrentName)"
    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($profileItem.Id)')/settings?`$expand=settingDefinitions"
        $settings = @(Get-IntuneSettingPages -Uri $uri)
        $displayNameMap = @{}
        $optionDisplayMap = @{}
        foreach ($setting in $settings) {
            $definitions = Get-IntuneSettingProperty $setting 'settingDefinitions'
            foreach ($definition in $definitions) {
                $definitionId = Get-IntuneSettingProperty $definition 'id'
                $displayName = Get-IntuneSettingProperty $definition 'displayName'
                if ($definitionId -and $displayName) { $displayNameMap[[string]$definitionId] = [string]$displayName }
                $options = Get-IntuneSettingProperty $definition 'options'
                foreach ($option in $options) {
                    $optionId = Get-IntuneSettingProperty $option 'itemId'
                    $optionName = Get-IntuneSettingProperty $option 'displayName'
                    if ($optionId -and $optionName) { $optionDisplayMap[[string]$optionId] = [string]$optionName }
                }
            }
        }

        foreach ($setting in $settings) {
            $instance = Get-IntuneSettingProperty $setting 'settingInstance'
            if (-not $instance) { continue }
            $leaves = @(Get-IntuneSettingLeaves -Instance $instance)
            if (-not $leaves.Count) { $leaves = @([PSCustomObject]@{ Instance = $instance; Depth = 0 }) }
            foreach ($leafItem in $leaves) {
                if (-not $leafItem) { continue }
                $leaf = $leafItem.Instance
                $definitionId = Get-IntuneSettingProperty $leaf 'settingDefinitionId'
                if (-not $definitionId) { continue }
                $definitionId = [string]$definitionId
                if ($displayNameMap.ContainsKey($definitionId)) { $settingName = $displayNameMap[$definitionId] }
                elseif ($definitionId -match '_([^_]+)$') { $settingName = $Matches[1] }
                else { $settingName = $definitionId }
                if (-not $settingName.ToLowerInvariant().Contains($searchLower) -and -not $definitionId.ToLowerInvariant().Contains($searchLower)) { continue }

                $value = '-'
                $choice = Get-IntuneSettingProperty $leaf 'choiceSettingValue'
                $simple = Get-IntuneSettingProperty $leaf 'simpleSettingValue'
                $simpleCollection = Get-IntuneSettingProperty $leaf 'simpleSettingCollectionValue'
                if ($choice) {
                    $choiceValue = Get-IntuneSettingProperty $choice 'value'
                    if ($choiceValue) {
                        $choiceText = [string]$choiceValue
                        $value = if ($optionDisplayMap.ContainsKey($choiceText)) { $optionDisplayMap[$choiceText] } else { $choiceText }
                    }
                } elseif ($simple) {
                    $simpleValue = Get-IntuneSettingProperty $simple 'value'
                    if ($null -ne $simpleValue) { $value = [string]$simpleValue }
                } elseif ($simpleCollection -and $simpleCollection -is [System.Collections.IEnumerable] -and -not ($simpleCollection -is [string])) {
                    $values = @($simpleCollection | ForEach-Object {
                        $entryValue = Get-IntuneSettingProperty $_ 'value'
                        if ($null -ne $entryValue) { [string]$entryValue }
                    })
                    $value = $values -join '; '
                }

                $results.Add([PSCustomObject]@{ ProfileName = [string]$profileItem.CurrentName; SettingName = [string]$settingName; Value = $value; DefinitionId = $definitionId })
            }
        }
    } catch { Write-FeatureLog "Failed to search '$($profileItem.CurrentName)': $($_.Exception.Message)" 'ERR' }
}
$level = if ($results.Count) { 'OK' } else { 'WARN' }
Write-FeatureLog "Search complete. Found $($results.Count) matching setting(s) across $($Profiles.Count) profile(s)." $level
if ($PassThru) { $results.ToArray() }
else { $results | Format-Table ProfileName, SettingName, Value, DefinitionId }
