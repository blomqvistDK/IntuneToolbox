# Intune Toolbox

The WPF interface and its standalone operation scripts are packaged together. Keep the whole `IntuneToolbox` folder intact when moving or copying it; paths resolve relative to the launcher, not the current working directory.

```text
IntuneToolbox/
|-- Invoke-IntuneToolbox-GUI.ps1
|-- Features/
|   |-- Get-IntuneProfiles.ps1
|   |-- Get-IntuneApps.ps1
|   |-- Invoke-IntuneAppAssignment.ps1
|   |-- Set-IntuneAppAvailableUninstall.ps1
|   |-- Invoke-IntuneProfileAssignment.ps1
|   |-- Rename-IntuneProfiles.ps1
|   |-- Remove-IntuneProfiles.ps1
|   |-- Invoke-AutopilotGroupTagOperation.ps1
|   |-- Search-IntuneAssignmentTargets.ps1
|   |-- Search-IntuneSettings.ps1
|   |-- Search-IntuneAuditEvents.ps1
|   |-- Get-IntuneDuplicateDevices.ps1
|   `-- Remove-IntuneDuplicateDevices.ps1
|-- Shared/
|   |-- Invoke-IntuneToolboxOperation.ps1
|   |-- Invoke-IntuneGraphWithRetry.ps1
|   |-- IntuneAppHelpers.ps1
|   `-- Get-IntuneDuplicateDeviceAnalysis.ps1
|-- Tests/
|   |-- Test-IntuneToolbox.ps1
|   |-- Test-IntuneApps.ps1
|   |-- Test-IntuneDuplicateDevices.ps1
|   `-- Invoke-IntuneDuplicateScanDiagnostic.ps1
`-- README.md
```

## PowerShell Modules

Graph features require `Microsoft.Graph.Authentication`. The toolbox loads it when an operation needs it and does not install modules automatically.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

The signed-in account still needs the delegated permissions and consent required by the selected operation. The scripts under `Shared/` are bundled with the toolbox and do not need to be installed separately.

The former `Intune/Invoke-IntuneToolbox-GUI.ps1` path forwards to the packaged launcher. The dated copies remain separate. The previous active-path file was retained under `Intune/old/` as a migration backup.

## Launch

Run in Windows PowerShell 5.1 or PowerShell 7, in STA mode:

```powershell
powershell.exe -NoProfile -STA -File .\Invoke-IntuneToolbox-GUI.ps1
pwsh.exe -NoProfile -STA -File .\Invoke-IntuneToolbox-GUI.ps1 -TenantId 'tenant-guid'
```

App, profile, and group lists keep the active theme's background colors while loading or disabled, including when switching between light and dark themes. Disabled controls remain non-interactive.

The GUI opens without authenticating. **Sign in** requests every delegated Microsoft Graph scope the toolbox uses in one consent prompt:

- `User.Read`
- `DeviceManagementConfiguration.ReadWrite.All` (profiles, assignments, filters, settings search, rename, delete)
- `DeviceManagementServiceConfig.ReadWrite.All` (Autopilot lookup and group tags)
- `DeviceManagementManagedDevices.ReadWrite.All` (duplicate-device scan and deletion, Autopilot device names)
- `DeviceManagementManagedDevices.PrivilegedOperations.All` (Autopilot Fresh Start)
- `DeviceManagementApps.ReadWrite.All` (app inventory/assignment management; also covers read-only audit lookup)
- `Group.Read.All` (group search for assignments)

A `ReadWrite.All` grant also satisfies the matching `Read.All` requirement. If a scope is still missing when an operation starts (for example, you signed in before this change or consent was partial), the toolbox requests it from the main window before the operation starts, so the Microsoft sign-in or consent window has a visible owner. If sign-in is cancelled or consent is not granted, the operation does not start and Activity shows the error. Background operations never open sign-in themselves; they receive only plain data, and their log lines and results are shown by the UI thread. The first successful operation pins the session to its tenant; signing out clears that pin. `-TenantId` optionally restricts sign-in to a specific tenant.

Install `Microsoft.Graph.Authentication` before using Graph features. The toolbox does not install PowerShell modules automatically. The signed-in user still needs appropriate consent, Intune roles, and licensing for the selected operation. App-only client-secret authentication is not supported.

### Troubleshooting sign-in

- **An operation seems stuck after you start it:** check the taskbar or Alt+Tab for a Microsoft sign-in or permissions window and complete or cancel it.
- **"Need admin approval":** the account cannot consent to the requested scope. An administrator must grant it (for example `DeviceManagementManagedDevices.Read.All`) to the Microsoft Graph Command Line Tools enterprise application.
- **"Missing Microsoft Graph permission(s)" in Activity:** the permission request failed or was cancelled. Run the action again and complete the sign-in window.

Build the WPF interface without authentication or showing the window:

```powershell
pwsh.exe -NoProfile -STA -File .\Invoke-IntuneToolbox-GUI.ps1 -ValidateOnly
```

## Feature Scripts

Every feature can be run independently. Without `-SkipGraphConnect`, the script imports Microsoft.Graph.Authentication and requests its delegated scopes. Use `-SkipGraphConnect` to require an existing session with the necessary scopes. `-PassThru` returns structured objects; without it, the script writes a table or informational output. Use `Get-Help` on each script for the full parameter list.

### Profile Inventory

```powershell
$profiles = & .\Features\Get-IntuneProfiles.ps1 -NameFilter 'Pilot' -ProfileType SettingsCatalog -AssignmentFilter Group -PassThru
```

Supports Settings Catalog, classic Device Configuration, and Administrative Templates profiles. All three types are enabled by default in the GUI and standalone inventory. Select `-ProfileType AdministrativeTemplates` to retrieve only Administrative Templates.

Administrative Templates are loaded from the tenant's `groupPolicyConfigurations` collection, including built-in and custom/imported ADMX profiles. They support assignment, rename, delete, name/platform/assignment filtering, assignment summaries, name previews, and CSV export. They are treated as Windows profiles. This is tenant inventory, not JSON import or restore.

Optional succeeded counts remain available for the existing profile types. Administrative Templates show `N/A` when succeeded counts are requested because their resource does not expose the reporting endpoint used by those types. Settings Search remains Settings Catalog-only. Name previews use `-Match`, `-Replace`, `-Prefix`, and `-Suffix`.

### Assignments

```powershell
& .\Features\Invoke-IntuneProfileAssignment.ps1 -Profiles $profiles -Action Add -TargetType Group -TargetGroup $group -WhatIf -PassThru

# The existing group parameters accept either one group or an array.
$groups = & .\Features\Search-IntuneAssignmentTargets.ps1 -Action Groups -SearchTerm 'Pilot' -PassThru
& .\Features\Invoke-IntuneProfileAssignment.ps1 -Profiles $profiles -Action Add -TargetType Group -TargetGroup $groups -WhatIf -PassThru
& .\Features\Invoke-IntuneProfileAssignment.ps1 -Profiles $profiles -Action Remove -RemoveScope Group -RemoveType both -RemoveGroup $groups -WhatIf -PassThru
```

`-Action` accepts `Add`, `Replace`, or `Remove`. Review the groups returned by a search before applying changes; remove `-WhatIf` only when ready.

In the GUI:

1. Select the profiles to change and choose an assignment action.
2. For group targets, search and select multiple results using Ctrl/Shift, then click **Add selected**.
3. Repeat searches as needed. Queued groups persist across searches and duplicate group IDs are ignored.
4. Use **Remove selected** or **Clear** to edit the queue. Assignment and group-removal queues are separate; form reset clears both.
5. Click **Apply assignment**. It uses **every queued group**, regardless of highlighted rows. Profiles must still be selected. A single include/exclude choice and optional device filter apply to all groups in an Add/Replace run.

Add preserves unrelated direct assignments. Replace replaces **all direct assignments** with the complete new target set in one request per profile, not one request per group. Group-specific Remove affects only the queued groups and selected include/exclude/both assignment types. **Remove ALL assignments** does not require queued groups and clears all direct assignments. Policy Set-owned assignments are excluded from submitted assignment payloads.

The GUI confirms the target groups and affected profile count before writes. Queues remain available after an operation for reuse.

### Rename And Delete

```powershell
& .\Features\Rename-IntuneProfiles.ps1 -Profiles $profiles -Prefix 'Pilot ' -WhatIf -PassThru
& .\Features\Remove-IntuneProfiles.ps1 -Profiles $profiles -WhatIf -PassThru
```

Both scripts support `-WhatIf` and emit per-profile status results. The GUI's rename and delete confirmation dialogs remain enabled.

### App Management

Open **App Management**, choose **All platforms** (default) or Windows, Android, iOS/iPadOS, or macOS, optionally enter a literal name substring, and click **Load apps**. Every Graph page is loaded; there is no app-type allowlist or result cap. Platform is derived from the app subtype. Unrecognized and platform-neutral types (such as generic web links) remain visible under All platforms. Inventory is limited by the administrator's RBAC visibility.

The sortable **Allow available uninstall** column shows **Yes/No** for Win32 apps, **Unknown** if Graph does not return a Boolean, and **N/A** for other app types. To change it, select apps and choose an explicit **Yes** or **No** under **Win32 app settings**, then click **Apply to selected Win32 apps**. No value is preselected. The confirmation lists the tenant, new value, eligible apps and skipped count. Mixed selections skip non-Win32 apps and report them in Activity.

This changes only the app-level `allowAvailableUninstall` property, which controls Company Portal uninstall for Win32 apps with an Available assignment. It does not uninstall apps, create/change assignments, or modify uninstall commands. Each app is freshly checked before updating; matching values are no-ops. Updates are read back before reporting success, and the column is refreshed with the observed value. Failed verification or missing outcomes show **Unknown**; a verified mismatching value is displayed but reported as a failure. See Activity and **Load apps** again before retrying uncertain changes. Successful changes are not rolled back if other apps fail.

Use **Filter loaded apps** above the results to narrow the list immediately without more Graph requests. For only Win32 apps, choose **App type = win32LobApp**. App name and Publisher match case-insensitive, literal substrings; Platform, App type, and Assignments (hint) use exact dropdown values from the loaded inventory. All criteria combine with **AND**. Counts show visible, loaded, and selected apps. **Clear filters** restores all loaded rows, but cannot restore apps excluded by the Name/Platform criteria used for **Load apps**. Changing result filters clears app selection and assignment inspection; sorting is preserved and apps are never selected automatically. A fresh load or sign-out clears result filters and cached results. Result filters are disabled while an app operation runs. Assignment status remains a Graph hint; after an assignment mutation it becomes **Refresh required**, which can remove an app from a status-filtered view. Changing Allow available uninstall leaves this hint unchanged.

Select apps with Ctrl/Shift. No apps are selected automatically. Choose Required, Available, or Uninstall; target Groups, All Users, or All Devices (default). To keep the right-hand panel short, group search and the queue appear only when the target is Groups, the filter picker appears only when a filter mode is chosen, and the existing-assignment list and its edit/remove buttons appear only after **Refresh assignments**. For groups, search, highlight results, and click **Add selected groups**. Repeat searches to build a persistent queue. **Every queued group is used**, regardless of highlighting; **Remove selected** and **Clear queue** edit that queue. Choose Include/Exclude and an optional device assignment filter. Filters cannot be attached to excluded groups; managed-app filters are not deployment filters.

**Add to selected apps** previews app names/IDs, tenant, targets, intent, and filter before writing. Exact duplicate assignments are no-ops; conflicting assignments for the same target require explicit inspection/editing or Replace. Unrelated assignments are never replaced.

**Replace this intent on selected apps** makes each selected app's **direct** assignments with the chosen intent and include/exclude type match exactly the queued targets (or All Users/All Devices):

- A queued target with no assignment is **added**. One with the same intent and filter is left unchanged. One with the same intent but a different filter has its **filter updated** in place, keeping its settings.
- Direct assignments with the same intent and include/exclude type whose target is not queued are **removed**.
- A queued target that already has a different intent is reported as a **conflict** and is not changed. Edit it explicitly.
- Other intents, the other include/exclude type, and Policy Set-owned assignments are never touched.

Replace first runs a read-only preview. The confirmation lists every add, update, removal, unchanged target and conflict per app; long previews are truncated in the dialog but complete in Activity. Nothing is written if you cancel or nothing changes. After confirmation, adds and updates run first. An app's removals run only if all of its adds and updates succeeded and it has no conflicts; otherwise its old assignments are kept and reported as skipped. Only apps with a successful preview are changed, and an assignment that changed after the preview is not updated or removed. Replace is not atomic across apps. To clear all assignments of an intent, use **Remove selected assignments** instead.

To inspect or change existing assignments:

1. Select exactly one app and click **Refresh assignments**.
2. Select an assignment to see its ID, group ID, filter, source, and platform-specific settings. Unavailable/deleted group names are explicitly marked with their IDs.
3. Click **Load selected assignment into form**, change intent/target/filter, then **Update selected assignment**. Editing requires exactly one target group when targeting groups.
4. Or select one or multiple direct assignments and click **Remove selected assignments**. This removes assignment records, not apps, and does not create an Uninstall assignment.

Policy Set-owned assignments are read-only. Existing `availableWithoutEnrollment` intents and unrecognized target types can be inspected/removed if directly owned, but cannot be edited in this basic editor. Edit preserves existing app-specific settings (including nested restart/schedule settings) and unrelated target properties. Setting filter mode to None explicitly clears its filter. Assignments are freshly re-read before changes; stale snapshots are rejected and require refresh. There is no atomic read/write transaction, so concurrent administrator changes cannot be entirely prevented.

New assignments omit app-specific settings and request **Intune service defaults**. This does not guarantee every app subtype accepts a settings-free create: app-specific settings/intent/filter requirements are enforced by Graph and failures are shown explicitly in Activity. Use Intune for subtypes requiring explicit settings or for advanced configuration. Available device targeting is allowed here only for Win32 and managed Google Play apps; other app types require a user target. For group targeting, select appropriate user/device groups for the app and intent; group membership type is not inferred. Other platform/management-mode restrictions still apply.

Writes use assignment-level POST/PATCH/DELETE, not replacement of the full assignment set. An uncertain create response is reconciled against fresh assignments and is **not blindly retried**. If reconciliation cannot confirm success, inspect/refresh before retrying. Bulk work is **not atomic**: successful targets remain changed when others fail. Activity reports per-app/per-target outcomes; assignment details are refreshed after writes for a single selected app. Inventory assignment status is only a Graph hint and is marked for refresh after mutations. Sign-out clears app inventory, inspection, and queues.

Inventory requests `DeviceManagementApps.Read.All`; mutations require `DeviceManagementApps.ReadWrite.All`. Group-name resolution/search uses `Group.Read.All`; loading filters uses `DeviceManagementConfiguration.Read.All`. Existing matching ReadWrite scopes satisfy Read scopes. The account additionally needs appropriate Intune Apps read/assign permissions, applicable licensing, and consent. The beta API is used for current app types and assignment-source metadata and may change. App creation/deletion, packaging, deployment reporting, and advanced assignment-settings editors are not included.

Changing **Allow available uninstall** additionally requires Intune app update permissions (and `DeviceManagementApps.ReadWrite.All`). Its minimal app-level PATCH leaves other properties untouched. Bulk operations are not atomic, and concurrent administrator changes cannot be entirely prevented.

Standalone examples (review selected IDs before removing `-WhatIf`):

```powershell
$inventory = & .\Features\Get-IntuneApps.ps1 -PlatformFilter Windows -NameFilter 'Pilot' -PassThru
$apps = @($inventory.Rows | Where-Object Id -eq '22222222-2222-2222-2222-222222222222')
# Explicitly enable Company Portal available uninstall (preview only)
& .\Features\Set-IntuneAppAvailableUninstall.ps1 -Apps $apps -AllowAvailableUninstall $true -TenantId $apps[0].TenantId -WhatIf -PassThru
$groups = & .\Features\Search-IntuneAssignmentTargets.ps1 -Action Groups -SearchTerm 'Pilot' -PassThru
$group = @($groups | Where-Object Id -eq '33333333-3333-3333-3333-333333333333')
& .\Features\Invoke-IntuneAppAssignment.ps1 -Apps $apps -Action Add -Intent required -TargetGroup $group -WhatIf -PassThru

$inspection = & .\Features\Get-IntuneApps.ps1 -Action Assignments -AppId $apps[0].Id -TenantId $apps[0].TenantId -PassThru
$assignment = @($inspection.Rows | Where-Object Id -eq '44444444-4444-4444-4444-444444444444')
& .\Features\Invoke-IntuneAppAssignment.ps1 -Apps $apps -Action Update -Assignments $assignment -Intent available -TargetGroup $group -WhatIf -PassThru
& .\Features\Invoke-IntuneAppAssignment.ps1 -Apps $apps -Action Remove -Assignments $assignment -WhatIf -PassThru

# Replace: preview, review, then apply only what was previewed
$plan = & .\Features\Invoke-IntuneAppAssignment.ps1 -Apps $apps -Action Replace -Intent required -TargetGroup $group -PlanOnly -PassThru
$plan | Format-Table AppName, Operation, Target, Message
& .\Features\Invoke-IntuneAppAssignment.ps1 -Apps $apps -Action Replace -Intent required -TargetGroup $group -Plan $plan -WhatIf -PassThru
```

`Get-IntuneApps.ps1 -PassThru` returns a complete report with `Rows`, `Action`, `AppId`, and `TenantId`, including empty successes. Failed reads publish no partial report. Mutations return one status per app/target or selected assignment. Run `Tests\Test-IntuneApps.ps1` for focused offline app/WPF checks; the main `Tests\Test-IntuneToolbox.ps1` suite includes these checks. Tests make no live tenant changes.

Inventory rows include nullable Boolean `AllowAvailableUninstall` and `AllowAvailableUninstallDisplay`. `Set-IntuneAppAvailableUninstall.ps1 -PassThru` returns one outcome per unique selected app: `AppId`, `AppName`, Boolean `RequestedValue`, nullable observed `AllowAvailableUninstall`, `AllowAvailableUninstallDisplay`, `Status`, and `Message`. Status is `OK`, `NoChange`, `Skipped`, `WhatIf`, or `Failed`; `OK` requires verified read-back. Use `$false` instead of `$true` to disable the setting.

### Autopilot Group Tags

The **Autopilot Group Tags** page looks up Autopilot devices by serial number (typed, pasted, or from a CSV with a serial-number column and an optional `GroupTag`/`NewGroupTag` column). It shows the current group tag, CSV tag, assigned deployment profile, profile assignment date, and Intune device name. Selected devices can receive the entered tag, their CSV tags, or have their tag cleared, optionally followed by a tenant-wide Autopilot sync. Results can be filtered, refreshed, and exported to CSV. **Fresh Start** reinstalls Windows and removes apps and user data (`keepUserData=false`); wait for the intended deployment profile assignment first, because a tag change does not guarantee profile readiness. Every change is previewed and confirmed before it is sent.

```powershell
& .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action Lookup -SerialNumber 'PF12345','PF67890' -PassThru

$rows = & .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action ImportCsv -Path .\devices.csv -PassThru
$devices = & .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action Lookup -Device $rows -PassThru
& .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action ApplyCsvTags -Device $devices -WhatIf -PassThru
& .\Features\Invoke-AutopilotGroupTagOperation.ps1 -Action ApplyTag -Device $devices -NewTag 'Production' -SyncAfterUpdate -WhatIf -PassThru
```

The same feature ships with the standalone `AutoPilot\AutopilotGroupTagTool`; its test suite covers the feature script in depth.

`-Action` accepts `GetTags`, `GetDevices`, or `UpdateTags`. Updates support `-WhatIf`; the GUI also confirms before applying tags.

### Assignment Targets And Settings Search

```powershell
& .\Features\Search-IntuneAssignmentTargets.ps1 -Action Groups -SearchTerm 'Pilot' -PassThru
& .\Features\Search-IntuneAssignmentTargets.ps1 -Action Filters -PassThru
& .\Features\Search-IntuneSettings.ps1 -SearchTerm 'Firewall' -Profiles $profiles -PassThru
```

### Audit Lookup

Open **Audit Lookup** to search historical Intune changes across **all categories**, including devices, policies, apps, and assignments. This feature is read-only and uses `GET /beta/deviceManagement/auditEvents`, not Entra directory audit logs.

- Enter **User** (initiating user's UPN or user ID), **Target** (historical object name), or both. At least one is required. Matching is case-insensitive, literal, and partial; both criteria must match when supplied. Application names are displayed but do not match the User filter.
- Dates default to today and the preceding six days in your computer's **local time zone**. Start and end dates are inclusive. Each local midnight boundary is converted separately to UTC, respecting daylight saving.
- The search follows every Graph page in the date window and matches user/target text locally. It does not impose a result cap or depend on undocumented nested substring filters. Larger date windows can take longer. Activity shows page/retry progress; failures publish no partial report and are not shown as "no matching events."
- Results are initially newest first and show local timestamps with UTC offsets, change type, operation, result, actor/application, targets, and category. Select an event and expand **Selected event details** for UTC time, event/correlation IDs, actor/IP information, and each target's changed properties with old/new values. Missing details are marked as not supplied.
- Each row is color-coded by change type with a stripe on its left edge and colored **Type**/**Operation** text: **green** = Create, **red** = Delete, **blue** = Update/Patch, **amber** = Other (for example device actions such as wipe or sync, or assignments). A legend above the list shows the colors, and the **Type** column can be sorted. The type comes from the event's operation type (`Create`/`Post`, `Delete`, `Patch`/`Update`/`Put`; any other value is Other). When that is missing, it comes from the first word of the activity name, for example "Delete DeviceConfiguration". Events with neither are **Unknown** and left uncolored. The colors adapt to the light and dark themes.
- **Export CSV** saves all events from the completed search, not just the selection. It includes ISO UTC timestamps, the `OperationKind` change type, identifiers, and JSON columns for actors and all targets/changes. Treat exported audit data as sensitive.

The toolbox requests `DeviceManagementApps.Read.All` for audit lookup (a matching `ReadWrite.All` grant also suffices). The administrator additionally needs Intune **Audit data / Read** access, for example through the Intune Administrator role, and applicable licensing. Results are limited to Graph's available retention and the account's accessible audit data; this does not query archived Azure Monitor logs. The explicitly selected beta API can change.

```powershell
& .\Features\Search-IntuneAuditEvents.ps1 -User 'admin@contoso.com' -Target 'Firewall' -PassThru
& .\Features\Search-IntuneAuditEvents.ps1 -Target 'Pilot' -StartDate '2026-11-01' -EndDate '2026-11-13' -PassThru

# Complete-report envelope, including successful searches with no matches.
$report = & .\Features\Search-IntuneAuditEvents.ps1 -Target 'Pilot' -AsReport -PassThru
```

See Microsoft's [audit API reference](https://learn.microsoft.com/en-us/graph/api/intune-auditing-auditevent-list?view=graph-rest-beta) and [Intune audit access documentation](https://learn.microsoft.com/en-us/intune/intune-service/fundamentals/monitor-audit-logs).

### Duplicate Windows Devices

Open **Duplicate Devices** and choose **Scan Windows devices**. Scanning is read-only and retrieves all pages of the visible managed-device inventory through Microsoft Graph v1.0. It asks Graph for Windows devices only (`operatingSystem eq 'Windows'`) and also checks the platform locally; if Graph rejects the filter, it downloads the full inventory and filters locally. If any page fails, no partial report is published. Activity shows progress for each downloaded page and for the analysis step.

Use **Cancel scan** to stop a running scan. Scans do not show a blocking overlay, so elapsed status, Activity, and cancellation remain accessible. A cancelled scan publishes no report, so deletion stays unavailable until a complete scan finishes. Deletion runs cannot be cancelled.

The toolbox **Sign in** already requests these scopes; the operation requests them only if they are still missing:

- Scan: `DeviceManagementManagedDevices.Read.All` (an existing `ReadWrite.All` grant also suffices).
- Delete: `DeviceManagementManagedDevices.ReadWrite.All`, requested only for deletion.

The account also needs the applicable Intune licensing, RBAC visibility, and **Managed devices/Delete** permission for deletion. Scope consent alone does not grant Intune administrative access.

#### Matching and keeper protection

- **Hardware:** same valid serial number and manufacturer.
- **Entra:** same valid Entra device ID; this indicates a shared registration, not independently proven hardware identity.
- **Name:** same device name without a common strong identifier; always report-only.

Matching ignores casing and leading/trailing whitespace but preserves meaningful punctuation. Blank/placeholder/all-zero serials and invalid/empty GUIDs are not matching evidence. Missing hardware identifiers, conflicting hardware/model values, known virtual-machine models, disagreeing overlapping groups, and missing/default/invalid/future sync timestamps require review. Missing manufacturer data therefore prevents a serial-only match from becoming eligible.

The unique most recently synced record is the suggested keeper and cannot be deleted. Ties require review. Only unambiguous strong-match non-keepers last synced **at least 60 full days ago**, evaluated in UTC, are eligible for manual deletion. The initial threshold is fixed at 60 days; there is no keeper override or automatic selection.

One device can belong to both hardware and Entra evidence groups, so membership row counts can exceed unique-device counts. A protection in any group applies to that device everywhere. Changing a filter clears selection; filters never narrow the data exported by **Export CSV**.

#### Manual deletion

Select only eligible records and choose **Delete selected**. The confirmation lists exact Intune IDs, names, and tenant. The operation fetches a complete fresh inventory, recomputes eligibility, and rechecks the target and supporting keepers immediately before each delete and retry. Changed evidence, a missing keeper, tenant changes, or read failures block the affected deletion. A fresh scan is required after the run before further deletion is available; failures and individual outcomes remain in Activity.

**Warning:** [Intune deletion on Windows triggers Retire](https://learn.microsoft.com/intune/device-management/actions/delete); it is not merely cosmetic inventory cleanup. It can remove management, company data/settings, and affect BitLocker key protectors. Back up recovery keys and local administrator credentials before intentionally removing managed devices. A successful DELETE confirms server-side removal, not completion of endpoint retirement.

Only the selected Intune managed-device records are targeted. No Entra or Autopilot registrations are deleted and no separate wipe/retire commands are sent. There is no atomic Graph duplicate-cleanup transaction: revalidation reduces, but cannot eliminate, the read/delete race. RBAC can hide other relevant records, and cloned/virtual machines or reused serial numbers can produce false matches. Operator review remains essential.

#### Standalone usage

```powershell
# Retain the report object (including fingerprints) for deletion revalidation.
$scan = & .\Features\Get-IntuneDuplicateDevices.ps1 -TenantId 'tenant-guid' -AsReport -PassThru
$scan.Rows | Format-Table GroupId, MatchType, DeviceName, Recommendation, ProtectionReason

# Copy specific reviewed Intune IDs; never automatically select all candidates.
$selectedIds = @('reviewed-managed-device-guid')
& .\Features\Remove-IntuneDuplicateDevices.ps1 -Id $selectedIds -Scan $scan -WhatIf -PassThru

# Only after reviewing the preview and retirement/BitLocker implications:
& .\Features\Remove-IntuneDuplicateDevices.ps1 -Id $selectedIds -Scan $scan -PassThru
```

Without `-AsReport`, `-PassThru` returns individual report rows. `-WhatIf` still needs read access for fresh validation but never requests write consent or sends write requests. Never reconstruct the deletion context from CSV: keep the original `$scan` report. Sign-out or tenant changes invalidate the GUI report.

Per-device outcomes are `Deleted`, `Protected`, `ChangedSinceScan`, `AlreadyAbsent`, `WhatIf`, or `Failed`, with explanatory messages. Initial authentication or full-inventory failures terminate the operation before any deletes.

CSV export always includes all loaded membership rows in this column order:

```text
GroupId, MatchType, MatchReason, Confidence, GroupSize, Id, DeviceName,
SerialNumber, Manufacturer, Model, AzureADDeviceId, UserPrincipalName,
OperatingSystem, OSVersion, EnrolledDateTime, LastSyncDateTime, DaysSinceLastSync,
ComplianceState, ManagementAgent, ManagementState, KeeperId, Recommendation,
CanDelete, ProtectionReason, TenantId, ScanTimeUtc
```

Start with a read-only pilot and review false positives before using manual deletion.

#### Troubleshooting large scans

Activity records the PowerShell/Graph Authentication versions, effective SDK request settings, each page request **before** it starts, page durations/counts, outer retry waits, and separate indexing/grouping/protection/report/summary timings. The GUI shows the last scan activity and elapsed seconds while the background worker runs; this is a waiting indicator, not proof that Graph is making progress. GUI filtering/publication is timed separately. Full-report device/group counts are cached, so selecting records does not repeatedly recount the entire report.

Interpret the last stage:

- **Requesting inventory page:** waiting inside Graph, potentially including SDK retries. Compare the eventual page duration with `ClientTimeout`, `MaxRetry`, and `RetriesTimeLimit`. The toolbox's logged outer retries are additional to SDK retries; neither policy is changed by diagnostics.
- **Analysis indexing/grouping/protections/report:** inventory is downloaded and local work is running. Coarse batch counts show progress without logging individual device records.
- **Scan complete but GUI publication is slow:** Graph and analysis have finished; inspect the separate GUI timing.
- **Scan failed:** the phase/page and error are logged. No partial report is published. A Windows-filter rejection may be followed by the documented full-inventory fallback.

From the toolbox directory, run this **read-only** diagnostic script to isolate Graph/analysis from GUI rendering:

```powershell
.\Tests\Invoke-IntuneDuplicateScanDiagnostic.ps1 -TenantId 'tenant-guid'
```

It signs in with read access only (or reuses the session with `-SkipGraphConnect`), never calls the removal feature, and writes two files to the current folder (`-OutputDirectory` to change):

- `DuplicateScanDiagnostic_<timestamp>.log`: every activity line with a timestamp and the seconds since the previous line.
- `DuplicateScanDiagnostic_<timestamp>_summary.txt`: status, inventory/row/device/group counts, pages, outer retries, slowest page, the longest wait and what preceded it, phase timings, SDK settings, and any error.

The summary is also written when the scan fails or you stop it with Ctrl+C, so a scan that hangs can be interrupted and still diagnosed. GUIDs, UPNs/e-mail addresses, and URLs (including continuation tokens) are masked; device records are never written. `-IncludeIdentifiers` disables masking for local use only. The script exits with code 1 when the scan does not complete.

Do not run the attached standalone cleanup script as a performance test: it has deletion paths, different matching rules, and extra per-duplicate Graph reads. The toolbox already uses hash indexes and a single inventory download.

For a stalled large-tenant scan, share only sanitized timestamps, counts, phase durations, SDK settings, and HTTP error status. Remove tenant IDs, device/user identifiers, URLs/continuation tokens, and any credentials from error text before sharing; do not share the inventory/report. The synthetic tests cannot establish a tenant-side root cause without this timing evidence. Existing timeout/retry behavior remains unchanged.

## Offline Validation

```powershell
powershell.exe -NoProfile -STA -File .\Tests\Test-IntuneToolbox.ps1
pwsh.exe -NoProfile -STA -File .\Tests\Test-IntuneToolbox.ps1

# Opt-in warmed scale measurements (1,000 / 7,000 / 14,000 devices and WPF rows).
pwsh.exe -NoProfile -STA -File .\Tests\Test-IntuneToolbox.ps1 -RunScaleBenchmark

# Duplicate analysis/pagination checks without WPF.
pwsh.exe -NoProfile -File .\Tests\Test-IntuneDuplicateDevices.ps1
```

The suite parses every package script, mocks Graph requests, checks returned profile/search data and assignment safeguards, verifies `-WhatIf` write protection and tenant mismatch handling, and builds the disconnected UI. Duplicate-device checks cover matching, pagination/retries, the server-side Windows filter and its fallback, progress logging, exact 60-day boundaries, keeper protection, changed/missing records, and deletion failures. GUI checks cover navigation, selection safeguards, filtering, CSV schema, scan success/failure/cancel states, refresh after deletion, actual scan-completion closures, stopping a running background worker, an end-to-end signed-in background run through the real runner (logs and results reach the UI thread), sign-in scope coverage, UI-thread permission requests (including cancelled sign-in and missing consent), and tenant invalidation. It also copies the package to a path containing spaces and launches it from a different working directory. No live tenant operations are performed.

Scale regressions scan 7,000 mocked devices over 28 pages, assert 14,000 hardware/Entra membership rows and exact keeper/eligibility decisions, reject scan writes, exercise late-page failures/exhausted retries, and check large name-only groups and distinct inventories. WPF checks exercise a 14,000-row report, count-cache invalidation, sorting/filter/selection preservation, and the elapsed waiting indicator. `-RunScaleBenchmark` additionally prints warmed repeated phase and WPF timings with exact output/request-count checks; timing thresholds are not imposed on routine tests because hardware/load vary.

Profile-management checks cover Administrative Templates inventory, filtering, API routing, rename/delete, N/A reporting, and CSV export. Multiple-group checks cover Add/Replace/Remove payloads for all three profile types, pagination, duplicate IDs, preservation of unrelated assignments and filters, Policy Set exclusion, retries, invalid inputs, and single-group compatibility. WPF tests exercise button availability while idle, busy, and after operations finish, both queues, actual search-completion closures, reset/clear/removal, all-queued-group snapshots, readiness without highlighted rows, and required profile/filter selection.

Audit checks cover user/target matching, multi-target events, all-page retrieval, UTC request bounds and daylight-saving transitions, missing optional data, consent/tenant errors, retries and late-page failures, successful-empty versus failed UI states, selected-event changes, change-type classification and row colors in both themes (including selected rows), Type/Operation sorting, full-result CSV/JSON preservation, and the real audit feature running through the background runner with a mocked Graph module. These are offline checks, not verification against a live tenant.
