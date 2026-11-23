<#PSScriptInfo

.VERSION 2.1
.AUTHOR Mikael Blomqvist
.EXTERNALMODULEDEPENDENCIES Microsoft.Graph.Authentication
.RELEASENOTES
v0.1 - Initial release (Intune Toolbox — three-tab GUI)
v0.2 - Added Delete Profiles tab; reordered tabs
v0.3 - Added confirmation dialogs to Bulk Assignment, Rename Profiles, and Autopilot Group Tags
v0.4 - Added fault resilience: Invoke-WithRetry with exponential backoff and Graph API 429/5xx retry handling for all bulk operations
v0.5 - Moved all bulk operations to background runspaces (Invoke-InBackground); UI remains responsive during bulk runs; buttons disabled while running
v0.6 - Redesigned theme: light mode is now the default; added dark/light toggle button; replaced purple accent with blue; secondary buttons use distinct sky-blue colour
v0.7 - Merged 'Rename Profiles' and 'Delete Profiles' tabs into a single 'Profile Management' tab with inner sub-tabs
v0.8 - Redesigned 'Profile Management' as a single unified tab: one shared profile list, optional match substring (blank = load all), Rename Selected and Delete Selected operate on the same loaded profiles
v0.9 - Profile Management: profiles no longer pre-selected after load; added Assignment Filter (Any / No assignments / Group assignment / All Devices / All Users) with per-profile assignment summary column
v1.0 - Merged 'Bulk Assignment' into 'Profile Management': tab is now a single unified view with three sub-tabs (Assignment, Rename, Delete) sharing one profile list and one output log; assignment engine rewritten to be apiType-aware (supports both configurationPolicies and deviceConfigurations); removed standalone Bulk Assignment tab
v1.1 - Profile Management layout: swapped columns so Profile Settings, Profiles list, and Output log are on the left; action sub-tabs (Assignment, Rename, Delete) moved to the right; added Prefix and Suffix fields to the Rename sub-tab
v1.2 - Added Export CSV button to the Profiles list; exports all loaded profiles (name, type, assignments, ID) to a user-chosen CSV file
v1.3 - Added optional 'Include succeeded count' checkbox; when enabled, fetches per-profile succeeded device count and shows it in a new Succeeded column
v1.4 - Added semi-transparent loading overlay shown during Load & Preview, Apply Assignment, Rename, and Delete operations; moved Load & Preview to background thread for responsive UI
v1.5 - Removed Naming Standard Validation from the Rename tab
v1.6 - Added Duplicate Devices: Windows duplicate evidence, CSV export, and guarded manual deletion
v1.7 - Added Administrative Templates management and persistent multi-group assignment/removal queues
v1.8 - Added read-only Intune Audit Lookup with user/target search, color-coded operation types, event details, and CSV export
v1.9 - Replaced Autopilot Group Tags with the Autopilot Group Tag Tool: serial/CSV lookup, profile assignment details, entered/CSV/clear tag actions, export, optional tenant-wide sync, and bulk Fresh Start
v2.0 - Added App Management: all-platform inventory, bulk app assignments, and settings-preserving direct assignment editing/removal
v2.1 - Added App Management Replace: previewed, add-first replacement of same-intent direct app assignments

.PRIVATEDATA
#>

<#
.SYNOPSIS
    Intune Toolbox — combined WPF GUI for everyday Intune and Autopilot tasks.

.DESCRIPTION
    A WPF-based GUI combining everyday Intune tools in a single tabbed window:
      Tab 1 - Profile Management  : Load, filter, assign, rename, and delete Settings Catalog, Device Configuration, and Administrative Templates profiles.
                                    Sub-tabs: Assignment (bulk assign/remove), Rename, Delete.
      Tab 2 - Autopilot Tags      : Look up Autopilot devices by serial number or CSV, review profile assignments, change or
                                    clear group tags, export results, and request bulk Fresh Start.
      Tab 3 - Setting Search     : Find settings across loaded Settings Catalog profiles.
      Tab 4 - Duplicate Devices  : Identify duplicate Windows records and manually remove eligible stale records.
      Tab 5 - Audit Lookup       : Search historical Intune changes by user and target name.
      Tab 6 - App Management     : Load apps, filter inventory columns, bulk assign or replace, and inspect/edit/remove direct assignments.

.EXAMPLE
    Interactive (user) authentication:
    .\Invoke-IntuneToolbox-GUI.ps1

    Pin delegated sign-in to a tenant:
    .\Invoke-IntuneToolbox-GUI.ps1 -TenantId <guid>

.EXAMPLE
    Build the interface without authentication:
    .\Invoke-IntuneToolbox-GUI.ps1 -ValidateOnly

.NOTES
    Version:        2.0
    Author:         Mikael Blomqvist
    Creation Date:  2026-03-13

        Delegated Microsoft Graph scopes are requested by operation. Required scopes include
        DeviceManagementConfiguration.Read/ReadWrite.All, DeviceManagementServiceConfig.Read/
        ReadWrite.All, DeviceManagementManagedDevices.Read/ReadWrite.All,
        DeviceManagementManagedDevices.PrivilegedOperations.All (Autopilot Fresh Start),
        DeviceManagementApps.Read/ReadWrite.All (apps and audit lookup), and Group.Read.All.
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [switch]$ValidateOnly
)

if ([Environment]::OSVersion.Platform -ne 'Win32NT') { throw 'Intune Toolbox requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Start PowerShell with -STA to use Intune Toolbox.' }

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$script:ToolVersion = '2.1'
$script:GraphState = [hashtable]::Synchronized(@{
    RequestedTenant = $TenantId
    ExpectedTenant = ''
    Tenant = ''
    Account = ''
    Connected = $false
})


#region XAML UI Definition

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Intune Toolbox" Height="880" Width="1400"
    MinHeight="700" MinWidth="1100"
    ResizeMode="CanResize" WindowStartupLocation="CenterScreen"
    Background="{DynamicResource BgBrush}">

    <Window.Resources>
        <SolidColorBrush x:Key="BgBrush"           Color="#191c22"/>
        <SolidColorBrush x:Key="AccentBrush"       Color="#3b82f6"/>
        <SolidColorBrush x:Key="AccentHover"       Color="#2563eb"/>
        <SolidColorBrush x:Key="SecondaryBtnBrush" Color="#252b34"/>
        <SolidColorBrush x:Key="SecondaryBtnHover" Color="#303946"/>
        <SolidColorBrush x:Key="DangerBrush"       Color="#b42332"/>
        <SolidColorBrush x:Key="DangerHoverBrush" Color="#dc2626"/>
        <SolidColorBrush x:Key="DangerPressedBrush" Color="#7f1d1d"/>
        <SolidColorBrush x:Key="SuccessBrush"     Color="#22a06b"/>
        <SolidColorBrush x:Key="AttentionBrush"   Color="#fbbf24"/>
        <SolidColorBrush x:Key="TextBrush"        Color="#e6e9ef"/>
        <SolidColorBrush x:Key="SubTextBrush"     Color="#9ca8b7"/>
        <SolidColorBrush x:Key="BorderBrush2"     Color="#39424e"/>
        <SolidColorBrush x:Key="InputBgBrush"     Color="#20252d"/>
        <SolidColorBrush x:Key="SurfaceBrush"     Color="#252b34"/>
        <SolidColorBrush x:Key="SecondaryBrush"   Color="#2c3541"/>
        <SolidColorBrush x:Key="SelectionBackground" Color="#20324b"/>
        <SolidColorBrush x:Key="SelectionHoverBackground" Color="#2a405d"/>
        <SolidColorBrush x:Key="SelectionAccent" Color="#3b82f6"/>
        <SolidColorBrush x:Key="SelectionForeground" Color="#e6e9ef"/>
        <SolidColorBrush x:Key="DisabledBackground" Color="#30343b"/>
        <SolidColorBrush x:Key="DisabledBorder" Color="#454b54"/>
        <SolidColorBrush x:Key="DisabledForeground" Color="#8b93a1"/>
        <SolidColorBrush x:Key="AuditCreateBrush" Color="#4ade80"/>
        <SolidColorBrush x:Key="AuditDeleteBrush" Color="#f87171"/>
        <SolidColorBrush x:Key="AuditUpdateBrush" Color="#60a5fa"/>
        <SolidColorBrush x:Key="AuditOtherBrush" Color="#fbbf24"/>

        <Style TargetType="DatePickerTextBox">
            <Setter Property="Background" Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/>
        </Style>

        <!-- Button -->
        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="{DynamicResource AccentBrush}"/>
            <Setter Property="Foreground"      Value="White"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="12,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="PrimaryButtonBorder" Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{DynamicResource AccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="PrimaryButtonBorder" Property="Background" Value="{DynamicResource DisabledBackground}"/>
                                <Setter TargetName="PrimaryButtonBorder" Property="BorderBrush" Value="{DynamicResource DisabledBorder}"/>
                                <Setter TargetName="PrimaryButtonBorder" Property="BorderThickness" Value="1"/>
                                <Setter Property="Foreground" Value="{DynamicResource DisabledForeground}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SecondaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="{DynamicResource SecondaryBtnBrush}"/>
            <Setter Property="Foreground"      Value="{DynamicResource TextBrush}"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="FontWeight"      Value="Normal"/>
            <Setter Property="Padding"         Value="10,5"/>
            <Setter Property="BorderBrush"     Value="{DynamicResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="SecondaryButtonBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{DynamicResource SecondaryBtnHover}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="SecondaryButtonBorder" Property="Background" Value="{DynamicResource DisabledBackground}"/>
                                <Setter TargetName="SecondaryButtonBorder" Property="BorderBrush" Value="{DynamicResource DisabledBorder}"/>
                                <Setter Property="Foreground" Value="{DynamicResource DisabledForeground}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="TertiaryBtn" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource SubTextBrush}"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="6,3"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Foreground" Value="{DynamicResource DisabledForeground}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="DangerBtn" TargetType="Button">
            <Setter Property="Background"      Value="{DynamicResource DangerBrush}"/>
            <Setter Property="Foreground"      Value="White"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="12,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="DangerButtonBorder" Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{DynamicResource DangerHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{DynamicResource DangerPressedBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="DangerButtonBorder" Property="Background" Value="{DynamicResource DisabledBackground}"/>
                                <Setter TargetName="DangerButtonBorder" Property="BorderBrush" Value="{DynamicResource DisabledBorder}"/>
                                <Setter TargetName="DangerButtonBorder" Property="BorderThickness" Value="1"/>
                                <Setter Property="Foreground" Value="{DynamicResource DisabledForeground}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox -->
        <Style TargetType="TextBox">
            <Setter Property="Background"      Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground"      Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush"     Value="{DynamicResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="6,4"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="CaretBrush"      Value="{DynamicResource TextBrush}"/>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{DynamicResource AccentBrush}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- ComboBox -->
        <Style TargetType="ComboBox">
            <Setter Property="Background"      Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground"      Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush"     Value="{DynamicResource BorderBrush2}"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="Padding"         Value="6,4"/>
            <Setter Property="MinHeight"       Value="26"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <Border x:Name="ComboBoxBorder" Background="{TemplateBinding Background}"
                                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                    CornerRadius="2"/>
                            <ToggleButton x:Name="DropDownToggle"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          Focusable="False" ClickMode="Press" Background="Transparent" BorderThickness="0">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="Transparent">
                                            <Grid>
                                                <Path Data="M 0 0 L 4 4 L 8 0 Z" Fill="{DynamicResource SubTextBrush}"
                                                      Width="8" Height="5" Stretch="Fill" HorizontalAlignment="Right"
                                                      VerticalAlignment="Center" Margin="0,0,9,0"/>
                                            </Grid>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              ContentStringFormat="{TemplateBinding SelectionBoxItemStringFormat}"
                                              TextElement.Foreground="{TemplateBinding Foreground}"
                                              HorizontalAlignment="Left" VerticalAlignment="Center" Margin="7,0,26,0"/>
                            <TextBox x:Name="PART_EditableTextBox" Visibility="Collapsed"
                                     Background="Transparent" Foreground="{TemplateBinding Foreground}"
                                     BorderThickness="0" CaretBrush="{DynamicResource TextBrush}"
                                     IsReadOnly="{TemplateBinding IsReadOnly}"
                                     FontSize="{TemplateBinding FontSize}"
                                     VerticalContentAlignment="Center" Margin="4,0,26,0"/>
                            <Popup x:Name="PART_Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="None">
                                <Border x:Name="DropDownBorder" Background="{DynamicResource InputBgBrush}"
                                        BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1" CornerRadius="2"
                                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                                    <ScrollViewer MaxHeight="{TemplateBinding MaxDropDownHeight}"
                                                  CanContentScroll="True" SnapsToDevicePixels="True">
                                        <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEditable" Value="True">
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Collapsed"/>
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsDropDownOpen" Value="True">
                                <Setter TargetName="ComboBoxBorder" Property="BorderBrush" Value="{DynamicResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="ComboBoxBorder" Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="Background" Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Padding"    Value="6,3"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="ItemBorder" Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter TextElement.Foreground="{TemplateBinding Foreground}"
                                              HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                              VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource SelectionBackground}"/>
                                <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource SecondaryBrush}"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsSelected" Value="True"/>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                </MultiTrigger.Conditions>
                                <Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource SelectionHoverBackground}"/>
                                <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                            </MultiTrigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- CheckBox -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground"               Value="{DynamicResource TextBrush}"/>
            <Setter Property="FontSize"                 Value="12"/>
            <Setter Property="Margin"                   Value="0,4,0,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <!-- ListBox -->
        <Style TargetType="ListBox">
            <Setter Property="Background"  Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground"  Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/>
            <Setter Property="FontSize"    Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBox">
                        <Border x:Name="ListBorder" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="1" SnapsToDevicePixels="True">
                            <ScrollViewer Focusable="False" Padding="{TemplateBinding Padding}">
                                <ItemsPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                            </ScrollViewer>
                        </Border>
                        <ControlTemplate.Triggers>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsGrouping" Value="True"/>
                                    <Condition Property="VirtualizingPanel.IsVirtualizingWhenGrouping" Value="False"/>
                                </MultiTrigger.Conditions>
                                <Setter Property="ScrollViewer.CanContentScroll" Value="False"/>
                            </MultiTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Padding" Value="6,3"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}"
                                SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                              VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource SelectionBackground}"/>
                    <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                </Trigger>
                <MultiTrigger>
                    <MultiTrigger.Conditions>
                        <Condition Property="IsMouseOver" Value="True"/>
                        <Condition Property="IsSelected" Value="False"/>
                    </MultiTrigger.Conditions>
                    <Setter Property="Background" Value="{DynamicResource SecondaryBrush}"/>
                </MultiTrigger>
                <MultiTrigger>
                    <MultiTrigger.Conditions>
                        <Condition Property="IsSelected" Value="True"/>
                        <Condition Property="IsMouseOver" Value="True"/>
                    </MultiTrigger.Conditions>
                    <Setter Property="Background" Value="{DynamicResource SelectionHoverBackground}"/>
                    <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                </MultiTrigger>
            </Style.Triggers>
        </Style>

        <!-- ListView -->
        <Style TargetType="ListView">
            <Setter Property="Background"  Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground"  Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/>
            <Setter Property="FontSize"    Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListView">
                        <Border x:Name="ListBorder" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                SnapsToDevicePixels="True">
                            <ScrollViewer Padding="{TemplateBinding Padding}"
                                          Style="{DynamicResource {x:Static GridView.GridViewScrollViewerStyleKey}}">
                                <ItemsPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                            </ScrollViewer>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocusWithin" Value="True">
                                <Setter TargetName="ListBorder" Property="BorderBrush" Value="{DynamicResource AccentBrush}"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsGrouping" Value="True"/>
                                    <Condition Property="VirtualizingPanel.IsVirtualizingWhenGrouping" Value="False"/>
                                </MultiTrigger.Conditions>
                                <Setter Property="ScrollViewer.CanContentScroll" Value="False"/>
                            </MultiTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="2,0,0,0"/>
            <!-- Own template replaces the system item chrome, whose selection gradient overrides the solid theme brush. -->
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}"
                                SnapsToDevicePixels="True">
                            <GridViewRowPresenter Columns="{TemplateBinding GridView.ColumnCollection}"
                                                  Content="{TemplateBinding Content}"
                                                  VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                                                  SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <MultiTrigger>
                    <MultiTrigger.Conditions>
                        <Condition Property="IsMouseOver" Value="True"/>
                        <Condition Property="IsSelected" Value="False"/>
                    </MultiTrigger.Conditions>
                    <Setter Property="Background" Value="{DynamicResource SecondaryBrush}"/>
                </MultiTrigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource SelectionBackground}"/>
                    <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                    <Setter Property="BorderBrush" Value="{DynamicResource SelectionAccent}"/>
                </Trigger>
                <MultiTrigger>
                    <MultiTrigger.Conditions>
                        <Condition Property="IsMouseOver" Value="True"/>
                        <Condition Property="IsSelected" Value="True"/>
                    </MultiTrigger.Conditions>
                    <Setter Property="Background" Value="{DynamicResource SelectionHoverBackground}"/>
                    <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                    <Setter Property="BorderBrush" Value="{DynamicResource SelectionAccent}"/>
                </MultiTrigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background"      Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="Foreground"      Value="{DynamicResource TextBrush}"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="8,5"/>
            <Setter Property="BorderBrush"     Value="{DynamicResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GridViewColumnHeader">
                        <Grid>
                            <Border x:Name="HeaderBorder" Background="{TemplateBinding Background}"
                                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                    Padding="{TemplateBinding Padding}">
                                <ContentPresenter RecognizesAccessKey="True"
                                                  HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                                  VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                                                  TextElement.Foreground="{TemplateBinding Foreground}"/>
                            </Border>
                            <Thumb x:Name="PART_HeaderGripper" Width="8" HorizontalAlignment="Right"
                                   Cursor="SizeWE">
                                <Thumb.Template>
                                    <ControlTemplate TargetType="Thumb">
                                        <Border Background="Transparent"/>
                                    </ControlTemplate>
                                </Thumb.Template>
                            </Thumb>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="HeaderBorder" Property="Background" Value="{DynamicResource SecondaryBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="HeaderBorder" Property="Background" Value="{DynamicResource SelectionBackground}"/>
                                <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                            </Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="HeaderBorder" Property="BorderBrush" Value="{DynamicResource SelectionAccent}"/>
                                <Setter TargetName="HeaderBorder" Property="BorderThickness" Value="0,0,0,2"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Foreground" Value="{DynamicResource DisabledForeground}"/>
                            </Trigger>
                            <Trigger Property="Role" Value="Padding">
                                <Setter TargetName="PART_HeaderGripper" Property="Visibility" Value="Collapsed"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SortableGridViewColumnHeader" TargetType="GridViewColumnHeader"
               BasedOn="{StaticResource {x:Type GridViewColumnHeader}}">
            <Setter Property="Cursor" Value="Hand"/>
        </Style>

        <!-- Label -->
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="{DynamicResource SubTextBrush}"/>
            <Setter Property="FontSize"   Value="11"/>
            <Setter Property="Padding"    Value="0,0,0,2"/>
        </Style>

        <!-- GroupBox -->
        <Style TargetType="GroupBox">
            <Setter Property="Foreground"      Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush"     Value="{DynamicResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="8"/>
            <Setter Property="Margin"          Value="0,0,0,10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,6">
                                <Rectangle Width="3" Fill="{DynamicResource BorderBrush2}"
                                           RadiusX="1" RadiusY="1"/>
                                <TextBlock Text="{TemplateBinding Header}"
                                           FontSize="12" FontWeight="SemiBold"
                                           Foreground="{DynamicResource TextBrush}"
                                           VerticalAlignment="Center"/>
                            </StackPanel>
                            <Border Grid.Row="1" BorderBrush="{DynamicResource BorderBrush2}"
                                    BorderThickness="1" CornerRadius="4"
                                    Padding="{TemplateBinding Padding}">
                                <ContentPresenter/>
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ToolTip -->
        <Style TargetType="ToolTip">
            <Setter Property="Background"      Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="Foreground"      Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush"     Value="{DynamicResource AccentBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="Padding"         Value="8,5"/>
        </Style>

        <!-- TabControl -->
        <Style TargetType="TabControl">
            <Setter Property="Background"  Value="{DynamicResource BgBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/>
        </Style>
        <Style x:Key="PageHostTabControl" TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid ClipToBounds="True">
                            <ContentPresenter ContentSource="SelectedContent" Margin="{TemplateBinding Padding}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="NavCategory" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource SubTextBrush}"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="8,14,0,6"/>
        </Style>
        <Style x:Key="NavItem" TargetType="RadioButton">
            <Setter Property="Foreground" Value="{DynamicResource SubTextBrush}"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="Margin" Value="0,1"/>
            <Setter Property="BorderThickness" Value="3,0,0,0"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="NavBorder" Background="Transparent" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="NavBorder" Property="Background" Value="{DynamicResource SecondaryBrush}"/>
                                <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="NavBorder" Property="Background" Value="{DynamicResource SelectionBackground}"/>
                                <Setter Property="BorderBrush" Value="{DynamicResource SelectionAccent}"/>
                                <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Background"      Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="Foreground"      Value="{DynamicResource SubTextBrush}"/>
            <Setter Property="BorderBrush"     Value="{DynamicResource BorderBrush2}"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="18,8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="{TemplateBinding Background}"
                                BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1,1,1,0"
                                CornerRadius="6,6,0,0" Margin="0,0,4,0" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header"
                                              HorizontalAlignment="Center" VerticalAlignment="Center"
                                              TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource SelectionBackground}"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="{DynamicResource SelectionAccent}"/>
                                <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                            </Trigger>
                            <!-- Hover on unselected tab: lighter blue -->
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                    <Condition Property="IsSelected"  Value="False"/>
                                </MultiTrigger.Conditions>
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource SecondaryBrush}"/>
                                <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                            </MultiTrigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                    <Condition Property="IsSelected"  Value="True"/>
                                </MultiTrigger.Conditions>
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource SelectionHoverBackground}"/>
                                <Setter Property="Foreground" Value="{DynamicResource SelectionForeground}"/>
                            </MultiTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <!-- Root grid: content + overlay -->
    <Grid>
        <!-- Loading Overlay (covers entire window) -->
        <Border x:Name="LoadingOverlay" Visibility="Collapsed"
                Panel.ZIndex="1000">
            <Border.Background>
                <SolidColorBrush Color="Black" Opacity="0.55"/>
            </Border.Background>
            <Border Background="{DynamicResource InputBgBrush}" CornerRadius="12"
                    Padding="40,28" HorizontalAlignment="Center" VerticalAlignment="Center"
                    BorderBrush="{DynamicResource AccentBrush}" BorderThickness="1.5">
                <Border.Effect>
                    <DropShadowEffect BlurRadius="30" ShadowDepth="0" Opacity="0.4" Color="Black"/>
                </Border.Effect>
                <StackPanel HorizontalAlignment="Center">
                    <TextBlock x:Name="TxtOverlayMessage" Text="Loading profiles..."
                               FontSize="16" FontWeight="SemiBold"
                               Foreground="{DynamicResource TextBrush}"
                               HorizontalAlignment="Center" Margin="0,0,0,12"/>
                    <ProgressBar IsIndeterminate="True" Width="220" Height="6"
                                 Foreground="{DynamicResource AccentBrush}"/>
                </StackPanel>
            </Border>
        </Border>

        <!-- Main content -->
        <Grid Margin="0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{DynamicResource SurfaceBrush}"
                BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="0,0,0,1" Padding="16,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock Text="Intune Toolbox" FontSize="21" FontWeight="SemiBold"
                               Foreground="{DynamicResource TextBrush}"/>
                    <TextBlock Text="Manage Intune profiles, Windows device records, and Autopilot devices."
                               FontSize="11" Foreground="{DynamicResource SubTextBrush}" Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock x:Name="TxtVersion" Text="" FontSize="11"
                               Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center" Margin="0,0,12,0"/>
                    <Button x:Name="BtnThemeToggle" Style="{StaticResource TertiaryBtn}"
                            Content="Light theme" ToolTip="Switch to the light theme"/>
                </StackPanel>
            </Grid>
        </Border>

        <Grid Grid.Row="1" Margin="8,8,12,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="196" MinWidth="176"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*" MinWidth="540"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="{DynamicResource SurfaceBrush}"
                    BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                        <StackPanel Margin="8,4">
                            <TextBlock Text="POLICY" Style="{StaticResource NavCategory}"/>
                            <RadioButton x:Name="NavProfileManagement" GroupName="ToolNavigation" Tag="0"
                                         Content="Profile Management" Style="{StaticResource NavItem}" IsChecked="True"/>
                            <RadioButton x:Name="NavSettingSearch" GroupName="ToolNavigation" Tag="2"
                                         Content="Setting Search" Style="{StaticResource NavItem}"/>
                            <TextBlock Text="APPS" Style="{StaticResource NavCategory}"/>
                            <RadioButton x:Name="NavAppManagement" GroupName="ToolNavigation" Tag="5"
                                         Content="App Management" Style="{StaticResource NavItem}"/>
                            <TextBlock Text="AUTOPILOT" Style="{StaticResource NavCategory}"/>
                            <RadioButton x:Name="NavGroupTags" GroupName="ToolNavigation" Tag="1"
                                         Content="Group Tags" Style="{StaticResource NavItem}"/>
                            <TextBlock Text="DEVICES" Style="{StaticResource NavCategory}"/>
                            <RadioButton x:Name="NavDuplicateDevices" GroupName="ToolNavigation" Tag="3"
                                         Content="Duplicate Devices" Style="{StaticResource NavItem}"/>
                            <TextBlock Text="AUDIT" Style="{StaticResource NavCategory}"/>
                            <RadioButton x:Name="NavAuditLookup" GroupName="ToolNavigation" Tag="4"
                                         Content="Audit Lookup" Style="{StaticResource NavItem}"/>
                        </StackPanel>
                    </ScrollViewer>
                    <Border Grid.Row="1" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="0,1,0,0" Padding="10,9">
                        <StackPanel>
                            <TextBlock Text="Microsoft Graph" FontSize="10" FontWeight="SemiBold"
                                       Foreground="{DynamicResource SubTextBrush}" Margin="0,0,0,5"/>
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,7">
                                <Ellipse x:Name="EllipseStatus" Width="8" Height="8"
                                         Fill="{DynamicResource SuccessBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                <TextBlock x:Name="TxtGraphConnectionStatus" Text="Connected" FontSize="11"
                                           Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center"/>
                            </StackPanel>
                            <TextBlock x:Name="TxtTenantName" Text="" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5"
                                       Foreground="{DynamicResource TextBrush}" TextTrimming="CharacterEllipsis"
                                       ToolTip=""/>
                            <Button x:Name="BtnSignInOut" Style="{StaticResource SecondaryBtn}"
                                    Content="Sign out" HorizontalAlignment="Left" ToolTip="Sign out of Microsoft Graph"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </Border>
            <Grid Grid.Column="2" Margin="4,2,0,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <StackPanel Grid.Row="0" Margin="4,0,4,10">
                    <TextBlock x:Name="TxtPageTitle" Text="Profile Management" FontSize="19" FontWeight="SemiBold"
                               Foreground="{DynamicResource TextBrush}"/>
                    <TextBlock x:Name="TxtPageDescription"
                               Text="Manage, rename, assign, export, and remove Intune configuration profiles."
                               FontSize="11" Foreground="{DynamicResource SubTextBrush}" Margin="0,3,0,0"/>
                </StackPanel>
                <TabControl Grid.Row="1" x:Name="MainTabs" Padding="0" Style="{StaticResource PageHostTabControl}">

            <!-- ════════════════════════════════════
                 TAB 1 — PROFILE MANAGEMENT
                 ════════════════════════════════════ -->
            <TabItem Header="Profile Management">
                <Grid Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="12"/>
                        <ColumnDefinition Width="370" MinWidth="320"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left column: Profile Settings + Profiles list + Output log -->
                    <Grid Grid.Column="0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*" MinHeight="60"/>
                            <RowDefinition Height="6"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <!-- Profile Settings (shared: search, types, assignment filter) -->
                        <GroupBox Grid.Row="0" Header="Filters" Margin="0,0,0,8">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="2*"/>
                                    <ColumnDefinition Width="12"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="12"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="12"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <Label Content="Profile name"/>
                                    <TextBox x:Name="TxtLoadFilter" ToolTip="Filter profiles by name; leave empty to load all matching profiles"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2">
                                    <Label Content="Assignment"/>
                                    <ComboBox x:Name="ComboAssignmentFilter">
                                        <ComboBoxItem Content="Any (no filter)"         Tag="Any"        IsSelected="True"/>
                                        <ComboBoxItem Content="No assignments"          Tag="None"/>
                                        <ComboBoxItem Content="Group assignment"        Tag="Group"/>
                                        <ComboBoxItem Content="Assigned to All Devices" Tag="AllDevices"/>
                                        <ComboBoxItem Content="Assigned to All Users"   Tag="AllUsers"/>
                                    </ComboBox>
                                </StackPanel>
                                <StackPanel Grid.Column="4">
                                    <Label Content="Platform"/>
                                    <ComboBox x:Name="ComboPlatformFilter">
                                        <ComboBoxItem Content="All platforms" Tag="All"     IsSelected="True"/>
                                        <ComboBoxItem Content="Windows"       Tag="Windows"/>
                                        <ComboBoxItem Content="iOS / iPadOS"  Tag="iOS"/>
                                        <ComboBoxItem Content="macOS"         Tag="macOS"/>
                                    </ComboBox>
                                </StackPanel>
                                <StackPanel Grid.Column="6" VerticalAlignment="Top">
                                    <Label Content=" "/>
                                    <Button x:Name="BtnPreview"
                                            Style="{StaticResource PrimaryBtn}"
                                            Content="Load profiles"
                                            Height="32"/>
                                </StackPanel>
                                <Grid Grid.Row="1" Grid.ColumnSpan="7" Margin="0,6,0,0">
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Profile types" FontSize="10" FontWeight="SemiBold"
                                               Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center" Margin="0,0,12,0"/>
                                    <WrapPanel Grid.Row="0" Grid.Column="1">
                                        <CheckBox x:Name="ChkSettingsCatalog" Content="Settings Catalog" IsChecked="True" Margin="0,0,18,0"/>
                                        <CheckBox x:Name="ChkDeviceConfig" Content="Device Configuration" IsChecked="True" Margin="0,0,18,0"/>
                                        <CheckBox x:Name="ChkAdministrativeTemplates" Content="Administrative Templates" IsChecked="True" Margin="0,0,18,0"/>
                                    </WrapPanel>
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Options" FontSize="10" FontWeight="SemiBold"
                                               Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center" Margin="0,0,12,0"/>
                                    <StackPanel Grid.Row="1" Grid.Column="1" Orientation="Horizontal">
                                        <CheckBox x:Name="ChkSucceededCount" Content="Include succeeded count"
                                                  ToolTip="Fetch per-profile succeeded device count (slower load)" Margin="0,0,18,0"/>
                                    </StackPanel>
                                </Grid>
                            </Grid>
                        </GroupBox>

                        <GroupBox Grid.Row="1" Header="Profiles">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Grid Grid.Row="0" Margin="0,0,0,6">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="8"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="8"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="6"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="6"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox Grid.Column="0" x:Name="TxtProfileFilter"
                                             ToolTip="Filter loaded profiles by current or preview name"/>
                                    <TextBlock Grid.Column="2" x:Name="TxtFilterCount" VerticalAlignment="Center"
                                               Foreground="{DynamicResource TextBrush}" FontSize="12"/>
                                    <Button  Grid.Column="4" x:Name="BtnRenameSelectAll"
                                             Style="{StaticResource SecondaryBtn}" Content="Select all"/>
                                    <Button  Grid.Column="6" x:Name="BtnRenameDeselectAll"
                                             Style="{StaticResource SecondaryBtn}" Content="Clear selection"/>
                                    <Button  Grid.Column="8" x:Name="BtnExportCsv"
                                             Style="{StaticResource SecondaryBtn}" Content="Export CSV"
                                             ToolTip="Export all loaded profiles to a CSV file"/>
                                </Grid>
                                <Grid Grid.Row="1">
                                    <ListView x:Name="ListProfiles" SelectionMode="Extended">
                                        <ListView.View>
                                            <GridView ColumnHeaderContainerStyle="{StaticResource SortableGridViewColumnHeader}">
                                                <GridViewColumn Header="Current Name" DisplayMemberBinding="{Binding CurrentName}" Width="220"/>
                                                <GridViewColumn Header="New Name" DisplayMemberBinding="{Binding NewName}" Width="190"/>
                                                <GridViewColumn Header="Type" DisplayMemberBinding="{Binding ProfileType}" Width="160"/>
                                                <GridViewColumn Header="Assignments" DisplayMemberBinding="{Binding Assignments}" Width="180"/>
                                                <GridViewColumn Header="Succeeded" DisplayMemberBinding="{Binding Succeeded}" Width="90"/>
                                                <GridViewColumn Header="Compliant" DisplayMemberBinding="{Binding Compliant}" Width="80"/>
                                            </GridView>
                                        </ListView.View>
                                    </ListView>
                                    <Border x:Name="ProfileEmptyState" Background="{DynamicResource InputBgBrush}"
                                            Visibility="Visible" IsHitTestVisible="False">
                                        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Margin="16">
                                            <TextBlock x:Name="TxtProfileEmptyTitle" Text="No profiles loaded" FontSize="13"
                                                       FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}"
                                                       HorizontalAlignment="Center"/>
                                            <TextBlock x:Name="TxtProfileEmptyDescription"
                                                       Text="Choose the filters above and select Load profiles to begin."
                                                       FontSize="11" Foreground="{DynamicResource SubTextBrush}"
                                                       TextAlignment="Center" TextWrapping="Wrap" Margin="0,4,0,0"/>
                                        </StackPanel>
                                    </Border>
                                </Grid>
                            </Grid>
                        </GroupBox>

                        <GridSplitter Grid.Row="2" Height="6" HorizontalAlignment="Stretch"
                                      Background="{DynamicResource BorderBrush2}" Cursor="SizeNS"
                                      ResizeDirection="Rows" ResizeBehavior="PreviousAndNext"/>

                        <Expander x:Name="ProfileActivityExpander" Grid.Row="3" IsExpanded="False" Margin="0,4,0,0">
                            <Expander.Header>
                                <Grid>
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                    <TextBlock Text="Activity" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}"/>
                                    <TextBlock x:Name="TxtProfileActivitySummary" Grid.Column="1" Text="No activity yet" FontSize="10"
                                               Foreground="{DynamicResource SubTextBrush}" TextTrimming="CharacterEllipsis"
                                               HorizontalAlignment="Stretch" TextAlignment="Left" Margin="12,0,0,0"/>
                                </Grid>
                            </Expander.Header>
                            <ScrollViewer x:Name="ProfileLogScroll" VerticalScrollBarVisibility="Auto" MaxHeight="150">
                                <RichTextBox x:Name="TxtProfileLog" IsReadOnly="True"
                                             Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}"
                                             BorderThickness="0" FontFamily="Consolas" FontSize="11"
                                             VerticalScrollBarVisibility="Disabled" HorizontalScrollBarVisibility="Disabled"/>
                            </ScrollViewer>
                        </Expander>
                    </Grid>

                    <!-- Right column: Action sub-tabs (Assignment / Rename / Delete) -->
                    <Grid Grid.Column="2">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <StackPanel Grid.Row="0" Margin="2,0,2,8">
                            <TextBlock x:Name="TxtProfileSelectionSummary" Text="No profiles selected" FontSize="12"
                                       FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}"/>
                            <TextBlock x:Name="TxtProfileSelectionHint"
                                       Text="Select one or more profiles to manage assignments, rename, or delete them."
                                       FontSize="10" Foreground="{DynamicResource SubTextBrush}" TextWrapping="Wrap" Margin="0,2,0,0"/>
                        </StackPanel>

                        <!-- Action sub-tabs -->
                        <TabControl Grid.Row="1" Margin="0,0,0,8">

                            <!-- Assignment sub-tab -->
                            <TabItem Header="Assignment">
                                <ScrollViewer VerticalScrollBarVisibility="Auto">
                                    <StackPanel Margin="0,4,0,0">

                                        <GroupBox Header="Action">
                                            <StackPanel>
                                                <Label Content="Assignment action"/>
                                                <ComboBox x:Name="ComboAction">
                                                    <ComboBoxItem Content="Add"     Tag="Add"/>
                                                    <ComboBoxItem Content="Replace" Tag="Replace"/>
                                                    <ComboBoxItem Content="Remove"  Tag="Remove"/>
                                                </ComboBox>
                                            </StackPanel>
                                        </GroupBox>

                                        <!-- Add / Replace options -->
                                        <StackPanel x:Name="PanelAddReplace">
                                            <GroupBox Header="Assignment Options">
                                                <Grid>
                                                    <Grid.RowDefinitions>
                                                        <RowDefinition Height="Auto"/>
                                                        <RowDefinition Height="Auto"/>
                                                    </Grid.RowDefinitions>
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="*"/>
                                                        <ColumnDefinition Width="8"/>
                                                        <ColumnDefinition Width="*"/>
                                                    </Grid.ColumnDefinitions>
                                                    <StackPanel Grid.Column="0" Grid.ColumnSpan="3">
                                                        <Label Content="Include / Exclude"/>
                                                        <ComboBox x:Name="ComboIncludeExclude">
                                                            <ComboBoxItem Content="Include" Tag="include" IsSelected="True"/>
                                                            <ComboBoxItem Content="Exclude" Tag="exclude"/>
                                                        </ComboBox>
                                                        <Label Content="Assignment target" Margin="0,8,0,0"/>
                                                        <ComboBox x:Name="ComboTargetType">
                                                            <ComboBoxItem Content="All Users"   Tag="Users"/>
                                                            <ComboBoxItem Content="All Devices" Tag="Devices"/>
                                                            <ComboBoxItem Content="Group"       Tag="Group"/>
                                                        </ComboBox>
                                                    </StackPanel>
                                                    <StackPanel Grid.Row="1" Grid.ColumnSpan="3" x:Name="PanelGroupSearch" Margin="0,8,0,0">
                                                        <Label Content="Search for groups"/>
                                                        <Grid>
                                                            <Grid.ColumnDefinitions>
                                                                <ColumnDefinition Width="*"/>
                                                                <ColumnDefinition Width="6"/>
                                                                <ColumnDefinition Width="Auto"/>
                                                            </Grid.ColumnDefinitions>
                                                            <TextBox Grid.Column="0" x:Name="TxtGroupSearch"/>
                                                            <Button  Grid.Column="2" x:Name="BtnSearchGroup"
                                                                     Style="{StaticResource SecondaryBtn}" Content="Search"/>
                                                        </Grid>
                                                        <Label Content="Search results (Ctrl/Shift to select multiple)" Margin="0,6,0,0"/>
                                                        <ListBox x:Name="ListGroups" Height="90" DisplayMemberPath="DisplayName" SelectionMode="Extended"/>
                                                        <Button x:Name="BtnQueueGroups" Content="Add selected" Style="{StaticResource SecondaryBtn}" Margin="0,4,0,0"/>
                                                        <TextBlock x:Name="TxtQueuedGroupCount" Text="0 groups queued" Foreground="{DynamicResource SubTextBrush}" Margin="0,8,0,4"/>
                                                        <ListBox x:Name="ListQueuedGroups" Height="90" DisplayMemberPath="DisplayName" SelectionMode="Extended"/>
                                                        <WrapPanel Margin="0,4,0,0">
                                                            <Button x:Name="BtnUnqueueGroups" Content="Remove selected" Style="{StaticResource SecondaryBtn}" Margin="0,0,6,0"/>
                                                            <Button x:Name="BtnClearQueuedGroups" Content="Clear" Style="{StaticResource SecondaryBtn}"/>
                                                        </WrapPanel>
                                                        <TextBlock Text="Apply uses every queued group; no highlighting required." TextWrapping="Wrap" Foreground="{DynamicResource SubTextBrush}" FontSize="10" Margin="0,4,0,0"/>
                                                    </StackPanel>
                                                </Grid>
                                            </GroupBox>

                                            <GroupBox x:Name="GroupFilter" Header="Device Filter (optional)">
                                                <StackPanel>
                                                    <Label Content="Filter direction"/>
                                                    <ComboBox x:Name="ComboFilterMode">
                                                        <ComboBoxItem Content="No filter"      Tag="None"    IsSelected="True"/>
                                                        <ComboBoxItem Content="Include filter" Tag="Include"/>
                                                        <ComboBoxItem Content="Exclude filter" Tag="Exclude"/>
                                                    </ComboBox>
                                                    <StackPanel x:Name="PanelFilterList" Margin="0,8,0,0">
                                                        <Label Content="Select filter"/>
                                                        <TextBox x:Name="TxtFilterSearch" Margin="0,0,0,4" ToolTip="Type to filter the list"/>
                                                        <ListBox x:Name="ListFilters" Height="90" DisplayMemberPath="DisplayName"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </GroupBox>
                                        </StackPanel>

                                        <!-- Remove options -->
                                        <GroupBox x:Name="GroupRemove" Header="Remove Options" Visibility="Collapsed">
                                            <Grid>
                                                <Grid.RowDefinitions>
                                                    <RowDefinition Height="Auto"/>
                                                    <RowDefinition Height="Auto"/>
                                                </Grid.RowDefinitions>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="8"/>
                                                    <ColumnDefinition Width="*"/>
                                                </Grid.ColumnDefinitions>
                                                <StackPanel Grid.Column="0" Grid.ColumnSpan="3">
                                                    <Label Content="Remove scope"/>
                                                    <ComboBox x:Name="ComboRemoveScope">
                                                        <ComboBoxItem Content="Remove group assignment" Tag="Group" IsSelected="True"/>
                                                        <ComboBoxItem Content="Remove ALL assignments"  Tag="All"/>
                                                    </ComboBox>
                                                </StackPanel>
                                                <StackPanel Grid.Row="1" Grid.ColumnSpan="3" x:Name="PanelRemoveGroup" Margin="0,8,0,0">
                                                    <Label Content="Search for groups"/>
                                                    <Grid>
                                                        <Grid.ColumnDefinitions>
                                                            <ColumnDefinition Width="*"/>
                                                            <ColumnDefinition Width="6"/>
                                                            <ColumnDefinition Width="Auto"/>
                                                        </Grid.ColumnDefinitions>
                                                        <TextBox Grid.Column="0" x:Name="TxtRemoveGroupSearch"/>
                                                        <Button  Grid.Column="2" x:Name="BtnSearchRemoveGroup"
                                                                 Style="{StaticResource SecondaryBtn}" Content="Search"/>
                                                    </Grid>
                                                    <Label Content="Search results (Ctrl/Shift to select multiple)" Margin="0,6,0,0"/>
                                                    <ListBox x:Name="ListRemoveGroups" Height="90" DisplayMemberPath="DisplayName" SelectionMode="Extended"/>
                                                    <Button x:Name="BtnQueueRemoveGroups" Content="Add selected" Style="{StaticResource SecondaryBtn}" Margin="0,4,0,0"/>
                                                    <TextBlock x:Name="TxtQueuedRemoveGroupCount" Text="0 groups queued" Foreground="{DynamicResource SubTextBrush}" Margin="0,8,0,4"/>
                                                    <ListBox x:Name="ListQueuedRemoveGroups" Height="90" DisplayMemberPath="DisplayName" SelectionMode="Extended"/>
                                                    <WrapPanel Margin="0,4,0,0">
                                                        <Button x:Name="BtnUnqueueRemoveGroups" Content="Remove selected" Style="{StaticResource SecondaryBtn}" Margin="0,0,6,0"/>
                                                        <Button x:Name="BtnClearQueuedRemoveGroups" Content="Clear" Style="{StaticResource SecondaryBtn}"/>
                                                    </WrapPanel>
                                                    <TextBlock Text="Apply removes assignments for every queued group; no highlighting required." TextWrapping="Wrap" Foreground="{DynamicResource SubTextBrush}" FontSize="10" Margin="0,4,0,0"/>
                                                    <Label Content="Assignment type to remove" Margin="0,8,0,0"/>
                                                    <ComboBox x:Name="ComboRemoveType">
                                                        <ComboBoxItem Content="Include" Tag="include"/>
                                                        <ComboBoxItem Content="Exclude" Tag="exclude" IsSelected="True"/>
                                                        <ComboBoxItem Content="Both"    Tag="both"/>
                                                    </ComboBox>
                                                </StackPanel>
                                            </Grid>
                                        </GroupBox>

                                        <Button x:Name="BtnAssignExecute" Style="{StaticResource PrimaryBtn}"
                                                Content="Apply Assignment" Height="32" Margin="0,8,0,0"/>

                                    </StackPanel>
                                </ScrollViewer>
                            </TabItem>

                            <!-- Rename sub-tab -->
                            <TabItem Header="Rename">
                                <ScrollViewer VerticalScrollBarVisibility="Auto">
                                    <StackPanel Margin="0,8,0,0">
                                        <GroupBox Header="Rename Settings">
                                            <StackPanel>
                                                <Label Content="Match substring"/>
                                                <TextBox x:Name="TxtMatch" ToolTip="Substring to find in profile names for replacement"/>
                                                <Label Content="Replace with" Margin="0,8,0,0"/>
                                                <TextBox x:Name="TxtReplace" ToolTip="Replaces the Match substring with this value"/>
                                                <Label Content="Prefix (prepend to name)" Margin="0,8,0,0"/>
                                                <TextBox x:Name="TxtPrefix" ToolTip="Text to add before the profile name"/>
                                                <Label Content="Suffix (append to name)" Margin="0,8,0,0"/>
                                                <TextBox x:Name="TxtSuffix" ToolTip="Text to add after the profile name"/>
                                                <CheckBox x:Name="ChkRemoveWhitespace" Content="Remove whitespace" Margin="0,10,0,0"
                                                          ToolTip="Remove all whitespace characters from profile names"/>
                                                <Button x:Name="BtnPreviewRename" Style="{StaticResource SecondaryBtn}"
                                                        Content="Preview rename" Height="32" Margin="0,10,0,0"
                                                        ToolTip="Preview how selected profiles will look after rename operations (updates the New Name column)"/>
                                            </StackPanel>
                                        </GroupBox>
                                        <Button x:Name="BtnRenameExecute" Style="{StaticResource PrimaryBtn}"
                                                Content="Rename selected" Height="32" Margin="0,8,0,0"/>
                                    </StackPanel>
                                </ScrollViewer>
                            </TabItem>

                            <!-- Delete sub-tab -->
                            <TabItem Header="Delete">
                                <StackPanel Margin="0,8,0,0">
                                    <Button x:Name="BtnDeleteExecute" Style="{StaticResource DangerBtn}"
                                            Content="Delete profiles" Height="32"/>
                                </StackPanel>
                            </TabItem>

                        </TabControl>

                        <!-- Reset button -->
                        <StackPanel Grid.Row="2" HorizontalAlignment="Right">
                            <Button x:Name="BtnRenameReset" Style="{StaticResource TertiaryBtn}"
                                    Content="Reset fields"/>
                        </StackPanel>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- ════════════════════════════════════
                 TAB 2 — AUTOPILOT GROUP TAGS
                 ════════════════════════════════════ -->
            <TabItem Header="Autopilot Group Tags">
                <Grid Margin="0,8,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*" MinHeight="160"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="12"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <GroupBox Grid.Column="0" Header="Find Devices">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Label Grid.Row="0" Content="Serial number(s)"/>
                                <Grid Grid.Row="1">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="8"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="TxtAPSerials" Grid.Column="0"
                                             ToolTip="Enter one or more serial numbers separated by commas, semicolons, spaces, or new lines"/>
                                    <Button x:Name="BtnAPLookup" Grid.Column="2" Style="{StaticResource PrimaryBtn}"
                                            Content="Look up" Height="28" MinWidth="96"/>
                                </Grid>
                                <Grid Grid.Row="2" Margin="0,10,0,0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="10"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Button x:Name="BtnAPImportCsv" Grid.Column="0" Style="{StaticResource SecondaryBtn}"
                                            Content="Import CSV..." ToolTip="Load serial numbers (and optional GroupTag values) from a CSV file"/>
                                    <TextBlock x:Name="TxtAPCsvFile" Grid.Column="2" Text="No CSV file loaded" FontSize="11"
                                               Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center"
                                               TextTrimming="CharacterEllipsis"/>
                                </Grid>
                                <TextBlock Grid.Row="3" Margin="0,8,0,0" FontSize="10" TextWrapping="Wrap"
                                           Foreground="{DynamicResource SubTextBrush}"
                                           Text="CSV: header 'SerialNumber' (or 'Serial Number' / 'Device Serial Number') and optional 'GroupTag' or 'NewGroupTag'. Comma or semicolon delimited."/>
                            </Grid>
                        </GroupBox>

                        <GroupBox Grid.Column="2" Header="Change Group Tag">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Label Grid.Row="0" Content="New group tag"/>
                                <Grid Grid.Row="1">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="8"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="TxtNewTag" Grid.Column="0" ToolTip="The group tag to apply to the selected devices"/>
                                    <Button x:Name="BtnAPApplyTag" Grid.Column="2" Style="{StaticResource PrimaryBtn}"
                                            Content="Apply to selected" Height="28" MinWidth="140"/>
                                </Grid>
                                <TextBlock x:Name="TxtAPTagHint" Grid.Row="2" Margin="0,3,0,0" FontSize="10"
                                           Foreground="{DynamicResource SubTextBrush}" Text="Applies one tag to every selected device."/>
                                <Grid Grid.Row="3" Margin="0,8,0,0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="8"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Button x:Name="BtnAPApplyCsvTags" Grid.Column="0" Style="{StaticResource SecondaryBtn}" VerticalAlignment="Top"
                                            Content="Apply CSV tags to selected"
                                            ToolTip="Set each selected device to the tag from its CSV row. Blank CSV tags are skipped."/>
                                    <StackPanel Grid.Column="2">
                                        <Button x:Name="BtnAPClearTag" Style="{StaticResource DangerBtn}"
                                                Content="Clear tag on selected" ToolTip="Remove the group tag from the selected devices"/>
                                        <Button x:Name="BtnAPFreshStart" Style="{StaticResource DangerBtn}" Margin="0,6,0,0"
                                                Content="Fresh Start selected" ToolTip="Reinstall Windows and remove apps and user data on selected Intune devices"/>
                                    </StackPanel>
                                </Grid>
                                <CheckBox x:Name="ChkAPSync" Grid.Row="4" Margin="0,8,0,0" IsChecked="False"
                                          Content="Request tenant-wide Autopilot sync after successful changes"
                                          ToolTip="Triggers a sync of all Autopilot devices in the tenant (not only the selected devices). Intune limits how often sync can run."/>
                            </Grid>
                        </GroupBox>
                    </Grid>

                    <GroupBox Grid.Row="1" Header="Devices" Margin="0,0,0,6">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*" MinHeight="60"/>
                            </Grid.RowDefinitions>
                            <Grid Grid.Row="0" Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="260"/>
                                    <ColumnDefinition Width="12"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="6"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="6"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="6"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="6"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="TxtAPDeviceFilter" Grid.Column="0"
                                         ToolTip="Filter by serial number, device name, group tag, profile, or status"/>
                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock x:Name="TxtCountLoaded" Text="0" FontSize="11" Foreground="{DynamicResource TextBrush}"/>
                                    <TextBlock Text=" loaded  •  " FontSize="11" Foreground="{DynamicResource SubTextBrush}"/>
                                    <TextBlock x:Name="TxtCountFound" Text="0" FontSize="11" Foreground="{DynamicResource TextBrush}"/>
                                    <TextBlock Text=" in Autopilot  •  " FontSize="11" Foreground="{DynamicResource SubTextBrush}"/>
                                    <TextBlock x:Name="TxtCountSelected" Text="0" FontSize="11" Foreground="{DynamicResource TextBrush}"/>
                                    <TextBlock Text=" selected" FontSize="11" Foreground="{DynamicResource SubTextBrush}"/>
                                </StackPanel>
                                <Button x:Name="BtnAPSelectAll" Grid.Column="3" Style="{StaticResource SecondaryBtn}"
                                        Content="Select all" ToolTip="Select every visible device that is uniquely registered in Autopilot"/>
                                <Button x:Name="BtnAPDeselectAll" Grid.Column="5" Style="{StaticResource SecondaryBtn}" Content="Clear selection"/>
                                <Button x:Name="BtnAPRefresh" Grid.Column="7" Style="{StaticResource SecondaryBtn}"
                                        Content="Refresh" ToolTip="Look up all loaded serial numbers again"/>
                                <Button x:Name="BtnAPExport" Grid.Column="9" Style="{StaticResource SecondaryBtn}"
                                        Content="Export CSV..." ToolTip="Export all loaded lookup and update results"/>
                                <Button x:Name="BtnAPClearResults" Grid.Column="11" Style="{StaticResource TertiaryBtn}"
                                        Content="Clear list"/>
                            </Grid>
                            <Grid Grid.Row="1">
                                <ListView x:Name="ListDevices" SelectionMode="Extended">
                                    <ListView.View>
                                        <GridView>
                                            <GridViewColumn Header="Serial Number" DisplayMemberBinding="{Binding RequestedSerial}" Width="140"/>
                                            <GridViewColumn Header="Device Name" DisplayMemberBinding="{Binding DeviceName}" Width="150"/>
                                            <GridViewColumn Header="Current Tag" DisplayMemberBinding="{Binding GroupTag}" Width="120"/>
                                            <GridViewColumn Header="CSV Tag" DisplayMemberBinding="{Binding CsvTag}" Width="100"/>
                                            <GridViewColumn Header="Assigned Profile" DisplayMemberBinding="{Binding ProfileName}" Width="180"/>
                                            <GridViewColumn Header="Profile Status" DisplayMemberBinding="{Binding ProfileAssignmentStatus}" Width="120"/>
                                            <GridViewColumn Header="Assigned Date" DisplayMemberBinding="{Binding ProfileAssignedDate}" Width="170"/>
                                            <GridViewColumn Header="Lookup" DisplayMemberBinding="{Binding LookupStatus}" Width="90"/>
                                            <GridViewColumn Header="Update" DisplayMemberBinding="{Binding UpdateStatus}" Width="160"/>
                                            <GridViewColumn Header="Fresh Start" DisplayMemberBinding="{Binding FreshStartStatus}" Width="180"/>
                                            <GridViewColumn Header="Details" DisplayMemberBinding="{Binding Details}" Width="320"/>
                                        </GridView>
                                    </ListView.View>
                                </ListView>
                                <Border x:Name="AutopilotEmptyState" Background="{DynamicResource InputBgBrush}"
                                        Visibility="Visible" IsHitTestVisible="False">
                                    <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Margin="16">
                                        <TextBlock x:Name="TxtAutopilotEmptyTitle" Text="No devices loaded" FontSize="13"
                                                   FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}"
                                                   HorizontalAlignment="Center"/>
                                        <TextBlock x:Name="TxtAutopilotEmptyDescription"
                                                   Text="Look up a serial number or import a CSV file to begin."
                                                   FontSize="11" Foreground="{DynamicResource SubTextBrush}"
                                                   TextAlignment="Center" TextWrapping="Wrap" Margin="0,4,0,0"/>
                                    </StackPanel>
                                </Border>
                            </Grid>
                        </Grid>
                    </GroupBox>
                    <Expander x:Name="AutopilotActivityExpander" Grid.Row="2" IsExpanded="False" Margin="0,4,0,0">
                            <Expander.Header>
                                <Grid>
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                    <TextBlock Text="Activity" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}"/>
                                    <TextBlock x:Name="TxtAutopilotActivitySummary" Grid.Column="1" Text="No activity yet" FontSize="10"
                                               Foreground="{DynamicResource SubTextBrush}" TextTrimming="CharacterEllipsis"
                                               HorizontalAlignment="Stretch" TextAlignment="Left" Margin="12,0,0,0"/>
                                </Grid>
                            </Expander.Header>
                            <ScrollViewer x:Name="AutopilotLogScroll" VerticalScrollBarVisibility="Auto" MaxHeight="150">
                                <RichTextBox x:Name="TxtAutopilotLog" IsReadOnly="True"
                                             Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}"
                                             BorderThickness="0" FontFamily="Consolas" FontSize="11"
                                             VerticalScrollBarVisibility="Disabled" HorizontalScrollBarVisibility="Disabled"/>
                            </ScrollViewer>
                    </Expander>
                </Grid>
            </TabItem>

            <!-- ════════════════════════════════════
                 TAB 3 — SETTING SEARCH
                 ════════════════════════════════════ -->
            <TabItem Header="Setting Search">
                <Grid Margin="0,8,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*" MinHeight="60"/>
                        <RowDefinition Height="6"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <GroupBox Grid.Row="0" Header="Search" Margin="0,0,0,8">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="8"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <Label Content="Setting name or definition ID (partial match)"/>
                                    <TextBox x:Name="TxtSettingSearchTerm"
                                             ToolTip="Enter a setting name or definition ID to search for across loaded profiles"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" VerticalAlignment="Bottom">
                                    <Button x:Name="BtnSettingSearchExecute" Style="{StaticResource PrimaryBtn}"
                                            Content="Search" Height="32" Width="110"/>
                                </StackPanel>
                            </Grid>
                            <Grid Margin="0,8,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Search scope" FontSize="10" FontWeight="SemiBold"
                                           Foreground="{DynamicResource SubTextBrush}" Margin="0,0,14,0"/>
                                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock x:Name="TxtSettingScopeTitle" Text="All loaded profiles"
                                               FontSize="11" Foreground="{DynamicResource TextBrush}"/>
                                    <TextBlock Text="  ·  " FontSize="11" Foreground="{DynamicResource SubTextBrush}"/>
                                    <TextBlock x:Name="TxtSettingScopeCount" Text="0 Settings Catalog profiles"
                                               FontSize="11" Foreground="{DynamicResource SubTextBrush}"/>
                                </StackPanel>
                            </Grid>
                        </StackPanel>
                    </GroupBox>

                    <GroupBox Grid.Row="1" Header="Results">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Grid Grid.Row="0" Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="8"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="8"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Grid.Column="0" x:Name="TxtSettingResultFilter"
                                         ToolTip="Type to filter results"/>
                                <TextBlock Grid.Column="2" x:Name="TxtSettingResultCount" VerticalAlignment="Center"
                                           Foreground="{DynamicResource TextBrush}" FontSize="12"/>
                                <Button Grid.Column="4" x:Name="BtnSettingSearchExport"
                                        Style="{StaticResource SecondaryBtn}" Content="Export CSV" IsEnabled="False"/>
                            </Grid>
                            <Grid Grid.Row="1">
                                <ListView x:Name="ListSettingResults" SelectionMode="Extended">
                                    <ListView.View>
                                        <GridView ColumnHeaderContainerStyle="{StaticResource SortableGridViewColumnHeader}">
                                            <GridViewColumn Header="Profile Name" DisplayMemberBinding="{Binding ProfileName}" Width="240"/>
                                            <GridViewColumn Header="Setting Name" DisplayMemberBinding="{Binding SettingName}" Width="300"/>
                                            <GridViewColumn Header="Value" DisplayMemberBinding="{Binding Value}" Width="220"/>
                                            <GridViewColumn Header="Definition ID" DisplayMemberBinding="{Binding DefinitionId}" Width="360"/>
                                        </GridView>
                                    </ListView.View>
                                </ListView>
                                <Border x:Name="SettingSearchEmptyState" Background="{DynamicResource InputBgBrush}"
                                        Visibility="Visible" IsHitTestVisible="False">
                                    <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Margin="16">
                                        <TextBlock x:Name="TxtSettingEmptyTitle" Text="No results" FontSize="13"
                                                   FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}"
                                                   HorizontalAlignment="Center"/>
                                        <TextBlock x:Name="TxtSettingEmptyDescription"
                                                   Text="Enter a setting name or definition ID to search across loaded Settings Catalog profiles."
                                                   FontSize="11" Foreground="{DynamicResource SubTextBrush}"
                                                   TextAlignment="Center" TextWrapping="Wrap" Margin="0,4,0,0"/>
                                    </StackPanel>
                                </Border>
                            </Grid>
                        </Grid>
                    </GroupBox>

                    <GridSplitter Grid.Row="2" Height="6" HorizontalAlignment="Stretch"
                                  Background="{DynamicResource BorderBrush2}" Cursor="SizeNS"
                                  ResizeDirection="Rows" ResizeBehavior="PreviousAndNext"/>

                    <Expander x:Name="SettingSearchActivityExpander" Grid.Row="3" IsExpanded="False" Margin="0,4,0,0">
                        <Expander.Header>
                            <Grid>
                                <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Text="Activity" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}"/>
                                <TextBlock x:Name="TxtSettingActivitySummary" Grid.Column="1" Text="No activity yet" FontSize="10"
                                           Foreground="{DynamicResource SubTextBrush}" TextTrimming="CharacterEllipsis"
                                           HorizontalAlignment="Stretch" TextAlignment="Left" Margin="12,0,0,0"/>
                            </Grid>
                        </Expander.Header>
                        <ScrollViewer x:Name="SettingSearchLogScroll" VerticalScrollBarVisibility="Auto" MaxHeight="150">
                            <RichTextBox x:Name="TxtSettingSearchLog" IsReadOnly="True"
                                         Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}"
                                         BorderThickness="0" FontFamily="Consolas" FontSize="11"
                                         VerticalScrollBarVisibility="Disabled" HorizontalScrollBarVisibility="Disabled"/>
                        </ScrollViewer>
                    </Expander>
                </Grid>
            </TabItem>

            <TabItem Header="Duplicate Devices">
                <Grid Margin="6">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <GroupBox Grid.Row="0" Header="Windows duplicate-device review">
                        <StackPanel>
                            <TextBlock Text="Deletion is manual: strong matches only, protected keeper, and at least 60 full days since last sync."
                                       TextWrapping="Wrap" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,6"/>
                            <TextBlock Text="WARNING: Windows deletion triggers Retire and can affect BitLocker. Back up recovery keys and local administrator credentials first."
                                       TextWrapping="Wrap" Foreground="{DynamicResource AttentionBrush}" Margin="0,0,0,8"/>
                            <WrapPanel>
                                <Button x:Name="BtnDuplicateScan" Content="Scan Windows devices" Style="{StaticResource PrimaryBtn}" Margin="0,0,8,0"/>
                                <ComboBox x:Name="ComboDuplicateMatch" Width="155" SelectedIndex="0" Margin="0,0,8,0">
                                    <ComboBoxItem Content="All match types" Tag="All"/>
                                    <ComboBoxItem Content="Hardware matches" Tag="Hardware"/>
                                    <ComboBoxItem Content="Entra ID matches" Tag="Entra"/>
                                    <ComboBoxItem Content="Name-only candidates" Tag="Name"/>
                                </ComboBox>
                                <TextBox x:Name="TxtDuplicateFilter" Width="190" ToolTip="Filter name, ID, serial, user, or protection reason" Margin="0,0,8,0"/>
                                <Button x:Name="BtnDuplicateExport" Content="Export CSV" Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0" IsEnabled="False"/>
                                <Button x:Name="BtnDuplicateDelete" Content="Delete selected" Style="{StaticResource DangerBtn}" Margin="0,0,8,0" IsEnabled="False"/>
                                <Button x:Name="BtnDuplicateCancel" Content="Cancel scan" Style="{StaticResource SecondaryBtn}" IsEnabled="False"
                                        ToolTip="Stop the running scan. Deletions cannot be cancelled."/>
                            </WrapPanel>
                            <TextBlock x:Name="TxtDuplicateCount" Text="Not scanned. No records are selected automatically."
                                       TextWrapping="Wrap" Foreground="{DynamicResource SubTextBrush}" Margin="0,8,0,0"/>
                        </StackPanel>
                    </GroupBox>
                    <ListView Grid.Row="1" x:Name="ListDuplicateDevices" SelectionMode="Extended">
                        <ListView.View>
                            <GridView ColumnHeaderContainerStyle="{StaticResource SortableGridViewColumnHeader}">
                                <GridViewColumn Header="Group" DisplayMemberBinding="{Binding GroupId}" Width="95"/>
                                <GridViewColumn Header="Match" DisplayMemberBinding="{Binding MatchType}" Width="80"/>
                                <GridViewColumn Header="Device name" DisplayMemberBinding="{Binding DeviceName}" Width="150"/>
                                <GridViewColumn Header="Recommendation" DisplayMemberBinding="{Binding Recommendation}" Width="165"/>
                                <GridViewColumn Header="Last sync" DisplayMemberBinding="{Binding LastSyncDateTime}" Width="190"/>
                                <GridViewColumn Header="Days" DisplayMemberBinding="{Binding DaysSinceLastSync}" Width="55"/>
                                <GridViewColumn Header="Protection / review reason" DisplayMemberBinding="{Binding ProtectionReason}" Width="300"/>
                                <GridViewColumn Header="Serial" DisplayMemberBinding="{Binding SerialNumber}" Width="140"/>
                                <GridViewColumn Header="Manufacturer" DisplayMemberBinding="{Binding Manufacturer}" Width="120"/>
                                <GridViewColumn Header="Model" DisplayMemberBinding="{Binding Model}" Width="140"/>
                                <GridViewColumn Header="User" DisplayMemberBinding="{Binding UserPrincipalName}" Width="200"/>
                                <GridViewColumn Header="Enrolled" DisplayMemberBinding="{Binding EnrolledDateTime}" Width="190"/>
                                <GridViewColumn Header="Intune device ID" DisplayMemberBinding="{Binding Id}" Width="270"/>
                                <GridViewColumn Header="Entra device ID" DisplayMemberBinding="{Binding AzureADDeviceId}" Width="270"/>
                                <GridViewColumn Header="Keeper ID" DisplayMemberBinding="{Binding KeeperId}" Width="270"/>
                                <GridViewColumn Header="Evidence" DisplayMemberBinding="{Binding MatchReason}" Width="280"/>
                                <GridViewColumn Header="Compliance" DisplayMemberBinding="{Binding ComplianceState}" Width="100"/>
                                <GridViewColumn Header="Management" DisplayMemberBinding="{Binding ManagementState}" Width="120"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Expander x:Name="DuplicateActivityExpander" Grid.Row="2" IsExpanded="False" Margin="0,6,0,0">
                        <Expander.Header>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Activity" Foreground="{DynamicResource TextBrush}" FontWeight="SemiBold"/>
                                <TextBlock x:Name="TxtDuplicateActivitySummary" Text="No activity yet" Margin="12,0,0,0"
                                           Foreground="{DynamicResource SubTextBrush}"/>
                            </StackPanel>
                        </Expander.Header>
                        <ScrollViewer x:Name="DuplicateLogScroll" VerticalScrollBarVisibility="Auto" MaxHeight="150">
                            <RichTextBox x:Name="TxtDuplicateLog" IsReadOnly="True" BorderThickness="0" FontFamily="Consolas" FontSize="11"
                                         Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}"/>
                        </ScrollViewer>
                    </Expander>
                </Grid>
            </TabItem>
            <TabItem Header="Audit Lookup">
                <Grid Margin="0,8,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <GroupBox Grid.Row="0" Header="Search Intune audit logs" Margin="0,0,0,8">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="12"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <Label Content="User (initiated by): UPN or user ID"/>
                                    <TextBox x:Name="TxtAuditUser" ToolTip="Case-insensitive partial match on the actor's UPN or user ID"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2">
                                    <Label Content="Target: object name"/>
                                    <TextBox x:Name="TxtAuditTarget" ToolTip="Case-insensitive partial match on any historical target name"/>
                                </StackPanel>
                            </Grid>
                            <TextBlock Text="Enter at least one search field. Both must match when supplied. Searches all Intune categories."
                                       Foreground="{DynamicResource SubTextBrush}" TextWrapping="Wrap" Margin="0,6,0,6"/>
                            <WrapPanel>
                                <StackPanel Margin="0,0,12,0">
                                    <Label Content="Start date (local)"/>
                                    <DatePicker x:Name="DateAuditStart" Width="150"
                                                Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}"/>
                                </StackPanel>
                                <StackPanel Margin="0,0,12,0">
                                    <Label Content="End date (local, inclusive)"/>
                                    <DatePicker x:Name="DateAuditEnd" Width="150"
                                                Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}"/>
                                </StackPanel>
                                <Button x:Name="BtnAuditSearch" Content="Search" Style="{StaticResource PrimaryBtn}"
                                        Width="110" Height="32" VerticalAlignment="Bottom"/>
                            </WrapPanel>
                            <TextBlock x:Name="TxtAuditTimeZone" Foreground="{DynamicResource SubTextBrush}" Margin="0,6,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </GroupBox>
                    <GroupBox Grid.Row="1" Header="Events">
                        <Grid>
                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                            <DockPanel Grid.Row="0" Margin="0,0,0,6">
                                <Button x:Name="BtnAuditExport" DockPanel.Dock="Right" Content="Export CSV" Style="{StaticResource SecondaryBtn}" IsEnabled="False"/>
                                <StackPanel x:Name="AuditLegend" DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center" Margin="12,0,12,0"
                                            ToolTip="Row stripe and Operation/Type text color show the kind of change">
                                    <Rectangle Width="10" Height="10" Fill="{DynamicResource AuditCreateBrush}" VerticalAlignment="Center"/>
                                    <TextBlock Text="Create" Margin="4,0,12,0" Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center"/>
                                    <Rectangle Width="10" Height="10" Fill="{DynamicResource AuditDeleteBrush}" VerticalAlignment="Center"/>
                                    <TextBlock Text="Delete" Margin="4,0,12,0" Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center"/>
                                    <Rectangle Width="10" Height="10" Fill="{DynamicResource AuditUpdateBrush}" VerticalAlignment="Center"/>
                                    <TextBlock Text="Update/Patch" Margin="4,0,12,0" Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center"/>
                                    <Rectangle Width="10" Height="10" Fill="{DynamicResource AuditOtherBrush}" VerticalAlignment="Center"/>
                                    <TextBlock Text="Other" Margin="4,0,0,0" Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center"/>
                                </StackPanel>
                                <TextBlock x:Name="TxtAuditCount" Text="Enter a user or target to search." VerticalAlignment="Center"
                                           Foreground="{DynamicResource TextBrush}" TextWrapping="Wrap"/>
                            </DockPanel>
                            <ListView Grid.Row="1" x:Name="ListAuditEvents" SelectionMode="Single">
                                <ListView.Resources>
                                    <Style x:Key="AuditKindText" TargetType="TextBlock">
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding OperationKind}" Value="Create">
                                                <Setter Property="Foreground" Value="{DynamicResource AuditCreateBrush}"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding OperationKind}" Value="Delete">
                                                <Setter Property="Foreground" Value="{DynamicResource AuditDeleteBrush}"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding OperationKind}" Value="Update">
                                                <Setter Property="Foreground" Value="{DynamicResource AuditUpdateBrush}"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding OperationKind}" Value="Other">
                                                <Setter Property="Foreground" Value="{DynamicResource AuditOtherBrush}"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </ListView.Resources>
                                <ListView.ItemContainerStyle>
                                    <!-- Kind triggers follow the shared selection triggers, so the stripe keeps its color when selected. -->
                                    <Style TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
                                        <Setter Property="BorderThickness" Value="4,0,0,0"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding OperationKind}" Value="Create">
                                                <Setter Property="BorderBrush" Value="{DynamicResource AuditCreateBrush}"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding OperationKind}" Value="Delete">
                                                <Setter Property="BorderBrush" Value="{DynamicResource AuditDeleteBrush}"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding OperationKind}" Value="Update">
                                                <Setter Property="BorderBrush" Value="{DynamicResource AuditUpdateBrush}"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding OperationKind}" Value="Other">
                                                <Setter Property="BorderBrush" Value="{DynamicResource AuditOtherBrush}"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </ListView.ItemContainerStyle>
                                <ListView.View>
                                    <GridView ColumnHeaderContainerStyle="{StaticResource SortableGridViewColumnHeader}">
                                        <GridViewColumn Header="Timestamp (local)" DisplayMemberBinding="{Binding TimestampLocal}" Width="195"/>
                                        <GridViewColumn Header="Type" Width="80">
                                            <GridViewColumn.CellTemplate>
                                                <DataTemplate>
                                                    <TextBlock Text="{Binding OperationKind}" FontWeight="SemiBold" Style="{StaticResource AuditKindText}"/>
                                                </DataTemplate>
                                            </GridViewColumn.CellTemplate>
                                        </GridViewColumn>
                                        <GridViewColumn Header="Operation" Width="230">
                                            <GridViewColumn.CellTemplate>
                                                <DataTemplate>
                                                    <TextBlock Text="{Binding Operation}" TextTrimming="CharacterEllipsis" Style="{StaticResource AuditKindText}"/>
                                                </DataTemplate>
                                            </GridViewColumn.CellTemplate>
                                        </GridViewColumn>
                                        <GridViewColumn Header="Result" DisplayMemberBinding="{Binding Result}" Width="90"/>
                                        <GridViewColumn Header="Initiated by" DisplayMemberBinding="{Binding InitiatedBy}" Width="220"/>
                                        <GridViewColumn Header="Application" DisplayMemberBinding="{Binding Application}" Width="180"/>
                                        <GridViewColumn Header="Target(s)" DisplayMemberBinding="{Binding TargetNames}" Width="270"/>
                                        <GridViewColumn Header="Category" DisplayMemberBinding="{Binding Category}" Width="120"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </GroupBox>
                    <Expander Grid.Row="2" x:Name="AuditDetailsExpander" Header="Selected event details" Margin="0,6,0,0"
                              Foreground="{DynamicResource TextBrush}">
                        <TextBox x:Name="TxtAuditDetails" IsReadOnly="True" Text="Select an event to inspect its actor, targets, and changes."
                                 AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" MaxHeight="220"
                                 FontFamily="Consolas" FontSize="11"/>
                    </Expander>
                    <Expander Grid.Row="3" x:Name="AuditActivityExpander" IsExpanded="False" Margin="0,6,0,0">
                        <Expander.Header>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Activity" Foreground="{DynamicResource TextBrush}" FontWeight="SemiBold"/>
                                <TextBlock x:Name="TxtAuditActivitySummary" Text="No activity yet" Margin="12,0,0,0" Foreground="{DynamicResource SubTextBrush}"/>
                            </StackPanel>
                        </Expander.Header>
                        <ScrollViewer x:Name="AuditLogScroll" VerticalScrollBarVisibility="Auto" MaxHeight="130">
                            <RichTextBox x:Name="TxtAuditLog" IsReadOnly="True" BorderThickness="0" FontFamily="Consolas" FontSize="11"
                                         Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}"/>
                        </ScrollViewer>
                    </Expander>
                </Grid>
            </TabItem>
            <TabItem Header="App Management">
                <Grid Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="12"/>
                        <ColumnDefinition Width="390" MinWidth="340"/>
                    </Grid.ColumnDefinitions>
                    <Grid Grid.Column="0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <GroupBox Header="App inventory">
                            <StackPanel>
                                <Label Content="Name contains (optional)"/>
                                <TextBox x:Name="TxtAppNameFilter"/>
                                <Label Content="Platform"/>
                                <ComboBox x:Name="ComboAppPlatform">
                                    <ComboBoxItem Content="All platforms" Tag="All" IsSelected="True"/>
                                    <ComboBoxItem Content="Windows" Tag="Windows"/>
                                    <ComboBoxItem Content="Android" Tag="Android"/>
                                    <ComboBoxItem Content="iOS / iPadOS" Tag="iOS"/>
                                    <ComboBoxItem Content="macOS" Tag="macOS"/>
                                </ComboBox>
                                <Button x:Name="BtnAppLoad" Content="Load apps" Style="{StaticResource PrimaryBtn}" Margin="0,8,0,0"/>
                                <TextBlock x:Name="TxtAppCount" Text="Load apps to begin." Margin="0,8,0,0"
                                           Foreground="{DynamicResource SubTextBrush}" TextWrapping="Wrap"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Grid.Row="1" Header="Apps - Ctrl/Shift to select multiple" Margin="0,8,0,0">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <StackPanel x:Name="PanelAppResultFilters" Margin="0,0,0,8">
                                    <TextBlock Text="Filter loaded apps (all criteria must match)" FontWeight="SemiBold"/>
                                    <WrapPanel>
                                        <StackPanel Width="160" Margin="0,0,8,0">
                                            <Label Content="App name contains"/>
                                            <TextBox x:Name="TxtAppResultName"/>
                                        </StackPanel>
                                        <StackPanel Width="140" Margin="0,0,8,0">
                                            <Label Content="Publisher contains"/>
                                            <TextBox x:Name="TxtAppResultPublisher"/>
                                        </StackPanel>
                                        <StackPanel Width="110" Margin="0,0,8,0">
                                            <Label Content="Platform"/>
                                            <ComboBox x:Name="ComboAppResultPlatform"/>
                                        </StackPanel>
                                        <StackPanel Width="170" Margin="0,0,8,0">
                                            <Label Content="App type"/>
                                            <ComboBox x:Name="ComboAppResultType"/>
                                        </StackPanel>
                                        <StackPanel Width="160" Margin="0,0,8,0">
                                            <Label Content="Assignments (hint)"/>
                                            <ComboBox x:Name="ComboAppResultStatus"/>
                                        </StackPanel>
                                        <Button x:Name="BtnAppClearResultFilters" Content="Clear filters" Style="{StaticResource TertiaryBtn}"
                                                VerticalAlignment="Bottom" ToolTip="Restore all loaded apps. Changing filters clears selection and assignment inspection."/>
                                    </WrapPanel>
                                </StackPanel>
                                <ListView Grid.Row="1" x:Name="ListApps" SelectionMode="Extended"
                                          VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling">
                                    <ListView.View>
                                        <GridView ColumnHeaderContainerStyle="{StaticResource SortableGridViewColumnHeader}">
                                            <GridViewColumn Header="App name" DisplayMemberBinding="{Binding DisplayName}" Width="230"/>
                                            <GridViewColumn Header="Platform" DisplayMemberBinding="{Binding Platform}" Width="95"/>
                                            <GridViewColumn Header="App type" DisplayMemberBinding="{Binding AppType}" Width="145"/>
                                            <GridViewColumn Header="Publisher" DisplayMemberBinding="{Binding Publisher}" Width="125"/>
                                            <GridViewColumn Header="Assignments (hint)" DisplayMemberBinding="{Binding AssignmentStatus}" Width="145"/>
                                        </GridView>
                                    </ListView.View>
                                </ListView>
                                <TextBlock Grid.Row="1" x:Name="TxtAppEmpty" Text="No apps loaded." IsHitTestVisible="False"
                                           HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="{DynamicResource SubTextBrush}"/>
                            </Grid>
                        </GroupBox>
                        <Expander Grid.Row="2" x:Name="AppActivityExpander" IsExpanded="False" Margin="0,6,0,0">
                            <Expander.Header>
                                <StackPanel Orientation="Horizontal">
                                    <TextBlock Text="Activity" Foreground="{DynamicResource TextBrush}" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="TxtAppActivitySummary" Text="No activity yet" Margin="12,0,0,0" Foreground="{DynamicResource SubTextBrush}"/>
                                </StackPanel>
                            </Expander.Header>
                            <ScrollViewer x:Name="AppLogScroll" VerticalScrollBarVisibility="Auto" MaxHeight="150">
                                <RichTextBox x:Name="TxtAppLog" IsReadOnly="True" BorderThickness="0" FontFamily="Consolas" FontSize="11"
                                             Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}"/>
                            </ScrollViewer>
                        </Expander>
                    </Grid>
                    <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="PanelAppActions">
                            <GroupBox Header="Assignment target and intent">
                                <StackPanel>
                                    <Label Content="Intent"/>
                                    <ComboBox x:Name="ComboAppIntent">
                                        <ComboBoxItem Content="Required" Tag="required" IsSelected="True"/>
                                        <ComboBoxItem Content="Available" Tag="available"/>
                                        <ComboBoxItem Content="Uninstall" Tag="uninstall"/>
                                    </ComboBox>
                                    <Label Content="Target"/>
                                    <ComboBox x:Name="ComboAppTarget">
                                        <ComboBoxItem Content="Groups" Tag="Group"/>
                                        <ComboBoxItem Content="All Users" Tag="Users"/>
                                        <ComboBoxItem Content="All Devices" Tag="Devices" IsSelected="True"/>
                                    </ComboBox>
                                    <ComboBox x:Name="ComboAppIncludeExclude" Margin="0,6,0,0">
                                        <ComboBoxItem Content="Include" Tag="include" IsSelected="True"/>
                                        <ComboBoxItem Content="Exclude group" Tag="exclude"/>
                                    </ComboBox>
                                    <StackPanel x:Name="PanelAppGroupSearch" Visibility="Collapsed">
                                        <Label Content="Group search"/>
                                        <TextBox x:Name="TxtAppGroupSearch"/>
                                        <Button x:Name="BtnAppGroupSearch" Content="Search groups" Style="{StaticResource SecondaryBtn}" Margin="0,6,0,0"/>
                                        <ListBox x:Name="ListAppGroups" DisplayMemberPath="DisplayName" SelectionMode="Extended" Height="80"/>
                                        <Button x:Name="BtnAppQueue" Content="Add selected groups" Style="{StaticResource SecondaryBtn}" Margin="0,6,0,0"/>
                                        <Label Content="Queued groups - every queued group is used"/>
                                        <ListBox x:Name="ListAppQueuedGroups" DisplayMemberPath="DisplayName" SelectionMode="Extended" Height="80"/>
                                        <WrapPanel>
                                            <Button x:Name="BtnAppUnqueue" Content="Remove selected" Style="{StaticResource TertiaryBtn}"/>
                                            <Button x:Name="BtnAppClearQueue" Content="Clear queue" Style="{StaticResource TertiaryBtn}"/>
                                        </WrapPanel>
                                    </StackPanel>
                                    <Label Content="Assignment filter"/>
                                    <ComboBox x:Name="ComboAppFilterMode">
                                        <ComboBoxItem Content="None" Tag="None" IsSelected="True"/>
                                        <ComboBoxItem Content="Include filter" Tag="include"/>
                                        <ComboBoxItem Content="Exclude filter" Tag="exclude"/>
                                    </ComboBox>
                                    <StackPanel x:Name="PanelAppFilter" Visibility="Collapsed">
                                        <Button x:Name="BtnAppFilters" Content="Load filters" Style="{StaticResource SecondaryBtn}" Margin="0,6,0,0"/>
                                        <ComboBox x:Name="ComboAppFilter" DisplayMemberPath="DisplayName"/>
                                    </StackPanel>
                                    <TextBlock Text="Existing app-specific settings are preserved when editing. New assignments request Intune defaults. Compatibility errors appear in Activity."
                                               TextWrapping="Wrap" Foreground="{DynamicResource SubTextBrush}" FontSize="11" Margin="0,8,0,0"/>
                                    <Button x:Name="BtnAppAdd" Content="Add to selected apps" Style="{StaticResource PrimaryBtn}" Margin="0,8,0,0"/>
                                    <Button x:Name="BtnAppReplace" Content="Replace this intent on selected apps" Style="{StaticResource DangerBtn}" Margin="0,6,0,0"
                                            ToolTip="Previews, then makes the direct assignments with this intent and include/exclude type match the queued targets."/>
                                    <TextBlock Text="Replace only changes direct assignments with the chosen intent and include/exclude type: it adds the queued targets, then removes the others. Removals are skipped for an app if any of its adds fail. Policy Set assignments are never changed."
                                               TextWrapping="Wrap" Foreground="{DynamicResource SubTextBrush}" FontSize="11" Margin="0,6,0,0"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Existing assignments - one app" Margin="0,8,0,0">
                                <StackPanel>
                                    <Button x:Name="BtnAppInspect" Content="Refresh assignments" Style="{StaticResource SecondaryBtn}"/>
                                    <TextBlock x:Name="TxtAppInspection" Text="Select one app and refresh assignments." TextWrapping="Wrap"
                                               Foreground="{DynamicResource SubTextBrush}" Margin="0,6,0,6"/>
                                    <StackPanel x:Name="PanelAppAssignmentDetails" Visibility="Collapsed">
                                    <ListView x:Name="ListAppAssignments" SelectionMode="Extended" Height="180">
                                        <ListView.View>
                                            <GridView>
                                                <GridViewColumn Header="Intent" DisplayMemberBinding="{Binding Intent}" Width="85"/>
                                                <GridViewColumn Header="Target" DisplayMemberBinding="{Binding TargetName}" Width="145"/>
                                                <GridViewColumn Header="Include/exclude" DisplayMemberBinding="{Binding IncludeExclude}" Width="95"/>
                                                <GridViewColumn Header="Source" DisplayMemberBinding="{Binding Source}" Width="80"/>
                                                <GridViewColumn Header="Filter mode" DisplayMemberBinding="{Binding FilterMode}" Width="80"/>
                                                <GridViewColumn Header="Filter ID" DisplayMemberBinding="{Binding FilterId}" Width="200"/>
                                            </GridView>
                                        </ListView.View>
                                    </ListView>
                                    <TextBox x:Name="TxtAppAssignmentDetails" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                                             MaxHeight="120" VerticalScrollBarVisibility="Auto"/>
                                    <Button x:Name="BtnAppEdit" Content="Load selected assignment into form" Style="{StaticResource SecondaryBtn}" Margin="0,6,0,0"/>
                                    <Button x:Name="BtnAppUpdate" Content="Update selected assignment" Style="{StaticResource PrimaryBtn}" Margin="0,6,0,0"/>
                                    <Button x:Name="BtnAppRemove" Content="Remove selected assignments" Style="{StaticResource DangerBtn}" Margin="0,6,0,0"/>
                                    <TextBlock Text="Policy Set-owned assignments are read-only. Removal deletes the assignment, not the app; it is not an Uninstall assignment."
                                               TextWrapping="Wrap" Foreground="{DynamicResource SubTextBrush}" FontSize="11" Margin="0,8,0,0"/>
                                    </StackPanel>
                                </StackPanel>
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
                </Grid>
            </TabItem>
        </TabControl>

            </Grid>
        </Grid>
        </Grid>
    </Grid>
</Window>
"@

#endregion XAML

#region Build Window

$reader = New-Object System.Xml.XmlNodeReader $xaml
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    [System.Windows.MessageBox]::Show(
        "Failed to load the UI window.`n`nError: $_",
        "Intune Toolbox — Startup Error",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    exit
}

# Shared header
$TxtTenantName      = $window.FindName("TxtTenantName")
$script:GraphUser = ''
$script:GraphTenantId = ''
$BtnSignInOut       = $window.FindName("BtnSignInOut")
$BtnThemeToggle     = $window.FindName("BtnThemeToggle")
$EllipseStatus      = $window.FindName("EllipseStatus")
$TxtGraphConnectionStatus = $window.FindName("TxtGraphConnectionStatus")
$TxtPageTitle       = $window.FindName("TxtPageTitle")
$TxtPageDescription = $window.FindName("TxtPageDescription")
$TxtVersion         = $window.FindName("TxtVersion")
$TxtVersion.Text    = "v$script:ToolVersion"

# Tab 1 — Profile Management (shared settings + right column)
$TxtMatch               = $window.FindName("TxtMatch")
$TxtLoadFilter          = $window.FindName("TxtLoadFilter")
$ChkSettingsCatalog     = $window.FindName("ChkSettingsCatalog")
$ChkDeviceConfig        = $window.FindName("ChkDeviceConfig")
$ChkAdministrativeTemplates = $window.FindName("ChkAdministrativeTemplates")
$ChkSucceededCount      = $window.FindName("ChkSucceededCount")
$ComboAssignmentFilter  = $window.FindName("ComboAssignmentFilter")
$ComboPlatformFilter    = $window.FindName("ComboPlatformFilter")


$BtnPreview             = $window.FindName("BtnPreview")
$BtnRenameReset         = $window.FindName("BtnRenameReset")
$TxtProfileFilter       = $window.FindName("TxtProfileFilter")
$TxtFilterCount         = $window.FindName("TxtFilterCount")
$ProfileEmptyState      = $window.FindName("ProfileEmptyState")
$TxtProfileEmptyTitle   = $window.FindName("TxtProfileEmptyTitle")
$TxtProfileEmptyDescription = $window.FindName("TxtProfileEmptyDescription")
$TxtProfileSelectionSummary = $window.FindName("TxtProfileSelectionSummary")
$TxtProfileSelectionHint = $window.FindName("TxtProfileSelectionHint")
$BtnRenameSelectAll     = $window.FindName("BtnRenameSelectAll")
$BtnRenameDeselectAll   = $window.FindName("BtnRenameDeselectAll")
$BtnExportCsv           = $window.FindName("BtnExportCsv")
$ListProfiles           = $window.FindName("ListProfiles")
$TxtProfileLog          = $window.FindName("TxtProfileLog")
$ProfileLogScroll       = $window.FindName("ProfileLogScroll")
$ProfileActivityExpander = $window.FindName("ProfileActivityExpander")
$TxtProfileActivitySummary = $window.FindName("TxtProfileActivitySummary")

# Loading overlay
$LoadingOverlay         = $window.FindName("LoadingOverlay")
$TxtOverlayMessage      = $window.FindName("TxtOverlayMessage")

# Tab 1 — Profile Management: Assignment sub-tab
$ComboAction          = $window.FindName("ComboAction")
$PanelAddReplace      = $window.FindName("PanelAddReplace")
$GroupRemove          = $window.FindName("GroupRemove")
$ComboIncludeExclude  = $window.FindName("ComboIncludeExclude")
$ComboTargetType      = $window.FindName("ComboTargetType")
$PanelGroupSearch     = $window.FindName("PanelGroupSearch")
$TxtGroupSearch       = $window.FindName("TxtGroupSearch")
$BtnSearchGroup       = $window.FindName("BtnSearchGroup")
$ListGroups           = $window.FindName("ListGroups")
$BtnQueueGroups       = $window.FindName("BtnQueueGroups")
$ListQueuedGroups     = $window.FindName("ListQueuedGroups")
$BtnUnqueueGroups     = $window.FindName("BtnUnqueueGroups")
$BtnClearQueuedGroups = $window.FindName("BtnClearQueuedGroups")
$TxtQueuedGroupCount  = $window.FindName("TxtQueuedGroupCount")
$ComboFilterMode      = $window.FindName("ComboFilterMode")
$PanelFilterList      = $window.FindName("PanelFilterList")
$TxtFilterSearch      = $window.FindName("TxtFilterSearch")
$ListFilters          = $window.FindName("ListFilters")
$ComboRemoveScope     = $window.FindName("ComboRemoveScope")
$PanelRemoveGroup     = $window.FindName("PanelRemoveGroup")
$TxtRemoveGroupSearch = $window.FindName("TxtRemoveGroupSearch")
$BtnSearchRemoveGroup = $window.FindName("BtnSearchRemoveGroup")
$ListRemoveGroups     = $window.FindName("ListRemoveGroups")
$BtnQueueRemoveGroups = $window.FindName("BtnQueueRemoveGroups")
$ListQueuedRemoveGroups = $window.FindName("ListQueuedRemoveGroups")
$BtnUnqueueRemoveGroups = $window.FindName("BtnUnqueueRemoveGroups")
$BtnClearQueuedRemoveGroups = $window.FindName("BtnClearQueuedRemoveGroups")
$TxtQueuedRemoveGroupCount = $window.FindName("TxtQueuedRemoveGroupCount")
$ComboRemoveType      = $window.FindName("ComboRemoveType")
$BtnAssignExecute     = $window.FindName("BtnAssignExecute")

# Tab 1 — Profile Management: Rename sub-tab
$TxtPrefix            = $window.FindName("TxtPrefix")
$TxtReplace           = $window.FindName("TxtReplace")
$TxtSuffix            = $window.FindName("TxtSuffix")
$BtnRenameExecute     = $window.FindName("BtnRenameExecute")
$ChkRemoveWhitespace  = $window.FindName("ChkRemoveWhitespace")
$BtnPreviewRename     = $window.FindName("BtnPreviewRename")

# Tab 1 — Profile Management: Delete sub-tab
$BtnDeleteExecute     = $window.FindName("BtnDeleteExecute")

# Tab 2 — Autopilot Group Tags
$MainTabs           = $window.FindName("MainTabs")
$script:ToolPageInfoByIndex = @{
    0 = @{ Title = 'Profile Management'; Description = 'Manage, rename, assign, export, and remove Intune configuration profiles.' }
    1 = @{ Title = 'Autopilot Group Tags'; Description = 'Look up devices by serial number or CSV, review profile assignments, and change group tags on selected devices.' }
    2 = @{ Title = 'Setting Search'; Description = 'Find settings across loaded Settings Catalog profiles.' }
    3 = @{ Title = 'Duplicate Devices'; Description = 'Review duplicate Windows records; protect keepers and manually remove eligible stale records.' }
    4 = @{ Title = 'Audit Lookup'; Description = 'Search historical Intune changes by initiating user and target name.' }
    5 = @{ Title = 'App Management'; Description = 'Load and filter Intune apps, bulk assign, and inspect or edit direct assignments.' }
}
$script:NavigationByIndex = @{}
foreach ($navigationName in @('NavProfileManagement', 'NavGroupTags', 'NavSettingSearch', 'NavDuplicateDevices', 'NavAuditLookup', 'NavAppManagement')) {
    $navigationItem = $window.FindName($navigationName)
    if (-not $navigationItem) { continue }
    $script:NavigationByIndex[[int]$navigationItem.Tag] = $navigationItem
    $navigationItem.Add_Checked({
        param($sender, $eventArgs)
        if ($sender.IsChecked -and $MainTabs.SelectedIndex -ne [int]$sender.Tag) {
            $MainTabs.SelectedIndex = [int]$sender.Tag
        }
    })
}
$MainTabs.Add_SelectionChanged({
    $pageInfo = $script:ToolPageInfoByIndex[[int]$MainTabs.SelectedIndex]
    if (-not $pageInfo) { return }
    $TxtPageTitle.Text = $pageInfo.Title
    $TxtPageDescription.Text = $pageInfo.Description
    $navigationItem = $script:NavigationByIndex[[int]$MainTabs.SelectedIndex]
    if ($navigationItem -and -not $navigationItem.IsChecked) { $navigationItem.IsChecked = $true }
})
$MainTabs.SelectedIndex = 0
$script:NavigationByIndex[0].IsChecked = $true
$TxtAPSerials       = $window.FindName("TxtAPSerials")
$BtnAPLookup        = $window.FindName("BtnAPLookup")
$BtnAPImportCsv     = $window.FindName("BtnAPImportCsv")
$TxtAPCsvFile       = $window.FindName("TxtAPCsvFile")
$TxtNewTag          = $window.FindName("TxtNewTag")
$TxtAPTagHint       = $window.FindName("TxtAPTagHint")
$BtnAPApplyTag      = $window.FindName("BtnAPApplyTag")
$BtnAPApplyCsvTags  = $window.FindName("BtnAPApplyCsvTags")
$BtnAPClearTag      = $window.FindName("BtnAPClearTag")
$BtnAPFreshStart    = $window.FindName("BtnAPFreshStart")
$ChkAPSync          = $window.FindName("ChkAPSync")
$TxtCountLoaded     = $window.FindName("TxtCountLoaded")
$TxtCountFound      = $window.FindName("TxtCountFound")
$TxtCountSelected   = $window.FindName("TxtCountSelected")
$AutopilotEmptyState = $window.FindName("AutopilotEmptyState")
$TxtAutopilotEmptyTitle = $window.FindName("TxtAutopilotEmptyTitle")
$TxtAutopilotEmptyDescription = $window.FindName("TxtAutopilotEmptyDescription")
$TxtAPDeviceFilter  = $window.FindName("TxtAPDeviceFilter")
$BtnAPSelectAll     = $window.FindName("BtnAPSelectAll")
$BtnAPDeselectAll   = $window.FindName("BtnAPDeselectAll")
$BtnAPRefresh       = $window.FindName("BtnAPRefresh")
$BtnAPExport        = $window.FindName("BtnAPExport")
$BtnAPClearResults  = $window.FindName("BtnAPClearResults")
$ListDevices        = $window.FindName("ListDevices")
$script:APActionButtons = @($BtnAPLookup, $BtnAPImportCsv, $BtnAPApplyTag, $BtnAPApplyCsvTags, $BtnAPClearTag,
    $BtnAPFreshStart, $BtnAPSelectAll, $BtnAPDeselectAll, $BtnAPRefresh, $BtnAPExport, $BtnAPClearResults)
$TxtAutopilotLog    = $window.FindName("TxtAutopilotLog")
$AutopilotLogScroll = $window.FindName("AutopilotLogScroll")
$AutopilotActivityExpander = $window.FindName("AutopilotActivityExpander")
$TxtAutopilotActivitySummary = $window.FindName("TxtAutopilotActivitySummary")

# Tab 3 — Setting Search
$TxtSettingSearchTerm    = $window.FindName("TxtSettingSearchTerm")
$BtnSettingSearchExecute = $window.FindName("BtnSettingSearchExecute")
$TxtSettingResultFilter  = $window.FindName("TxtSettingResultFilter")
$TxtSettingResultCount   = $window.FindName("TxtSettingResultCount")
$TxtSettingScopeTitle    = $window.FindName("TxtSettingScopeTitle")
$TxtSettingScopeCount    = $window.FindName("TxtSettingScopeCount")
$SettingSearchEmptyState = $window.FindName("SettingSearchEmptyState")
$TxtSettingEmptyTitle    = $window.FindName("TxtSettingEmptyTitle")
$TxtSettingEmptyDescription = $window.FindName("TxtSettingEmptyDescription")
$BtnSettingSearchExport  = $window.FindName("BtnSettingSearchExport")
$ListSettingResults      = $window.FindName("ListSettingResults")
$TxtSettingSearchLog     = $window.FindName("TxtSettingSearchLog")
$SettingSearchLogScroll  = $window.FindName("SettingSearchLogScroll")
$SettingSearchActivityExpander = $window.FindName("SettingSearchActivityExpander")
$TxtSettingActivitySummary = $window.FindName("TxtSettingActivitySummary")

$BtnDuplicateScan = $window.FindName('BtnDuplicateScan')
$BtnDuplicateExport = $window.FindName('BtnDuplicateExport')
$BtnDuplicateDelete = $window.FindName('BtnDuplicateDelete')
$BtnDuplicateCancel = $window.FindName('BtnDuplicateCancel')
$ComboDuplicateMatch = $window.FindName('ComboDuplicateMatch')
$TxtDuplicateFilter = $window.FindName('TxtDuplicateFilter')
$TxtDuplicateCount = $window.FindName('TxtDuplicateCount')
$ListDuplicateDevices = $window.FindName('ListDuplicateDevices')
$DuplicateActivityExpander = $window.FindName('DuplicateActivityExpander')
$TxtDuplicateActivitySummary = $window.FindName('TxtDuplicateActivitySummary')
$DuplicateLogScroll = $window.FindName('DuplicateLogScroll')
$TxtDuplicateLog = $window.FindName('TxtDuplicateLog')

# Start disconnected; operations acquire delegated scopes as needed.
$script:IsConnected = $false
$script:GraphUser = ''
$script:GraphTenantId = ''
$TxtTenantName.Text = ''
$TxtTenantName.ToolTip = ''
$TxtGraphConnectionStatus.Text = 'Not connected'
$BtnSignInOut.Content = 'Sign in'
$EllipseStatus.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'DangerBrush')

#endregion Build Window

#region Helpers

function Update-ActivitySummary {
    param(
        [System.Windows.Controls.Expander]$Expander,
        [System.Windows.Controls.TextBlock]$Summary,
        [string]$Timestamp,
        [string]$Message,
        [string]$Level
    )

    $status = switch ($Level) {
        'OK'   { @{ Glyph = '✓'; BrushKey = 'SuccessBrush' } }
        'WARN' { @{ Glyph = '⚠'; BrushKey = 'AttentionBrush' } }
        'ERR'  { @{ Glyph = '✕'; BrushKey = 'DangerBrush' } }
        default { @{ Glyph = 'ⓘ'; BrushKey = 'SubTextBrush' } }
    }

    $Summary.Text = ''
    $Summary.Inlines.Clear()
    $timeRun = [System.Windows.Documents.Run]::new("$Timestamp  ")
    $timeRun.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'SubTextBrush')
    $Summary.Inlines.Add($timeRun)
    $statusRun = [System.Windows.Documents.Run]::new("$($status.Glyph)  ")
    $statusRun.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe UI Symbol')
    $statusRun.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, $status.BrushKey)
    $Summary.Inlines.Add($statusRun)
    $messageRun = [System.Windows.Documents.Run]::new($Message)
    $messageRun.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'TextBrush')
    $Summary.Inlines.Add($messageRun)
    $Summary.ToolTip = "$Timestamp [$Level] $Message"
    if ($Level -in @('WARN', 'ERR')) { $Expander.IsExpanded = $true }
}

function Write-ActivityEntry {
    param(
        [System.Windows.Controls.RichTextBox]$LogBox,
        [System.Windows.Controls.ScrollViewer]$LogScroll,
        [System.Windows.Controls.Expander]$Expander,
        [System.Windows.Controls.TextBlock]$Summary,
        [string]$Message,
        [string]$Level
    )

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $status = switch ($Level) {
        'OK'   { @{ Glyph = '✓'; BrushKey = 'SuccessBrush' } }
        'WARN' { @{ Glyph = '⚠'; BrushKey = 'AttentionBrush' } }
        'ERR'  { @{ Glyph = '✕'; BrushKey = 'DangerBrush' } }
        default { @{ Glyph = 'ⓘ'; BrushKey = 'SubTextBrush' } }
    }

    $LogBox.Dispatcher.Invoke([action]{
        $paragraph = [System.Windows.Documents.Paragraph]::new()
        $paragraph.Margin = [System.Windows.Thickness]::new(0)
        $timeRun = [System.Windows.Documents.Run]::new("$timestamp  ")
        $timeRun.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'SubTextBrush')
        $paragraph.Inlines.Add($timeRun)
        $statusRun = [System.Windows.Documents.Run]::new("$($status.Glyph)  ")
        $statusRun.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe UI Symbol')
        $statusRun.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, $status.BrushKey)
        $paragraph.Inlines.Add($statusRun)
        $messageRun = [System.Windows.Documents.Run]::new($Message)
        $messageRun.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'TextBrush')
        $paragraph.Inlines.Add($messageRun)
        $LogBox.Document.Blocks.Add($paragraph)
        $LogScroll.ScrollToBottom()
        Update-ActivitySummary -Expander $Expander -Summary $Summary -Timestamp $timestamp -Message $Message -Level $Level
    })
}

function Write-ProfileLog {
    param([string]$Message, [string]$Level = 'INFO')
    Write-ActivityEntry -LogBox $TxtProfileLog -LogScroll $ProfileLogScroll `
        -Expander $ProfileActivityExpander -Summary $TxtProfileActivitySummary -Message $Message -Level $Level
}

function Write-AssignLog {
    param([string]$Message, [string]$Level = "INFO")
    Write-ProfileLog -Message $Message -Level $Level
}

function Write-AutopilotLog {
    param([string]$Message, [string]$Level = 'INFO')
    Write-ActivityEntry -LogBox $TxtAutopilotLog -LogScroll $AutopilotLogScroll `
        -Expander $AutopilotActivityExpander -Summary $TxtAutopilotActivitySummary -Message $Message -Level $Level
}


function Get-SelectedTag {
    param($ComboBox)
    $item = $ComboBox.SelectedItem
    if ($item -and $item.Tag) { return $item.Tag }
    return $null
}

function Show-Panel {
    param($Panel, [bool]$Visible)
    $Panel.Visibility = if ($Visible) { 'Visible' } else { 'Collapsed' }
}

function Show-Overlay {
    param([string]$Message = 'Working...')
    $TxtOverlayMessage.Text    = $Message
    $LoadingOverlay.Visibility = 'Visible'
}

function Hide-Overlay {
    $LoadingOverlay.Visibility = 'Collapsed'
}

function Update-TenantInfo {
    $ctx = Get-MgContext
    if ($ctx) {
        $script:IsConnected = $true
        $tid  = $ctx.TenantId
        if ($script:AppState -and $script:AppState.Tenant -and $script:AppState.Tenant -ne $tid) { Clear-AppManagement }
        if ($script:AppState) { $script:AppState.Tenant = [string]$tid }
        $name = try {
            $org = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/organization" -Method Get
            $org.value[0].displayName
        } catch {
            if ($ctx.Account) { $ctx.Account } else { "Connected" }
        }
        $TxtTenantName.Text   = $name
        $script:GraphUser = if ($ctx.Account) { [string]$ctx.Account } else { '' }
        $script:GraphTenantId = [string]$tid
        $script:GraphState.ExpectedTenant = [string]$tid
        $script:GraphState.Tenant = [string]$tid
        $script:GraphState.Account = [string]$ctx.Account
        $script:GraphState.Connected = $true
        $TxtTenantName.ToolTip = "User: $script:GraphUser`nTenant ID: $script:GraphTenantId"
        $TxtGraphConnectionStatus.Text = "Connected"
        $BtnSignInOut.Content = "Sign out"
        $EllipseStatus.Fill   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#16a34a'))
    } else {
        $script:IsConnected   = $false
        $script:GraphState.Connected = $false
        $TxtTenantName.Text   = ""
        $TxtTenantName.ToolTip = ''
        $script:GraphUser = ''
        $script:GraphTenantId = ''
        $TxtGraphConnectionStatus.Text = "Not connected"
        $BtnSignInOut.Content = "Sign in"
        $EllipseStatus.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'DangerBrush')
    }
}

function Show-APMessage {
    param([string]$Message, [string]$Title, [string]$Icon = 'Information', [string]$Buttons = 'OK', [string]$DefaultResult = 'None')
    if ($ValidateOnly) { return 'Yes' }
    return [string][System.Windows.MessageBox]::Show($window, $Message, $Title,
        [System.Windows.MessageBoxButton]$Buttons, [System.Windows.MessageBoxImage]$Icon, [System.Windows.MessageBoxResult]$DefaultResult)
}

function Get-APNormalizedSerial {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToUpperInvariant()
}

function Get-APRowKey {
    param($Row)
    $key = Get-APNormalizedSerial $Row.RequestedSerial
    if ($key) { return $key }
    return "#invalid:$($Row.Details)"
}

function Get-APSerialInput {
    param([string]$Text)
    $seen = @{}
    foreach ($value in @($Text -split '[\s,;]+')) {
        $serial = $value.Trim().Trim('"').Trim()
        if (-not $serial) { continue }
        $key = Get-APNormalizedSerial $serial
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $serial
    }
}

function Test-APNewTagInput {
    param([string]$Tag)
    $value = ([string]$Tag).Trim()
    if (-not $value) { return 'Enter a group tag. Use Clear tag to remove tags.' }
    if ($value.Length -gt $script:APMaxGroupTagLength) { return "Group tag exceeds $script:APMaxGroupTagLength characters ($($value.Length))." }
    if ($value -match '[\x00-\x1F\x7F]') { return 'Group tag contains control characters.' }
    return ''
}

function Merge-APDeviceRows {
    param([object[]]$NewRows, [switch]$ReplaceAll)
    if ($ReplaceAll) { $script:APRows.Clear() }
    foreach ($newRow in @($NewRows | Where-Object { $_ -and $_.RecordType -eq 'Device' })) {
        $key = Get-APRowKey $newRow
        $replaced = $false
        for ($i = 0; $i -lt $script:APRows.Count; $i++) {
            if ((Get-APRowKey $script:APRows[$i]) -eq $key) { $script:APRows[$i] = $newRow; $replaced = $true; break }
        }
        if (-not $replaced) { $script:APRows.Add($newRow) }
    }
}

function Get-APSelectedEligibleRows {
    @($ListDevices.SelectedItems | Where-Object { $_.Eligible -and $_.LookupStatus -eq 'Found' -and $_.AutopilotId })
}

function Get-APSelectedFreshStartRows {
    @(Get-APSelectedEligibleRows | Where-Object {
        $_.ManagedDeviceId -and $_.ManagedDeviceId -ne '00000000-0000-0000-0000-000000000000' -and
        $_.UpdateStatus -ne 'Pending verification' -and $_.UpdateStatus -notlike 'Failed*' -and
        $_.UpdateStatus -notlike 'Stale*' -and $_.FreshStartStatus -ne 'Requested'
    })
}

function Update-APDeviceView {
    $term = $TxtAPDeviceFilter.Text.Trim()
    $ListDevices.SelectedItems.Clear()
    $ListDevices.Items.Clear()
    foreach ($row in $script:APRows) {
        if ($term) {
            $haystack = @($row.RequestedSerial, $row.DeviceName, $row.GroupTag, $row.CsvTag, $row.ProfileName,
                $row.LookupStatus, $row.UpdateStatus, $row.FreshStartStatus, $row.Details) -join "`n"
            if ($haystack.IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        }
        [void]$ListDevices.Items.Add($row)
    }
    Update-APCounters
}

function Update-APCounters {
    $selected = @($ListDevices.SelectedItems)
    $eligibleSelected = @(Get-APSelectedEligibleRows)
    $csvReady = @($eligibleSelected | Where-Object { -not $_.CsvIssue -and -not [string]::IsNullOrWhiteSpace($_.CsvTag) })
    $visibleEligible = @($ListDevices.Items | Where-Object { $_.Eligible })
    $TxtCountLoaded.Text = $script:APRows.Count.ToString()
    $TxtCountFound.Text = @($script:APRows | Where-Object { $_.Eligible }).Count.ToString()
    $TxtCountSelected.Text = $selected.Count.ToString()

    if ($script:APRows.Count -eq 0) {
        $TxtAutopilotEmptyTitle.Text = 'No devices loaded'
        $TxtAutopilotEmptyDescription.Text = 'Look up a serial number or import a CSV file to begin.'
        $AutopilotEmptyState.Visibility = 'Visible'
    } elseif ($ListDevices.Items.Count -eq 0) {
        $TxtAutopilotEmptyTitle.Text = 'No matching devices'
        $TxtAutopilotEmptyDescription.Text = 'Change the filter to show loaded devices.'
        $AutopilotEmptyState.Visibility = 'Visible'
    } else {
        $AutopilotEmptyState.Visibility = 'Collapsed'
    }

    $tagError = Test-APNewTagInput $TxtNewTag.Text
    if ($TxtNewTag.Text -and $tagError) {
        $TxtAPTagHint.Text = $tagError
        $TxtAPTagHint.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AttentionBrush')
    } else {
        $TxtAPTagHint.Text = 'Applies one tag to every selected device.'
        $TxtAPTagHint.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubTextBrush')
    }

    # Busy buttons keep their temporary label until Invoke-InBackground restores them.
    $busy = $script:BusyButtonCounts -and $script:BusyButtonCounts.ContainsKey($BtnAPLookup.Name)
    if (-not $busy) {
        $count = $eligibleSelected.Count
        $BtnAPApplyTag.Content = if ($count) { "Apply to $count selected" } else { 'Apply to selected' }
        $BtnAPApplyCsvTags.Content = if ($csvReady.Count) { "Apply CSV tags to $($csvReady.Count) selected" } else { 'Apply CSV tags to selected' }
        $BtnAPClearTag.Content = if ($count) { "Clear tag on $count selected" } else { 'Clear tag on selected' }
    }

    Set-ButtonAvailability -Button $BtnAPLookup -Available (@(Get-APSerialInput $TxtAPSerials.Text).Count -gt 0)
    Set-ButtonAvailability -Button $BtnAPImportCsv -Available $true
    Set-ButtonAvailability -Button $BtnAPApplyTag -Available ($eligibleSelected.Count -gt 0 -and -not $tagError)
    Set-ButtonAvailability -Button $BtnAPApplyCsvTags -Available ($csvReady.Count -gt 0)
    Set-ButtonAvailability -Button $BtnAPClearTag -Available ($eligibleSelected.Count -gt 0)
    Set-ButtonAvailability -Button $BtnAPFreshStart -Available (@(Get-APSelectedFreshStartRows).Count -gt 0)
    Set-ButtonAvailability -Button $BtnAPSelectAll -Available ($visibleEligible.Count -gt 0)
    Set-ButtonAvailability -Button $BtnAPDeselectAll -Available ($selected.Count -gt 0)
    Set-ButtonAvailability -Button $BtnAPRefresh -Available ($script:APRows.Count -gt 0)
    Set-ButtonAvailability -Button $BtnAPExport -Available ($script:APRows.Count -gt 0)
    Set-ButtonAvailability -Button $BtnAPClearResults -Available ($script:APRows.Count -gt 0)
    $TxtAPSerials.IsEnabled = -not $busy
    $TxtNewTag.IsEnabled = -not $busy
    $ChkAPSync.IsEnabled = -not $busy
    $TxtAPDeviceFilter.IsEnabled = -not $busy
}

function Get-APUpdatePlan {
    param([object[]]$Rows, [ValidateSet('ApplyTag', 'ApplyCsvTags', 'ClearTag')][string]$Mode, [string]$Tag)
    $changes = [System.Collections.Generic.List[object]]::new()
    $skipped = [System.Collections.Generic.List[string]]::new()
    foreach ($row in $Rows) {
        $target = switch ($Mode) { 'ApplyTag' { $Tag.Trim() } 'ClearTag' { '' } 'ApplyCsvTags' { ([string]$row.CsvTag).Trim() } }
        if ($Mode -eq 'ApplyCsvTags' -and $row.CsvIssue) { $skipped.Add("$($row.RequestedSerial): CSV issue"); continue }
        if ($Mode -eq 'ApplyCsvTags' -and -not $target) { $skipped.Add("$($row.RequestedSerial): no CSV tag"); continue }
        if ([string]$row.GroupTag -ceq $target) { $skipped.Add("$($row.RequestedSerial): already set"); continue }
        $changes.Add([PSCustomObject]@{ Row = $row; Serial = $row.RequestedSerial; OldTag = [string]$row.GroupTag; NewTag = $target })
    }
    [PSCustomObject]@{ Changes = $changes.ToArray(); Skipped = $skipped.ToArray() }
}

function Export-APDeviceRows {
    param([Parameter(Mandatory)][string]$Path)
    $columns = 'RequestedSerial', 'SerialNumber', 'DeviceName', 'GroupTag', 'CsvTag', 'CsvIssue', 'ProfileName',
        'ProfileAssignmentStatus', 'ProfileAssignedDate', 'ProfileAssignedDateTimeUtc', 'LookupStatus', 'ManagedDeviceStatus',
        'UpdateStatus', 'FreshStartStatus', 'PreviousTag', 'RequestedTag', 'LastAction', 'Details', 'AutopilotId', 'ManagedDeviceId',
        'AutopilotDisplayName', 'Model', 'Manufacturer'
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $script:APRows | Select-Object $columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding $encoding
}

function Test-AssignmentReady {
    if ($ListProfiles.SelectedItems.Count -eq 0) { return $false }
    $action = Get-SelectedTag $ComboAction
    if ($action -in @('Add', 'Replace')) {
        $targetType = Get-SelectedTag $ComboTargetType
        $filterMode = Get-SelectedTag $ComboFilterMode
        if (-not $targetType -or -not $filterMode -or -not (Get-SelectedTag $ComboIncludeExclude)) { return $false }
        if ($targetType -eq 'Group' -and $ListQueuedGroups.Items.Count -eq 0) { return $false }
        if ($filterMode -ne 'None' -and -not $ListFilters.SelectedItem) { return $false }
        return $true
    }
    if ($action -eq 'Remove') {
        $removeScope = Get-SelectedTag $ComboRemoveScope
        if (-not $removeScope) { return $false }
        if ($removeScope -eq 'Group' -and ($ListQueuedRemoveGroups.Items.Count -eq 0 -or -not (Get-SelectedTag $ComboRemoveType))) { return $false }
        return $true
    }
    return $false
}

function Update-ProfileCounters {
    $loadedCount = $script:AllProfiles.Count
    $visibleCount = $ListProfiles.Items.Count
    $selectedCount = $ListProfiles.SelectedItems.Count
    $profileNoun = if ($selectedCount -eq 1) { 'profile' } else { 'profiles' }
    $loadedNoun = if ($loadedCount -eq 1) { 'profile' } else { 'profiles' }
    $filterActive = -not [string]::IsNullOrWhiteSpace($TxtProfileFilter.Text)

    $TxtFilterCount.Text = if ($filterActive) {
        "$visibleCount shown / $loadedCount loaded  •  $selectedCount selected"
    } else {
        "$loadedCount $loadedNoun loaded  •  $selectedCount selected"
    }
    $TxtProfileSelectionSummary.Text = if ($selectedCount -gt 0) {
        "$selectedCount $profileNoun selected"
    } else {
        'No profiles selected'
    }
    $TxtProfileSelectionHint.Text = if ($selectedCount -gt 0) {
        'Actions apply only to the selected profiles.'
    } else {
        'Select one or more profiles to manage assignments, rename, or delete them.'
    }

    if ($loadedCount -eq 0) {
        $TxtProfileEmptyTitle.Text = 'No profiles loaded'
        $TxtProfileEmptyDescription.Text = 'Choose the filters above and select Load profiles to begin.'
        $ProfileEmptyState.Visibility = 'Visible'
    } elseif ($visibleCount -eq 0) {
        $TxtProfileEmptyTitle.Text = 'No matching profiles'
        $TxtProfileEmptyDescription.Text = 'Change the profile list filter to show loaded profiles.'
        $ProfileEmptyState.Visibility = 'Visible'
    } else {
        $ProfileEmptyState.Visibility = 'Collapsed'
    }

    $hasSelection = $selectedCount -gt 0
    Set-ButtonAvailability -Button $BtnRenameSelectAll -Available ($visibleCount -gt 0)
    Set-ButtonAvailability -Button $BtnRenameDeselectAll -Available $hasSelection
    Set-ButtonAvailability -Button $BtnExportCsv -Available ($loadedCount -gt 0)
    Set-ButtonAvailability -Button $BtnAssignExecute -Available (Test-AssignmentReady)
    Set-ButtonAvailability -Button $BtnQueueGroups -Available ($ListGroups.SelectedItems.Count -gt 0)
    Set-ButtonAvailability -Button $BtnUnqueueGroups -Available ($ListQueuedGroups.SelectedItems.Count -gt 0)
    Set-ButtonAvailability -Button $BtnClearQueuedGroups -Available ($ListQueuedGroups.Items.Count -gt 0)
    Set-ButtonAvailability -Button $BtnQueueRemoveGroups -Available ($ListRemoveGroups.SelectedItems.Count -gt 0)
    Set-ButtonAvailability -Button $BtnUnqueueRemoveGroups -Available ($ListQueuedRemoveGroups.SelectedItems.Count -gt 0)
    Set-ButtonAvailability -Button $BtnClearQueuedRemoveGroups -Available ($ListQueuedRemoveGroups.Items.Count -gt 0)
    $TxtQueuedGroupCount.Text = "$($ListQueuedGroups.Items.Count) group(s) queued"
    $TxtQueuedRemoveGroupCount.Text = "$($ListQueuedRemoveGroups.Items.Count) group(s) queued"
    $hasRenameOptions = -not [string]::IsNullOrWhiteSpace($TxtMatch.Text) -or
                        -not [string]::IsNullOrWhiteSpace($TxtPrefix.Text) -or
                        -not [string]::IsNullOrWhiteSpace($TxtSuffix.Text) -or
                        [bool]$ChkRemoveWhitespace.IsChecked
    Set-ButtonAvailability -Button $BtnPreviewRename -Available ($hasSelection -and $hasRenameOptions)
    Set-ButtonAvailability -Button $BtnRenameExecute -Available ($hasSelection -and $hasRenameOptions)
    Set-ButtonAvailability -Button $BtnDeleteExecute -Available $hasSelection
    $BtnAssignExecute.Content = if ($hasSelection) { "Apply assignment to $selectedCount $profileNoun" } else { 'Apply assignment' }
    $BtnRenameExecute.Content = if ($hasSelection) { "Rename $selectedCount $profileNoun" } else { 'Rename selected' }
    $BtnDeleteExecute.Content = if ($hasSelection) { "Delete $selectedCount $profileNoun" } else { 'Delete profiles' }
    Update-SettingSearchScope
}

function Update-SettingSearchScope {
    $selectedCount = @($ListProfiles.SelectedItems | Where-Object { $_.ApiType -eq 'configurationPolicies' }).Count
    if ($ListProfiles.SelectedItems.Count -gt 0) {
        $TxtSettingScopeTitle.Text = 'Selected profiles'
    } else {
        $loadedCount = @($script:AllProfiles | Where-Object { $_.ApiType -eq 'configurationPolicies' }).Count
        $TxtSettingScopeTitle.Text = 'All loaded profiles'
    }
    $scopeCount = if ($ListProfiles.SelectedItems.Count -gt 0) { $selectedCount } else { $loadedCount }
    $profileNoun = if ($scopeCount -eq 1) { 'profile' } else { 'profiles' }
    $TxtSettingScopeCount.Text = "$scopeCount Settings Catalog $profileNoun"
    $hasSearchTerm = -not [string]::IsNullOrWhiteSpace($TxtSettingSearchTerm.Text)
    Set-ButtonAvailability -Button $BtnSettingSearchExecute -Available (($scopeCount -gt 0) -and $hasSearchTerm)
}

function Update-SettingSearchResultState {
    $totalCount = $script:AllSearchResults.Count
    $visibleCount = $ListSettingResults.Items.Count
    $filterActive = -not [string]::IsNullOrWhiteSpace($TxtSettingResultFilter.Text)
    $resultNoun = if ($totalCount -eq 1) { 'result' } else { 'results' }
    $TxtSettingResultCount.Text = if ($filterActive) {
        "$visibleCount / $totalCount $resultNoun"
    } else {
        "$totalCount $resultNoun"
    }
    Set-ButtonAvailability -Button $BtnSettingSearchExport -Available ($totalCount -gt 0)

    if ($visibleCount -gt 0) {
        $SettingSearchEmptyState.Visibility = 'Collapsed'
    } elseif (-not $script:HasSearchedSettings) {
        $TxtSettingEmptyTitle.Text = 'No results'
        $TxtSettingEmptyDescription.Text = 'Enter a setting name or definition ID to search across loaded Settings Catalog profiles.'
        $SettingSearchEmptyState.Visibility = 'Visible'
    } elseif ($totalCount -eq 0) {
        $TxtSettingEmptyTitle.Text = 'No matching settings found'
        $TxtSettingEmptyDescription.Text = 'Try another setting name or definition ID.'
        $SettingSearchEmptyState.Visibility = 'Visible'
    } elseif ($visibleCount -eq 0) {
        $TxtSettingEmptyTitle.Text = 'No matching results'
        $TxtSettingEmptyDescription.Text = 'Change the result filter to show search results.'
        $SettingSearchEmptyState.Visibility = 'Visible'
    }
}

function Set-ButtonAvailability {
    param(
        [Parameter(Mandatory)][System.Windows.Controls.Button]$Button,
        [Parameter(Mandatory)][bool]$Available
    )

    $isBusy = $script:BusyButtonCounts.ContainsKey($Button.Name)
    $Button.IsEnabled = $Available -and -not $isBusy
}

function Receive-BackgroundOutput {
    # Runs on the UI thread: drains log entries and result items queued by a background worker.
    param($LogQueue, $OutputQueue, $OutputAction, $LogBox, $LogScroll, $Expander, $Summary)
    $entry = $null
    if ($LogQueue -and $LogBox) {
        while ($LogQueue.TryDequeue([ref]$entry)) {
            if ($LogBox.Name -eq 'TxtDuplicateLog' -and $script:DuplicateState.Busy -and $script:DuplicateState.Cancel) {
                $script:DuplicateState.Message = [string]$entry.Message
            }
            Write-ActivityEntry -LogBox $LogBox -LogScroll $LogScroll -Expander $Expander -Summary $Summary `
                -Message ([string]$entry.Message) -Level ([string]$entry.Level)
        }
    }
    $item = $null
    if ($OutputQueue) {
        while ($OutputQueue.TryDequeue([ref]$item)) {
            if (-not $OutputAction) { continue }
            try { & $OutputAction $item }
            catch {
                if ($LogBox) {
                    Write-ActivityEntry -LogBox $LogBox -LogScroll $LogScroll -Expander $Expander -Summary $Summary `
                        -Message "Could not process a result: $($_.Exception.Message)" -Level 'ERR'
                }
            }
        }
    }
}

function Invoke-InBackground {
    <#
    .SYNOPSIS
        Runs $Worker on a background thread so the WPF UI stays responsive.
        Disables $Buttons for the run, labels the primary button with $BusyLabel,
        then re-enables everything when done.

        $Worker is a feature descriptor from New-IntuneToolboxFeatureWorker, or a self-contained
        script block (recreated from its text in the background runspace, so closures are not
        available). Script-block workers receive $WriteLog to push log lines to the UI.
    #>
    param(
        [Parameter(Mandatory)]$Worker,
        [Parameter(Mandatory)]$Buttons,
        [Parameter(Mandatory)][string]$BusyLabel,
        [Parameter(Mandatory)]$LogBox,
        [Parameter(Mandatory)]$LogScroll,
        [scriptblock]$OnComplete  = $null,
        [bool]$DarkMode           = $false,
        [string]$OverlayMessage   = '',
        # Optional: receives a Stop scriptblock that cancels the running worker.
        [hashtable]$CancelHandle  = $null
    )

    $activityExpander = $null
    $activitySummary = $null
    switch ($LogBox.Name) {
        'TxtProfileLog' {
            $activityExpander = $ProfileActivityExpander
            $activitySummary = $TxtProfileActivitySummary
        }
        'TxtAutopilotLog' {
            $activityExpander = $AutopilotActivityExpander
            $activitySummary = $TxtAutopilotActivitySummary
        }
        'TxtSettingSearchLog' {
            $activityExpander = $SettingSearchActivityExpander
            $activitySummary = $TxtSettingActivitySummary
        }
        'TxtDuplicateLog' {
            $activityExpander = $DuplicateActivityExpander
            $activitySummary = $TxtDuplicateActivitySummary
        }
        'TxtAuditLog' {
            $activityExpander = $AuditActivityExpander
            $activitySummary = $TxtAuditActivitySummary
        }
        'TxtAppLog' {
            $activityExpander = $AppUI.AppActivityExpander
            $activitySummary = $AppUI.TxtAppActivitySummary
        }
    }

    # Show overlay if a message was provided
    if ($OverlayMessage) { $LoadingOverlay.Dispatcher.Invoke([action]{ $TxtOverlayMessage.Text = $OverlayMessage; $LoadingOverlay.Visibility = 'Visible' }) }

    # Disable buttons and update label
    $savedContent = @{}
    $savedEnabled = @{}
    foreach ($b in $Buttons) {
        $savedContent[$b] = $b.Content
        $savedEnabled[$b] = $b.IsEnabled
        if (-not $script:BusyButtonCounts.ContainsKey($b.Name)) { $script:BusyButtonCounts[$b.Name] = 0 }
        $script:BusyButtonCounts[$b.Name]++
        $b.IsEnabled = $false
    }
    $Buttons[0].Content = $BusyLabel

    # The background runspace receives only plain data and thread-safe queues. Script blocks created
    # on the UI runspace (worker factories, output actions, log writers) must never run on the
    # background thread: in PowerShell 7 that can stall the worker before it logs anything. Logs and
    # output items are drained by the UI-thread timer below.
    $logQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $outputQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $job = @{ LogQueue = $logQueue; OutputQueue = $outputQueue }
    $outputAction = $null
    if ($Worker -is [scriptblock]) {
        $job.WorkerText = $Worker.ToString()
    } else {
        foreach ($key in @('FeaturePath', 'RunnerPath', 'Parameters', 'Scopes', 'State', 'PermissionNote', 'PermissionError')) {
            $job[$key] = $Worker.$key
        }
        $outputAction = $Worker.OutputAction
    }

    $ps   = [System.Management.Automation.PowerShell]::Create()
    $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 1)
    $pool.Open()
    $ps.RunspacePool = $pool

    # Here-string keeps this code out of the outer script's parse tree and its runspace.
    $workerScript = @'
Param($job)
$logQueue = $job.LogQueue
$outputQueue = $job.OutputQueue
$WriteLog = {
    param([string]$Msg, [string]$Lvl = 'INFO')
    $logQueue.Enqueue([PSCustomObject]@{ Message = $Msg; Level = $Lvl })
}.GetNewClosure()
if ($job.WorkerText) {
    & ([scriptblock]::Create($job.WorkerText)) $WriteLog
    return
}
if ($job.PermissionError) { & $WriteLog "Microsoft Graph permission request failed: $($job.PermissionError)" 'ERR'; return }
if ($job.PermissionNote) { & $WriteLog $job.PermissionNote 'OK' }
$featureParameters = @{}
foreach ($key in @($job.Parameters.Keys)) { $featureParameters[$key] = $job.Parameters[$key] }
$featureParameters.SkipGraphConnect = $true
$featureParameters.PassThru = $true
$featureParameters.LogAction = $WriteLog
try {
    $items = @(& $job.RunnerPath -Path $job.FeaturePath -Parameters $featureParameters -NonInteractive `
        -Modules @('Microsoft.Graph.Authentication') -Scopes $job.Scopes -State $job.State -LogAction $WriteLog)
    foreach ($item in $items) { $outputQueue.Enqueue($item) }
} catch {
    & $WriteLog "Feature failed: $($_.Exception.Message)" 'ERR'
}
'@
    [void]$ps.AddScript($workerScript)
    [void]$ps.AddParameter('job', $job)

    $handle = $ps.BeginInvoke()
    if ($CancelHandle) {
        $CancelHandle.Stop = { if (-not $handle.IsCompleted) { [void]$ps.BeginStop($null, $null) } }.GetNewClosure()
    }

    # Poll on a DispatcherTimer so we can update the UI thread when done
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [timespan]::FromMilliseconds(200)
    $capturedButtons    = $Buttons
    $capturedSaved      = $savedContent
    $capturedEnabled    = $savedEnabled
    $capturedOnComplete = $OnComplete
    $capturedOverlay    = [bool]$OverlayMessage
    $capturedOverlayCtl = $LoadingOverlay
    $capturedBusyButtonCounts = $script:BusyButtonCounts
    $capturedGraphState = $script:GraphState
    $capturedIsConnected = Get-Variable -Name IsConnected -Scope Script
    $capturedGraphUser = Get-Variable -Name GraphUser -Scope Script
    $capturedGraphTenantId = Get-Variable -Name GraphTenantId -Scope Script
    $capturedTenantName = $TxtTenantName
    $capturedConnectionStatus = $TxtGraphConnectionStatus
    $capturedSignInOut = $BtnSignInOut
    $capturedStatusIndicator = $EllipseStatus
    $capturedCancelHandle = $CancelHandle
    $capturedQueues = @{
        LogQueue = $logQueue; OutputQueue = $outputQueue; OutputAction = $outputAction
        LogBox = $LogBox; LogScroll = $LogScroll; Expander = $activityExpander; Summary = $activitySummary
    }
    $capturedDuplicateState = $script:DuplicateState
    $capturedDuplicateCount = $TxtDuplicateCount
    $timer.Add_Tick({
        Receive-BackgroundOutput @capturedQueues
        if ($capturedQueues.LogBox.Name -eq 'TxtDuplicateLog' -and $capturedDuplicateState.Busy -and
            $capturedDuplicateState.Cancel -and $capturedDuplicateState.ScanTimer) {
            $capturedDuplicateCount.Text = "$($capturedDuplicateState.Message) Elapsed: $([math]::Floor($capturedDuplicateState.ScanTimer.Elapsed.TotalSeconds)) s."
        }
        if ($handle.IsCompleted) {
            $timer.Stop()
            if ($capturedCancelHandle) { $capturedCancelHandle.Stop = $null }
            try   { $ps.EndInvoke($handle) } catch {
                if ($capturedQueues.LogBox) {
                    $capturedQueues.LogQueue.Enqueue([PSCustomObject]@{ Message = "Background task stopped: $($_.Exception.Message)"; Level = 'WARN' })
                }
            }
            Receive-BackgroundOutput @capturedQueues
            $ps.Dispose(); $pool.Close(); $pool.Dispose()
            foreach ($b in $capturedButtons) {
                $b.IsEnabled = $capturedEnabled[$b]
                $b.Content   = $capturedSaved[$b]
                $capturedBusyButtonCounts[$b.Name]--
                if ($capturedBusyButtonCounts[$b.Name] -le 0) { [void]$capturedBusyButtonCounts.Remove($b.Name) }
            }
            if ($capturedOverlay) { $capturedOverlayCtl.Visibility = 'Collapsed' }
            if ($capturedOnComplete) { & $capturedOnComplete }
            Update-ProfileCounters
            Update-APCounters
            Update-SettingSearchResultState
            Update-DuplicateDeviceState
            if ($capturedGraphState.Connected) {
                $capturedIsConnected.Value = $true
                $capturedGraphUser.Value = [string]$capturedGraphState.Account
                $capturedGraphTenantId.Value = [string]$capturedGraphState.Tenant
                $capturedTenantName.Text = if ($capturedGraphUser.Value) { $capturedGraphUser.Value } else { $capturedGraphTenantId.Value }
                $capturedTenantName.ToolTip = "User: $($capturedGraphUser.Value)`nTenant ID: $($capturedGraphTenantId.Value)"
                $capturedConnectionStatus.Text = 'Connected'
                $capturedSignInOut.Content = 'Sign out'
                $capturedStatusIndicator.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#16a34a'))
            }
        }
    }.GetNewClosure())
    $timer.Start()
}

function Get-IntuneToolboxMissingScopes {
    param([string[]]$Scopes, $Context)
    # Mirrors the scope check in Shared\Invoke-IntuneToolboxOperation.ps1: *.ReadWrite.All covers *.Read.All.
    @($Scopes | Where-Object {
        -not $Context -or ($_ -notin $Context.Scopes -and
            -not ($_ -like '*.Read.All' -and ($_ -replace '\.Read\.All$', '.ReadWrite.All') -in $Context.Scopes))
    })
}

function Request-IntuneToolboxScopes {
    <#
    .SYNOPSIS
        Requests missing Graph permissions on the UI thread so the Microsoft sign-in
        window has a visible owner. Returns a note when permissions were added, '' when
        nothing was needed, and throws when sign-in fails or is cancelled.
    #>
    param([Parameter(Mandatory)][string[]]$Scopes)
    $context = Get-MgContext
    $missing = @(Get-IntuneToolboxMissingScopes -Scopes $Scopes -Context $context)
    if ($missing.Count -eq 0) { return '' }
    $tenant = if ($script:GraphState.ExpectedTenant) { [string]$script:GraphState.ExpectedTenant }
        elseif ($context -and $context.TenantId) { [string]$context.TenantId }
        else { [string]$script:GraphState.RequestedTenant }
    $previousStatus = $TxtGraphConnectionStatus.Text
    $TxtGraphConnectionStatus.Text = "Waiting for Microsoft sign-in ($($missing -join ', '))..."
    # Let WPF render the status before the blocking sign-in call.
    $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $requested = @(@($context.Scopes) + $Scopes | Where-Object { $_ } | Select-Object -Unique)
        $connectParameters = @{ Scopes = $requested; NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
        if ($tenant) { $connectParameters.TenantId = $tenant }
        Connect-MgGraph @connectParameters | Out-Null
    } catch {
        $TxtGraphConnectionStatus.Text = $previousStatus
        throw
    }
    $context = Get-MgContext
    if ($script:GraphState.ExpectedTenant -and (-not $context -or $context.TenantId -ne $script:GraphState.ExpectedTenant)) {
        $TxtGraphConnectionStatus.Text = $previousStatus
        throw 'Sign-in returned a different tenant. Sign out and reconnect before continuing.'
    }
    Update-TenantInfo
    $stillMissing = @(Get-IntuneToolboxMissingScopes -Scopes $Scopes -Context $context)
    if ($stillMissing.Count) { throw "Sign-in did not grant: $($stillMissing -join ', '). An administrator may need to grant consent." }
    return "Granted Microsoft Graph permission(s): $($missing -join ', ')."
}

function New-IntuneToolboxFeatureWorker {
    param(
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)][hashtable]$Parameters,
        [Parameter(Mandatory)][string[]]$Scopes,
        [Parameter(Mandatory)][scriptblock]$OutputAction
    )

    $featurePath = Join-Path $PSScriptRoot (Join-Path 'Features' $RelativePath)
    $runnerPath = Join-Path $PSScriptRoot 'Shared\Invoke-IntuneToolboxOperation.ps1'
    $permissionNote = ''
    $permissionError = ''
    try { $permissionNote = Request-IntuneToolboxScopes -Scopes $Scopes } catch { $permissionError = $_.Exception.Message }
    # Plain data only: Invoke-InBackground runs the shared runner inside its own runspace and
    # calls OutputAction on the UI thread.
    return [PSCustomObject]@{
        FeaturePath     = $featurePath
        RunnerPath      = $runnerPath
        Parameters      = $Parameters
        Scopes          = $Scopes
        State           = $script:GraphState
        PermissionNote  = $permissionNote
        PermissionError = $permissionError
        OutputAction    = $OutputAction
    }
}

#endregion Helpers

#region Script state
$script:ToolboxGraphScopes = @(
    'User.Read'
    'DeviceManagementConfiguration.ReadWrite.All'
    'DeviceManagementServiceConfig.ReadWrite.All'
    'DeviceManagementManagedDevices.ReadWrite.All'
    'DeviceManagementManagedDevices.PrivilegedOperations.All'
    'DeviceManagementApps.ReadWrite.All'
    'Group.Read.All'
)
$script:AllProfiles        = [System.Collections.ArrayList]::new()
$script:AllPolicies        = @()
$script:AllFilters         = @()
$script:APRows             = [System.Collections.Generic.List[object]]::new()
$script:APOperationResults = [System.Collections.Generic.List[object]]::new()
$script:APMaxGroupTagLength = 128
$script:APFeaturePath      = Join-Path $PSScriptRoot 'Features\Invoke-AutopilotGroupTagOperation.ps1'
$script:AllSearchResults   = [System.Collections.ArrayList]::new()
$script:HasSearchedSettings = $false
$script:BusyButtonCounts = @{}
Update-ProfileCounters
Update-APCounters
Update-SettingSearchResultState

# ── Naming Standard: settingDefinitionId → Product / App / Category mappings ──
$script:AppKeyMap = @{
    'outlk16'  = @{ Product = 'M365 Apps'; App = 'Microsoft Outlook 2016' }
    'word16'   = @{ Product = 'M365 Apps'; App = 'Microsoft Word 2016' }
    'excel16'  = @{ Product = 'M365 Apps'; App = 'Microsoft Excel 2016' }
    'ppt16'    = @{ Product = 'M365 Apps'; App = 'Microsoft PowerPoint 2016' }
    'access16' = @{ Product = 'M365 Apps'; App = 'Microsoft Access 2016' }
    'pub16'    = @{ Product = 'M365 Apps'; App = 'Microsoft Publisher 2016' }
    'onent16'  = @{ Product = 'M365 Apps'; App = 'Microsoft OneNote 2016' }
    'proj16'   = @{ Product = 'M365 Apps'; App = 'Microsoft Project 2016' }
    'visio16'  = @{ Product = 'M365 Apps'; App = 'Microsoft Visio 2016' }
    'lync16'   = @{ Product = 'M365 Apps'; App = 'Skype for Business 2016' }
    'office16' = @{ Product = 'M365 Apps'; App = 'Common Settings' }
    'msedge'   = @{ Product = 'Microsoft Edge'; App = 'Microsoft Edge' }
    'chrome'   = @{ Product = 'Google Chrome'; App = 'Google Chrome' }
}

$script:CategoryMap = @{
    # Outlook
    'l_toolsaccounts'                       = 'Account Settings'
    'l_security'                            = 'Security'
    'l_miscellaneous'                       = 'Miscellaneous'
    'l_mailformat'                          = 'Mail Format'
    'l_mailsetup'                           = 'Mail Setup'
    'l_preferences'                         = 'Preferences'
    'l_calendaroptions'                     = 'Calendar'
    'l_outlooktoday'                        = 'Outlook Today'
    'l_searchfolders'                       = 'Search'
    'l_outlooksocialnconnector'             = 'Outlook Social Connector'
    'l_infopathsettoolsoptionsgeneralcategory' = 'General'
    'l_meetingworkspacescategory'           = 'Meeting Workspaces'
    # Common Office
    'l_microsoftofficesecurity'             = 'Security'
    'l_trustcenter'                         = 'Trust Center'
    'l_filetab'                             = 'File Tab'
    'l_privacy'                             = 'Privacy'
    'l_advanced'                            = 'Advanced'
    'l_toolscustomize'                      = 'Customization'
    'l_coauthoring'                         = 'Co-Authoring'
    'l_serversettings'                      = 'Server Settings'
    'l_converters'                          = 'Converters'
    'l_publishingoptions'                   = 'Publishing'
    'l_proofing'                            = 'Proofing'
    'l_printingsettings'                    = 'Printing'
    'l_savesettings'                        = 'Save'
    'l_collaborationsettings'               = 'Collaboration'
    # Edge / Chrome
    'l_contentssettings'                    = 'Content Settings'
    'l_extensions'                          = 'Extensions'
    'l_passwordmanager'                     = 'Password Manager'
    'l_startup'                             = 'Startup'
}

$script:PlatformMap = @{
    'windows10'       = 'Win'
    'windows10AndLater' = 'Win'
    'iOS'             = 'iOS'
    'macOS'           = 'macOS'
    'android'         = 'Android'
    'androidForWork'  = 'Android'
    'linux'           = 'Linux'
}

function Resolve-NamingComponents {
    <#
    .SYNOPSIS  Parse a settingDefinitionId and return Product, App, Category.
    #>
    param([string]$SettingDefinitionId)

    $segments = $SettingDefinitionId -split '~'
    if ($segments.Count -lt 4) { return $null }

    # Context (Device vs User)
    $context = if ($segments[0] -match '^device_') { 'Device' } elseif ($segments[0] -match '^user_') { 'User' } else { $null }

    # App key: everything after _config_ up to ~ boundary
    $appKeyRaw = $segments[0] -replace '^.*_config_', ''
    if (-not $appKeyRaw -or $appKeyRaw -eq $segments[0]) { return $null }

    # Match app key against known prefixes (strip version suffix like v2, v3)
    $product = $null; $appName = $null
    foreach ($prefix in ($script:AppKeyMap.Keys | Sort-Object { $_.Length } -Descending)) {
        if ($appKeyRaw -match "^$([regex]::Escape($prefix))") {
            $product = $script:AppKeyMap[$prefix].Product
            $appName = $script:AppKeyMap[$prefix].App
            break
        }
    }
    if (-not $product) { return $null }

    # Category: segment index 3 (after app_key ~ policy ~ l_appDisplayName ~ l_category)
    # Strip any appended leaf-setting suffix (_l_…) so we keep only the top-level category
    $categorySeg = ($segments[3] -split '_l_', 2)[0]
    $categoryName = $script:CategoryMap[$categorySeg]
    if (-not $categoryName) {
        # Fallback: strip l_ prefix and insert spaces before uppercase or at word boundaries
        $raw = $categorySeg -replace '^l_', ''
        # Try to split on known word patterns by inserting spaces
        $raw = $raw -creplace '([a-z])([A-Z])', '$1 $2'
        $categoryName = (Get-Culture).TextInfo.ToTitleCase($raw)
    }

    return [PSCustomObject]@{
        Context  = $context
        Product  = $product
        App      = $appName
        Category = $categoryName
    }
}
#endregion

#region Tab 1 — Profile Management handlers

$BtnPreview.Add_Click({
    $loadFilter   = $TxtLoadFilter.Text
    $matchStr     = $TxtMatch.Text
    $assignFilter   = Get-SelectedTag $ComboAssignmentFilter
    if (-not $assignFilter) { $assignFilter = 'Any' }
    $platformFilter = Get-SelectedTag $ComboPlatformFilter
    if (-not $platformFilter) { $platformFilter = 'All' }

    $includeSettingsCatalog = $ChkSettingsCatalog.IsChecked
    $includeDeviceConfig    = $ChkDeviceConfig.IsChecked
    $includeAdministrativeTemplates = $ChkAdministrativeTemplates.IsChecked
    $includeSucceeded       = $ChkSucceededCount.IsChecked
    if (-not $includeSettingsCatalog -and -not $includeDeviceConfig -and -not $includeAdministrativeTemplates) {
        Write-ProfileLog "Select at least one Profile Type." "WARN"
        return
    }

    $replaceStr            = $TxtReplace.Text
    $prefixStr             = $TxtPrefix.Text
    $suffixStr             = $TxtSuffix.Text
    $script:AllProfiles.Clear()
    $ListProfiles.Items.Clear()
    $TxtProfileFilter.Text = ""

    # Shared thread-safe list for the background worker to populate
    $previewBag = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()

    

    $_listView    = $ListProfiles
        $profileTypes = @()
        if ($includeSettingsCatalog) { $profileTypes += 'SettingsCatalog' }
        if ($includeDeviceConfig) { $profileTypes += 'DeviceConfiguration' }
        if ($includeAdministrativeTemplates) { $profileTypes += 'AdministrativeTemplates' }
        $profileParameters = @{
            NameFilter = $loadFilter
            ProfileType = $profileTypes
            AssignmentFilter = $assignFilter
            PlatformFilter = $platformFilter
            IncludeSucceeded = [bool]$includeSucceeded
            Match = $matchStr
            Replace = $replaceStr
            Prefix = $prefixStr
            Suffix = $suffixStr
        }
        $profileOutputAction = {
            param($profileItem)
            $succeededCount = 0
            $sortValue = if ([int]::TryParse([string]$profileItem.Succeeded, [ref]$succeededCount) -and $succeededCount -ge 0) {
                $succeededCount
            } else { -1 }
            $profileItem | Add-Member -NotePropertyName SucceededSortValue -NotePropertyValue $sortValue -Force
            $previewBag.Add($profileItem)
        }.GetNewClosure()
        $profileScopes = @('DeviceManagementConfiguration.Read.All')
        if ($includeSucceeded) { $profileScopes += 'DeviceManagementManagedDevices.Read.All' }
        $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Get-IntuneProfiles.ps1' `
            -Parameters $profileParameters -Scopes $profileScopes -OutputAction $profileOutputAction

        $_listView    = $ListProfiles
    $_filterBox  = $TxtProfileFilter
    $_bag        = $previewBag
    $_allProfiles = $script:AllProfiles
    $onComplete = {
        $sorted = @($_bag.ToArray() | Sort-Object ProfileType, CurrentName)
        $_allProfiles.Clear()
        [void]$_allProfiles.AddRange($sorted)
        $_filterBox.Text = ""
        $_listView.Items.Clear()
        foreach ($p in $_allProfiles) { $_listView.Items.Add($p) }
        Update-ProfileCounters
    }.GetNewClosure()

    Invoke-InBackground -Worker $worker `
        -Buttons        @($BtnPreview, $BtnRenameExecute, $BtnDeleteExecute, $BtnAssignExecute, $BtnRenameReset) `
        -BusyLabel      "Loading..." `
        -LogBox         $TxtProfileLog `
        -LogScroll      $ProfileLogScroll `
        -DarkMode       $script:IsDarkMode `
        -OverlayMessage 'Loading profiles...' `
        -OnComplete     $onComplete
})

function Set-ProfileTableColumnWidths {
    if ($ListProfiles.ActualWidth -le 0) { return }
    $columns = $ListProfiles.View.Columns
    if ($columns.Count -lt 6) { return }

    $fixedWidth = $columns[3].Width + $columns[4].Width + $columns[5].Width
    $flexWidth = [Math]::Max(0, $ListProfiles.ActualWidth - 34 - $fixedWidth)
    $typeWidth = [Math]::Max(160, [Math]::Floor($flexWidth * 0.15))
    $currentWidth = [Math]::Max(190, [Math]::Floor($flexWidth * 0.45))
    $newWidth = [Math]::Max(180, [Math]::Floor($flexWidth - $currentWidth - $typeWidth))
    $columns[0].Width = $currentWidth
    $columns[1].Width = $newWidth
    $columns[2].Width = $typeWidth
}

function Enable-ListViewSorting {
    param([System.Windows.Controls.ListView]$ListView, [hashtable]$TemplateSortPaths = @{})

    $ListView.Tag = [PSCustomObject]@{
        Headers = @($ListView.View.Columns | ForEach-Object { [string]$_.Header })
        Property = $null
        Direction = $null
        TemplateSortPaths = $TemplateSortPaths
    }
    $ListView.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, [System.Windows.RoutedEventHandler]{
        param($sender, $eventArgs)

        $header = $eventArgs.OriginalSource
        if ($header -isnot [System.Windows.Controls.GridViewColumnHeader]) {
            $header = $eventArgs.Source
        }
        while ($header -and $header -isnot [System.Windows.Controls.GridViewColumnHeader]) {
            try { $header = [System.Windows.Media.VisualTreeHelper]::GetParent($header) } catch { $header = $null }
        }
        if (-not $header -or -not $header.Column) { return }

        if ($header.Column.DisplayMemberBinding) {
            $sortProperty = $header.Column.DisplayMemberBinding.Path.Path
        } else {
            # Template columns have no binding to sort by; they declare their property by base header text.
            $columnIndex = $sender.View.Columns.IndexOf($header.Column)
            if ($columnIndex -lt 0) { return }
            $sortProperty = $sender.Tag.TemplateSortPaths[$sender.Tag.Headers[$columnIndex]]
            if (-not $sortProperty) { return }
        }
        if ($sortProperty -eq 'Succeeded') { $sortProperty = 'SucceededSortValue' }
        if ($sortProperty -eq 'TimestampLocal') { $sortProperty = 'Timestamp' }
        $sortState = $sender.Tag
        $direction = if ($sortState.Property -eq $sortProperty -and
                         $sortState.Direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
            [System.ComponentModel.ListSortDirection]::Descending
        } else {
            [System.ComponentModel.ListSortDirection]::Ascending
        }

        $sortState.Property = $sortProperty
        $sortState.Direction = $direction
        $sender.Items.SortDescriptions.Clear()
        [void]$sender.Items.SortDescriptions.Add(
            [System.ComponentModel.SortDescription]::new($sortProperty, $direction))
        $sender.Items.Refresh()

        for ($index = 0; $index -lt $sender.View.Columns.Count; $index++) {
            $column = $sender.View.Columns[$index]
            $baseHeader = $sortState.Headers[$index]
            $column.Header = $baseHeader
            if ($column -eq $header.Column) {
                $indicator = if ($direction -eq [System.ComponentModel.ListSortDirection]::Ascending) { '▲' } else { '▼' }
                $column.Header = "$baseHeader $indicator"
            }
        }
    }.GetNewClosure())
}

$ListProfiles.Add_SizeChanged({ Set-ProfileTableColumnWidths })
$ListProfiles.Add_SelectionChanged({ Update-ProfileCounters })
Enable-ListViewSorting -ListView $ListProfiles
Enable-ListViewSorting -ListView $ListSettingResults
$TxtMatch.Add_TextChanged({ Update-ProfileCounters })
$TxtPrefix.Add_TextChanged({ Update-ProfileCounters })
$TxtSuffix.Add_TextChanged({ Update-ProfileCounters })
$ChkRemoveWhitespace.Add_Click({ Update-ProfileCounters })

$TxtProfileFilter.Add_TextChanged({
    $term = $TxtProfileFilter.Text.Trim().ToLower()
    $ListProfiles.Items.Clear()
    $filtered = if ($term -eq '') { $script:AllProfiles } else {
        $script:AllProfiles | Where-Object {
            $_.CurrentName.ToLower().Contains($term) -or $_.NewName.ToLower().Contains($term)
        }
    }
    foreach ($p in $filtered) { $ListProfiles.Items.Add($p) }
    Update-ProfileCounters
})

$BtnRenameSelectAll.Add_Click({ $ListProfiles.SelectAll() })

$BtnRenameDeselectAll.Add_Click({ $ListProfiles.SelectedItems.Clear() })

$BtnExportCsv.Add_Click({
    if ($script:AllProfiles.Count -eq 0) {
        Write-ProfileLog "No profiles loaded — use 'Load profiles' first." "WARN"
        return
    }

    $dlg              = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title        = "Export Profiles to CSV"
    $dlg.Filter       = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dlg.DefaultExt   = "csv"
    $dlg.FileName     = "IntuneProfiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -ne $true) { return }

    try {
        $script:AllProfiles |
            Select-Object CurrentName, NewName, ProfileType, Assignments, Succeeded, Compliant, Id |
            Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Write-ProfileLog "Exported $($script:AllProfiles.Count) profile(s) to '$($dlg.FileName)'." "OK"
    } catch {
        Write-ProfileLog "Export failed: $_" "ERR"
    }
})

$BtnPreviewRename.Add_Click({
    $matchStr         = $TxtMatch.Text
    $replaceStr       = $TxtReplace.Text
    $prefixStr        = $TxtPrefix.Text
    $suffixStr        = $TxtSuffix.Text
    $removeWhitespace = $ChkRemoveWhitespace.IsChecked

    if ([string]::IsNullOrEmpty($prefixStr) -and [string]::IsNullOrEmpty($suffixStr) -and [string]::IsNullOrEmpty($matchStr) -and -not $removeWhitespace) {
        Write-ProfileLog "Nothing to preview — provide a Prefix, Suffix, Match substring, or enable Remove whitespace." "WARN"
        return
    }

    $selectedProfiles = @($ListProfiles.SelectedItems)
    if ($selectedProfiles.Count -eq 0) {
        Write-ProfileLog "Select at least one profile to preview." "WARN"
        return
    }

    $updated = 0
    foreach ($p in $selectedProfiles) {
        $baseName = if (-not [string]::IsNullOrEmpty($matchStr)) { $p.CurrentName.Replace($matchStr, $replaceStr) } else { $p.CurrentName }
        $newName  = $prefixStr + $baseName + $suffixStr
        if ($removeWhitespace) { $newName = $newName -replace '\s', '' }
        $p.NewName = $newName
        $updated++
    }

    # Refresh the list view showing only the previewed profiles
    $TxtProfileFilter.Text = ''
    $ListProfiles.Items.Clear()
    foreach ($p in $selectedProfiles) { $ListProfiles.Items.Add($p) }
    $ListProfiles.SelectAll()
    Update-ProfileCounters

    Write-ProfileLog "Preview: updated New Name for $updated profile(s)." "OK"
})

$BtnRenameExecute.Add_Click({
    $matchStr        = $TxtMatch.Text
    $replaceStr      = $TxtReplace.Text
    $prefixStr       = $TxtPrefix.Text
    $suffixStr       = $TxtSuffix.Text
    $removeWhitespace = $ChkRemoveWhitespace.IsChecked

    if ([string]::IsNullOrEmpty($prefixStr) -and [string]::IsNullOrEmpty($suffixStr) -and [string]::IsNullOrEmpty($matchStr) -and -not $removeWhitespace) {
        Write-ProfileLog "Nothing to rename — provide a Prefix, Suffix, Match substring, or enable Remove whitespace." "WARN"
        return
    }

    $selectedProfiles = @($ListProfiles.SelectedItems)
    if ($selectedProfiles.Count -eq 0) {
        Write-ProfileLog "Select at least one profile to rename." "WARN"
        return
    }

    $ops = @()
    if ($prefixStr)        { $ops += "prefix '$prefixStr'" }
    if ($matchStr)         { $ops += "replace '$matchStr' \u2192 '$replaceStr'" }
    if ($suffixStr)        { $ops += "suffix '$suffixStr'" }
    if ($removeWhitespace) { $ops += 'remove whitespace' }
    $confirmMsg = "You are about to rename $($selectedProfiles.Count) profile(s) ($($ops -join '; ')).`n`nThis will modify profile display names in Intune immediately. Continue?"
    $result = [System.Windows.MessageBox]::Show($confirmMsg, "Confirm Rename", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-ProfileLog "Rename cancelled by user." "WARN"
        return
    }

    Write-ProfileLog "Starting rename for $($selectedProfiles.Count) profile(s)..."

    $renameParameters = @{
        Profiles = $selectedProfiles
        Match = $matchStr
        Replace = $replaceStr
        Prefix = $prefixStr
        Suffix = $suffixStr
        RemoveWhitespace = [bool]$removeWhitespace
    }
    $ignoreRenameResult = { param($result) }.GetNewClosure()
    $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Rename-IntuneProfiles.ps1' `
        -Parameters $renameParameters -Scopes @('DeviceManagementConfiguration.ReadWrite.All') -OutputAction $ignoreRenameResult

    Invoke-InBackground -Worker $worker `
        -Buttons        @($BtnRenameExecute, $BtnDeleteExecute, $BtnRenameReset, $BtnPreview) `
        -BusyLabel      "Renaming..." `
        -LogBox         $TxtProfileLog `
        -LogScroll      $ProfileLogScroll `
        -DarkMode       $script:IsDarkMode `
        -OverlayMessage 'Renaming profiles...'
})

$BtnDeleteExecute.Add_Click({
    $selectedProfiles = @($ListProfiles.SelectedItems)
    if ($selectedProfiles.Count -eq 0) {
        Write-ProfileLog "Select at least one profile to delete." "WARN"
        return
    }

    $confirmMsg = "You are about to delete $($selectedProfiles.Count) selected Intune profile(s).`n`nThis operation removes the selected profiles from Intune immediately. Continue?"
    $result = [System.Windows.MessageBox]::Show($confirmMsg, "Confirm Delete", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-ProfileLog "Delete cancelled by user." "WARN"
        return
    }

    Write-ProfileLog "Starting deletion of $($selectedProfiles.Count) profile(s)..."

    $script:DeleteSucceededIds = [System.Collections.Generic.HashSet[string]]::new()

    $successfulDeleteIds = $script:DeleteSucceededIds
    $deleteOutputAction = {
        param($result)
        if ($result.Status -eq 'OK') { [void]$successfulDeleteIds.Add([string]$result.Id) }
    }.GetNewClosure()
    $deleteParameters = @{ Profiles = $selectedProfiles }
    $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Remove-IntuneProfiles.ps1' `
        -Parameters $deleteParameters -Scopes @('DeviceManagementConfiguration.ReadWrite.All') -OutputAction $deleteOutputAction

    $_profileList = $ListProfiles
    $_filterBox   = $TxtProfileFilter
    $_allProfiles = $script:AllProfiles
    $_succeededIds = $script:DeleteSucceededIds
    $onComplete = {
        $remaining = @($_allProfiles | Where-Object { $_.Id -notin $_succeededIds })
        $_allProfiles.Clear()
        [void]$_allProfiles.AddRange($remaining)
        $_filterBox.Text = ""
        $_profileList.Items.Clear()
        foreach ($p in $_allProfiles) { $_profileList.Items.Add($p) }
        Update-ProfileCounters
    }.GetNewClosure()

    Invoke-InBackground -Worker     $worker `
        -Buttons        @($BtnDeleteExecute, $BtnRenameExecute, $BtnRenameReset, $BtnPreview) `
        -BusyLabel      "Deleting..." `
        -LogBox         $TxtProfileLog `
        -LogScroll      $ProfileLogScroll `
        -DarkMode       $script:IsDarkMode `
        -OnComplete     $onComplete `
        -OverlayMessage 'Deleting profiles...'
})

$BtnRenameReset.Add_Click({
    $TxtMatch.Text         = ""
    $TxtLoadFilter.Text    = ""
    $TxtPrefix.Text        = ""
    $TxtReplace.Text       = ""
    $TxtSuffix.Text        = ""
    $TxtProfileFilter.Text = ""
    $ComboAssignmentFilter.SelectedIndex  = 0
    $ComboPlatformFilter.SelectedIndex    = 0
    # Assignment sub-tab
    $ComboAction.SelectedIndex      = 0
    $ComboFilterMode.SelectedIndex  = 0
    $ListGroups.SelectedItem        = $null
    $ListRemoveGroups.SelectedItem  = $null
    $ListGroups.Items.Clear()
    $ListRemoveGroups.Items.Clear()
    $ListQueuedGroups.Items.Clear()
    $ListQueuedRemoveGroups.Items.Clear()
    $ListFilters.SelectedItem       = $null
    $TxtGroupSearch.Text            = ""
    $TxtRemoveGroupSearch.Text      = ""
    $TxtFilterSearch.Text           = ""
    $script:AllFilters              = @()
    $ListFilters.Items.Clear()
    # Profiles list and log
    $script:AllProfiles.Clear()
    $ListProfiles.Items.Clear()
    $TxtProfileLog.Document.Blocks.Clear()
    Update-ProfileCounters
    Write-ProfileLog "Form reset."
})

#endregion Tab 1 — Profile Management

$ComboAction.Add_SelectionChanged({
    $action = Get-SelectedTag $ComboAction
    if ($action -eq 'Remove') {
        Show-Panel $PanelAddReplace $false
        Show-Panel $GroupRemove     $true
    } else {
        Show-Panel $PanelAddReplace $true
        Show-Panel $GroupRemove     $false
    }
    Update-ProfileCounters
})

$ComboTargetType.Add_SelectionChanged({
    Show-Panel $PanelGroupSearch ((Get-SelectedTag $ComboTargetType) -eq 'Group')
    Update-ProfileCounters
})

$ComboFilterMode.Add_SelectionChanged({
    $mode = Get-SelectedTag $ComboFilterMode
    Show-Panel $PanelFilterList ($mode -ne 'None')
    if ($mode -ne 'None' -and $script:AllFilters.Count -eq 0) {
        Write-AssignLog "Loading device filters..."
        $ListFilters.Items.Clear()
        $TxtFilterSearch.Text = ""
        $filterBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
        $filterOutputAction = { param($filter) $filterBag.Add($filter) }.GetNewClosure()
        $filterWorker = New-IntuneToolboxFeatureWorker -RelativePath 'Search-IntuneAssignmentTargets.ps1' `
            -Parameters @{ Action = 'Filters' } -Scopes @('DeviceManagementConfiguration.Read.All') -OutputAction $filterOutputAction
        $_filterBag = $filterBag
        $onFilterLoad = {
            $script:AllFilters = @($_filterBag.ToArray())
            foreach ($filter in $script:AllFilters) { $ListFilters.Items.Add($filter) }
            Write-AssignLog "Loaded $($script:AllFilters.Count) filter(s)." 'OK'
        }.GetNewClosure()
        Invoke-InBackground -Worker $filterWorker -Buttons @($BtnAssignExecute) -BusyLabel 'Loading...' `
            -LogBox $TxtProfileLog -LogScroll $ProfileLogScroll -DarkMode $script:IsDarkMode -OnComplete $onFilterLoad
    }
    Update-ProfileCounters
})

$TxtFilterSearch.Add_TextChanged({
    $term = $TxtFilterSearch.Text.Trim().ToLower()
    $ListFilters.Items.Clear()
    $filtered = if ($term -eq '') { $script:AllFilters } else {
        $script:AllFilters | Where-Object { $_.DisplayName.ToLower().Contains($term) }
    }
    foreach ($f in $filtered) { $ListFilters.Items.Add($f) }
    Update-ProfileCounters
})

$ListGroups.Add_SelectionChanged({ Update-ProfileCounters })
$ListQueuedGroups.Add_SelectionChanged({ Update-ProfileCounters })
$ListFilters.Add_SelectionChanged({ Update-ProfileCounters })
$ListRemoveGroups.Add_SelectionChanged({ Update-ProfileCounters })
$ListQueuedRemoveGroups.Add_SelectionChanged({ Update-ProfileCounters })
$ComboRemoveType.Add_SelectionChanged({ Update-ProfileCounters })

$ComboRemoveScope.Add_SelectionChanged({
    Show-Panel $PanelRemoveGroup ((Get-SelectedTag $ComboRemoveScope) -eq 'Group')
    Update-ProfileCounters
})

# Initial panel states for assignment sub-tab
Show-Panel $PanelGroupSearch $false
Show-Panel $PanelFilterList  $false

function Add-AssignmentQueueGroups {
    param([System.Windows.Controls.ListBox]$Results, [System.Windows.Controls.ListBox]$Queue)
    foreach ($group in @($Results.SelectedItems)) {
        if ([string]::IsNullOrWhiteSpace([string]$group.Id)) {
            Write-AssignLog "Cannot queue a group without an ID." 'ERR'
            continue
        }
        if ([string]$group.Id -notin @($Queue.Items | ForEach-Object { [string]$_.Id })) {
            [void]$Queue.Items.Add($group)
        }
    }
    Update-ProfileCounters
}

function Remove-AssignmentQueueGroups {
    param([System.Windows.Controls.ListBox]$Queue)
    foreach ($group in @($Queue.SelectedItems)) { $Queue.Items.Remove($group) }
    Update-ProfileCounters
}

$BtnQueueGroups.Add_Click({ Add-AssignmentQueueGroups -Results $ListGroups -Queue $ListQueuedGroups })
$BtnQueueRemoveGroups.Add_Click({ Add-AssignmentQueueGroups -Results $ListRemoveGroups -Queue $ListQueuedRemoveGroups })
$BtnUnqueueGroups.Add_Click({ Remove-AssignmentQueueGroups -Queue $ListQueuedGroups })
$BtnUnqueueRemoveGroups.Add_Click({ Remove-AssignmentQueueGroups -Queue $ListQueuedRemoveGroups })
$BtnClearQueuedGroups.Add_Click({ $ListQueuedGroups.Items.Clear(); Update-ProfileCounters })
$BtnClearQueuedRemoveGroups.Add_Click({ $ListQueuedRemoveGroups.Items.Clear(); Update-ProfileCounters })

$BtnSearchGroup.Add_Click({
    $term = $TxtGroupSearch.Text.Trim()
    if ($term.Length -eq 0) { Write-AssignLog "Enter a group search term." "WARN"; return }
    Write-AssignLog "Searching for groups matching '$term'..."
    $ListGroups.Items.Clear()
    $groupBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $groupOutputAction = { param($group) $groupBag.Add($group) }.GetNewClosure()
    $groupWorker = New-IntuneToolboxFeatureWorker -RelativePath 'Search-IntuneAssignmentTargets.ps1' `
        -Parameters @{ Action = 'Groups'; SearchTerm = $term } -Scopes @('Group.Read.All') -OutputAction $groupOutputAction
    $_groupBag = $groupBag; $_groupList = $ListGroups
    $onGroupSearch = {
        $_groupList.Items.Clear()
        foreach ($group in $_groupBag.ToArray() | Sort-Object DisplayName) { $_groupList.Items.Add($group) }
        Update-ProfileCounters
    }.GetNewClosure()
    Invoke-InBackground -Worker $groupWorker -Buttons @($BtnSearchGroup) -BusyLabel 'Searching...' `
        -LogBox $TxtProfileLog -LogScroll $ProfileLogScroll -DarkMode $script:IsDarkMode -OnComplete $onGroupSearch
})

$BtnSearchRemoveGroup.Add_Click({
    $term = $TxtRemoveGroupSearch.Text.Trim()
    if ($term.Length -eq 0) { Write-AssignLog "Enter a group search term." "WARN"; return }
    Write-AssignLog "Searching for groups matching '$term'..."
    $ListRemoveGroups.Items.Clear()
    $groupBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $groupOutputAction = { param($group) $groupBag.Add($group) }.GetNewClosure()
    $groupWorker = New-IntuneToolboxFeatureWorker -RelativePath 'Search-IntuneAssignmentTargets.ps1' `
        -Parameters @{ Action = 'Groups'; SearchTerm = $term } -Scopes @('Group.Read.All') -OutputAction $groupOutputAction
    $_groupBag = $groupBag; $_groupList = $ListRemoveGroups
    $onGroupSearch = {
        $_groupList.Items.Clear()
        foreach ($group in $_groupBag.ToArray() | Sort-Object DisplayName) { $_groupList.Items.Add($group) }
        Update-ProfileCounters
    }.GetNewClosure()
    Invoke-InBackground -Worker $groupWorker -Buttons @($BtnSearchRemoveGroup) -BusyLabel 'Searching...' `
        -LogBox $TxtProfileLog -LogScroll $ProfileLogScroll -DarkMode $script:IsDarkMode -OnComplete $onGroupSearch
})

function Get-AssignmentParameters {
    $parameters = @{
        Profiles = @($ListProfiles.SelectedItems)
        Action = Get-SelectedTag $ComboAction
    }
    if ($parameters.Action -in @('Add', 'Replace')) {
        $parameters.TargetType = Get-SelectedTag $ComboTargetType
        $parameters.IncludeExclude = Get-SelectedTag $ComboIncludeExclude
        $parameters.FilterMode = Get-SelectedTag $ComboFilterMode
        if ($parameters.TargetType -eq 'Group') { $parameters.TargetGroup = @($ListQueuedGroups.Items) }
        if ($parameters.FilterMode -ne 'None') { $parameters.AssignmentFilter = $ListFilters.SelectedItem }
    } else {
        $parameters.RemoveScope = Get-SelectedTag $ComboRemoveScope
        if ($parameters.RemoveScope -eq 'Group') {
            $parameters.RemoveGroup = @($ListQueuedRemoveGroups.Items)
            $parameters.RemoveType = Get-SelectedTag $ComboRemoveType
        }
    }
    return $parameters
}

$BtnAssignExecute.Add_Click({
    if (-not (Test-AssignmentReady)) {
        Write-AssignLog "Select profiles, an action and target, queue at least one group for group operations, and select any enabled filter." 'WARN'
        return
    }
    # Use the same UI-thread snapshot for confirmation and the background operation.
    $assignmentParameters = Get-AssignmentParameters
    $action = $assignmentParameters.Action
    $selectedProfiles = $assignmentParameters.Profiles
    $targetDetail = if ($action -eq 'Remove') {
        if ($assignmentParameters.RemoveScope -eq 'All') { 'ALL direct assignments' }
        else { "$($assignmentParameters.RemoveType) assignments for $($assignmentParameters.RemoveGroup.Count) group(s):`n$($assignmentParameters.RemoveGroup.DisplayName -join "`n")" }
    } elseif ($assignmentParameters.TargetType -eq 'Group') {
        "$($assignmentParameters.IncludeExclude) $($assignmentParameters.TargetGroup.Count) group(s):`n$($assignmentParameters.TargetGroup.DisplayName -join "`n")"
    } else { "All $($assignmentParameters.TargetType)" }
    if ($assignmentParameters.FilterMode -and $assignmentParameters.FilterMode -ne 'None') {
        $targetDetail += "`nFilter: $($assignmentParameters.FilterMode) $($assignmentParameters.AssignmentFilter.DisplayName)"
    }
    $confirmImpact = if ($action -eq 'Replace') { "`nAll existing direct assignments will be replaced by the complete new target set." } else { '' }
    $confirmMsg = "$action assignments on $($selectedProfiles.Count) profile(s).`nTargets: $targetDetail$confirmImpact`n`nThis action will take effect in Intune immediately. Continue?"
    $result = [System.Windows.MessageBox]::Show($confirmMsg, "Confirm Bulk Assignment", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-AssignLog "Bulk assignment cancelled by user." "WARN"
        return
    }
    Write-AssignLog "Starting '$action' for $($selectedProfiles.Count) profile(s): $targetDetail"

    $ignoreAssignmentResult = { param($result) }.GetNewClosure()
    $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Invoke-IntuneProfileAssignment.ps1' `
        -Parameters $assignmentParameters -Scopes @('DeviceManagementConfiguration.ReadWrite.All') -OutputAction $ignoreAssignmentResult

    Invoke-InBackground -Worker $worker `
        -Buttons        @($BtnAssignExecute, $BtnRenameReset, $BtnPreview) `
        -BusyLabel      "Running..." `
        -LogBox         $TxtProfileLog `
        -LogScroll      $ProfileLogScroll `
        -DarkMode       $script:IsDarkMode `
        -OverlayMessage 'Applying assignments...'
})

#endregion Tab 1 — Profile Management: Assignment

#region Tab 2 — Autopilot Group Tags handlers

function Start-APOperation {
    <#
    .SYNOPSIS
        Runs one Invoke-AutopilotGroupTagOperation action in the background. Returned rows are collected in
        $script:APOperationResults; $OnComplete runs on the UI thread and reads $script:APOperationContext.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Parameters,
        [Parameter(Mandatory)][string[]]$Scopes,
        [Parameter(Mandatory)][System.Windows.Controls.Button]$PrimaryButton,
        [Parameter(Mandatory)][string]$BusyLabel,
        [Parameter(Mandatory)][string]$OverlayMessage,
        [Parameter(Mandatory)][scriptblock]$OnComplete,
        [hashtable]$Context = @{}
    )
    if ($script:BusyButtonCounts.ContainsKey($BtnAPLookup.Name)) {
        Write-AutopilotLog 'Another Autopilot operation is still running.' 'WARN'
        return
    }
    $script:APOperationContext = $Context
    $script:APOperationResults = [System.Collections.Generic.List[object]]::new()
    $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Invoke-AutopilotGroupTagOperation.ps1' `
        -Parameters $Parameters -Scopes $Scopes -OutputAction { param($item) $script:APOperationResults.Add($item) }
    $buttons = @($PrimaryButton) + @($script:APActionButtons | Where-Object { $_ -ne $PrimaryButton })
    Invoke-InBackground -Worker $worker -Buttons $buttons -BusyLabel $BusyLabel `
        -LogBox $TxtAutopilotLog -LogScroll $AutopilotLogScroll -DarkMode $script:IsDarkMode `
        -OverlayMessage $OverlayMessage -OnComplete $OnComplete
}

function Complete-APDeviceLookup {
    $context = $script:APOperationContext
    $rows = @($script:APOperationResults | Where-Object { $_.RecordType -eq 'Device' })
    if ($rows.Count -eq 0) {
        Write-AutopilotLog 'The lookup returned no results; the device list was not changed.' 'WARN'
        return
    }
    Merge-APDeviceRows -NewRows $rows -ReplaceAll:([bool]$context.ReplaceAll)
    if ($context.ImportPath) {
        $TxtAPCsvFile.Text = "$(Split-Path $context.ImportPath -Leaf) - $($context.ImportedCount) serial number(s)"
        $TxtAPCsvFile.ToolTip = $context.ImportPath
    }
    Update-APDeviceView
}

function Start-APDeviceLookup {
    param(
        [object[]]$Devices,
        [string[]]$Serials,
        [string]$ImportPath,
        [switch]$ReplaceAll,
        [System.Windows.Controls.Button]$PrimaryButton = $BtnAPLookup,
        [string]$OverlayMessage = 'Looking up devices...'
    )
    $parameters = @{ Action = 'Lookup' }
    if ($Devices) { $parameters.Device = @($Devices) }
    if ($Serials) { $parameters.SerialNumber = @($Serials) }
    Start-APOperation -Parameters $parameters `
        -Scopes @('DeviceManagementServiceConfig.Read.All', 'DeviceManagementManagedDevices.Read.All') `
        -PrimaryButton $PrimaryButton -BusyLabel 'Looking up...' -OverlayMessage $OverlayMessage `
        -Context @{ ReplaceAll = [bool]$ReplaceAll; ImportPath = $ImportPath; ImportedCount = @($Devices).Count } `
        -OnComplete { Complete-APDeviceLookup }
}

function Import-APCsvFile {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $logAction = { param($Msg, $Lvl = 'INFO') Write-AutopilotLog $Msg $Lvl }
        # CSV parsing is local (no Graph), so it runs on the UI thread before the background lookup.
        $imported = @(& $script:APFeaturePath -Action ImportCsv -Path $Path -SkipGraphConnect -PassThru -LogAction $logAction)
    } catch {
        Write-AutopilotLog "CSV import failed: $($_.Exception.Message)" 'ERR'
        Show-APMessage "The CSV file could not be loaded.`n`n$($_.Exception.Message)" 'Import failed' 'Warning' | Out-Null
        return
    }
    if ($imported.Count -eq 0) {
        Write-AutopilotLog 'The CSV file contained no serial numbers.' 'WARN'
        return
    }
    Start-APDeviceLookup -Devices $imported -ImportPath $Path -ReplaceAll -PrimaryButton $BtnAPImportCsv `
        -OverlayMessage 'Importing CSV and looking up devices...'
}

function Start-APTagUpdate {
    param([ValidateSet('ApplyTag', 'ApplyCsvTags', 'ClearTag')][string]$Mode)
    $selectedAll = @($ListDevices.SelectedItems)
    $selection = @(Get-APSelectedEligibleRows)
    $ineligible = $selectedAll.Count - $selection.Count
    if ($selection.Count -eq 0) { Write-AutopilotLog 'Select at least one device that was found in Autopilot.' 'WARN'; return }

    $tag = ''
    if ($Mode -eq 'ApplyTag') {
        $tagError = Test-APNewTagInput $TxtNewTag.Text
        if ($tagError) { Write-AutopilotLog $tagError 'WARN'; return }
        $tag = $TxtNewTag.Text.Trim()
    }

    $plan = Get-APUpdatePlan -Rows $selection -Mode $Mode -Tag $tag
    if ($plan.Changes.Count -eq 0) {
        Write-AutopilotLog "No selected device needs a change ($($plan.Skipped.Count) skipped)." 'WARN'
        Show-APMessage 'None of the selected devices need a change.' 'Nothing to update' | Out-Null
        return
    }

    $previewLines = @($plan.Changes | Select-Object -First 15 | ForEach-Object {
        $old = if ($_.OldTag) { "'$($_.OldTag)'" } else { '(none)' }
        $new = if ($_.NewTag) { "'$($_.NewTag)'" } else { '(cleared)' }
        "  $($_.Serial): $old -> $new"
    })
    if ($plan.Changes.Count -gt 15) { $previewLines += "  ... and $($plan.Changes.Count - 15) more" }
    $title, $intro = switch ($Mode) {
        'ApplyTag' { 'Confirm Group Tag Update', "Set group tag '$tag' on $($plan.Changes.Count) device(s)?" }
        'ApplyCsvTags' { 'Confirm CSV Group Tag Update', "Apply the CSV group tags to $($plan.Changes.Count) device(s)?" }
        'ClearTag' { 'Confirm Clear Group Tag', "REMOVE the group tag from $($plan.Changes.Count) device(s)?`nDevices without a group tag may no longer match dynamic groups that target their current tag." }
    }
    $notes = @()
    if ($plan.Skipped.Count) { $notes += "$($plan.Skipped.Count) selected device(s) will be skipped (no change needed or no valid CSV tag)." }
    if ($ineligible -gt 0) { $notes += "$ineligible selected row(s) are not uniquely registered in Autopilot and will be ignored." }
    if ($ChkAPSync.IsChecked) { $notes += 'A tenant-wide Autopilot sync will be requested after successful changes.' }
    $notes += 'Group tag changes do not immediately reassign deployment profiles; dynamic group membership and profile assignment can take time.'
    $message = "$intro`n`n$($previewLines -join "`n")`n`n$($notes -join "`n")`n`nThis modifies Autopilot devices in Intune immediately. Continue?"
    if ((Show-APMessage $message $title 'Warning' 'YesNo') -ne 'Yes') {
        Write-AutopilotLog 'Group tag change cancelled.' 'WARN'
        return
    }

    $targets = @($plan.Changes | ForEach-Object { $_.Row })
    $parameters = @{ Action = $Mode; Device = $targets }
    if ($Mode -eq 'ApplyTag') { $parameters.NewTag = $tag }
    if ($ChkAPSync.IsChecked) { $parameters.SyncAfterUpdate = $true }
    Write-AutopilotLog "Updating $($targets.Count) device(s)..."
    $primaryButton = switch ($Mode) { 'ApplyTag' { $BtnAPApplyTag } 'ApplyCsvTags' { $BtnAPApplyCsvTags } 'ClearTag' { $BtnAPClearTag } }
    Start-APOperation -Parameters $parameters -Scopes @('DeviceManagementServiceConfig.ReadWrite.All') `
        -PrimaryButton $primaryButton -BusyLabel 'Updating...' -OverlayMessage 'Updating group tags...' `
        -OnComplete { Complete-APTagUpdate }
}

function Complete-APTagUpdate {
    $results = @($script:APOperationResults)
    $deviceResults = @($results | Where-Object { $_.RecordType -eq 'Device' })
    if ($deviceResults.Count -eq 0) {
        Write-AutopilotLog 'The group tag update returned no results. Check the activity log and refresh the device list.' 'WARN'
        Show-APMessage 'The group tag update did not return any results. Review the activity log and refresh the devices before trying again.' 'Group tag update finished' 'Warning' | Out-Null
        return
    }
    Merge-APDeviceRows -NewRows $deviceResults
    Update-APDeviceView
    $updated = @($deviceResults | Where-Object { $_.UpdateStatus -eq 'Updated' }).Count
    $pending = @($deviceResults | Where-Object { $_.UpdateStatus -eq 'Pending verification' }).Count
    $failed = @($deviceResults | Where-Object { $_.UpdateStatus -like 'Failed*' -or $_.UpdateStatus -like 'Stale*' }).Count
    $sync = @($results | Where-Object { $_.RecordType -eq 'Sync' }) | Select-Object -First 1
    $summary = "Updated: $updated`nPending verification: $pending`nFailed or stale: $failed"
    if ($sync) { $summary += "`nAutopilot sync: $($sync.UpdateStatus)$(if ($sync.Details) { " ($($sync.Details))" })" }
    $icon = if ($failed -or ($sync -and $sync.UpdateStatus -eq 'Failed')) { 'Warning' } else { 'Information' }
    Show-APMessage $summary 'Group tag update finished' $icon | Out-Null
    $updatedKeys = @($deviceResults | ForEach-Object { Get-APRowKey $_ })
    foreach ($item in @($ListDevices.Items)) {
        if ((Get-APRowKey $item) -in $updatedKeys) { [void]$ListDevices.SelectedItems.Add($item) }
    }
    Update-APCounters
}

function Start-APFreshStart {
    $targets = @(Get-APSelectedFreshStartRows)
    if (-not $targets.Count) { Write-AutopilotLog 'Select devices with a resolved Intune device and no unresolved tag update.' 'WARN'; return }
    $skipped = $ListDevices.SelectedItems.Count - $targets.Count
    $preview = @($targets | Select-Object -First 15 | ForEach-Object { "  $($_.RequestedSerial) - $($_.DeviceName) - tag '$($_.GroupTag)'" })
    if ($targets.Count -gt 15) { $preview += "  ... and $($targets.Count - 15) more" }
    $message = "Request Fresh Start on $($targets.Count) selected device(s)?`n`n$($preview -join "`n")`n`nWARNING: Windows will be reinstalled. Apps and user data will be removed (keepUserData=false). Back up required data before continuing.`n`nWait for dynamic group membership and the intended Autopilot deployment profile assignment BEFORE proceeding. A successful tag change or sync does not confirm profile readiness. Fresh Start does not guarantee the new profile will be applied.`n`n$skipped selected row(s) will be skipped. Requests run when devices check in; acceptance does not mean completion.`n`nContinue?"
    if ((Show-APMessage $message 'Confirm Bulk Fresh Start' 'Warning' 'YesNo' 'No') -ne 'Yes') {
        Write-AutopilotLog 'Fresh Start cancelled.' 'WARN'
        return
    }
    Start-APOperation -Parameters @{ Action = 'FreshStart'; Device = $targets } `
        -Scopes @('DeviceManagementServiceConfig.Read.All', 'DeviceManagementManagedDevices.Read.All', 'DeviceManagementManagedDevices.PrivilegedOperations.All') `
        -PrimaryButton $BtnAPFreshStart -BusyLabel 'Requesting...' -OverlayMessage 'Requesting Fresh Start...' `
        -Context @{ TargetCount = $targets.Count } -OnComplete { Complete-APFreshStart }
}

function Complete-APFreshStart {
    $results = @($script:APOperationResults | Where-Object { $_.RecordType -eq 'Device' })
    if ($results.Count -eq 0) {
        Write-AutopilotLog 'Fresh Start returned no results. Check the activity log and refresh the device list.' 'WARN'
        Show-APMessage 'Fresh Start did not return any results. Review the activity log and refresh the devices before trying again.' 'Fresh Start requests finished' 'Warning' | Out-Null
        return
    }
    Merge-APDeviceRows -NewRows $results
    Update-APDeviceView
    $requested = @($results | Where-Object { $_.FreshStartStatus -eq 'Requested' }).Count
    $summary = "Fresh Start requested: $requested`nFailed, stale or skipped: $($results.Count - $requested)`n`nMonitor completion in Intune. Requests are not proof that Windows has been reinstalled."
    $icon = if ($requested -ne $results.Count) { 'Warning' } else { 'Information' }
    Show-APMessage $summary 'Fresh Start requests finished' $icon | Out-Null
}

$TxtAPSerials.Add_TextChanged({ Update-APCounters })
$TxtNewTag.Add_TextChanged({ Update-APCounters })
$ListDevices.Add_SelectionChanged({ Update-APCounters })
$TxtAPDeviceFilter.Add_TextChanged({ Update-APDeviceView })
$TxtAPSerials.Add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.Key -eq [System.Windows.Input.Key]::Enter -and $BtnAPLookup.IsEnabled) {
        $eventArgs.Handled = $true
        $BtnAPLookup.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    }
})

$BtnAPLookup.Add_Click({
    $serials = @(Get-APSerialInput $TxtAPSerials.Text)
    if ($serials.Count -eq 0) { Write-AutopilotLog 'Enter at least one serial number.' 'WARN'; return }
    Write-AutopilotLog "Looking up $($serials.Count) serial number(s)..."
    Start-APDeviceLookup -Serials $serials
})

$BtnAPImportCsv.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dialog.Title = 'Select a CSV file with serial numbers'
    if (-not $dialog.ShowDialog($window)) { return }
    if ($script:APRows.Count -gt 0) {
        $answer = Show-APMessage "Replace the $($script:APRows.Count) loaded device(s) with the contents of the CSV file?" 'Import CSV' 'Question' 'YesNo'
        if ($answer -ne 'Yes') { return }
    }
    Write-AutopilotLog "Importing $($dialog.FileName)..."
    Import-APCsvFile -Path $dialog.FileName
})

$BtnAPRefresh.Add_Click({
    if ($script:APRows.Count -eq 0) { return }
    Write-AutopilotLog "Refreshing $($script:APRows.Count) loaded device(s)..."
    Start-APDeviceLookup -Devices @($script:APRows) -ReplaceAll -PrimaryButton $BtnAPRefresh -OverlayMessage 'Refreshing devices...'
})

$BtnAPClearResults.Add_Click({
    $script:APRows.Clear()
    $TxtAPCsvFile.Text = 'No CSV file loaded'
    $TxtAPCsvFile.ToolTip = $null
    $TxtAPDeviceFilter.Text = ''
    Update-APDeviceView
    Write-AutopilotLog 'Device list cleared.'
})

$BtnAPSelectAll.Add_Click({
    $ListDevices.SelectedItems.Clear()
    foreach ($item in @($ListDevices.Items)) { if ($item.Eligible) { [void]$ListDevices.SelectedItems.Add($item) } }
    Update-APCounters
})

$BtnAPDeselectAll.Add_Click({
    $ListDevices.SelectedItems.Clear()
    Update-APCounters
})

$BtnAPApplyTag.Add_Click({ Start-APTagUpdate -Mode ApplyTag })
$BtnAPApplyCsvTags.Add_Click({ Start-APTagUpdate -Mode ApplyCsvTags })
$BtnAPClearTag.Add_Click({ Start-APTagUpdate -Mode ClearTag })
$BtnAPFreshStart.Add_Click({ Start-APFreshStart })

$BtnAPExport.Add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.FileName = "AutopilotGroupTags_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if (-not $dialog.ShowDialog($window)) { return }
    try {
        Export-APDeviceRows -Path $dialog.FileName
        Write-AutopilotLog "Exported $($script:APRows.Count) row(s) to $($dialog.FileName)." 'OK'
    } catch {
        Write-AutopilotLog "Export failed: $($_.Exception.Message)" 'ERR'
    }
})
#endregion Tab 2 — Autopilot

#region Tab 3 — Setting Search

function Write-SettingSearchLog {
    param([string]$Message, [string]$Level = "INFO")
    Write-ActivityEntry -LogBox $TxtSettingSearchLog -LogScroll $SettingSearchLogScroll `
        -Expander $SettingSearchActivityExpander -Summary $TxtSettingActivitySummary -Message $Message -Level $Level
}

$TxtSettingSearchTerm.Add_TextChanged({ Update-SettingSearchScope })

$BtnSettingSearchExecute.Add_Click({
    $searchTerm = $TxtSettingSearchTerm.Text.Trim()
    if ([string]::IsNullOrEmpty($searchTerm)) {
        Write-SettingSearchLog "Enter a search term." "WARN"
        return
    }

    $selectedItems = @($ListProfiles.SelectedItems)
    $profilesToSearch = if ($selectedItems.Count -gt 0) {
        @($selectedItems | Where-Object { $_.ApiType -eq 'configurationPolicies' })
    } else {
        @($script:AllProfiles | Where-Object { $_.ApiType -eq 'configurationPolicies' })
    }

    if ($profilesToSearch.Count -eq 0) {
        Write-SettingSearchLog "No Settings Catalog profiles available. Load profiles in the Profile Management tab first." "WARN"
        return
    }

    $script:AllSearchResults.Clear()
    $ListSettingResults.Items.Clear()
    $TxtSettingResultFilter.Text = ""
    $script:HasSearchedSettings = $false
    Update-SettingSearchResultState

    $searchBag = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()

    

    $_listView   = $ListSettingResults
        $searchParameters = @{ SearchTerm = $searchTerm; Profiles = $profilesToSearch }
        $searchOutputAction = { param($result) $searchBag.Add($result) }.GetNewClosure()
        $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Search-IntuneSettings.ps1' `
            -Parameters $searchParameters -Scopes @('DeviceManagementConfiguration.Read.All') -OutputAction $searchOutputAction

        $_listView   = $ListSettingResults
    $_filterBox  = $TxtSettingResultFilter
    $_bag        = $searchBag
    $_allResults = $script:AllSearchResults
    $onComplete = {
        $sorted = @($_bag.ToArray() | Sort-Object ProfileName, SettingName)
        $_allResults.Clear()
        [void]$_allResults.AddRange($sorted)
        $_filterBox.Text = ""
        $_listView.Items.Clear()
        foreach ($r in $_allResults) { $_listView.Items.Add($r) }
        $script:HasSearchedSettings = $true
        Update-SettingSearchResultState
    }.GetNewClosure()

    Invoke-InBackground -Worker $worker `
        -Buttons        @($BtnSettingSearchExecute) `
        -BusyLabel      "Searching..." `
        -LogBox         $TxtSettingSearchLog `
        -LogScroll      $SettingSearchLogScroll `
        -DarkMode       $script:IsDarkMode `
        -OverlayMessage 'Searching settings...' `
        -OnComplete     $onComplete
})

$TxtSettingResultFilter.Add_TextChanged({
    $term = $TxtSettingResultFilter.Text.Trim().ToLower()
    $ListSettingResults.Items.Clear()
    $filtered = if ($term -eq '') { $script:AllSearchResults } else {
        $script:AllSearchResults | Where-Object {
            $_.ProfileName.ToLower().Contains($term) -or
            $_.SettingName.ToLower().Contains($term) -or
            $_.Value.ToLower().Contains($term) -or
            $_.DefinitionId.ToLower().Contains($term)
        }
    }
    foreach ($r in $filtered) { $ListSettingResults.Items.Add($r) }
    Update-SettingSearchResultState
})

$BtnSettingSearchExport.Add_Click({
    if ($script:AllSearchResults.Count -eq 0) {
        Write-SettingSearchLog "No results to export — run a search first." "WARN"
        return
    }
    $dlg            = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title      = "Export Search Results to CSV"
    $dlg.Filter     = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dlg.DefaultExt = "csv"
    $dlg.FileName   = "SettingSearch_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -ne $true) { return }
    try {
        $script:AllSearchResults |
            Select-Object ProfileName, SettingName, Value, DefinitionId |
            Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Write-SettingSearchLog "Exported $($script:AllSearchResults.Count) result(s) to '$($dlg.FileName)'." "OK"
    } catch {
        Write-SettingSearchLog "Export failed: $_" "ERR"
    }
})

#endregion Tab 3 — Setting Search

#region Duplicate Devices

$script:DuplicateState = @{
    Report = $null; Valid = $false; Busy = $false; Outcomes = @()
    Cancel = $null; CancelRequested = $false
    Message = 'Not scanned. No records are selected automatically.'
    CountsFor = $null; UniqueDeviceCount = 0; GroupCount = 0; Refreshing = $false; ScanTimer = $null
}
$script:DuplicateExportColumns = @(
    'GroupId', 'MatchType', 'MatchReason', 'Confidence', 'GroupSize', 'Id', 'DeviceName',
    'SerialNumber', 'Manufacturer', 'Model', 'AzureADDeviceId', 'UserPrincipalName',
    'OperatingSystem', 'OSVersion', 'EnrolledDateTime', 'LastSyncDateTime', 'DaysSinceLastSync',
    'ComplianceState', 'ManagementAgent', 'ManagementState', 'KeeperId', 'Recommendation',
    'CanDelete', 'ProtectionReason', 'TenantId', 'ScanTimeUtc'
)

function Write-DuplicateLog {
    param([string]$Message, [string]$Level = 'INFO')
    Write-ActivityEntry -LogBox $TxtDuplicateLog -LogScroll $DuplicateLogScroll `
        -Expander $DuplicateActivityExpander -Summary $TxtDuplicateActivitySummary -Message $Message -Level $Level
}

function Clear-DuplicateReport {
    param([string]$Message)
    $script:DuplicateState.Report = $null
    $script:DuplicateState.Valid = $false
    $script:DuplicateState.CountsFor = $null
    $script:DuplicateState.UniqueDeviceCount = 0
    $script:DuplicateState.GroupCount = 0
    $script:DuplicateState.Message = $Message
    $ListDuplicateDevices.Items.Clear()
}

function Update-DuplicateDeviceState {
    $state = $script:DuplicateState
    if (-not $state -or $state.Refreshing) { return }
    if ($state.Report -and $state.Report.TenantId -ne $script:GraphState.ExpectedTenant) {
        Clear-DuplicateReport 'Tenant changed. Scan again before reviewing or deleting devices.'
        Write-DuplicateLog $state.Message 'WARN'
    }
    $selected = @($ListDuplicateDevices.SelectedItems)
    if ($state.Valid -and $state.Report -and -not [object]::ReferenceEquals($state.CountsFor, $state.Report)) {
        $deviceIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $groupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $state.Report.Rows) {
            [void]$deviceIds.Add($row.Id)
            [void]$groupIds.Add($row.GroupId)
        }
        $state.UniqueDeviceCount = $deviceIds.Count
        $state.GroupCount = $groupIds.Count
        $state.CountsFor = $state.Report
    }
    $selectedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $selected) { [void]$selectedIds.Add($row.Id) }
    $eligible = $state.Valid -and $selected.Count -gt 0 -and @($selected | Where-Object { -not $_.CanDelete }).Count -eq 0
    Set-ButtonAvailability -Button $BtnDuplicateScan -Available (-not $state.Busy)
    Set-ButtonAvailability -Button $BtnDuplicateExport -Available (-not $state.Busy -and $state.Valid -and $state.Report.Rows.Count -gt 0)
    Set-ButtonAvailability -Button $BtnDuplicateDelete -Available (-not $state.Busy -and $eligible)
    Set-ButtonAvailability -Button $BtnDuplicateCancel -Available ($state.Busy -and $null -ne $state.Cancel -and -not $state.CancelRequested)
    $TxtDuplicateCount.Text = if ($state.Valid) {
        "$($state.UniqueDeviceCount) unique devices / $($state.GroupCount) groups / $($state.Report.Rows.Count) membership rows ($($ListDuplicateDevices.Items.Count) visible); $($selectedIds.Count) selected. Tenant: $($state.Report.TenantId)"
    } else { $state.Message }
}

function Update-DuplicateDeviceFilter {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $state = $script:DuplicateState
    $state.Refreshing = $true
    try {
        $ListDuplicateDevices.UnselectAll()
        $ListDuplicateDevices.Items.Clear()
        if ($state.Report) {
            $type = [string]$ComboDuplicateMatch.SelectedItem.Tag
            $term = $TxtDuplicateFilter.Text.Trim()
            foreach ($row in $state.Report.Rows) {
                if ($type -ne 'All' -and $row.MatchType -ne $type) { continue }
                if ($term) {
                    $text = "$($row.DeviceName) $($row.Id) $($row.SerialNumber) $($row.UserPrincipalName) $($row.ProtectionReason)"
                    if ($text.IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                }
                [void]$ListDuplicateDevices.Items.Add($row)
            }
        }
    } finally { $state.Refreshing = $false }
    Update-DuplicateDeviceState
    if ($state.Valid) { Write-DuplicateLog "GUI report filtering/publication finished in $([math]::Round($timer.Elapsed.TotalSeconds, 2)) s ($($ListDuplicateDevices.Items.Count) visible membership rows)." 'INFO' }
}

function Start-IntuneDuplicateScan {
    if ($script:DuplicateState.Busy) { Write-DuplicateLog 'An operation is already running.' 'WARN'; return }
    Clear-DuplicateReport 'Scanning Windows devices...'
    $script:DuplicateState.Busy = $true
    $script:DuplicateState.ScanTimer = [Diagnostics.Stopwatch]::StartNew()
    $cancelHandle = @{ Stop = $null }
    $script:DuplicateState.Cancel = $cancelHandle
    $script:DuplicateState.CancelRequested = $false
    $bag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $outputAction = { param($report) $bag.Add($report) }.GetNewClosure()
    $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Get-IntuneDuplicateDevices.ps1' `
        -Parameters @{ AsReport = $true } -Scopes @('DeviceManagementManagedDevices.Read.All') -OutputAction $outputAction
    $state = $script:DuplicateState
    $graphState = $script:GraphState
    $onComplete = {
        $state.Busy = $false
        if ($state.ScanTimer) { $state.ScanTimer.Stop() }
        $wasCancelled = [bool]$state.CancelRequested
        $state.Cancel = $null
        $state.CancelRequested = $false
        $reports = @($bag.ToArray())
        if ($wasCancelled) {
            Clear-DuplicateReport 'Scan cancelled. No report is available; scan again to review devices.'
            Write-DuplicateLog $state.Message 'WARN'
        } elseif ($reports.Count -eq 1 -and $reports[0].TenantId -eq $graphState.ExpectedTenant) {
            $state.Report = $reports[0]
            $state.Valid = $true
            $state.Message = if ($state.Report.Rows.Count -eq 0) { 'Scan complete: no duplicate candidates found.' } else { 'Scan complete.' }
            Write-DuplicateLog $state.Message 'OK'
        } else {
            Clear-DuplicateReport 'Scan failed or tenant changed. No deletion is available; see Activity.'
        }
        Update-DuplicateDeviceFilter
    }.GetNewClosure()
    Update-DuplicateDeviceState
    Invoke-InBackground -Worker $worker -Buttons @($BtnDuplicateScan, $BtnDuplicateDelete, $BtnDuplicateExport, $BtnSignInOut) `
        -BusyLabel 'Scanning...' -LogBox $TxtDuplicateLog -LogScroll $DuplicateLogScroll `
        -DarkMode $script:IsDarkMode -OnComplete $onComplete `
        -CancelHandle $cancelHandle
}

$BtnDuplicateScan.Add_Click({ Start-IntuneDuplicateScan })
$BtnDuplicateCancel.Add_Click({
    $state = $script:DuplicateState
    if (-not $state.Busy -or -not $state.Cancel -or -not $state.Cancel.Stop) { return }
    $state.CancelRequested = $true
    Write-DuplicateLog 'Cancelling scan...' 'WARN'
    & $state.Cancel.Stop
    Update-DuplicateDeviceState
})
$TxtDuplicateFilter.Add_TextChanged({ Update-DuplicateDeviceFilter })
$ComboDuplicateMatch.Add_SelectionChanged({ Update-DuplicateDeviceFilter })
$ListDuplicateDevices.Add_SelectionChanged({ Update-DuplicateDeviceState })

$BtnDuplicateExport.Add_Click({
    if (-not $script:DuplicateState.Valid -or $script:DuplicateState.Busy) {
        Write-DuplicateLog 'A completed scan is required before exporting.' 'WARN'; return
    }
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Export complete duplicate-device report'
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.DefaultExt = 'csv'
    $dialog.FileName = "IntuneDuplicateDevices_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dialog.ShowDialog() -ne $true) { return }
    try {
        $script:DuplicateState.Report.Rows | Select-Object -Property $script:DuplicateExportColumns |
            Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-DuplicateLog "Exported all $($script:DuplicateState.Report.Rows.Count) membership rows to '$($dialog.FileName)'." 'OK'
    } catch { Write-DuplicateLog "Export failed: $($_.Exception.Message)" 'ERR' }
})

$BtnDuplicateDelete.Add_Click({
    Update-DuplicateDeviceState
    $selected = @($ListDuplicateDevices.SelectedItems)
    if (-not $script:DuplicateState.Valid -or $script:DuplicateState.Busy -or $selected.Count -eq 0 -or
        @($selected | Where-Object { -not $_.CanDelete }).Count) {
        Write-DuplicateLog 'Select only eligible stale strong-match records. Keepers and review candidates cannot be deleted.' 'WARN'
        return
    }
    $ids = @($selected.Id | Select-Object -Unique)
    $details = @($selected | Sort-Object Id -Unique | ForEach-Object { "$($_.DeviceName) [$($_.Id)]" }) -join "`n"
    $message = "Delete $($ids.Count) selected Intune managed-device record(s) in tenant $($script:DuplicateState.Report.TenantId)?`n`n$details`n`nOnly unambiguous non-keepers last synced at least 60 full days ago qualify. Fresh validation will run first.`n`nWARNING: Windows deletion triggers Retire, removes management, and can affect company data and BitLocker key protectors. Back up recovery keys and local administrator credentials first. Entra and Autopilot registrations will not be deleted.`n`nContinue?"
    $answer = [System.Windows.MessageBox]::Show($message, 'Confirm device deletion', [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning, [System.Windows.MessageBoxResult]::No)
    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { Write-DuplicateLog 'Deletion cancelled.' 'WARN'; return }
    $parameters = @{ Id = $ids; Scan = $script:DuplicateState.Report }
    $script:DuplicateState.Valid = $false
    $script:DuplicateState.Busy = $true
    $script:DuplicateState.Message = 'Revalidating selected records before deletion...'
    $ListDuplicateDevices.SelectedItems.Clear()
    $bag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $outputAction = { param($result) $bag.Add($result) }.GetNewClosure()
    $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Remove-IntuneDuplicateDevices.ps1' -Parameters $parameters `
        -Scopes @('DeviceManagementManagedDevices.ReadWrite.All') -OutputAction $outputAction
    $state = $script:DuplicateState
    $onComplete = {
        $state.Outcomes = @($bag.ToArray())
        $state.Busy = $false
        Write-DuplicateLog "Deletion run ended with $($state.Outcomes.Count) per-device outcomes. Refreshing inventory; inspect Activity for failures." 'INFO'
        Start-IntuneDuplicateScan
    }.GetNewClosure()
    Update-DuplicateDeviceState
    Invoke-InBackground -Worker $worker -Buttons @($BtnDuplicateDelete, $BtnDuplicateScan, $BtnDuplicateExport, $BtnSignInOut) `
        -BusyLabel 'Revalidating / deleting...' -LogBox $TxtDuplicateLog -LogScroll $DuplicateLogScroll `
        -DarkMode $script:IsDarkMode -OverlayMessage 'Revalidating and deleting selected Intune records...' -OnComplete $onComplete
})
Update-DuplicateDeviceState

#endregion Duplicate Devices

#region Audit Lookup

foreach ($controlName in @(
    'TxtAuditUser', 'TxtAuditTarget', 'DateAuditStart', 'DateAuditEnd', 'BtnAuditSearch', 'BtnAuditExport',
    'TxtAuditTimeZone', 'TxtAuditCount', 'ListAuditEvents', 'AuditDetailsExpander', 'TxtAuditDetails',
    'AuditActivityExpander', 'TxtAuditActivitySummary', 'AuditLogScroll', 'TxtAuditLog'
)) { Set-Variable -Name $controlName -Value $window.FindName($controlName) }
$DateAuditStart.Language = [Windows.Markup.XmlLanguage]::GetLanguage([Globalization.CultureInfo]::CurrentCulture.IetfLanguageTag)
$DateAuditEnd.Language = $DateAuditStart.Language
$DateAuditStart.SelectedDate = [datetime]::Today.AddDays(-6)
$DateAuditEnd.SelectedDate = [datetime]::Today
$TxtAuditTimeZone.Text = "Dates and displayed times use $([TimeZoneInfo]::Local.DisplayName). UTC timestamps are retained in details and CSV."
$script:AuditState = @{ Report = $null; Busy = $false; DateError = $false }
$script:AuditExportColumns = @(
    'TimestampUtc', 'OperationKind', 'Operation', 'ActivityType', 'HttpOperation', 'Result', 'InitiatedBy', 'Application',
    'UserPrincipalName', 'UserId', 'ApplicationId', 'TargetNames', 'Category', 'Component',
    'Id', 'CorrelationId', 'TenantId', 'ActorJson', 'ResourcesJson'
)
Enable-ListViewSorting -ListView $ListAuditEvents -TemplateSortPaths @{ 'Type' = 'OperationKind'; 'Operation' = 'Operation' }

function Write-AuditLog {
    param([string]$Message, [string]$Level = 'INFO')
    Write-ActivityEntry -LogBox $TxtAuditLog -LogScroll $AuditLogScroll `
        -Expander $AuditActivityExpander -Summary $TxtAuditActivitySummary -Message $Message -Level $Level
}

function Clear-IntuneAuditReport {
    param([string]$Message)
    $script:AuditState.Report = $null
    $ListAuditEvents.Items.Clear()
    $ListAuditEvents.Items.SortDescriptions.Clear()
    $ListAuditEvents.Tag.Property = $null
    $ListAuditEvents.Tag.Direction = $null
    for ($index = 0; $index -lt $ListAuditEvents.View.Columns.Count; $index++) {
        $ListAuditEvents.View.Columns[$index].Header = $ListAuditEvents.Tag.Headers[$index]
    }
    $TxtAuditDetails.Text = 'Select an event to inspect its actor, targets, and changes.'
    $BtnAuditExport.IsEnabled = $false
    $TxtAuditCount.Text = $Message
}

function Format-AuditDetailValue {
    param($Value)
    if ($null -eq $Value) { return 'Not supplied' }
    if ([string]$Value -eq '') { return '(empty)' }
    [string]$Value
}

function Get-IntuneAuditDetails {
    param([Parameter(Mandatory)]$Row)
    $event = $Row.Event
    $lines = @(
        "Event ID: $(Format-AuditDetailValue $Row.Id)"
        "Tenant ID: $(Format-AuditDetailValue $Row.TenantId)"
        "UTC timestamp: $($Row.TimestampUtc)"
        "Operation: $(Format-AuditDetailValue $Row.Operation)"
        "Change type: $(Format-AuditDetailValue $Row.OperationKind)"
        "Activity type: $(Format-AuditDetailValue $Row.ActivityType)"
        "HTTP operation: $(Format-AuditDetailValue $Row.HttpOperation)"
        "Result: $(Format-AuditDetailValue $Row.Result)"
        "Category: $(Format-AuditDetailValue $Row.Category)"
        "Component: $(Format-AuditDetailValue $Row.Component)"
        "Correlation ID: $(Format-AuditDetailValue $Row.CorrelationId)"
        ''
        'Actor (including identifiers and IP address):'
        $(if ($null -eq $event.actor) { 'Not supplied' } else { ConvertTo-Json -InputObject $event.actor -Depth 100 })
    )
    $resources = @($event.resources | Where-Object { $null -ne $_ })
    if (-not $resources.Count) { $lines += @('', 'Targets: Not supplied') }
    foreach ($resource in $resources) {
        $lines += @(
            ''
            "Target: $(Format-AuditDetailValue $resource.displayName)"
            "Resource ID: $(Format-AuditDetailValue $resource.resourceId)"
            "Type: $(Format-AuditDetailValue $resource.type)"
            "Audit resource type: $(Format-AuditDetailValue $resource.auditResourceType)"
        )
        $properties = @($resource.modifiedProperties | Where-Object { $null -ne $_ })
        if (-not $properties.Count) { $lines += 'Changed properties: Not supplied' }
        foreach ($property in $properties) {
            $lines += @(
                "Property: $(Format-AuditDetailValue $property.displayName)"
                "  Old value: $(Format-AuditDetailValue $property.oldValue)"
                "  New value: $(Format-AuditDetailValue $property.newValue)"
            )
        }
    }
    $lines -join [Environment]::NewLine
}

function Start-IntuneAuditSearch {
    if ($script:AuditState.Busy) { return }
    if ($script:AuditState.DateError) {
        $script:AuditState.DateError = $false
        Write-AuditLog 'An invalid date was reset by the calendar control. Review the dates before searching again.' 'WARN'
        return
    }
    $userTerm = $TxtAuditUser.Text.Trim()
    $targetTerm = $TxtAuditTarget.Text.Trim()
    if (-not $userTerm -and -not $targetTerm) { Write-AuditLog 'Enter a User (UPN or ID), a Target name, or both.' 'WARN'; return }
    $startDate = [datetime]::MinValue
    $endDate = [datetime]::MinValue
    if (-not [datetime]::TryParse($DateAuditStart.Text, [ref]$startDate) -or
        -not [datetime]::TryParse($DateAuditEnd.Text, [ref]$endDate) -or
        $startDate.Date -gt $endDate.Date -or $endDate.Date -eq [datetime]::MaxValue.Date) {
        Write-AuditLog 'Enter valid start/end dates, with start on or before end.' 'WARN'
        return
    }
    $state = $script:AuditState
    $state.Busy = $true
    Clear-IntuneAuditReport 'Searching all pages in the selected date range...'
    $parameters = @{ User = $userTerm; Target = $targetTerm; StartDate = $startDate.Date; EndDate = $endDate.Date; AsReport = $true }
    $outputAction = {
        param($Report)
        if ($Report.Completed -ne $true -or $null -eq $Report.Rows) { throw 'The audit worker did not return a complete report.' }
        $state.Report = $Report
    }.GetNewClosure()
    $worker = New-IntuneToolboxFeatureWorker -RelativePath 'Search-IntuneAuditEvents.ps1' `
        -Parameters $parameters -Scopes @('DeviceManagementApps.Read.All') -OutputAction $outputAction
    $list = $ListAuditEvents
    $count = $TxtAuditCount
    $export = $BtnAuditExport
    $onComplete = {
        $state.Busy = $false
        if ($null -eq $state.Report) { $count.Text = 'Search failed. See Activity for details.'; return }
        foreach ($row in $state.Report.Rows) { [void]$list.Items.Add($row) }
        $count.Text = if ($state.Report.Rows.Count) { "$($state.Report.Rows.Count) matching event(s), newest first." } else { 'No matching events in the selected date range.' }
        $export.IsEnabled = $state.Report.Rows.Count -gt 0
    }.GetNewClosure()
    Invoke-InBackground -Worker $worker -Buttons @($BtnAuditSearch) -BusyLabel 'Searching...' `
        -LogBox $TxtAuditLog -LogScroll $AuditLogScroll -OnComplete $onComplete
}

$BtnAuditSearch.Add_Click({ Start-IntuneAuditSearch })
foreach ($datePicker in @($DateAuditStart, $DateAuditEnd)) {
    $datePicker.Add_DateValidationError({
        param($sender, $eventArgs)
        $eventArgs.ThrowException = $false
        $script:AuditState.DateError = $true
        Write-AuditLog 'Enter a valid calendar date. The invalid date will be reset; review it before searching.' 'WARN'
    })
}
$ListAuditEvents.Add_SelectionChanged({
    if ($ListAuditEvents.SelectedItem) { $TxtAuditDetails.Text = Get-IntuneAuditDetails -Row $ListAuditEvents.SelectedItem }
    else { $TxtAuditDetails.Text = 'Select an event to inspect its actor, targets, and changes.' }
})
$BtnAuditExport.Add_Click({
    if ($script:AuditState.Busy -or $null -eq $script:AuditState.Report -or -not $script:AuditState.Report.Rows.Count) {
        Write-AuditLog 'Run a successful search with matching events before exporting.' 'WARN'
        return
    }
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Export Intune Audit Events'
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.DefaultExt = 'csv'
    $dialog.FileName = "IntuneAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dialog.ShowDialog() -ne $true) { return }
    try {
        $script:AuditState.Report.Rows | Select-Object -Property $script:AuditExportColumns |
            Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-AuditLog "Exported $($script:AuditState.Report.Rows.Count) audit event(s) to '$($dialog.FileName)'." 'OK'
    } catch { Write-AuditLog "Export failed: $($_.Exception.Message)" 'ERR' }
})

#endregion Audit Lookup

#region Shared — Sign In/Out

$script:IsDarkMode = $true

$BtnThemeToggle.Add_Click({
    $res = $window.Resources
    function Set-Brush { param($key, $hex)
        $brush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($hex))
        $brush.Freeze()
        $res[$key] = $brush
    }
    if ($script:IsDarkMode) {
        # Switch to Light
        Set-Brush 'BgBrush'        '#f1f5f9'
        Set-Brush 'InputBgBrush'   '#ffffff'
        Set-Brush 'TextBrush'      '#1e293b'
        Set-Brush 'SubTextBrush'   '#64748b'
        Set-Brush 'BorderBrush2'   '#cbd5e1'
        Set-Brush 'SurfaceBrush'   '#e8edf3'
        Set-Brush 'SecondaryBrush' '#dbe3ed'
        Set-Brush 'SecondaryBtnBrush' '#f8fafc'
        Set-Brush 'SecondaryBtnHover' '#e2e8f0'
        Set-Brush 'SelectionBackground' '#dbeafe'
        Set-Brush 'SelectionHoverBackground' '#bfdbfe'
        Set-Brush 'SelectionAccent' '#2563eb'
        Set-Brush 'SelectionForeground' '#1e293b'
        Set-Brush 'DisabledBackground' '#e5e7eb'
        Set-Brush 'DisabledBorder' '#cbd5e1'
        Set-Brush 'DisabledForeground' '#64748b'
        Set-Brush 'SuccessBrush' '#15803d'
        Set-Brush 'AttentionBrush' '#b45309'
        Set-Brush 'DangerBrush' '#b91c1c'
        Set-Brush 'DangerHoverBrush' '#dc2626'
        Set-Brush 'DangerPressedBrush' '#991b1b'
        Set-Brush 'AccentBrush'       '#2563eb'
        Set-Brush 'AccentHover'       '#1d4ed8'
        Set-Brush 'AuditCreateBrush' '#15803d'
        Set-Brush 'AuditDeleteBrush' '#b91c1c'
        Set-Brush 'AuditUpdateBrush' '#1d4ed8'
        Set-Brush 'AuditOtherBrush' '#b45309'
        $BtnThemeToggle.Content = 'Dark theme'
        $BtnThemeToggle.ToolTip = 'Switch to the dark theme'
        $script:IsDarkMode = $false
    } else {
        # Switch to Dark
        Set-Brush 'BgBrush'        '#191c22'
        Set-Brush 'InputBgBrush'   '#20252d'
        Set-Brush 'TextBrush'      '#e6e9ef'
        Set-Brush 'SubTextBrush'   '#9ca8b7'
        Set-Brush 'BorderBrush2'   '#39424e'
        Set-Brush 'SurfaceBrush'   '#252b34'
        Set-Brush 'SecondaryBrush' '#2c3541'
        Set-Brush 'SelectionBackground' '#20324b'
        Set-Brush 'SelectionHoverBackground' '#2a405d'
        Set-Brush 'SelectionAccent' '#3b82f6'
        Set-Brush 'SelectionForeground' '#e6e9ef'
        Set-Brush 'DisabledBackground' '#30343b'
        Set-Brush 'DisabledBorder' '#454b54'
        Set-Brush 'DisabledForeground' '#8b93a1'
        Set-Brush 'AccentBrush'       '#3b82f6'
        Set-Brush 'AccentHover'       '#2563eb'
        Set-Brush 'SecondaryBtnBrush' '#252b34'
        Set-Brush 'SecondaryBtnHover' '#303946'
        Set-Brush 'SuccessBrush' '#4ade80'
        Set-Brush 'AttentionBrush' '#fbbf24'
        Set-Brush 'DangerBrush' '#b42332'
        Set-Brush 'DangerHoverBrush' '#dc2626'
        Set-Brush 'DangerPressedBrush' '#7f1d1d'
        Set-Brush 'AuditCreateBrush' '#4ade80'
        Set-Brush 'AuditDeleteBrush' '#f87171'
        Set-Brush 'AuditUpdateBrush' '#60a5fa'
        Set-Brush 'AuditOtherBrush' '#fbbf24'
        $BtnThemeToggle.Content = 'Light theme'
        $BtnThemeToggle.ToolTip = 'Switch to the light theme'
        $script:IsDarkMode = $true
    }
})

$script:AppUI = @{}
foreach ($name in @(
    'TxtAppNameFilter', 'ComboAppPlatform', 'BtnAppLoad', 'TxtAppCount', 'ListApps', 'TxtAppEmpty',
    'PanelAppResultFilters', 'TxtAppResultName', 'TxtAppResultPublisher', 'ComboAppResultPlatform',
    'ComboAppResultType', 'ComboAppResultStatus', 'BtnAppClearResultFilters',
    'AppActivityExpander', 'TxtAppActivitySummary', 'AppLogScroll', 'TxtAppLog', 'PanelAppActions',
    'ComboAppIntent', 'ComboAppTarget', 'ComboAppIncludeExclude', 'TxtAppGroupSearch', 'BtnAppGroupSearch',
    'ListAppGroups', 'BtnAppQueue', 'ListAppQueuedGroups', 'BtnAppUnqueue', 'BtnAppClearQueue',
    'ComboAppFilterMode', 'BtnAppFilters', 'ComboAppFilter', 'BtnAppAdd', 'BtnAppReplace', 'BtnAppInspect',
    'TxtAppInspection', 'ListAppAssignments', 'TxtAppAssignmentDetails', 'BtnAppEdit', 'BtnAppUpdate', 'BtnAppRemove',
    'PanelAppGroupSearch', 'PanelAppFilter', 'PanelAppAssignmentDetails'
)) { $script:AppUI[$name] = $window.FindName($name) }
$script:AppState = @{
    Busy = $false; Tenant = ''; InspectedAppId = ''; EditingAssignmentId = ''; Revision = 0
    Inventory = @(); InventoryLoaded = $false; RefreshingFilters = $false; InventoryMessage = 'No apps loaded.'
    Results = [System.Collections.Generic.List[object]]::new()
    PlanRows = [System.Collections.Generic.List[object]]::new(); ReplaceParameters = $null
}
Enable-ListViewSorting -ListView $AppUI.ListApps

function Write-AppLog {
    param([string]$Message, [string]$Level = 'INFO')
    Write-ActivityEntry -LogBox $AppUI.TxtAppLog -LogScroll $AppUI.AppLogScroll `
        -Expander $AppUI.AppActivityExpander -Summary $AppUI.TxtAppActivitySummary -Message $Message -Level $Level
}

function Clear-AppInspection {
    $script:AppState.InspectedAppId = ''
    $script:AppState.EditingAssignmentId = ''
    $script:AppState.Revision++
    $AppUI.ListAppAssignments.Items.Clear()
    $AppUI.TxtAppAssignmentDetails.Text = ''
    $AppUI.TxtAppInspection.Text = 'Select one app and refresh assignments.'
    Show-Panel $AppUI.PanelAppAssignmentDetails $false
}

function Clear-AppManagement {
    Clear-AppInspection
    $script:AppState.Tenant = ''
    Reset-AppInventory
    foreach ($control in @('ListAppGroups', 'ListAppQueuedGroups', 'ComboAppFilter')) { $AppUI[$control].Items.Clear() }
    Update-AppState
}

function Update-AppResultFilterOptions {
    param([switch]$Reset)
    $script:AppState.RefreshingFilters = $true
    try {
        if ($Reset) {
            $AppUI.TxtAppResultName.Clear()
            $AppUI.TxtAppResultPublisher.Clear()
        }
        foreach ($option in @(
            @{ Control = 'ComboAppResultPlatform'; Property = 'Platform' },
            @{ Control = 'ComboAppResultType'; Property = 'AppType' },
            @{ Control = 'ComboAppResultStatus'; Property = 'AssignmentStatus' }
        )) {
            $combo = $AppUI[$option.Control]
            $selected = if ($Reset) { '' } else { [string](Get-SelectedTag $combo) }
            $values = @($script:AppState.Inventory | ForEach-Object { [string]$_.$($option.Property) } | Where-Object { $_ })
            # Retain an active status criterion even if a mutation changes its last matching row.
            if ($selected) { $values += $selected }
            $combo.Items.Clear()
            $all = [System.Windows.Controls.ComboBoxItem]::new()
            $all.Content = 'All'; $all.Tag = ''
            [void]$combo.Items.Add($all)
            $combo.SelectedItem = $all
            foreach ($value in @($values | Sort-Object -Unique)) {
                $item = [System.Windows.Controls.ComboBoxItem]::new()
                $item.Content = $value; $item.Tag = $value
                [void]$combo.Items.Add($item)
                if ($value -eq $selected) { $combo.SelectedItem = $item }
            }
        }
    } finally { $script:AppState.RefreshingFilters = $false }
}

function Update-AppResultFilter {
    param([switch]$PreserveSelection)
    if ($script:AppState.RefreshingFilters) { return }
    $selectedIds = @()
    if ($PreserveSelection) { $selectedIds = @($AppUI.ListApps.SelectedItems | ForEach-Object Id) }
    $sorts = @($AppUI.ListApps.Items.SortDescriptions)
    $name = $AppUI.TxtAppResultName.Text
    $publisher = $AppUI.TxtAppResultPublisher.Text
    $platform = Get-SelectedTag $AppUI.ComboAppResultPlatform
    $type = Get-SelectedTag $AppUI.ComboAppResultType
    $status = Get-SelectedTag $AppUI.ComboAppResultStatus
    Clear-AppInspection
    $AppUI.ListApps.UnselectAll()
    $AppUI.ListApps.Items.Clear()
    foreach ($row in $script:AppState.Inventory) {
        if ($name -and ([string]$row.DisplayName).IndexOf($name, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        if ($publisher -and ([string]$row.Publisher).IndexOf($publisher, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        if ($platform -and $row.Platform -ne $platform) { continue }
        if ($type -and $row.AppType -ne $type) { continue }
        if ($status -and $row.AssignmentStatus -ne $status) { continue }
        [void]$AppUI.ListApps.Items.Add($row)
        if ($row.Id -in $selectedIds) { [void]$AppUI.ListApps.SelectedItems.Add($row) }
    }
    $AppUI.ListApps.Items.SortDescriptions.Clear()
    foreach ($sort in $sorts) { [void]$AppUI.ListApps.Items.SortDescriptions.Add($sort) }
    Update-AppState
}

function Reset-AppInventory {
    param([string]$Message = 'No apps loaded.')
    $script:AppState.Inventory = @()
    $script:AppState.InventoryLoaded = $false
    $script:AppState.InventoryMessage = $Message
    Update-AppResultFilterOptions -Reset
    Update-AppResultFilter
}

function Test-AppFormReady {
    $groupTarget = (Get-SelectedTag $AppUI.ComboAppTarget) -eq 'Group'
    $filterMode = Get-SelectedTag $AppUI.ComboAppFilterMode
    return ((-not $groupTarget -or $AppUI.ListAppQueuedGroups.Items.Count -gt 0) -and
        ($filterMode -eq 'None' -or $null -ne $AppUI.ComboAppFilter.SelectedItem) -and
        ((Get-SelectedTag $AppUI.ComboAppIncludeExclude) -ne 'exclude' -or ($groupTarget -and $filterMode -eq 'None')))
}

function Update-AppState {
    $selected = @($AppUI.ListApps.SelectedItems)
    $assignments = @($AppUI.ListAppAssignments.SelectedItems)
    $inspected = $selected.Count -eq 1 -and $script:AppState.InspectedAppId -eq $selected[0].Id
    $direct = $assignments.Count -gt 0 -and @($assignments | Where-Object { -not $_.Editable }).Count -eq 0
    $idle = -not $script:AppState.Busy
    $AppUI.PanelAppResultFilters.IsEnabled = $idle
    $AppUI.TxtAppCount.Text = if ($script:AppState.InventoryLoaded) {
        "$($AppUI.ListApps.Items.Count) visible / $($script:AppState.Inventory.Count) loaded / $($selected.Count) selected. Assignment status is a Graph hint."
    } else { $script:AppState.InventoryMessage }
    $AppUI.TxtAppEmpty.Text = if (-not $script:AppState.InventoryLoaded) { $script:AppState.InventoryMessage } `
        elseif (-not $script:AppState.Inventory.Count) { 'No apps returned by the load criteria.' } else { 'No apps match the result filters.' }
    $AppUI.TxtAppEmpty.Visibility = if ($AppUI.ListApps.Items.Count) { 'Collapsed' } else { 'Visible' }
    $AppUI.BtnAppInspect.IsEnabled = $idle -and $selected.Count -eq 1
    $AppUI.BtnAppAdd.IsEnabled = $idle -and $selected.Count -gt 0 -and (Test-AppFormReady)
    $AppUI.BtnAppReplace.IsEnabled = $AppUI.BtnAppAdd.IsEnabled
    $supportedIntent = $assignments.Count -eq 1 -and $assignments[0].Intent -in @('required', 'available', 'uninstall')
    $supportedTarget = $assignments.Count -eq 1 -and
        ([string]$assignments[0].Assignment.target.'@odata.type' -replace '^#', '') -in @(
            'microsoft.graph.groupAssignmentTarget', 'microsoft.graph.exclusionGroupAssignmentTarget',
            'microsoft.graph.allLicensedUsersAssignmentTarget', 'microsoft.graph.allDevicesAssignmentTarget'
        )
    $AppUI.BtnAppEdit.IsEnabled = $idle -and $inspected -and $direct -and $supportedIntent -and $supportedTarget
    $AppUI.BtnAppUpdate.IsEnabled = $AppUI.BtnAppEdit.IsEnabled -and
        $script:AppState.EditingAssignmentId -eq $assignments[0].Id -and (Test-AppFormReady) -and
        ((Get-SelectedTag $AppUI.ComboAppTarget) -ne 'Group' -or $AppUI.ListAppQueuedGroups.Items.Count -eq 1)
    $AppUI.BtnAppRemove.IsEnabled = $idle -and $inspected -and $direct
}

function Start-AppWorker {
    param([string]$Feature, [hashtable]$Parameters, [string[]]$Scopes, [scriptblock]$OutputAction, [scriptblock]$Complete)
    if ($script:AppState.Busy) { Write-AppLog 'Wait for the current app operation.' 'WARN'; return }
    $script:AppState.Busy = $true
    $AppUI.PanelAppActions.IsEnabled = $false
    $AppUI.ListApps.IsEnabled = $false
    Update-AppState
    $worker = New-IntuneToolboxFeatureWorker -RelativePath $Feature -Parameters $Parameters -Scopes $Scopes -OutputAction $OutputAction
    $state = $script:AppState
    $controls = $AppUI
    $completion = {
        $state.Busy = $false
        $controls.PanelAppActions.IsEnabled = $true
        $controls.ListApps.IsEnabled = $true
        Update-AppState
        if ($Complete) { & $Complete }
    }.GetNewClosure()
    Invoke-InBackground -Worker $worker -Buttons @($AppUI.BtnAppLoad) -BusyLabel 'Working...' `
        -LogBox $AppUI.TxtAppLog -LogScroll $AppUI.AppLogScroll -OnComplete $completion
}

function Start-AppInspection {
    if ($AppUI.ListApps.SelectedItems.Count -ne 1) { Write-AppLog 'Select exactly one app.' 'WARN'; return }
    $app = $AppUI.ListApps.SelectedItem
    Clear-AppInspection
    $revision = $script:AppState.Revision
    $state = $script:AppState
    $controls = $AppUI
    $AppUI.TxtAppInspection.Text = 'Loading complete assignment collection...'
    Show-Panel $AppUI.PanelAppAssignmentDetails $true
    $output = {
        param($report)
        if ($state.Revision -ne $revision -or $controls.ListApps.SelectedItems.Count -ne 1 -or $controls.ListApps.SelectedItem.Id -ne $app.Id) { return }
        foreach ($row in $report.Rows) { [void]$controls.ListAppAssignments.Items.Add($row) }
        $state.InspectedAppId = $app.Id
        $state.Tenant = $report.TenantId
        $controls.TxtAppInspection.Text = "$($report.Rows.Count) assignments for $($app.DisplayName). Policy Set assignments are read-only."
    }.GetNewClosure()
    $complete = {
        if (-not $script:AppState.InspectedAppId) {
            $AppUI.TxtAppInspection.Text = 'Inspection failed or became stale. Refresh assignments; see Activity.'
            Show-Panel $AppUI.PanelAppAssignmentDetails $false
        }
        Update-AppState
    }
    Start-AppWorker -Feature 'Get-IntuneApps.ps1' -Parameters @{ Action = 'Assignments'; AppId = $app.Id; TenantId = $app.TenantId } `
        -Scopes @('DeviceManagementApps.Read.All') -OutputAction $output -Complete $complete
}

function Select-AppComboTag {
    param($Combo, [string]$Tag)
    foreach ($item in $Combo.Items) { if ($item.Tag -eq $Tag) { $Combo.SelectedItem = $item; return } }
    throw "Unsupported assignment option '$Tag'."
}

function Set-AppAssignmentForm {
    $row = $AppUI.ListAppAssignments.SelectedItem
    if (-not $row -or -not $row.Editable) { Write-AppLog 'Select a direct assignment.' 'WARN'; return }
    if ($row.Intent -notin @('required', 'available', 'uninstall')) { Write-AppLog "Intent '$($row.Intent)' is read-only in this editor." 'WARN'; return }
    $type = ([string]$row.Assignment.target.'@odata.type' -replace '^#', '')
    $targetType = switch ($type) {
        'microsoft.graph.groupAssignmentTarget' { 'Group' }
        'microsoft.graph.exclusionGroupAssignmentTarget' { 'Group' }
        'microsoft.graph.allLicensedUsersAssignmentTarget' { 'Users' }
        'microsoft.graph.allDevicesAssignmentTarget' { 'Devices' }
        default { Write-AppLog "Target '$type' is not supported by this editor." 'WARN'; return }
    }
    if (-not $targetType) { return }
    Select-AppComboTag $AppUI.ComboAppIntent $row.Intent
    Select-AppComboTag $AppUI.ComboAppTarget $targetType
    Select-AppComboTag $AppUI.ComboAppIncludeExclude $row.IncludeExclude
    $AppUI.ListAppQueuedGroups.Items.Clear()
    if ($row.GroupId) { [void]$AppUI.ListAppQueuedGroups.Items.Add([PSCustomObject]@{ Id = $row.GroupId; DisplayName = $row.TargetName }) }
    Select-AppComboTag $AppUI.ComboAppFilterMode $row.FilterMode
    $AppUI.ComboAppFilter.SelectedIndex = -1
    if ($row.FilterId) {
        $filter = @($AppUI.ComboAppFilter.Items | Where-Object Id -eq $row.FilterId | Select-Object -First 1)
        if (-not $filter.Count) {
            $filter = @([PSCustomObject]@{ Id = $row.FilterId; DisplayName = "Existing filter ($($row.FilterId))"; Platform = '' })
            [void]$AppUI.ComboAppFilter.Items.Add($filter[0])
        }
        $AppUI.ComboAppFilter.SelectedItem = $filter[0]
    }
    $script:AppState.EditingAssignmentId = $row.Id
    Update-AppState
}

function Invoke-AppAssignmentAction {
    param([ValidateSet('Add', 'Update', 'Remove', 'Replace')][string]$Action)
    $button = $AppUI["BtnApp$Action"]
    Update-AppState
    if (-not $button.IsEnabled) { Write-AppLog 'Select apps/assignments and valid target options before applying.' 'WARN'; return }
    $apps = @($AppUI.ListApps.SelectedItems)
    $parameters = @{
        Apps = $apps; Action = $Action; TenantId = $apps[0].TenantId
        Intent = Get-SelectedTag $AppUI.ComboAppIntent
        TargetType = Get-SelectedTag $AppUI.ComboAppTarget
        TargetGroup = @($AppUI.ListAppQueuedGroups.Items)
        IncludeExclude = Get-SelectedTag $AppUI.ComboAppIncludeExclude
        FilterMode = Get-SelectedTag $AppUI.ComboAppFilterMode
    }
    if ($parameters.FilterMode -ne 'None') { $parameters.AssignmentFilter = $AppUI.ComboAppFilter.SelectedItem }
    if ($Action -in @('Update', 'Remove')) { $parameters.Assignments = @($AppUI.ListAppAssignments.SelectedItems) }
    if ($Action -eq 'Replace') { Start-AppReplacePreview -Parameters $parameters; return }
    Write-AppLog "$Action confirmation requested for $($apps.Count) app(s):`n$(@($apps | ForEach-Object { "  $($_.DisplayName) [$($_.Id)]" }) -join "`n")"
    $targetDetail = if ($parameters.TargetType -eq 'Group') {
        (@($parameters.TargetGroup | ForEach-Object { "$($_.DisplayName) [$($_.Id)]" }) -join "`n")
    } else { "All $($parameters.TargetType)" }
    $details = if ($Action -eq 'Remove') {
        (@($parameters.Assignments | ForEach-Object { "$($_.Id): $($_.Intent) / $($_.TargetName)" }) -join "`n")
    } else {
        "$($parameters.Intent) / $($parameters.IncludeExclude)`n$targetDetail`nFilter: $($parameters.FilterMode) $($parameters.AssignmentFilter.Id)"
    }
    $warning = if ($Action -ne 'Remove' -and ($parameters.Intent -in @('required', 'uninstall') -or $parameters.TargetType -ne 'Group')) {
        "`nWARNING: Required/Uninstall or tenant-wide targets can affect many devices."
    } else { '' }
    if ($Action -eq 'Update') { $details = "Assignment ID: $($parameters.Assignments[0].Id)`n$details" }
    $message = "Tenant: $($apps[0].TenantId)`n$Action assignments for $($apps.Count) app(s) (listed in Activity).`n`n$details$warning`n`nChanges take effect in Intune immediately. Bulk operations are not atomic. Continue?"
    if (-not (Confirm-AppAssignment -Message $message -Action $Action)) {
        Write-AppLog 'App assignment operation cancelled.' 'WARN'; return
    }
    Start-AppMutation -Parameters $parameters
}

function Start-AppMutation {
    param([hashtable]$Parameters)
    $apps = @($Parameters.Apps)
    Clear-AppInspection
    $script:AppState.Results.Clear()
    $output = { param($row) $script:AppState.Results.Add($row) }
    $state = $script:AppState
    $controls = $AppUI
    $complete = {
        $failed = @($state.Results | Where-Object Status -eq 'Failed').Count
        Write-AppLog "Operation returned $($state.Results.Count) outcome(s), $failed failed. Successful targets are not rolled back." $(if ($failed) { 'WARN' } else { 'INFO' })
        foreach ($app in $apps) { $app.AssignmentStatus = 'Refresh required' }
        Update-AppResultFilterOptions
        Update-AppResultFilter -PreserveSelection
        if ($controls.ListApps.SelectedItems.Count -eq 1) { Start-AppInspection }
    }.GetNewClosure()
    Start-AppWorker -Feature 'Invoke-IntuneAppAssignment.ps1' -Parameters $Parameters -Scopes @('DeviceManagementApps.ReadWrite.All') `
        -OutputAction $output -Complete $complete
}

function Start-AppReplacePreview {
    param([hashtable]$Parameters)
    $script:AppState.ReplaceParameters = $Parameters
    $script:AppState.PlanRows.Clear()
    $preview = $Parameters.Clone()
    $preview.PlanOnly = $true
    Write-AppLog "Previewing Replace for $(@($Parameters.Apps).Count) app(s). Nothing is changed yet."
    Start-AppWorker -Feature 'Invoke-IntuneAppAssignment.ps1' -Parameters $preview -Scopes @('DeviceManagementApps.Read.All') `
        -OutputAction { param($row) $script:AppState.PlanRows.Add($row) } -Complete { Complete-AppReplacePreview }
}

function Format-AppReplacePreview {
    param([hashtable]$Parameters, [object[]]$Plan, [int]$MaxLines = 40)
    $symbols = @{ Add = '+ Add'; Update = '~ Update filter'; Remove = '- REMOVE'; NoChange = '= Unchanged'; Conflict = '! Conflict' }
    $targetDetail = if ($Parameters.TargetType -eq 'Group') {
        (@($Parameters.TargetGroup | ForEach-Object { "  $($_.DisplayName) [$($_.Id)]" }) -join "`n")
    } else { "  All $($Parameters.TargetType)" }
    $planned = @($Plan | Where-Object Status -eq 'Planned')
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($targetGroup in @($planned | Group-Object Operation, Target)) {
        $row = $targetGroup.Group[0]
        $appCount = @($targetGroup.Group | ForEach-Object AppId | Select-Object -Unique).Count
        $lines.Add("  $($symbols[$row.Operation]): $($row.Target) ($appCount app$(if ($appCount -ne 1) { 's' }))")
    }
    $shown = @($lines | Select-Object -First $MaxLines)
    if ($lines.Count -gt $MaxLines) { $shown += "... $($lines.Count - $MaxLines) more line(s); the full preview is in Activity." }
    $count = { param($operation) @($planned | Where-Object Operation -eq $operation).Count }
    $failed = @($Plan | Where-Object Status -eq 'Failed')
    $changedApps = @($planned | Where-Object Operation -in @('Add', 'Update', 'Remove') | ForEach-Object AppId | Select-Object -Unique).Count
    $text = "Tenant: $($Parameters.TenantId)`nReplace '$($Parameters.Intent)' $($Parameters.IncludeExclude) assignments with:`n$targetDetail`n" +
        "Filter: $($Parameters.FilterMode) $($Parameters.AssignmentFilter.Id)`n`nApps to change: $changedApps (per-app preview is in Activity)`n" +
        "$($shown -join "`n")`n`n" +
        "Totals: $(& $count 'Add') add, $(& $count 'Update') update, $(& $count 'Remove') REMOVE, $(& $count 'Conflict') conflict(s)."
    if ($failed.Count) {
        $text += "`nPreview failed for $(@($failed | ForEach-Object AppId | Select-Object -Unique).Count) app(s) (excluded; see Activity)."
    }
    $text + "`n`nAdds and updates run first. An app's removals are skipped if any of its adds, updates or conflicts are not resolved. " +
        "Changes take effect in Intune immediately and are not atomic. Continue?"
}

function Complete-AppReplacePreview {
    $parameters = $script:AppState.ReplaceParameters
    $script:AppState.ReplaceParameters = $null
    $plan = @($script:AppState.PlanRows)
    if (-not $parameters) { return }
    if (-not $plan.Count) { Write-AppLog 'Replace preview did not complete. Nothing was changed; see Activity.' 'WARN'; return }
    $changedAppIds = @($plan | Where-Object { $_.Status -eq 'Planned' -and $_.Operation -in @('Add', 'Update', 'Remove') } | ForEach-Object AppId | Select-Object -Unique)
    if (-not $changedAppIds.Count) { Write-AppLog 'Replace preview found nothing to change. Nothing was written.' 'INFO'; return }
    $symbols = @{ Add = '+ Add'; Update = '~ Update filter'; Remove = '- REMOVE'; NoChange = '= Unchanged'; Conflict = '! Conflict' }
    $previewLines = foreach ($appGroup in @($plan | Group-Object AppId)) {
        "$($appGroup.Group[0].AppName) [$($appGroup.Name)]"
        foreach ($row in $appGroup.Group) {
            $label = if ($row.Status -eq 'Failed') { 'Preview failed' } else { $symbols[$row.Operation] }
            $suffix = if ($row.Status -eq 'Failed' -or $row.Operation -eq 'Conflict') { " - $($row.Message)" } else { '' }
            "  ${label}: $($row.Target)$suffix"
        }
    }
    Write-AppLog "Replace preview per app:`n$(@($previewLines) -join "`n")"
    $message = Format-AppReplacePreview -Parameters $parameters -Plan $plan
    if (-not (Confirm-AppAssignment -Message $message -Action 'Replace')) {
        Write-AppLog 'Replace cancelled after preview. Nothing was changed.' 'WARN'; return
    }
    $execute = $parameters.Clone()
    # Only apps that were previewed successfully are changed; the plan pins what may be updated or removed.
    $execute.Apps = @($parameters.Apps | Where-Object { $_.Id -in $changedAppIds })
    $execute.Plan = $plan
    Start-AppMutation -Parameters $execute
}

function Confirm-AppAssignment {
    param([string]$Message, [string]$Action)
    if ($ValidateOnly) { return $false }
    return [System.Windows.MessageBox]::Show($window, $Message, "Confirm app assignment $Action", 'YesNo', 'Warning', 'No') -eq 'Yes'
}

$AppUI.BtnAppLoad.Add_Click({
    if ($script:AppState.Busy) { Write-AppLog 'Wait for the current app operation.' 'WARN'; return }
    Reset-AppInventory -Message 'Loading apps...'
    $output = {
        param($report)
        $script:AppState.Tenant = $report.TenantId
        $script:AppState.Inventory = @($report.Rows)
        $script:AppState.InventoryLoaded = $true
        Update-AppResultFilterOptions -Reset
        Update-AppResultFilter
    }
    $complete = {
        if (-not $script:AppState.InventoryLoaded) {
            $script:AppState.InventoryMessage = 'App load failed. See Activity and try Load apps again.'
            Update-AppState
        }
    }
    Start-AppWorker -Feature 'Get-IntuneApps.ps1' `
        -Parameters @{ NameFilter = $AppUI.TxtAppNameFilter.Text; PlatformFilter = Get-SelectedTag $AppUI.ComboAppPlatform } `
        -Scopes @('DeviceManagementApps.Read.All') -OutputAction $output -Complete $complete
})
foreach ($name in @('TxtAppResultName', 'TxtAppResultPublisher')) {
    $AppUI[$name].Add_TextChanged({ Update-AppResultFilter })
}
foreach ($name in @('ComboAppResultPlatform', 'ComboAppResultType', 'ComboAppResultStatus')) {
    $AppUI[$name].Add_SelectionChanged({ Update-AppResultFilter })
}
$AppUI.BtnAppClearResultFilters.Add_Click({ Update-AppResultFilterOptions -Reset; Update-AppResultFilter })
$AppUI.ListApps.Add_SelectionChanged({ Clear-AppInspection; Update-AppState })
$AppUI.ListAppAssignments.Add_SelectionChanged({
    $script:AppState.EditingAssignmentId = ''
    $AppUI.TxtAppAssignmentDetails.Text = if ($AppUI.ListAppAssignments.SelectedItems.Count -eq 1) {
        $row = $AppUI.ListAppAssignments.SelectedItem
        "ID: $($row.Id)`nTarget: $($row.TargetName) [$($row.GroupId)]`nFilter: $($row.FilterMode) $($row.FilterId)`nSource: $($row.Source)`nSettings:`n$($row.Assignment.settings | ConvertTo-Json -Depth 50)"
    } else { '' }
    Update-AppState
})
$AppUI.BtnAppInspect.Add_Click({ Start-AppInspection })
$AppUI.BtnAppGroupSearch.Add_Click({
    $term = $AppUI.TxtAppGroupSearch.Text.Trim()
    if (-not $term) { Write-AppLog 'Enter a group search term.' 'WARN'; return }
    $AppUI.ListAppGroups.Items.Clear()
    Start-AppWorker -Feature 'Search-IntuneAssignmentTargets.ps1' -Parameters @{ Action = 'Groups'; SearchTerm = $term } `
        -Scopes @('Group.Read.All') -OutputAction { param($row) [void]$AppUI.ListAppGroups.Items.Add($row) }
})
$AppUI.BtnAppFilters.Add_Click({
    $AppUI.ComboAppFilter.Items.Clear()
    Start-AppWorker -Feature 'Search-IntuneAssignmentTargets.ps1' -Parameters @{ Action = 'Filters' } `
        -Scopes @('DeviceManagementConfiguration.Read.All') -OutputAction {
            param($row)
            if ($row.ManagementType -ne 'apps') { [void]$AppUI.ComboAppFilter.Items.Add($row) }
        }
})
$AppUI.BtnAppQueue.Add_Click({ Add-AssignmentQueueGroups -Results $AppUI.ListAppGroups -Queue $AppUI.ListAppQueuedGroups; Update-AppState })
$AppUI.BtnAppUnqueue.Add_Click({ Remove-AssignmentQueueGroups -Queue $AppUI.ListAppQueuedGroups; Update-AppState })
$AppUI.BtnAppClearQueue.Add_Click({ $AppUI.ListAppQueuedGroups.Items.Clear(); Update-AppState })
$AppUI.BtnAppEdit.Add_Click({ Set-AppAssignmentForm })
$AppUI.BtnAppAdd.Add_Click({ Invoke-AppAssignmentAction Add })
$AppUI.BtnAppReplace.Add_Click({ Invoke-AppAssignmentAction Replace })
$AppUI.BtnAppUpdate.Add_Click({ Invoke-AppAssignmentAction Update })
$AppUI.BtnAppRemove.Add_Click({ Invoke-AppAssignmentAction Remove })
foreach ($name in @('ComboAppIntent', 'ComboAppTarget', 'ComboAppIncludeExclude', 'ComboAppFilterMode', 'ComboAppFilter')) {
    $AppUI[$name].Add_SelectionChanged({ Update-AppState })
}
$AppUI.ComboAppTarget.Add_SelectionChanged({ Show-Panel $AppUI.PanelAppGroupSearch ((Get-SelectedTag $AppUI.ComboAppTarget) -eq 'Group') })
$AppUI.ComboAppFilterMode.Add_SelectionChanged({ Show-Panel $AppUI.PanelAppFilter ((Get-SelectedTag $AppUI.ComboAppFilterMode) -ne 'None') })
Reset-AppInventory

$BtnSignInOut.Add_Click({
    if ($script:AppState.Busy) { Write-AppLog 'Wait for the app operation before changing sign-in.' 'WARN'; return }
    if ($script:DuplicateState.Busy) { Write-DuplicateLog 'Wait for the duplicate-device operation before changing sign-in.' 'WARN'; return }
    if ($script:AuditState.Busy) { Write-AuditLog 'Wait for the audit search before changing sign-in.' 'WARN'; return }
    if ($script:IsConnected) {
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
        } catch { }
        $script:IsConnected   = $false
        $TxtTenantName.Text   = ""
        $TxtTenantName.ToolTip = ''
        $script:GraphUser = ''
        $script:GraphTenantId = ''
        $script:GraphState.Connected = $false
        $script:GraphState.ExpectedTenant = ''
        $script:GraphState.Tenant = ''
        $script:GraphState.Account = ''
        Clear-DuplicateReport 'Signed out. Scan again after reconnecting.'
        Clear-IntuneAuditReport 'Signed out. Search again after reconnecting.'
        Clear-AppManagement
        Update-DuplicateDeviceState
        $TxtGraphConnectionStatus.Text = "Not connected"
        $BtnSignInOut.Content = "Sign in"
        $EllipseStatus.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'DangerBrush')
        Write-ProfileLog   "Signed out from Microsoft Graph." "INFO"
        Write-AutopilotLog "Signed out from Microsoft Graph." "INFO"
    } else {
        Write-ProfileLog   "Signing in to Microsoft Graph..." "INFO"
        Write-AutopilotLog "Signing in to Microsoft Graph..." "INFO"
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            # Request every permission the toolbox uses so later actions do not need another sign-in.
            $connectParameters = @{
                Scopes = $script:ToolboxGraphScopes
                NoWelcome = $true
                ContextScope = 'Process'
                ErrorAction = 'Stop'
            }
            if ($script:GraphState.RequestedTenant) { $connectParameters.TenantId = $script:GraphState.RequestedTenant }
            Connect-MgGraph @connectParameters | Out-Null
            Update-TenantInfo
            Update-DuplicateDeviceState
            $tidLabel = $script:GraphTenantId
            Write-ProfileLog   "Signed in to tenant $tidLabel." "OK"
            Write-AutopilotLog "Signed in to tenant $tidLabel." "OK"
        } catch {
            $TxtGraphConnectionStatus.Text = "Connection failed"
            $EllipseStatus.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#dc2626'))
            Write-ProfileLog   "Sign in failed: $_" "ERR"
            Write-AutopilotLog "Sign in failed: $_" "ERR"
        }
    }
})

#endregion Shared

if ($ValidateOnly) {
    $featurePaths = @(
        'Features\Get-IntuneProfiles.ps1',
        'Features\Get-IntuneApps.ps1',
        'Features\Invoke-IntuneAppAssignment.ps1',
        'Shared\IntuneAppHelpers.ps1',
        'Features\Invoke-IntuneProfileAssignment.ps1',
        'Features\Rename-IntuneProfiles.ps1',
        'Features\Remove-IntuneProfiles.ps1',
        'Features\Invoke-AutopilotGroupTagOperation.ps1',
        'Features\Search-IntuneAssignmentTargets.ps1',
        'Features\Search-IntuneSettings.ps1',
        'Features\Search-IntuneAuditEvents.ps1',
        'Features\Get-IntuneDuplicateDevices.ps1',
        'Features\Remove-IntuneDuplicateDevices.ps1',
        'Shared\Get-IntuneDuplicateDeviceAnalysis.ps1'
    ) | ForEach-Object { Join-Path $PSScriptRoot $_ }
    $missingFeatures = @($featurePaths | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    [PSCustomObject]@{ Window = $window; FeaturePaths = $featurePaths; MissingFeatures = $missingFeatures; Connected = $script:IsConnected }
    return
}

[void]$window.ShowDialog()
