<#
==============================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2025/08/18
  UPDATED: 2026/02/19 - WPF GUI version

  Script: New User Provisioning - GUI Version
  Scope : WPF form replaces all Read-Host prompts. Creates AD user,
          syncs to Entra ID, sets UsageLocation, assigns M365 license.
          Live progress log displayed in-window. Logs written to .\Logs\
==============================================================================
#>

#Requires -Modules ActiveDirectory

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# --- Config ---
$PreferredSkuOrder              = @('SPE_E3','ENTERPRISEPACK')
$InitialSyncWaitSeconds         = 300
$UserDiscoveryRetryCount        = 60
$UserDiscoveryRetryDelaySeconds = 30
$LogRetentionDays               = 90
$UPNSuffix                      = "@hanskissle.com"

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
$LogDir    = Join-Path $ScriptDir 'Logs'
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

try {
    Get-ChildItem -Path $LogDir -File -Include *.log,*.txt -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
} catch {}

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogPath   = Join-Path $LogDir ("NewUserProvisioning_{0}.log" -f $Timestamp)

$OUs = @(
    "OU=Accounting,OU=Administration,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Customer Service,OU=Administration,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Human Resources,OU=Administration,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Information Technology,OU=Administration,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Purchasing,OU=Administration,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=R&D,OU=Administration,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Sales,OU=Administration,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Maintenance,OU=Production,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Production,OU=Production,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Quality Control,OU=Production,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Receiving,OU=Production,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Shipping,OU=Production,OU=Employees,OU=Hans Kissle North,DC=hk,DC=lan",
    "OU=Maintenance,OU=Users,OU=Hans Kissle South-Dallas,DC=hk,DC=lan",
    "OU=Office,OU=Users,OU=Hans Kissle South-Dallas,DC=hk,DC=lan",
    "OU=Production,OU=Users,OU=Hans Kissle South-Dallas,DC=hk,DC=lan",
    "OU=Quality Control,OU=Users,OU=Hans Kissle South-Dallas,DC=hk,DC=lan",
    "OU=Ship-Rec,OU=Users,OU=Hans Kissle South-Dallas,DC=hk,DC=lan"
)

function Get-OUFriendlyName {
    param([string]$DN)
    $parts = $DN -split ',' | Where-Object { $_ -match '^OU=' }
    ($parts | ForEach-Object { $_ -replace '^OU=','' }) -join ' > '
}

[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="New User Provisioning - Hans Kissle"
    Height="820" Width="780"
    MinHeight="820" MinWidth="780"
    WindowStartupLocation="CenterScreen"
    Background="#1E1E2E"
    FontFamily="Segoe UI"
    FontSize="13">
    <Window.Resources>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="0,4,0,2"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#313244"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="BorderBrush" Value="#45475A"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="CaretBrush" Value="#CDD6F4"/>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="#313244"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="BorderBrush" Value="#45475A"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Height" Value="28"/>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#89B4FA"/>
            <Setter Property="Foreground" Value="#1E1E2E"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Height" Value="36"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="16,0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#B4D0FF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#45475A"/>
                                <Setter Property="Foreground" Value="#6C7086"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="#45475A"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Height" Value="36"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="16,0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#585B70"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DropdownCombo" TargetType="ComboBox">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#111111"/>
            <Setter Property="BorderBrush" Value="#45475A"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Background="#FFFFFF" BorderBrush="#45475A"
                                          BorderThickness="1" Foreground="#111111"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          Padding="6,4">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="20"/>
                                    </Grid.ColumnDefinitions>
                                    <ContentPresenter Grid.Column="0"
                                        Content="{TemplateBinding SelectionBoxItem}"
                                        ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                        VerticalAlignment="Center" Margin="2,0,0,0"/>
                                    <TextBlock Grid.Column="1" Text="v" Foreground="#111111"
                                               VerticalAlignment="Center" HorizontalAlignment="Center"
                                               FontSize="10"/>
                                </Grid>
                            </ToggleButton>
                            <Popup IsOpen="{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}}"
                                   Placement="Bottom" AllowsTransparency="True">
                                <Border Background="#FFFFFF" BorderBrush="#89B4FA" BorderThickness="1" MaxHeight="200">
                                    <ScrollViewer>
                                        <ItemsPresenter/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#111111"/>
            <Setter Property="Padding" Value="6,4"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#CDD6F4"/>
                    <Setter Property="Foreground" Value="#1E1E2E"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#CDD6F4"/>
                    <Setter Property="Foreground" Value="#1E1E2E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#111111"/>
            <Setter Property="Padding" Value="6,4"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#CDD6F4"/>
                    <Setter Property="Foreground" Value="#1E1E2E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Margin="0,0,0,16">
            <TextBlock Text="New User Provisioning" FontSize="20" FontWeight="Bold"
                       Foreground="#89B4FA" Margin="0,0,0,4"/>
            <TextBlock Text="Hans Kissle - Active Directory + Microsoft 365"
                       FontSize="12" Foreground="#6C7086"/>
            <Separator Background="#45475A" Margin="0,10,0,0"/>
        </StackPanel>
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <StackPanel Grid.Row="0" Grid.Column="0">
                <Label Content="First Name *"/>
                <TextBox x:Name="txtFirstName"/>
            </StackPanel>
            <StackPanel Grid.Row="0" Grid.Column="2">
                <Label Content="Last Name *"/>
                <TextBox x:Name="txtLastName"/>
            </StackPanel>
            <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,8,0,0">
                <Label Content="Username * (e.g. jdoe)"/>
                <TextBox x:Name="txtUsername"/>
            </StackPanel>
            <StackPanel Grid.Row="1" Grid.Column="2" Margin="0,8,0,0">
                <Label Content="Initial Password *"/>
                <PasswordBox x:Name="pwdPassword"/>
            </StackPanel>
            <StackPanel Grid.Row="2" Grid.Column="0" Margin="0,8,0,0">
                <Label Content="Email (optional)"/>
                <TextBox x:Name="txtEmail"/>
            </StackPanel>
            <StackPanel Grid.Row="2" Grid.Column="2" Margin="0,8,0,0">
                <Label Content="Telephone (optional)"/>
                <TextBox x:Name="txtPhone"/>
            </StackPanel>
            <StackPanel Grid.Row="3" Grid.Column="0" Margin="0,8,0,0">
                <Label Content="Job Title (optional)"/>
                <TextBox x:Name="txtTitle"/>
            </StackPanel>
            <StackPanel Grid.Row="3" Grid.Column="2" Margin="0,8,0,0">
                <Label Content="Department (optional)"/>
                <TextBox x:Name="txtDepartment"/>
            </StackPanel>
            <StackPanel Grid.Row="4" Grid.Column="0" Margin="0,8,0,0">
                <Label Content="Company (optional)"/>
                <TextBox x:Name="txtCompany"/>
            </StackPanel>
            <StackPanel Grid.Row="4" Grid.Column="2" Margin="0,8,0,0">
                <Label Content="Office (optional)"/>
                <TextBox x:Name="txtOffice"/>
            </StackPanel>
            <StackPanel Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,8,0,0">
                <Label Content="Description (optional)"/>
                <TextBox x:Name="txtDescription"/>
            </StackPanel>
            <StackPanel Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,8,0,0">
                <Label Content="Organizational Unit (OU) *"/>
                <ComboBox x:Name="cboOU" Style="{StaticResource DropdownCombo}"/>
            </StackPanel>
            <StackPanel Grid.Row="7" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,8,0,0">
                <Label Content="Manager (optional) - type to search AD"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="txtManagerSearch" Grid.Column="0"/>
                    <Button x:Name="btnSearchManager" Grid.Column="2"
                            Content="Search" Width="80" Height="28"
                            Style="{StaticResource SecondaryButton}"/>
                </Grid>
                <Border x:Name="borderManagerResults" Background="#FFFFFF"
                        BorderBrush="#89B4FA" BorderThickness="1"
                        CornerRadius="4" Margin="0,2,0,0"
                        Visibility="Collapsed" MaxHeight="120">
                    <ListBox x:Name="lstManagerResults"
                             Background="Transparent"
                             Foreground="#111111"
                             BorderThickness="0"
                             MaxHeight="120"/>
                </Border>
                <TextBlock x:Name="lblManagerSelected" Foreground="#A6E3A1"
                           FontSize="11" Margin="2,4,0,0" Text="" Visibility="Collapsed"/>
            </StackPanel>
            <StackPanel Grid.Row="8" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,8,0,0">
                <Label Content="License Assignment"/>
                <ComboBox x:Name="cboLicense" Style="{StaticResource DropdownCombo}">
                    <ComboBoxItem Content="Auto-assign preferred (M365 E3 / O365 E3)" IsSelected="True"/>
                    <ComboBoxItem Content="Prompt me to choose after sync"/>
                </ComboBox>
            </StackPanel>
        </Grid>
        <Grid Grid.Row="2" Margin="0,16,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="Progress Log" FontSize="12" FontWeight="SemiBold"
                       Foreground="#6C7086" Margin="0,0,0,4"/>
            <Border Grid.Row="1" Background="#181825" BorderBrush="#45475A"
                    BorderThickness="1" CornerRadius="4">
                <ScrollViewer x:Name="logScroller" VerticalScrollBarVisibility="Auto">
                    <TextBlock x:Name="txtLog"
                               Foreground="#CDD6F4"
                               FontFamily="Consolas"
                               FontSize="11"
                               Padding="8"
                               TextWrapping="Wrap"/>
                </ScrollViewer>
            </Border>
        </Grid>
        <Grid Grid.Row="3" Margin="0,12,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="lblStatus" Grid.Column="0"
                       VerticalAlignment="Center" FontSize="12"
                       Foreground="#6C7086" Text="Ready"/>
            <Button x:Name="btnClear" Grid.Column="2" Content="Clear Form"
                    Width="110" Style="{StaticResource SecondaryButton}"/>
            <Button x:Name="btnProvision" Grid.Column="4" Content="Provision User"
                    Width="150" Style="{StaticResource PrimaryButton}"/>
        </Grid>
    </Grid>
</Window>
"@

$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)

$txtFirstName         = $Window.FindName('txtFirstName')
$txtLastName          = $Window.FindName('txtLastName')
$txtUsername          = $Window.FindName('txtUsername')
$pwdPassword          = $Window.FindName('pwdPassword')
$txtEmail             = $Window.FindName('txtEmail')
$txtPhone             = $Window.FindName('txtPhone')
$txtTitle             = $Window.FindName('txtTitle')
$txtDepartment        = $Window.FindName('txtDepartment')
$txtCompany           = $Window.FindName('txtCompany')
$txtOffice            = $Window.FindName('txtOffice')
$txtDescription       = $Window.FindName('txtDescription')
$cboOU                = $Window.FindName('cboOU')
$txtManagerSearch     = $Window.FindName('txtManagerSearch')
$btnSearchManager     = $Window.FindName('btnSearchManager')
$lstManagerResults    = $Window.FindName('lstManagerResults')
$borderManagerResults = $Window.FindName('borderManagerResults')
$lblManagerSelected   = $Window.FindName('lblManagerSelected')
$cboLicense           = $Window.FindName('cboLicense')
$txtLog               = $Window.FindName('txtLog')
$logScroller          = $Window.FindName('logScroller')
$btnProvision         = $Window.FindName('btnProvision')
$btnClear             = $Window.FindName('btnClear')
$lblStatus            = $Window.FindName('lblStatus')

$script:SelectedManagerDN   = $null
$script:SelectedManagerName = $null

# --- Populate OU dropdown (hardcoded list matching original script) ---
foreach ($ou in $OUs) {
    $item         = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = Get-OUFriendlyName $ou
    $item.Tag     = $ou
    $cboOU.Items.Add($item) | Out-Null
}
$cboOU.SelectedIndex = 0

$btnSearchManager.Add_Click({
    $query = $txtManagerSearch.Text.Trim()
    if (-not $query) {
        [System.Windows.MessageBox]::Show("Please enter a name to search.", "Search", "OK", "Warning")
        return
    }
    $lstManagerResults.Items.Clear()
    $script:SelectedManagerDN   = $null
    $script:SelectedManagerName = $null
    $lblManagerSelected.Visibility = 'Collapsed'
    try {
        $escaped = ($query -replace '\\','\\') -replace "'","''"
        $users = Get-ADUser -Filter "Name -like '*$escaped*'" -Properties SamAccountName, UserPrincipalName -ErrorAction Stop |
                 Sort-Object Name | Select-Object -First 30
        if (-not $users) {
            [System.Windows.MessageBox]::Show("No users found matching '$query'.", "Search Results", "OK", "Information")
            $borderManagerResults.Visibility = 'Collapsed'
            return
        }
        foreach ($u in $users) {
            $item         = New-Object System.Windows.Controls.ListBoxItem
            $item.Content = "{0}  ({1})" -f $u.Name, $u.SamAccountName
            $item.Tag     = $u
            $lstManagerResults.Items.Add($item) | Out-Null
        }
        $borderManagerResults.Visibility = 'Visible'
    } catch {
        [System.Windows.MessageBox]::Show("AD search failed: $($_.Exception.Message)", "Error", "OK", "Error")
    }
})

$lstManagerResults.Add_SelectionChanged({
    $sel = $lstManagerResults.SelectedItem
    if ($sel) {
        $u = $sel.Tag
        $script:SelectedManagerDN   = $u.DistinguishedName
        $script:SelectedManagerName = $u.Name
        $txtManagerSearch.Text           = $u.Name
        $lblManagerSelected.Text         = "Selected: $($u.Name) ($($u.SamAccountName))"
        $lblManagerSelected.Visibility   = 'Visible'
        $borderManagerResults.Visibility = 'Collapsed'
    }
})

$txtManagerSearch.Add_KeyDown({
    if ($_.Key -eq 'Return') {
        $btnSearchManager.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
})

$btnClear.Add_Click({
    $txtFirstName.Text     = ''
    $txtLastName.Text      = ''
    $txtUsername.Text      = ''
    $pwdPassword.Password  = ''
    $txtEmail.Text         = ''
    $txtPhone.Text         = ''
    $txtTitle.Text         = ''
    $txtDepartment.Text    = ''
    $txtCompany.Text       = ''
    $txtOffice.Text        = ''
    $txtDescription.Text   = ''
    $txtManagerSearch.Text = ''
    $txtLog.Inlines.Clear()
    $lblManagerSelected.Visibility   = 'Collapsed'
    $borderManagerResults.Visibility = 'Collapsed'
    $script:SelectedManagerDN        = $null
    $script:SelectedManagerName      = $null
    $cboOU.SelectedIndex             = 0
    $cboLicense.SelectedIndex        = 0
    $lblStatus.Text                  = 'Ready'
    $lblStatus.Foreground            = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#6C7086')
    $btnProvision.IsEnabled          = $true
})

$btnProvision.Add_Click({

    $errs = @()
    if (-not $txtFirstName.Text.Trim()) { $errs += "First Name is required." }
    if (-not $txtLastName.Text.Trim())  { $errs += "Last Name is required." }
    if (-not $txtUsername.Text.Trim())  { $errs += "Username is required." }
    if (-not $pwdPassword.Password)     { $errs += "Password is required." }
    if ($cboOU.SelectedItem -eq $null)  { $errs += "Please select an OU." }
    if ($errs.Count -gt 0) {
        [System.Windows.MessageBox]::Show(($errs -join "`n"), "Validation Error", "OK", "Warning")
        return
    }

    $FirstName   = $txtFirstName.Text.Trim()
    $LastName    = $txtLastName.Text.Trim()
    $Username    = $txtUsername.Text.Trim()
    $Password    = $pwdPassword.Password
    $Email       = $txtEmail.Text.Trim()
    $Phone       = $txtPhone.Text.Trim()
    $Title       = $txtTitle.Text.Trim()
    $Department  = $txtDepartment.Text.Trim()
    $Company     = $txtCompany.Text.Trim()
    $Office      = $txtOffice.Text.Trim()
    $Description = $txtDescription.Text.Trim()
    $SelectedOU  = $cboOU.SelectedItem.Tag
    $ManagerDN   = $script:SelectedManagerDN
    $AutoLicense = ($cboLicense.SelectedIndex -eq 0)
    $UPN         = "$Username$UPNSuffix"

    $btnProvision.IsEnabled = $false
    $lblStatus.Text         = 'Provisioning...'
    $lblStatus.Foreground   = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#89B4FA')
    $txtLog.Inlines.Clear()

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()

    $rs.SessionStateProxy.SetVariable('FirstName',                  $FirstName)
    $rs.SessionStateProxy.SetVariable('LastName',                   $LastName)
    $rs.SessionStateProxy.SetVariable('Username',                   $Username)
    $rs.SessionStateProxy.SetVariable('Password',                   $Password)
    $rs.SessionStateProxy.SetVariable('Email',                      $Email)
    $rs.SessionStateProxy.SetVariable('Phone',                      $Phone)
    $rs.SessionStateProxy.SetVariable('Title',                      $Title)
    $rs.SessionStateProxy.SetVariable('Department',                 $Department)
    $rs.SessionStateProxy.SetVariable('Company',                    $Company)
    $rs.SessionStateProxy.SetVariable('Office',                     $Office)
    $rs.SessionStateProxy.SetVariable('Description',                $Description)
    $rs.SessionStateProxy.SetVariable('SelectedOU',                 $SelectedOU)
    $rs.SessionStateProxy.SetVariable('ManagerDN',                  $ManagerDN)
    $rs.SessionStateProxy.SetVariable('AutoLicense',                $AutoLicense)
    $rs.SessionStateProxy.SetVariable('UPN',                        $UPN)
    $rs.SessionStateProxy.SetVariable('LogPath',                    $LogPath)
    $rs.SessionStateProxy.SetVariable('PreferredSkuOrder',          $PreferredSkuOrder)
    $rs.SessionStateProxy.SetVariable('InitialSyncWaitSeconds',     $InitialSyncWaitSeconds)
    $rs.SessionStateProxy.SetVariable('UserDiscoveryRetryCount',    $UserDiscoveryRetryCount)
    $rs.SessionStateProxy.SetVariable('UserDiscoveryRetryDelaySeconds', $UserDiscoveryRetryDelaySeconds)
    $rs.SessionStateProxy.SetVariable('Window',                     $Window)
    $rs.SessionStateProxy.SetVariable('txtLog',                     $txtLog)
    $rs.SessionStateProxy.SetVariable('logScroller',                $logScroller)
    $rs.SessionStateProxy.SetVariable('btnProvision',               $btnProvision)
    $rs.SessionStateProxy.SetVariable('lblStatus',                  $lblStatus)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs

    $ps.AddScript({
        $ErrorActionPreference = 'Stop'

        function AppendLog {
            param([string]$Level = 'INFO', [string]$Message)
            $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $line = "{0} [{1}] {2}" -f $ts, $Level, $Message
            Add-Content -Path $LogPath -Value $line -Encoding UTF8
            $color = switch ($Level) {
                'INFO'    { '#CDD6F4' }
                'WARN'    { '#F9E2AF' }
                'ERROR'   { '#F38BA8' }
                'SUCCESS' { '#A6E3A1' }
                default   { '#CDD6F4' }
            }
            $Window.Dispatcher.Invoke([action]{
                $run            = New-Object System.Windows.Documents.Run
                $run.Text       = $line + "`n"
                $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($color)
                $txtLog.Inlines.Add($run)
                $logScroller.ScrollToBottom()
            })
        }

        function SetStatus {
            param([string]$Text, [string]$Color = '#89B4FA')
            $Window.Dispatcher.Invoke([action]{
                $lblStatus.Text       = $Text
                $lblStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
            })
        }

        $success = $false

        try {
            Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
            AppendLog INFO "Starting provisioning for $UPN"
            AppendLog INFO "Selected OU: $SelectedOU"

            $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

            AppendLog INFO "Creating AD user..."
            $splat = @{
                Name              = "$FirstName $LastName"
                Path              = $SelectedOU
                GivenName         = $FirstName
                Surname           = $LastName
                SamAccountName    = $Username
                UserPrincipalName = $UPN
                AccountPassword   = $SecurePassword
                Enabled           = $true
            }
            if ($Description) { $splat.Description  = $Description }
            if ($Phone)        { $splat.OfficePhone  = $Phone }
            if ($Email)        { $splat.EmailAddress = $Email }
            if ($Office)       { $splat.Office       = $Office }
            if ($Title)        { $splat.Title        = $Title }
            if ($Department)   { $splat.Department   = $Department }
            if ($Company)      { $splat.Company      = $Company }
            if ($ManagerDN)    { $splat.Manager      = $ManagerDN }

            New-ADUser @splat -PassThru | Enable-ADAccount | Out-Null
            AppendLog SUCCESS "User $UPN created in AD and enabled."

            if ($ManagerDN) {
                AppendLog INFO "Manager assigned: $ManagerDN"
            } else {
                AppendLog WARN "No manager assigned."
            }

            AppendLog INFO ("Waiting {0}s (~{1} min) for AD Connect sync..." -f $InitialSyncWaitSeconds, [math]::Round($InitialSyncWaitSeconds/60,1))
            SetStatus "Waiting for AD Connect sync (~5 min)..."

            $syncEnd = (Get-Date).AddSeconds($InitialSyncWaitSeconds)
            while ((Get-Date) -lt $syncEnd) {
                $remaining = [math]::Round(($syncEnd - (Get-Date)).TotalSeconds)
                $r = $remaining
                $Window.Dispatcher.Invoke([action]{ $lblStatus.Text = "Waiting for AD Connect sync... ($r s remaining)" })
                Start-Sleep -Seconds 5
            }

            AppendLog INFO "Connecting to Microsoft Graph..."
            SetStatus "Connecting to Microsoft Graph..."
            Connect-MgGraph -Scopes "User.ReadWrite.All,Directory.ReadWrite.All" -NoWelcome
            AppendLog INFO "Connected to Microsoft Graph."

            AppendLog INFO "Polling for user in Entra ID..."
            SetStatus "Polling Entra ID for synced user..."
            $found = $false
            for ($i = 1; $i -le $UserDiscoveryRetryCount; $i++) {
                try {
                    $aadUser = Get-MgUser -UserId $UPN -ErrorAction Stop
                    if ($aadUser) {
                        AppendLog SUCCESS "User $UPN found in Entra ID on attempt $i of $UserDiscoveryRetryCount."
                        $found = $true
                        break
                    }
                } catch {
                    $attempt = $i
                    $maxAtt  = $UserDiscoveryRetryCount
                    AppendLog WARN "Attempt $attempt of $maxAtt - user not yet in Entra ID. Retrying in $($UserDiscoveryRetryDelaySeconds)s..."
                    $Window.Dispatcher.Invoke([action]{ $lblStatus.Text = "Polling Entra ID... attempt $attempt of $maxAtt" })
                }
                Start-Sleep -Seconds $UserDiscoveryRetryDelaySeconds
            }

            if (-not $found) {
                throw "User $UPN not found in Entra ID after polling timeout."
            }

            Set-MgUser -UserId $UPN -UsageLocation "US"
            AppendLog SUCCESS "UsageLocation set to US."

            $skus = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber,
                @{n='Available';e={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}}
            AppendLog INFO "Retrieved $($skus.Count) subscribed SKUs."

            $selectedSku = $null
            if ($AutoLicense) {
                foreach ($p in $PreferredSkuOrder) {
                    $c = $skus | Where-Object { $_.SkuPartNumber -eq $p -and $_.Available -gt 0 } | Select-Object -First 1
                    if ($c) { $selectedSku = $c; break }
                }
                if (-not $selectedSku) { throw "No preferred SKU (M365 E3 / O365 E3) available." }
                AppendLog SUCCESS "Auto-selected license: $($selectedSku.SkuPartNumber) (Available: $($selectedSku.Available))"
            } else {
                $availableSkus = $skus | Where-Object { $_.Available -gt 0 }
                if (-not $availableSkus) { throw "No licenses available." }
                $skuList  = ($availableSkus | ForEach-Object { "$($_.SkuPartNumber) (Available: $($_.Available))" }) -join "`n"
                $firstSku = $availableSkus[0].SkuPartNumber
                $chosen   = $null
                $Window.Dispatcher.Invoke([action]{
                    $chosen = [System.Windows.MessageBox]::Show(
                        "Available licenses:`n`n$skuList`n`nClick OK to assign $firstSku, or Cancel to skip.",
                        "License Selection", "OKCancel", "Question")
                })
                if ($chosen -eq 'OK') {
                    $selectedSku = $availableSkus[0]
                } else {
                    throw "License assignment cancelled by operator."
                }
            }

            $addLicenses = @(@{ SkuId = $selectedSku.SkuId })
            Set-MgUserLicense -UserId $UPN -AddLicenses $addLicenses -RemoveLicenses @()
            AppendLog SUCCESS "License $($selectedSku.SkuPartNumber) assigned to $UPN."

            $success = $true
            AppendLog SUCCESS "============================================"
            AppendLog SUCCESS "Provisioning COMPLETE for $UPN"
            AppendLog SUCCESS "============================================"

        } catch {
            AppendLog ERROR "Provisioning FAILED: $($_.Exception.Message)"
        } finally {
            if ($success) {
                AppendLog SUCCESS "Log saved to: $LogPath"
                SetStatus "Provisioning complete!" '#A6E3A1'
            } else {
                SetStatus "Provisioning failed - see log above." '#F38BA8'
            }
            $Window.Dispatcher.Invoke([action]{ $btnProvision.IsEnabled = $true })
        }
    }) | Out-Null

    $ps.BeginInvoke() | Out-Null
})

$Window.ShowDialog() | Out-Null
