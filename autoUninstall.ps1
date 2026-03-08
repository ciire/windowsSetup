# ---------------------------------------------------------
# DEPENDENCIES (Required for the Checkbox GUI)
# ---------------------------------------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ---------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------

function Test-Admin {
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-NOT $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "Elevating privileges..." -ForegroundColor Yellow
        Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Exit
    }
}

# --- RESTORED FUNCTION ---
function Get-UnifiedAppList {
    Write-Host "Aggregating user-removable applications..." -ForegroundColor Cyan

    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    [array]$regApps = Get-ItemProperty $registryPaths -ErrorAction SilentlyContinue | 
        Where-Object { $_.DisplayName -and ($_.UninstallString -or $_.QuietUninstallString) -and ($_.SystemComponent -ne 1) } |
        Select-Object @{Name="DisplayName"; Expression={$_.DisplayName}}, 
                      @{Name="Id"; Expression={$_.PSChildName}}, 
                      @{Name="Type"; Expression={"Win32/64"}},
                      UninstallString

    [array]$appxApps = Get-AppxPackage | 
        Where-Object { 
            $_.IsFramework -eq $false -and 
            $_.IsResourcePackage -eq $false -and 
            $_.IsBundle -eq $false -and 
            $_.NonRemovable -eq $false -and
            $_.SignatureKind -ne "System" -and 
            $_.Name -notmatch "Extension" -and
            $_.Status -eq "Ok"
        } |
        Select-Object @{Name="DisplayName"; Expression={
                          $n = $_.Name -replace 'Microsoft\.', '' -replace 'Windows\.', ''
                          $n = [regex]::Replace($n, '([a-z])([A-Z])', '$1 $2')
                          $n.Replace('.', ' ')
                      }}, 
                      @{Name="Id"; Expression={$_.PackageFullName}}, 
                      @{Name="Type"; Expression={"UWP"}},
                      @{Name="UninstallString"; Expression={"Remove-AppxPackage"}}

    $combined = @()
    if ($regApps) { $combined += $regApps }
    if ($appxApps) { $combined += $appxApps }

    return $combined | Sort-Object DisplayName
}

# --- CORRECTED Show-Checklist (With UseTemplate Logic) ---
function Show-Checklist {
    param(
        $AppList,
        $PreSelectedNames = @()
    )

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="App Uninstaller" Height="700" Width="400" Background="#121212" WindowStartupLocation="CenterScreen">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="*"/>    <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> <RowDefinition Height="Auto"/> </Grid.RowDefinitions>
        
        <TextBlock Text="Select Bloatware" Foreground="#00fbff" FontSize="18" Margin="0,0,0,10" FontWeight="Bold"/>
        
        <CheckBox x:Name="UseTemplate" Grid.Row="1" Content="Use saved template (apps_to_remove.txt)" 
                  Foreground="#00fbff" Margin="0,0,0,10" IsChecked="False" VerticalAlignment="Center"/>

        <ListBox x:Name="AppListBox" Grid.Row="2" Background="#1e1e1e" Foreground="White" BorderThickness="0">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <StackPanel Orientation="Horizontal" Margin="2">
                        <CheckBox IsChecked="{Binding IsChecked}" Margin="0,0,10,0" VerticalAlignment="Center"/>
                        <TextBlock Text="{Binding DisplayName}" VerticalAlignment="Center" FontSize="13"/>
                    </StackPanel>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>

        <CheckBox x:Name="SaveToggle" Grid.Row="3" Content="Update apps_to_remove.txt with current selection" 
                  Foreground="Gray" Margin="0,15,0,0" IsChecked="True" VerticalAlignment="Center"/>

        <CheckBox x:Name="WidgetToggle" Grid.Row="4" Content="Completely Remove Windows Widgets" 
                  Foreground="#ff00ff" Margin="0,10,0,0" IsChecked="False" VerticalAlignment="Center"/>

        <CheckBox x:Name="AutoRunToggle" Grid.Row="5" Content="Run this script automatically after restart" 
                  Foreground="Yellow" Margin="0,10,0,0" IsChecked="False" VerticalAlignment="Center"/>
        
        <Button x:Name="BtnStart" Grid.Row="6" Content="UNINSTALL SELECTED" Height="35" Margin="0,15,0,0" 
                Background="#ff3333" Foreground="White" FontWeight="Bold"/>
    </Grid>
</Window>
"@
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $wrappedList = foreach($app in $AppList) { 
        [PSCustomObject]@{ DisplayName = $app.DisplayName; IsChecked = $false; Original = $app } 
    }
    
    $listBox = $window.FindName("AppListBox")
    $listBox.ItemsSource = $wrappedList
    
    $templateToggle = $window.FindName("UseTemplate")
    
    if ($PreSelectedNames.Count -eq 0) {
        $templateToggle.IsEnabled = $false
        $templateToggle.Content = "No template found (apps_to_remove.txt)"
        $templateToggle.Foreground = [System.Windows.Media.Brushes]::Gray
    }

    $templateToggle.Add_Click({
        foreach ($item in $wrappedList) {
            if ($templateToggle.IsChecked) {
                if ($PreSelectedNames -contains $item.DisplayName) { $item.IsChecked = $true }
            } else {
                if ($PreSelectedNames -contains $item.DisplayName) { $item.IsChecked = $false }
            }
        }
        $listBox.Items.Refresh()
    })

    ($window.FindName("BtnStart")).Add_Click({ $window.DialogResult = $true; $window.Close() })
    
    if ($window.ShowDialog()) { 
        return [PSCustomObject]@{
            AppsToUninst  = $wrappedList | Where-Object { $_.IsChecked } | Select-Object -ExpandProperty Original
            SaveRequested = ($window.FindName("SaveToggle")).IsChecked
            RemoveWidgets = ($window.FindName("WidgetToggle")).IsChecked
            AutoRun       = ($window.FindName("AutoRunToggle")).IsChecked # NEW RETURN VALUE
        }
    }
}
function Disable-WindowsWidgets {
    Write-Host "[ACTION] Disabling Widgets for ALL USERS (Machine Policy)..." -ForegroundColor Cyan
    
    # Path for the System-Wide Policy
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
    $name = "AllowNewsAndInterests"

    try {
        # 1. Ensure the Policy directory exists
        if (-not (Test-Path $registryPath)) {
            New-Object -TypeName PSObject | Out-Null # Placeholder for clarity
            New-Item -Path $registryPath -Force | Out-Null
            Write-Host " [+] Created Policy directory." -ForegroundColor Gray
        }

        # 2. Set the 'Allow' value to 0 (Disabled/No)
        # This overrides individual user settings in the Settings app
        Set-ItemProperty -Path $registryPath -Name $name -Value 0 -ErrorAction Stop
        Write-Host " [+] Machine policy set: Widgets disabled for all users." -ForegroundColor Green

        # 3. Force UI Refresh
        # Policy changes often require an Explorer restart to show immediately
        Write-Host " [+] Refreshing Taskbar..." -ForegroundColor Gray
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    } 
    catch {
        Write-Host " [!] Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host " [!] Ensure you are running as Administrator." -ForegroundColor Yellow
    }
}

function Invoke-SimulatedManualRemoval {
    param ([Parameter(Mandatory=$true)] $App)
    $name = $App.DisplayName
    Write-Host "`n[ACTION] Triggering uninstaller for: $name" -ForegroundColor Cyan
    
    if ($name -match "Acer" -or $name -match "Quick Access") {
        Write-Host " [!] Stopping background services..." -ForegroundColor Yellow
        $acerServices = "QASvc", "AcerAgentService", "AOPClickSvc"
        foreach ($svc in $acerServices) {
            if (Get-Service $svc -ErrorAction SilentlyContinue) {
                Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            }
        }
        Get-Process -Name "QAAdmin", "QuickAccess*", "AcerConfig*" -ErrorAction SilentlyContinue | Stop-Process -Force
    }

    try {
        if ($App.Type -eq "UWP") {
            try {
                # 1. Get the Package Name (e.g., Microsoft.MicrosoftSolitaireCollection)
                $pName = ($App.Id -split "_")[0]

                # 2. Kill the process if running (prevents 'in-use' blocks)
                Get-Process | Where-Object { $_.Name -match $pName } | Stop-Process -Force -ErrorAction SilentlyContinue

                # 3. PURGE FOR ALL USERS
                # This removes the app from every existing account on the PC
                Write-Host " [>] Removing from all user profiles..." -ForegroundColor Gray
                Get-AppxPackage -Name "*$pName*" -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue

                # 4. DEPROVISION (The "Master Copy")
                # This prevents the app from ever being installed on new accounts created later
                Write-Host " [>] De-provisioning master package..." -ForegroundColor Gray
                Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -match $pName } | 
                    Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
                    
                Write-Host " [SUCCESS] $name removed for ALL current and FUTURE users." -ForegroundColor Green
            } catch {
                Write-Host " [!] Failed to fully purge $name. It may be a protected System app." -ForegroundColor Red
            }
        } else {
            if ($App.UninstallString -match '^"(.*)"\s*(.*)$') {
                $path = $matches[1]; $args = $matches[2]
            } else {
                $path = ($App.UninstallString -split ".exe")[0] + ".exe"
                $args = ($App.UninstallString -split ".exe")[1]
            }

            $silentArgs = "$args /S /silent /quiet /norestart".Trim()
            $proc = Start-Process -FilePath $path -ArgumentList $silentArgs -Verb RunAs -PassThru -Wait -ErrorAction SilentlyContinue
            
            Start-Sleep -Seconds 5 
            $currentSystemApps = Get-UnifiedAppList
            $stillExists = $currentSystemApps | Where-Object { $_.Id -eq $App.Id }

            if ($stillExists) {
                Write-Host " [!] Silent failed. Launching standard GUI..." -ForegroundColor White
                Start-Process -FilePath $path -ArgumentList $args -Verb RunAs -Wait
            } else {
                Write-Host " [SUCCESS] $name removed silently." -ForegroundColor Green
            }
        }
    } catch { 
        Write-Host " [!] Critical Error: Could not launch $name uninstaller." -ForegroundColor Red 
    }
}

function Set-RunOnce {
    # Using $PSCommandPath is often more reliable for the full file path
    Add-Content -Path "C:\temp\run_log.txt" -Value "Script ran at $(Get-Date)"   
    $scriptPath = $PSCommandPath 
    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    $name = "BloatwareCleanupTest"
    
    $command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -File `"$scriptPath`""
    try {
        Set-ItemProperty -Path $registryPath -Name $name -Value $command -ErrorAction Stop
        Write-Host "[AUTO-RUN] Script scheduled to run once after next login." -ForegroundColor Green
    } catch {
        Write-Host "[!] Failed to set Auto-Run. Ensure you have Admin rights." -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------
# MAIN EXECUTION
# ---------------------------------------------------------

Test-Admin

$configPath = Join-Path $PSScriptRoot "apps_to_remove.txt"
$preSelected = @()

if (Test-Path $configPath) {
    $preSelected = Get-Content $configPath
}

$allApps = Get-UnifiedAppList
$guiResult = Show-Checklist -AppList $allApps -PreSelectedNames $preSelected

if ($guiResult) {
    # 1. Handle Auto-Run Schedule
    if ($guiResult.AutoRun) { 
        Set-RunOnce 
    }

    # 2. Handle Widgets
    if ($guiResult.RemoveWidgets) { 
        Disable-WindowsWidgets 
    }
    
    # 3. Handle Template Saving
    if ($guiResult.SaveRequested -and $guiResult.AppsToUninst) {
        $guiResult.AppsToUninst.DisplayName | Out-File $configPath -Force
        Write-Host "[CONFIG] Template updated." -ForegroundColor Gray
    }

    # 4. Run Uninstalls
    if ($guiResult.AppsToUninst) {
        foreach ($item in $guiResult.AppsToUninst) {
            Invoke-SimulatedManualRemoval -App $item
        }
    }
    
    Write-Host "`nCleanup completed." -ForegroundColor White
}

Pause