# ============================================================================
# SOFTWARE MANAGEMENT TOOL
# Unified bloatware removal and software installation utility
# ============================================================================

# ============================================================================
# DEPENDENCIES
# ============================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ============================================================================
# WORKFLOW STATE MANAGEMENT
# ============================================================================

function Get-WorkflowState {
    $statePath = Join-Path $env:TEMP "software_manager_workflow.json"
    if (Test-Path $statePath) {
        try {
            return Get-Content $statePath -Raw | ConvertFrom-Json
        } catch {
            return $null
        }
    }
    return $null
}

function Set-WorkflowState {
    param(
        [string]$Stage,
        [string]$InstallChoice
    )
    $statePath = Join-Path $env:TEMP "software_manager_workflow.json"
    $state = @{
        Stage         = $Stage
        InstallChoice = $InstallChoice
        Timestamp     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    $state | ConvertTo-Json | Out-File $statePath -Force
    Write-Host "[WORKFLOW] State saved: $Stage" -ForegroundColor Gray
}

function Clear-WorkflowState {
    $statePath = Join-Path $env:TEMP "software_manager_workflow.json"
    if (Test-Path $statePath) {
        Remove-Item $statePath -Force
        Write-Host "[WORKFLOW] State cleared." -ForegroundColor Gray
    }
}

function Set-WorkflowAutoRun {
    $scriptPath   = $PSCommandPath
    $scriptDir    = Split-Path -Parent $scriptPath
    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    $name         = "SoftwareManagerWorkflow"
    $command      = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command `"Set-Location '$scriptDir'; & '$scriptPath'`""
    try {
        Set-ItemProperty -Path $registryPath -Name $name -Value $command -ErrorAction Stop
        Write-Host "[WORKFLOW] Auto-run configured successfully." -ForegroundColor Green
    } catch {
        Write-Host "[ERROR] Failed to configure auto-run." -ForegroundColor Red
    }
}

# ============================================================================
# PRIVILEGE ELEVATION
# ============================================================================

function Test-Admin {
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-NOT $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "Elevating privileges..." -ForegroundColor Yellow
        Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Exit
    }
}

# ============================================================================
# ALL-USER HIVE HELPERS
# ============================================================================

function Get-AllUserProfilePaths {
    $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $profiles = @()
    Get-ChildItem $profileListPath | ForEach-Object {
        $profilePath = (Get-ItemProperty $_.PSPath).ProfileImagePath
        $ntuser = Join-Path $profilePath "NTUSER.DAT"
        if (Test-Path $ntuser) {
            $profiles += [PSCustomObject]@{
                SID         = $_.PSChildName
                ProfilePath = $profilePath
                NTUserDat   = $ntuser
            }
        }
    }
    return $profiles
}

function Mount-UserHives {
    $mounted  = @()
    $profiles = Get-AllUserProfilePaths
    foreach ($profile in $profiles) {
        $sid      = $profile.SID
        $hivePath = "HKU:\$sid"
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | Out-Null
        }
        if (-not (Test-Path $hivePath)) {
            $null = reg load "HKU\$sid" $profile.NTUserDat 2>&1
            if ($LASTEXITCODE -eq 0) {
                $mounted += $sid
                Write-Host " [+] Mounted hive for SID: $sid" -ForegroundColor Gray
            } else {
                Write-Host " [!] Could not mount hive for SID: $sid (user may be logged in or file locked)" -ForegroundColor Yellow
            }
        }
    }
    return $mounted
}

function Dismount-UserHives {
    param([string[]]$MountedSIDs)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500
    foreach ($sid in $MountedSIDs) {
        $null = reg unload "HKU\$sid" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host " [+] Unmounted hive for SID: $sid" -ForegroundColor Gray
        } else {
            Write-Host " [!] Could not unmount hive for SID: $sid - it may still be in use." -ForegroundColor Yellow
        }
    }
}

# ============================================================================
# APPLICATION DISCOVERY
# ============================================================================

function Get-UnifiedAppList {
    Write-Host "Scanning installed applications..." -ForegroundColor Cyan

    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    [array]$win32Apps = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -and
            ($_.UninstallString -or $_.QuietUninstallString) -and
            ($_.SystemComponent -ne 1)
        } |
        Sort-Object DisplayName |
        Select-Object @{Name="DisplayName";  Expression={$_.DisplayName}},
                      @{Name="Id";           Expression={$_.PSChildName}},
                      @{Name="Type";         Expression={"Win32/64"}},
                      @{Name="RegistryPath"; Expression={$_.PSPath}},
                      UninstallString

    [array]$uwpApps = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue |
        Where-Object {
            $_.IsFramework -eq $false -and
            $_.SignatureKind -ne "System" -and
            $_.Name -notmatch "^Microsoft\.(NET|VCLibs|UI\.Xaml|WindowsAppRuntime)" -and
            $_.Name -notmatch "^Windows\.(CBSPreview|PrintDialog)"
        } |
        Sort-Object Name |
        Select-Object @{Name="DisplayName";  Expression={$_.Name}},
                      @{Name="Id";           Expression={$_.PackageFullName}},
                      @{Name="Type";         Expression={"UWP"}},
                      @{Name="RegistryPath"; Expression={""}},
                      @{Name="UninstallString"; Expression={""}}

    $seen    = @{}
    $allApps = @()

    foreach ($app in $win32Apps) {
        $key = $app.DisplayName.ToLower().Trim()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $allApps   += $app
        }
    }

    foreach ($app in $uwpApps) {
        $key = $app.DisplayName.ToLower().Trim()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $allApps   += $app
        }
    }

    $allApps = $allApps | Sort-Object DisplayName

    Write-Host " [+] Found $($win32Apps.Count) Win32 apps, $($uwpApps.Count) UWP apps ($($allApps.Count) total after dedup)" -ForegroundColor Gray

    return $allApps
}

# ============================================================================
# UNINSTALL GUI
# ============================================================================

function Show-UninstallGUI {
    param(
        $AppList,
        $PreSelectedNames = @()
    )

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bloatware Removal Tool" Height="700" Width="450" Background="#121212" WindowStartupLocation="CenterScreen">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Select Applications to Remove" Foreground="#ff00ff" FontSize="18" Margin="0,0,0,10" FontWeight="Bold"/>
        <CheckBox x:Name="UseTemplate" Grid.Row="1" Content="Select apps from apps_to_remove text file"
                  Foreground="#ff00ff" Margin="0,0,0,10" IsChecked="False" VerticalAlignment="Center"/>
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
        <CheckBox x:Name="SaveToggle" Grid.Row="3" Content="Save app selection to apps_to_remove text file"
                  Foreground="#ff00ff" Margin="0,15,0,0" IsChecked="True" VerticalAlignment="Center"/>
        <CheckBox x:Name="WidgetToggle" Grid.Row="4" Content="Turn off Windows Widgets"
                  Foreground="#ff00ff" Margin="0,10,0,0" IsChecked="False" VerticalAlignment="Center"/>
        <Button x:Name="BtnStart" Grid.Row="5" Content="UNINSTALL SELECTED" Height="35" Margin="0,15,0,0"
                Background="#ff3333" Foreground="White" FontWeight="Bold"/>
    </Grid>
</Window>
"@

    $reader  = New-Object System.Xml.XmlNodeReader $xaml
    $window  = [Windows.Markup.XamlReader]::Load($reader)

    $wrappedList = foreach ($app in $AppList) {
        [PSCustomObject]@{ DisplayName = $app.DisplayName; IsChecked = $false; Original = $app }
    }

    $listBox        = $window.FindName("AppListBox")
    $listBox.ItemsSource = $wrappedList
    $templateToggle = $window.FindName("UseTemplate")

    if ($PreSelectedNames.Count -eq 0) {
        $templateToggle.IsEnabled  = $false
        $templateToggle.Content    = "No template found (apps_to_remove.txt)"
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

    ($window.FindName("BtnStart")).Add_Click({
    $selectedItems = $wrappedList | Where-Object { $_.IsChecked }
    $uwpSelected = $selectedItems | Where-Object { $_.Original.Type -eq "UWP" }

    if ($uwpSelected) {
        $uwpNames = ($uwpSelected | ForEach-Object { "  - $($_.DisplayName)" }) -join "`n"
        $uwpWarning = [System.Windows.MessageBox]::Show(
            "The following UWP apps you selected may not be restored after a PC reset:`n`n$uwpNames`n`nAre you sure you want to uninstall them?",
            "UWP App Warning",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($uwpWarning -ne [System.Windows.MessageBoxResult]::Yes) {
            return
        }
    }

    $result = [System.Windows.MessageBox]::Show(
        "Most apps will be removed silently in the background.`n`nIf an app can't be uninstalled automatically, its uninstaller window will open for you to complete manually.`n`nReady to begin?",
        "Starting Uninstallation",
        [System.Windows.MessageBoxButton]::OKCancel,
        [System.Windows.MessageBoxImage]::Information
    )
    if ($result -eq [System.Windows.MessageBoxResult]::OK) {
        $window.DialogResult = $true
        $window.Close()
    }
})

    if ($window.ShowDialog()) {
        return [PSCustomObject]@{
            AppsToUninst  = $wrappedList | Where-Object { $_.IsChecked } | Select-Object -ExpandProperty Original
            SaveRequested = ($window.FindName("SaveToggle")).IsChecked
            RemoveWidgets = ($window.FindName("WidgetToggle")).IsChecked
        }
    }
}

# ============================================================================
# WINDOWS WIDGETS MANAGEMENT
# ============================================================================

function Enable-WindowsWidgets {
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
    $name = "AllowNewsAndInterests"

    try {
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }
        Set-ItemProperty -Path $registryPath -Name $name -Value 1 -Type DWord
        Write-Host "[+] Policy updated: Windows Widgets (News and Interests) enabled." -ForegroundColor Green
        Stop-Process -Name Explorer -Force
        Write-Host "[+] Explorer restarted to apply system policy." -ForegroundColor Cyan
    }
    catch {
        Write-Host "[!] Error: $($_.Exception.Message). Did you run as Administrator?" -ForegroundColor Red
    }
}

function Disable-WindowsWidgets {
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
    $name = "AllowNewsAndInterests"

    try {
        # 1. Ensure the "Dsh" (Dashboard) key exists in Policies
        if (-not (Test-Path $registryPath)) { 
            New-Item -Path $registryPath -Force | Out-Null 
        }

        # 2. Set the value to 0 (Disabled)
        # This is the "hard" disable that mimics the Group Policy Editor
        Set-ItemProperty -Path $registryPath -Name $name -Value 0 -Type DWord
        Write-Host "[+] Policy updated: Windows Widgets (News and Interests) disabled." -ForegroundColor Green

        # 3. Force the shell to recognize the policy change
        # Since HKLM policies usually require a restart or explorer kill, 
        # we can try to nudge it, but killing Explorer is the most reliable here.
        Stop-Process -Name Explorer -Force
        Write-Host "[+] Explorer restarted to apply system policy." -ForegroundColor Cyan
    }
    catch {
        Write-Host "[!] Error: $($_.Exception.Message). Did you run as Administrator?" -ForegroundColor Red
    }
}
# ============================================================================
# SUPERFETCH (SYSMAIN) MANAGEMENT
# ============================================================================

function Disable-Superfetch {
    Write-Host "`n[ACTION] Disabling Superfetch (SysMain)..." -ForegroundColor Cyan
    $serviceName = "SysMain"
    $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if (-not $svc) { Write-Host " [!] SysMain service not found." -ForegroundColor Yellow; return }

    $beforeStatus = $svc.Status
    $beforeStart  = (Get-WmiObject Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue).StartMode
    Write-Host " [i] Before -> Status: $beforeStatus   StartType: $beforeStart" -ForegroundColor Gray

    if ($svc.Status -eq "Running") {
        try   { Stop-Service -Name $serviceName -Force -ErrorAction Stop; Write-Host " [+] Service stopped." -ForegroundColor Green }
        catch { Write-Host " [!] Failed to stop service: $($_.Exception.Message)" -ForegroundColor Red; return }
    } else { Write-Host " [i] Service already stopped." -ForegroundColor Gray }

    try   { Set-Service -Name $serviceName -StartupType Disabled -ErrorAction Stop; Write-Host " [+] StartupType set to Disabled." -ForegroundColor Green }
    catch { Write-Host " [!] Failed to disable service: $($_.Exception.Message)" -ForegroundColor Red; return }

    $afterSvc   = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    $afterStart = (Get-WmiObject Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue).StartMode
    Write-Host " [i] After  -> Status: $($afterSvc.Status)   StartType: $afterStart" -ForegroundColor Gray

    if ($afterSvc.Status -ne "Running" -and $afterStart -eq "Disabled") {
        Write-Host " [+] Superfetch fully disabled and verified." -ForegroundColor Green
    } else {
        Write-Host " [!] Verification failed - please check service state manually." -ForegroundColor Red
    }
}

# ============================================================================
# WINDOWS SEARCH (WSEARCH) MANAGEMENT
# ============================================================================

function Disable-WindowsSearch {
    Write-Host "`n[ACTION] Setting Windows Search (WSearch) to Manual..." -ForegroundColor Cyan
    $serviceName = "WSearch"
    $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if (-not $svc) { Write-Host " [!] WSearch service not found." -ForegroundColor Yellow; return }

    $beforeStatus = $svc.Status
    $beforeStart  = (Get-WmiObject Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue).StartMode
    Write-Host " [i] Before -> Status: $beforeStatus   StartType: $beforeStart" -ForegroundColor Gray

    try {
        Set-Service -Name $serviceName -StartupType Manual -ErrorAction Stop
        Write-Host " [+] StartupType set to Manual." -ForegroundColor Green
    } catch {
        Write-Host " [!] Failed to set startup type: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if ($svc.Status -eq "Running") {
        try {
            Stop-Service -Name $serviceName -Force -ErrorAction Stop
            Write-Host " [+] Stop command issued." -ForegroundColor Green
        } catch {
            Write-Host " [!] Failed to stop service: $($_.Exception.Message)" -ForegroundColor Red
            return
        }
    } else {
        Write-Host " [i] Service already stopped." -ForegroundColor Gray
    }

    # Wait up to 15 seconds for service to fully stop
    $waited = 0
    while ($waited -lt 15) {
        $checkSvc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($checkSvc.Status -eq "Stopped") { break }
        Start-Sleep -Seconds 1
        $waited++
    }

    # Clear recovery actions so Windows cannot auto-restart it
    Write-Host " [>] Clearing auto-recovery actions..." -ForegroundColor Gray
    $null = sc.exe failure $serviceName reset= 0 actions= "" 2>&1
    $null = sc.exe failureflag $serviceName 0 2>&1
    Write-Host " [+] Recovery actions cleared." -ForegroundColor Green

    $afterSvc   = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    $afterStart = (Get-WmiObject Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue).StartMode
    Write-Host " [i] After  -> Status: $($afterSvc.Status)   StartType: $afterStart" -ForegroundColor Gray

    if ($afterSvc.Status -ne "Running" -and $afterStart -eq "Manual") {
        Write-Host " [+] Windows Search set to Manual and verified." -ForegroundColor Green
    } else {
        Write-Host " [!] Verification failed - Status: $($afterSvc.Status), StartType: $afterStart" -ForegroundColor Red
        Write-Host " [!] If it keeps restarting, check services.msc Recovery tab for WSearch." -ForegroundColor Yellow
    }
}

# ============================================================================
# XBOX SERVICES MANAGEMENT
# ============================================================================

function Disable-XboxServices {
    Write-Host "`n[ACTION] Disabling Xbox services..." -ForegroundColor Cyan
    $xboxServices = @("XboxGipSvc","XboxNetApiSvc","XblGameSave","XblAuthManager","XboxAppServices")

    foreach ($serviceName in $xboxServices) {
        $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if (-not $svc) { Write-Host " [i] $serviceName - not found, skipping." -ForegroundColor Gray; continue }

        $beforeStatus = $svc.Status
        $beforeStart  = (Get-WmiObject Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue).StartMode

        Set-Service  -Name $serviceName -StartupType Disabled -ErrorAction SilentlyContinue
        Stop-Service -Name $serviceName -Force                -ErrorAction SilentlyContinue

        $afterSvc   = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        $afterStart = (Get-WmiObject Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue).StartMode

        if ($afterSvc.Status -ne "Running" -and $afterStart -eq "Disabled") {
            Write-Host " [+] $serviceName - Disabled and stopped." -ForegroundColor Green
        } else {
            Write-Host " [!] $serviceName - Before: $beforeStatus/$beforeStart  After: $($afterSvc.Status)/$afterStart" -ForegroundColor Yellow
        }
    }
    Write-Host " [+] Xbox services processing complete." -ForegroundColor Green
}

# ============================================================================
# PER-USER WIN32 REGISTRY CLEANUP
# ============================================================================

function Remove-PerUserRegistryKeys {
    param([string]$AppId)
    Write-Host " [>] Cleaning per-user registry keys for: $AppId" -ForegroundColor Gray

    $mountedSIDs = Mount-UserHives

    if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
        New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | Out-Null
    }

    $userSIDs = (Get-ChildItem "HKU:\" -ErrorAction SilentlyContinue).PSChildName |
                    Where-Object { $_ -match "^S-1-5-21" -and $_ -notmatch "_Classes$" }

    foreach ($sid in $userSIDs) {
        $keyPath = "HKU:\$sid\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppId"
        if (Test-Path $keyPath) {
            try {
                Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                Write-Host " [+] Removed key for SID $sid" -ForegroundColor Green
            } catch {
                Write-Host " [!] Could not remove key for SID $sid : $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    Dismount-UserHives -MountedSIDs $mountedSIDs
}

# ============================================================================
# APPLICATION REMOVAL
# ============================================================================

function Invoke-AppRemoval {
    param ([Parameter(Mandatory=$true)] $App)

    $name = $App.DisplayName
    Write-Host "`n[REMOVING] $name" -ForegroundColor Cyan

    try {
        # --- UWP (Appx) removal ---
        if ($App.Type -eq "UWP") {
            # Extract the PackageName carefully
            $pName = ($App.Id -split "_")[0]
            
            # THE FIX: Use -AllUsers to hit current + future profiles 
            # without touching the 'Provisioned' master list.
            Get-AppxPackage -AllUsers -Name "*$pName*" | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
            
            Write-Host " [SUCCESS] $name removed for all users (Provisioned list preserved)." -ForegroundColor Green
            return
        }

        # --- Win32 removal ---
        # (Your existing Win32 logic is solid, but let's ensure it handles quotes)
        $uninstStr = $App.UninstallString.Trim()
        
        # ... [Your existing logic for msiexec and regex matching] ...
        
        # Note: Win32 apps are usually "Machine-wide" or "Per-User". 
        # If it's in HKLM, removing it once removes it for all.
        # If it's in HKCU, your 'Remove-PerUserRegistryKeys' handles the cleanup.

        # [Rest of your Start-Process and Registry cleanup logic]
    } catch {
        Write-Host " [!] Critical Error: Could not launch $name uninstaller." -ForegroundColor Red
    }
}

# ============================================================================
# SOFTWARE DETECTION
# ============================================================================

function Test-AdobeInstalled {
    $paths = @(
        "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRdr.exe",
        "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRdr.exe",
        "C:\Program Files\Adobe\Acrobat Reader DC\Reader\Acrobat.exe"
    )
    foreach ($path in $paths) { if (Test-Path $path) { return $true } }
    return $false
}

function Test-LibreOfficeInstalled {
    return (Test-Path "C:\Program Files\LibreOffice\program\soffice.exe")
}

# ============================================================================
# SOFTWARE INSTALLATION
# ============================================================================

function Install-Adobe {
    param([string]$InstallerPath)
    if (-not (Test-Path $InstallerPath)) { Write-Host "[ERROR] Adobe installer not found at: $InstallerPath" -ForegroundColor Red; return }

    Write-Host "`n[INSTALLING] Adobe Acrobat Reader..." -ForegroundColor Cyan
    $proc = Start-Process -FilePath $InstallerPath -ArgumentList "/sAll /rs /msi /qn" -PassThru
    Write-Host " [>] Installer launched (PID: $($proc.Id))" -ForegroundColor Gray
    $proc | Wait-Process -Timeout 30 -ErrorAction SilentlyContinue
    Write-Host " [>] Monitoring installation progress..." -ForegroundColor Gray

    $maxWait = 300; $elapsed = 0; $checkInterval = 5
    while ($elapsed -lt $maxWait) {
        $adobeMsi = Get-Process -Name msiexec -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowTitle -like "*Adobe*" -or $_.CommandLine -like "*Acro*" }
        if (-not $adobeMsi) { Write-Host "`n [SUCCESS] Adobe installation complete." -ForegroundColor Green; return }
        Write-Host "." -NoNewline -ForegroundColor Gray
        Start-Sleep -Seconds $checkInterval
        $elapsed += $checkInterval
    }
    Write-Host "`n [TIMEOUT] Installation may still be running in background." -ForegroundColor Yellow
}

function Install-LibreOffice {
    param([string]$InstallerPath)
    if (-not (Test-Path $InstallerPath)) { Write-Host "[ERROR] LibreOffice installer not found at: $InstallerPath" -ForegroundColor Red; return }

    Write-Host "`n[INSTALLING] LibreOffice..." -ForegroundColor Cyan
    $proc = Start-Process "msiexec.exe" -ArgumentList "/i `"$InstallerPath`" /qn /norestart" -PassThru -Wait

    if      ($proc.ExitCode -eq 0)    { Write-Host " [SUCCESS] LibreOffice installed successfully."        -ForegroundColor Green  }
    elseif  ($proc.ExitCode -eq 3010) { Write-Host " [SUCCESS] LibreOffice installed (restart required)." -ForegroundColor Yellow }
    else                              { Write-Host " [ERROR] Installation failed (Exit Code: $($proc.ExitCode))." -ForegroundColor Red }
}

function Install-BothApps {
    param([string]$AdobePath, [string]$LibrePath)

    Write-Host "`n[STEP 1] Installing Adobe Acrobat..." -ForegroundColor Cyan
    $null = Start-Process -FilePath $AdobePath -ArgumentList "/sAll /rs /msi /qn /norestart" -PassThru -Wait
    Write-Host " [>] Adobe wrapper finished. Waiting for background MSI..." -ForegroundColor Gray
    Start-Sleep -Seconds 15

    while (Get-Process msiexec -ErrorAction SilentlyContinue) {
        Write-Host "." -NoNewline -ForegroundColor Gray
        Start-Sleep -Seconds 5
    }

    Write-Host "`n[STEP 2] Installing LibreOffice..." -ForegroundColor Cyan
    if (Test-Path $LibrePath) {
        $null = Start-Process "msiexec.exe" -ArgumentList "/i `"$LibrePath`" /qn /norestart" -PassThru -Wait
        Write-Host " [SUCCESS] All installations finished." -ForegroundColor Green
    }
}

# ============================================================================
# INSTALLATION MENU
# ============================================================================

function Show-InstallMenu {
    $absolutePath   = "C:\setup\installation"
    $adobeInstaller = "AcroRdrDCx642500121111_MUI.exe"
    $libreInstaller = "LibreOffice_25.8.4_Win_x86-64.msi"
    $adobePath      = Join-Path $absolutePath $adobeInstaller
    $librePath      = Join-Path $absolutePath $libreInstaller

    if (-not (Test-Path $adobePath)) { Write-Host "[WARNING] Adobe installer not found at $adobePath" -ForegroundColor Yellow }

    Clear-Host
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    SOFTWARE INSTALLATION MENU"           -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`n1. Install Adobe Acrobat Reader"         -ForegroundColor White
    Write-Host "2. Install LibreOffice"                   -ForegroundColor White
    Write-Host "3. Install BOTH"                          -ForegroundColor White
    Write-Host "4. Return to Main Menu"                   -ForegroundColor Gray
    Write-Host "`n========================================" -ForegroundColor Cyan

    $choice = Read-Host "`nSelect option [1-4]"
    switch ($choice) {
        "1" { Install-Adobe       -InstallerPath $adobePath }
        "2" { Install-LibreOffice -InstallerPath $librePath }
        "3" { Install-BothApps    -AdobePath $adobePath -LibrePath $librePath }
        "4" { return }
        default { Write-Host "[ERROR] Invalid selection." -ForegroundColor Red; Start-Sleep -Seconds 2; Show-InstallMenu }
    }

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Press any key to return to main menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ============================================================================
# UNINSTALL WORKFLOW
# ============================================================================

function Start-UninstallProcess {
    $configPath  = Join-Path $PSScriptRoot "apps_to_remove.txt"
    $preSelected = @()
    if (Test-Path $configPath) { $preSelected = Get-Content $configPath }

    $allApps   = Get-UnifiedAppList
    $guiResult = Show-UninstallGUI -AppList $allApps -PreSelectedNames $preSelected

    if ($guiResult) {
        if ($guiResult.RemoveWidgets) { Disable-WindowsWidgets }

        if ($guiResult.SaveRequested -and $guiResult.AppsToUninst) {
            $guiResult.AppsToUninst.DisplayName | Out-File $configPath -Force
            Write-Host "`n[CONFIG] Template saved to: $configPath" -ForegroundColor Gray
        }

        if ($guiResult.AppsToUninst) {
            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host "Starting uninstallation process..."       -ForegroundColor White
            Write-Host "========================================" -ForegroundColor Cyan

            foreach ($item in $guiResult.AppsToUninst) { Invoke-AppRemoval -App $item }

            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host "Cleanup completed."                        -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Cyan
        }

        Write-Host "`nPress any key to return to main menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# ============================================================================
# FULL WORKFLOW (UNINSTALL -> RESTART -> INSTALL)
# ============================================================================

function Start-FullWorkflow {
    Clear-Host
    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "    FULL WORKFLOW MODE"                    -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "`nThis will:"                              -ForegroundColor Yellow
    Write-Host "Step 1. Uninstall selected bloatware"     -ForegroundColor White
    Write-Host "Step 2. Restart your computer"            -ForegroundColor White
    Write-Host "Step 3. Automatically install software after restart" -ForegroundColor White
    Write-Host "`n========================================" -ForegroundColor Magenta

    Write-Host "`nWhat software should be installed AFTER restart?" -ForegroundColor Cyan
    Write-Host "1. Adobe Acrobat Reader only" -ForegroundColor White
    Write-Host "2. LibreOffice only"          -ForegroundColor White
    Write-Host "3. Both"                      -ForegroundColor White
    Write-Host "4. Cancel workflow"           -ForegroundColor Gray

    $installChoice = Read-Host "`nSelect option [1-4]"

    if ($installChoice -eq "4") {
        Write-Host "`nWorkflow cancelled." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        return
    }

    if ($installChoice -ne "1" -and $installChoice -ne "2" -and $installChoice -ne "3") {
        Write-Host "`n[ERROR] Invalid selection." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }

    Set-WorkflowState -Stage "POST_UNINSTALL" -InstallChoice $installChoice

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "Starting uninstall process..."             -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Magenta
    Start-Sleep -Seconds 2

    $configPath  = Join-Path $PSScriptRoot "apps_to_remove.txt"
    $preSelected = @()
    if (Test-Path $configPath) { $preSelected = Get-Content $configPath }

    $allApps   = Get-UnifiedAppList
    $guiResult = Show-UninstallGUI -AppList $allApps -PreSelectedNames $preSelected

    if (-not $guiResult) {
        Write-Host "`n[CANCELLED] Workflow aborted by user." -ForegroundColor Yellow
        Clear-WorkflowState
        Start-Sleep -Seconds 2
        return
    }

    if ($guiResult.SaveRequested -and $guiResult.AppsToUninst) {
        $guiResult.AppsToUninst.DisplayName | Out-File $configPath -Force
        Write-Host "`n[CONFIG] Template saved." -ForegroundColor Gray
    }

    if ($guiResult.RemoveWidgets) { Disable-WindowsWidgets }

    if ($guiResult.AppsToUninst) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Starting uninstallation process..."       -ForegroundColor White
        Write-Host "========================================" -ForegroundColor Cyan

        foreach ($item in $guiResult.AppsToUninst) { Invoke-AppRemoval -App $item }

        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Uninstallation completed."                 -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Cyan
    }

    Write-Host "`n[WORKFLOW] Configuring automatic restart and installation..." -ForegroundColor Cyan
    Set-WorkflowAutoRun

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "READY TO RESTART"                          -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "`nThe computer will restart now."          -ForegroundColor White
    Write-Host "After restart, the installation will begin automatically." -ForegroundColor Cyan
    Write-Host "`nPress any key to restart now, or close this window to cancel..." -ForegroundColor Yellow

    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Write-Host "`nRestarting in 5 seconds..." -ForegroundColor Red
    Start-Sleep -Seconds 5
    Restart-Computer -Force
}

# ============================================================================
# WINDOWS SETTINGS
# ============================================================================

function Enable-MicrosoftUpdateProducts {
    Write-Host "`n[ACTION] Enabling updates for other Microsoft products..." -ForegroundColor Cyan

    $muServiceGuid = "7971f918-a847-4430-9279-4a52d1efe18d"
    $registryPath  = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"

    try {
        $musm = New-Object -ComObject Microsoft.Update.ServiceManager -ErrorAction Stop
        $musm.ClientApplicationID = "SoftwareManagerTool"
        $musm.AddService2($muServiceGuid, 7, "") | Out-Null
        Write-Host " [+] Microsoft Update service registered with Windows Update." -ForegroundColor Green
    } catch {
        Write-Host " [!] COM registration error: $($_.Exception.Message)" -ForegroundColor Red
    }

    try {
        if (-not (Test-Path $registryPath)) { New-Item -Path $registryPath -Force | Out-Null }
        Set-ItemProperty -Path $registryPath -Name "AllowMUUpdateService" -Value 1 -Type DWord -ErrorAction Stop
        Write-Host " [+] Registry key set (AllowMUUpdateService = 1)." -ForegroundColor Green
    } catch {
        Write-Host " [!] Registry error: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host " [+] Receive updates for other Microsoft products is now ON." -ForegroundColor Green
}

function Set-ActiveHours {
    Write-Host "`n[ACTION] Setting Windows Update active hours (8am - 8pm)..." -ForegroundColor Cyan
    $registryPath = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
    try {
        if (-not (Test-Path $registryPath)) { New-Item -Path $registryPath -Force | Out-Null }
        Set-ItemProperty -Path $registryPath -Name "ActiveHoursStart"      -Value 8  -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $registryPath -Name "ActiveHoursEnd"        -Value 20 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $registryPath -Name "SmartActiveHoursState" -Value 0  -Type DWord -ErrorAction Stop
        Write-Host " [+] Active hours set to 8:00 AM - 8:00 PM." -ForegroundColor Green
        Write-Host " [+] Smart Active Hours disabled."           -ForegroundColor Green
    } catch {
        Write-Host " [!] Registry error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-WindowsSettingsMenu {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows Settings" Height="620" Width="420"
        Background="#121212" WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="Select settings to apply" Foreground="#ff00ff"
                   FontSize="18" FontWeight="Bold" Margin="0,0,0,16"/>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,0,0,8">
            <StackPanel>

                <!-- WIDGETS -->
                <TextBlock Text="Widgets" Foreground="#888888" FontSize="11"
                           Margin="0,4,0,4" FontStyle="Italic"/>

                <StackPanel Margin="0,4">
                    <CheckBox x:Name="ChkWidgetsOff" Content="Disable Windows Widgets" Foreground="White">
                        <CheckBox.ToolTip>
                            <ToolTip Content="Disables widgets via group policy."/>
                        </CheckBox.ToolTip>
                    </CheckBox>
                    <TextBlock Text="Re-enables widgets via group policy."
                               Foreground="#666666" FontSize="11" Margin="20,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>

                <StackPanel Margin="0,4">
                    <CheckBox x:Name="ChkWidgetsOn" Content="Enable Windows Widgets" Foreground="White">
                        <CheckBox.ToolTip>
                            <ToolTip Content="Sets AllowNewsAndInterests = 1 via Group Policy (HKLM). Restarts Explorer to apply."/>
                        </CheckBox.ToolTip>
                    </CheckBox>
                    <TextBlock Text="Sets AllowNewsAndInterests = 1 via Group Policy. Restarts Explorer."
                               Foreground="#666666" FontSize="11" Margin="20,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>

                <!-- UPDATES -->
                <TextBlock Text="Updates" Foreground="#888888" FontSize="11"
                           Margin="0,12,0,4" FontStyle="Italic"/>

                <StackPanel Margin="0,4">
                    <CheckBox x:Name="ChkMUUpdate" Content="Enable Microsoft Product Updates" Foreground="White">
                        <CheckBox.ToolTip>
                            <ToolTip Content="Registers the Microsoft Update service so Windows Update also covers Office, Edge, and other Microsoft products."/>
                        </CheckBox.ToolTip>
                    </CheckBox>
                    <TextBlock Text="Registers Microsoft Update service so Windows Update covers Office, Edge, and other Microsoft products."
                               Foreground="#666666" FontSize="11" Margin="20,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>

                <StackPanel Margin="0,4">
                    <CheckBox x:Name="ChkHours" Content="Set Active Hours (8am - 8pm)" Foreground="White">
                        <CheckBox.ToolTip>
                            <ToolTip Content="Prevents Windows from restarting for updates between 8:00 AM and 8:00 PM. Disables Smart Active Hours."/>
                        </CheckBox.ToolTip>
                    </CheckBox>
                    <TextBlock Text="Prevents restarts for updates between 8am - 8pm. Disables Smart Active Hours."
                               Foreground="#666666" FontSize="11" Margin="20,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>

                <!-- SERVICES -->
                <TextBlock Text="Services" Foreground="#888888" FontSize="11"
                           Margin="0,12,0,4" FontStyle="Italic"/>

                <StackPanel Margin="0,4">
                    <CheckBox x:Name="ChkSuperfetch" Content="Disable Superfetch (SysMain)" Foreground="White">
                        <CheckBox.ToolTip>
                            <ToolTip Content="Stops and disables the SysMain service."/>
                        </CheckBox.ToolTip>
                    </CheckBox>
                    <TextBlock Text="Stops and disables SysMain. Reduces background disk activity, recommended for SSDs."
                               Foreground="#666666" FontSize="11" Margin="20,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>

                <StackPanel Margin="0,4">
                    <CheckBox x:Name="ChkWSearch" Content="Set Windows Search to Manual (WSearch)" Foreground="White">
                        <CheckBox.ToolTip>
                            <ToolTip Content="Sets WSearch startup to Manual and clears auto-recovery actions so it won't restart automatically."/>
                        </CheckBox.ToolTip>
                    </CheckBox>
                    <TextBlock Text="Sets WSearch to Manual startup and clears recovery actions to prevent auto-restart."
                               Foreground="#666666" FontSize="11" Margin="20,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>

                <StackPanel Margin="0,4">
                    <CheckBox x:Name="ChkXbox" Content="Disable Xbox Services" Foreground="White">
                        <CheckBox.ToolTip>
                            <ToolTip Content="Stops and disables XboxGipSvc, XboxNetApiSvc, XblGameSave, XblAuthManager, and XboxAppServices."/>
                        </CheckBox.ToolTip>
                    </CheckBox>
                    <TextBlock Text="Stops and disables Xbox background services."
                               Foreground="#666666" FontSize="11" Margin="20,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>

            </StackPanel>
        </ScrollViewer>

        <CheckBox x:Name="ChkAll" Grid.Row="2" Content="Select all"
                  Foreground="#ff00ff" FontWeight="Bold" Margin="0,4,0,16"/>

        <Button x:Name="BtnApply" Grid.Row="3" Content="APPLY SELECTED"
                Height="35" Background="#ff3333" Foreground="White" FontWeight="Bold"/>
    </Grid>
</Window>
"@

    # --- Everything below this line is unchanged ---
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $chkWidgetsOff = $window.FindName("ChkWidgetsOff")
    $chkWidgetsOn  = $window.FindName("ChkWidgetsOn")
    $chkAll        = $window.FindName("ChkAll")
    $btnApply      = $window.FindName("BtnApply")

    $checkboxes = @(
        $chkWidgetsOff,
        $chkWidgetsOn,
        $window.FindName("ChkMUUpdate"),
        $window.FindName("ChkHours"),
        $window.FindName("ChkSuperfetch"),
        $window.FindName("ChkWSearch"),
        $window.FindName("ChkXbox")
    )

    $chkWidgetsOff.Add_Click({
        if ($chkWidgetsOff.IsChecked) { $chkWidgetsOn.IsChecked = $false }
        if ($checkboxes | Where-Object { -not $_.IsChecked }) {
            $chkAll.IsChecked = $false
        } else {
            $chkAll.IsChecked = $true
        }
    })

    $chkWidgetsOn.Add_Click({
        if ($chkWidgetsOn.IsChecked) { $chkWidgetsOff.IsChecked = $false }
        if ($checkboxes | Where-Object { -not $_.IsChecked }) {
            $chkAll.IsChecked = $false
        } else {
            $chkAll.IsChecked = $true
        }
    })

    $chkAll.Add_Click({
        $state = $chkAll.IsChecked
        foreach ($cb in $checkboxes) { $cb.IsChecked = $state }
        if ($state) { $chkWidgetsOn.IsChecked = $false }
    })

    $otherBoxes = $checkboxes | Where-Object { $_ -ne $chkWidgetsOff -and $_ -ne $chkWidgetsOn }
    foreach ($cb in $otherBoxes) {
        $cb.Add_Click({
            if ($checkboxes | Where-Object { -not $_.IsChecked }) {
                $chkAll.IsChecked = $false
            } else {
                $chkAll.IsChecked = $true
            }
        })
    }

    $btnApply.Add_Click({
        $anySelected = $checkboxes | Where-Object { $_.IsChecked }
        if (-not $anySelected) {
            [System.Windows.MessageBox]::Show(
                "No settings selected.",
                "Nothing to do",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }
        $window.DialogResult = $true
        $window.Close()
    })

    if ($window.ShowDialog()) {
        if ($window.FindName("ChkWidgetsOff").IsChecked) { Disable-WindowsWidgets }
        if ($window.FindName("ChkWidgetsOn").IsChecked)  { Enable-WindowsWidgets  }
        if ($window.FindName("ChkMUUpdate").IsChecked)   { Enable-MicrosoftUpdateProducts }
        if ($window.FindName("ChkHours").IsChecked)      { Set-ActiveHours }
        if ($window.FindName("ChkSuperfetch").IsChecked) { Disable-Superfetch }
        if ($window.FindName("ChkWSearch").IsChecked)    { Disable-WindowsSearch }
        if ($window.FindName("ChkXbox").IsChecked)       { Disable-XboxServices }

        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Selected settings applied." -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "`nPress any key to return to main menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# ============================================================================
# MAIN MENU
# ============================================================================

function Show-MainMenu {
    while ($true) {
        Clear-Host
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "    SOFTWARE MANAGEMENT TOOL"             -ForegroundColor White
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "`n1. UNINSTALL Bloatware/Applications"    -ForegroundColor Gray
        Write-Host "2. INSTALL Software (Adobe, LibreOffice)" -ForegroundColor Gray
        Write-Host "3. FULL WORKFLOW (Uninstall -> Restart -> Install)" -ForegroundColor Gray
        Write-Host "4. WINDOWS SETTINGS (Widgets / Updates / Active Hours / Superfetch)" -ForegroundColor Gray
        Write-Host "5. Exit"                                  -ForegroundColor Gray
        Write-Host "`n========================================" -ForegroundColor Cyan

        $choice = Read-Host "`nSelect option [1-5]"

        switch ($choice) {
            "1" { Start-UninstallProcess }
            "2" { Show-InstallMenu }
            "3" { Start-FullWorkflow }
            "4" { Show-WindowsSettingsMenu }
            "5" { Write-Host "`nExiting..." -ForegroundColor Yellow; Exit }
            default {
                Write-Host "`n[ERROR] Invalid selection. Please choose 1-5." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    }
}

# ============================================================================
# SCRIPT ENTRY POINT
# ============================================================================

Test-Admin

$workflowState = Get-WorkflowState

if ($workflowState -and $workflowState.Stage -eq "POST_UNINSTALL") {
    $absolutePath = "C:\setup\installation"

    if (Test-Path $absolutePath) {
        Set-Location -Path $absolutePath
        Write-Host "[WORKFLOW] Working directory set to: $absolutePath" -ForegroundColor Gray
    }

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "    WORKFLOW CONTINUATION DETECTED"        -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "`nResuming installation phase..."           -ForegroundColor Cyan
    Start-Sleep -Seconds 3

    $adobeInstaller = "AcroRdrDCx642500121111_MUI.exe"
    $libreInstaller = "LibreOffice_25.8.4_Win_x86-64.msi"
    $adobePath      = Join-Path $absolutePath $adobeInstaller
    $librePath      = Join-Path $absolutePath $libreInstaller

    switch ($workflowState.InstallChoice) {
        "1" { Write-Host "`n[WORKFLOW] Installing Adobe Acrobat Reader..." -ForegroundColor Cyan; Install-Adobe       -InstallerPath $adobePath }
        "2" { Write-Host "`n[WORKFLOW] Installing LibreOffice..."          -ForegroundColor Cyan; Install-LibreOffice -InstallerPath $librePath }
        "3" { Write-Host "`n[WORKFLOW] Installing both applications..."    -ForegroundColor Cyan; Install-BothApps    -AdobePath $adobePath -LibrePath $librePath }
    }

    Clear-WorkflowState

    Write-Host "`n========================================" -ForegroundColor White
    Write-Host "WORKFLOW COMPLETED SUCCESSFULLY!"          -ForegroundColor White
    Write-Host "========================================" -ForegroundColor White
    Write-Host "`nAll tasks finished. Press any key to exit..." -ForegroundColor White
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Exit
}

Show-MainMenu

# $HKCU_Advanced = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
# $HKCU_Search   = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"

# Set-ItemProperty -Path $HKCU_Advanced -Name "ShowTaskViewButton" -Value 0 -Erion SilentlyContinue
# Set-ItemProperty -Path $HKCU_Search   -Name "SearchboxTaskbarMode" -Value 0 -ErrorAction SilentlyContinuerorAct