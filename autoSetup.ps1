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
# Manages multi-step workflows that span system restarts
# ============================================================================

function Get-WorkflowState {
    <#
    .SYNOPSIS
    Retrieves the current workflo+w state from temp storage
    #>
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
    <#
    .SYNOPSIS
    Saves workflow state for post-restart continuation
    #>
    param(
        [string]$Stage,
        [string]$InstallChoice
    )
    
    $statePath = Join-Path $env:TEMP "software_manager_workflow.json"
    $state = @{
        Stage = $Stage
        InstallChoice = $InstallChoice
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    $state | ConvertTo-Json | Out-File $statePath -Force
    Write-Host "[WORKFLOW] State saved: $Stage" -ForegroundColor Gray
}

function Clear-WorkflowState {
    <#
    .SYNOPSIS
    Removes workflow state file after completion
    #>
    $statePath = Join-Path $env:TEMP "software_manager_workflow.json"
    if (Test-Path $statePath) {
        Remove-Item $statePath -Force
        Write-Host "[WORKFLOW] State cleared." -ForegroundColor Gray
    }
}

function Set-WorkflowAutoRun {
    <#
    .SYNOPSIS
    Configures script to auto-run after restart via RunOnce registry key
    #>
    $scriptPath = $PSCommandPath
    $scriptDir = Split-Path -Parent $scriptPath
    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    $name = "SoftwareManagerWorkflow"
    
    # Ensure working directory is correct when script auto-runs
    $command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command `"Set-Location '$scriptDir'; & '$scriptPath'`""
    
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
    <#
    .SYNOPSIS
    Checks for admin privileges and elevates if needed
    #>
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-NOT $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "Elevating privileges..." -ForegroundColor Yellow
        Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Exit
    }
}

# ============================================================================
# APPLICATION DISCOVERY
# ============================================================================

function Get-UnifiedAppList {
    <#
    .SYNOPSIS
    Scans for both Win32/64 and UWP applications
    .DESCRIPTION
    Combines registry-based apps and AppX packages into a unified list
    #>
    Write-Host "Scanning installed applications..." -ForegroundColor Cyan

    # Scan registry for Win32/64 applications
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    [array]$regApps = Get-ItemProperty $registryPaths -ErrorAction SilentlyContinue | 
        Where-Object { 
            $_.DisplayName -and 
            ($_.UninstallString -or $_.QuietUninstallString) -and 
            ($_.SystemComponent -ne 1) 
        } |
        Select-Object @{Name="DisplayName"; Expression={$_.DisplayName}}, 
                      @{Name="Id"; Expression={$_.PSChildName}}, 
                      @{Name="Type"; Expression={"Win32/64"}},
                      UninstallString

    # Scan for UWP/AppX packages
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
                          # Humanize package names
                          $n = $_.Name -replace 'Microsoft\.', '' -replace 'Windows\.', ''
                          $n = [regex]::Replace($n, '([a-z])([A-Z])', '$1 $2')
                          $n.Replace('.', ' ')
                      }}, 
                      @{Name="Id"; Expression={$_.PackageFullName}}, 
                      @{Name="Type"; Expression={"UWP"}},
                      @{Name="UninstallString"; Expression={"Remove-AppxPackage"}}

    # Combine and sort
    $combined = @()
    if ($regApps) { $combined += $regApps }
    if ($appxApps) { $combined += $appxApps }

    return $combined | Sort-Object DisplayName
}

# ============================================================================
# UNINSTALL GUI
# ============================================================================

function Show-UninstallGUI {
    <#
    .SYNOPSIS
    Displays interactive application selection interface
    .PARAMETER AppList
    Array of applications to display
    .PARAMETER PreSelectedNames
    Application names to pre-check (from template)
    #>
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

    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Wrap apps with IsChecked property for data binding
    $wrappedList = foreach($app in $AppList) { 
        [PSCustomObject]@{ DisplayName = $app.DisplayName; IsChecked = $false; Original = $app } 
    }
    
    $listBox = $window.FindName("AppListBox")
    $listBox.ItemsSource = $wrappedList
    
    # Configure template toggle
    $templateToggle = $window.FindName("UseTemplate")
    
    if ($PreSelectedNames.Count -eq 0) {
        $templateToggle.IsEnabled = $false
        $templateToggle.Content = "No template found (apps_to_remove.txt)"
        $templateToggle.Foreground = [System.Windows.Media.Brushes]::Gray
    }

    # Template toggle event handler
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
        # Show friendly informational message about the process
        $result = [System.Windows.MessageBox]::Show(
            "Most apps will be removed silently in the background.`n`n" +
            "If an app can't be uninstalled automatically, its uninstaller window will open for you to complete manually.`n`n" +
            "Ready to begin?",
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
# WINDOWS WIDGETS REMOVAL
# ============================================================================

function Disable-WindowsWidgets {
    <#
    .SYNOPSIS
    Disables Windows Widgets via group policy registry key
    #>
    Write-Host "`n[ACTION] Disabling Widgets for all users..." -ForegroundColor Cyan
    
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
    $name = "AllowNewsAndInterests"

    try {
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
            Write-Host " [+] Created policy directory." -ForegroundColor Gray
        }

        Set-ItemProperty -Path $registryPath -Name $name -Value 0 -ErrorAction Stop
        Write-Host " [+] Widgets disabled for all users." -ForegroundColor Green

        # Restart Explorer to apply changes
        Write-Host " [+] Refreshing taskbar..." -ForegroundColor Gray
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    } 
    catch {
        Write-Host " [!] Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ============================================================================
# APPLICATION REMOVAL
# ============================================================================

function Invoke-AppRemoval {
    <#
    .SYNOPSIS
    Removes an application using appropriate uninstall method
    .DESCRIPTION
    Handles both UWP (AppX) and Win32/MSI applications with silent uninstall fallback
    #>
    param ([Parameter(Mandatory=$true)] $App)
    
    $name = $App.DisplayName
    Write-Host "`n[REMOVING] $name" -ForegroundColor Cyan

    try {
        if ($App.Type -eq "UWP") {
            # Remove UWP/AppX package for all users
            $pName = ($App.Id -split "_")[0]
            Get-AppxPackage -Name "*$pName*" -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
            Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -match $pName } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
            Write-Host " [SUCCESS] $name removed." -ForegroundColor Green
        } 
        else {
            # Handle Win32/MSI applications
            $uninstStr = $App.UninstallString.Trim()
            $exePath = ""
            $args = ""

            # Parse uninstall string and determine silent arguments
            if ($uninstStr -imatch "msiexec") {
                # MSI-based installer
                $exePath = "msiexec.exe"
                $args = ($uninstStr -ireplace ".*msiexec\.exe\s*", "" -ireplace "/I", "/X").Trim()
                $silentArgs = "$args /qn /norestart"
            } 
            elseif ($uninstStr -match '^"(.*)"\s*(.*)$') {
                # Quoted executable path
                $exePath = $matches[1]
                $args = $matches[2]
                $silentArgs = "$args /S /silent /quiet /norestart".Trim()
            } 
            else {
                # Unquoted or complex path
                if ($uninstStr -like "*.exe*") {
                    $split = $uninstStr -split ".exe", 2
                    $exePath = ($split[0] + ".exe").Replace('"','')
                    $args = $split[1].Trim()
                    $silentArgs = "$args /S /silent /quiet /norestart".Trim()
                } else {
                    $exePath = "cmd.exe"
                    $silentArgs = "/c $uninstStr /S /silent /quiet /norestart"
                }
            }

            # Attempt silent removal
            Write-Host " [>] Attempting silent removal..." -ForegroundColor Gray
            $proc = Start-Process -FilePath $exePath -ArgumentList $silentArgs -Verb RunAs -PassThru -Wait -ErrorAction SilentlyContinue
            
            Start-Sleep -Seconds 5 

            # Verify removal
            $stillExists = (Get-UnifiedAppList) | Where-Object { $_.Id -eq $App.Id }

            if ($stillExists) {
                # Fallback to GUI uninstaller
                Write-Host " [!] Silent removal failed. Launching standard GUI..." -ForegroundColor White
                Start-Process -FilePath $exePath -ArgumentList $args -Verb RunAs -Wait
            } else {
                Write-Host " [SUCCESS] $name removed." -ForegroundColor Green
            }
        }
    } 
    catch { 
        Write-Host " [!] Critical Error: Could not launch $name uninstaller." -ForegroundColor Red 
        Write-Host " [!] Details: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ============================================================================
# SOFTWARE DETECTION
# ============================================================================

function Test-AdobeInstalled {
    <#
    .SYNOPSIS
    Checks if Adobe Acrobat Reader is installed
    #>
    $paths = @(
        "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRdr.exe",
        "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRdr.exe",
        "C:\Program Files\Adobe\Acrobat Reader DC\Reader\Acrobat.exe"
    )
    foreach ($path in $paths) { 
        if (Test-Path $path) { return $true } 
    }
    return $false
}

function Test-LibreOfficeInstalled {
    <#
    .SYNOPSIS
    Checks if LibreOffice is installed
    #>
    return (Test-Path "C:\Program Files\LibreOffice\program\soffice.exe")
}

# ============================================================================
# SOFTWARE INSTALLATION
# ============================================================================

function Install-Adobe {
    <#
    .SYNOPSIS
    Installs Adobe Acrobat Reader silently
    .PARAMETER InstallerPath
    Full path to Adobe installer executable
    #>
    param([string]$InstallerPath)

    if (-not (Test-Path $InstallerPath)) { 
        Write-Host "[ERROR] Adobe installer not found at: $InstallerPath" -ForegroundColor Red
        return 
    }

    Write-Host "`n[INSTALLING] Adobe Acrobat Reader..." -ForegroundColor Cyan
    
    # Launch installer with silent parameters
    $proc = Start-Process -FilePath $InstallerPath -ArgumentList "/sAll /rs /msi /qn" -PassThru
    Write-Host " [>] Installer launched (PID: $($proc.Id))" -ForegroundColor Gray
    
    # Wait for launcher wrapper
    $proc | Wait-Process -Timeout 30 -ErrorAction SilentlyContinue
    
    # Monitor actual MSI installation
    Write-Host " [>] Monitoring installation progress..." -ForegroundColor Gray
    
    $maxWait = 300  # 5 minutes timeout
    $elapsed = 0
    $checkInterval = 5
    
    while ($elapsed -lt $maxWait) {
        $adobeMsi = Get-Process -Name msiexec -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.MainWindowTitle -like "*Adobe*" -or 
                $_.CommandLine -like "*Acro*"
            }
        
        if (-not $adobeMsi) {
            Write-Host "`n [SUCCESS] Adobe installation complete." -ForegroundColor Green
            return
        }
        
        Write-Host "." -NoNewline -ForegroundColor Gray
        Start-Sleep -Seconds $checkInterval
        $elapsed += $checkInterval
    }
    
    Write-Host "`n [TIMEOUT] Installation may still be running in background." -ForegroundColor Yellow
}

function Install-LibreOffice {
    <#
    .SYNOPSIS
    Installs LibreOffice silently via MSI
    .PARAMETER InstallerPath
    Full path to LibreOffice MSI installer
    #>
    param([string]$InstallerPath)

    if (-not (Test-Path $InstallerPath)) { 
        Write-Host "[ERROR] LibreOffice installer not found at: $InstallerPath" -ForegroundColor Red
        return 
    }

    Write-Host "`n[INSTALLING] LibreOffice..." -ForegroundColor Cyan
    
    $proc = Start-Process "msiexec.exe" -ArgumentList "/i `"$InstallerPath`" /qn /norestart" -PassThru -Wait
    
    if ($proc.ExitCode -eq 0) {
        Write-Host " [SUCCESS] LibreOffice installed successfully." -ForegroundColor Green
    } elseif ($proc.ExitCode -eq 3010) {
        Write-Host " [SUCCESS] LibreOffice installed (restart required)." -ForegroundColor Yellow
    } else {
        Write-Host " [ERROR] Installation failed (Exit Code: $($proc.ExitCode))." -ForegroundColor Red
    }
}

function Install-BothApps {
    <#
    .SYNOPSIS
    Installs both Adobe and LibreOffice sequentially
    .DESCRIPTION
    Waits for Adobe MSI to complete before starting LibreOffice to avoid conflicts
    #>
    param([string]$AdobePath, [string]$LibrePath)
    
    Write-Host "`n[STEP 1] Installing Adobe Acrobat..." -ForegroundColor Cyan
    
    # Start Adobe and wait for extraction wrapper
    $adobeProc = Start-Process -FilePath $AdobePath -ArgumentList "/sAll /rs /msi /qn /norestart" -PassThru -Wait
    
    Write-Host " [>] Adobe wrapper finished. Waiting for background MSI..." -ForegroundColor Gray
    Start-Sleep -Seconds 15 
    
    # Wait for MSI installer to finish
    while (Get-Process msiexec -ErrorAction SilentlyContinue) {
        Write-Host "." -NoNewline -ForegroundColor Gray
        Start-Sleep -Seconds 5
    }

    Write-Host "`n[STEP 2] Installing LibreOffice..." -ForegroundColor Cyan
    if (Test-Path $LibrePath) {
        $libreProc = Start-Process "msiexec.exe" -ArgumentList "/i `"$LibrePath`" /qn /norestart" -PassThru -Wait
        Write-Host " [SUCCESS] All installations finished." -ForegroundColor Green
    }
}

# ============================================================================
# INSTALLATION MENU
# ============================================================================

function Show-InstallMenu {
    <#
    .SYNOPSIS
    Displays software installation options
    #>
    # Define installer paths
    $absolutePath = "C:\setup\installation"
    $adobeInstaller = "AcroRdrDCx642500121111_MUI.exe"
    $libreInstaller = "LibreOffice_25.8.4_Win_x86-64.msi"
    
    $adobePath = Join-Path $absolutePath $adobeInstaller
    $librePath = Join-Path $absolutePath $libreInstaller

    # Validate installer availability
    if (-not (Test-Path $adobePath)) {
        Write-Host "[WARNING] Adobe installer not found at $adobePath" -ForegroundColor Yellow
    }
    
    Clear-Host
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    SOFTWARE INSTALLATION MENU" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`n1. Install Adobe Acrobat Reader" -ForegroundColor White
    Write-Host "2. Install LibreOffice" -ForegroundColor White
    Write-Host "3. Install BOTH" -ForegroundColor White
    Write-Host "4. Return to Main Menu" -ForegroundColor Gray
    Write-Host "`n========================================" -ForegroundColor Cyan
    
    $choice = Read-Host "`nSelect option [1-4]"
    
    switch ($choice) {
        "1" { Install-Adobe -InstallerPath $adobePath }
        "2" { Install-LibreOffice -InstallerPath $librePath }
        "3" { Install-BothApps -AdobePath $adobePath -LibrePath $librePath }
        "4" { return }
        default { 
            Write-Host "[ERROR] Invalid selection." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-InstallMenu
        }
    }
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Press any key to return to main menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ============================================================================
# UNINSTALL WORKFLOW
# ============================================================================

function Start-UninstallProcess {
    <#
    .SYNOPSIS
    Initiates the application uninstall workflow
    #>
    $configPath = Join-Path $PSScriptRoot "apps_to_remove.txt"
    $preSelected = @()

    # Load template if available
    if (Test-Path $configPath) {
        $preSelected = Get-Content $configPath
    }

    $allApps = Get-UnifiedAppList
    $guiResult = Show-UninstallGUI -AppList $allApps -PreSelectedNames $preSelected

    if ($guiResult) {
        # Remove Windows Widgets if requested
        if ($guiResult.RemoveWidgets) { 
            Disable-WindowsWidgets 
        }
        
        # Save selection template
        if ($guiResult.SaveRequested -and $guiResult.AppsToUninst) {
            $guiResult.AppsToUninst.DisplayName | Out-File $configPath -Force
            Write-Host "`n[CONFIG] Template saved to: $configPath" -ForegroundColor Gray
        }

        # Execute uninstalls
        if ($guiResult.AppsToUninst) {
            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host "Starting uninstallation process..." -ForegroundColor White
            Write-Host "========================================" -ForegroundColor Cyan
            
            foreach ($item in $guiResult.AppsToUninst) {
                Invoke-AppRemoval -App $item
            }
            
            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host "Cleanup completed." -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Cyan
        }
        
        Write-Host "`nPress any key to return to main menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# ============================================================================
# FULL WORKFLOW (UNINSTALL → RESTART → INSTALL)
# ============================================================================

function Start-FullWorkflow {
    <#
    .SYNOPSIS
    Executes complete workflow: uninstall bloatware, restart, then install software
    .DESCRIPTION
    Auto-restart is MANDATORY in this workflow - it always configures auto-run after restart
    #>
    Clear-Host
    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "    FULL WORKFLOW MODE" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "`nThis will:" -ForegroundColor Yellow
    Write-Host "Step 1. Uninstall selected bloatware" -ForegroundColor White
    Write-Host "Step 2. Restart your computer" -ForegroundColor White
    Write-Host "Step 3. Automatically install software after restart" -ForegroundColor White
    Write-Host "`n========================================" -ForegroundColor Magenta
    
    # Select post-restart software installation
    Write-Host "`nWhat software should be installed AFTER restart?" -ForegroundColor Cyan
    Write-Host "1. Adobe Acrobat Reader only" -ForegroundColor White
    Write-Host "2. LibreOffice only" -ForegroundColor White
    Write-Host "3. Both" -ForegroundColor White
    Write-Host "4. Cancel workflow" -ForegroundColor Gray
    
    $installChoice = Read-Host "`nSelect option [1-4]"
    
    if ($installChoice -eq "4") {
        Write-Host "`nWorkflow cancelled." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        return
    }
    
    if ($installChoice -notmatch '^[1-3]$') {
        Write-Host "`n[ERROR] Invalid selection." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }
    
    # Save workflow state for post-restart continuation
    Set-WorkflowState -Stage "POST_UNINSTALL" -InstallChoice $installChoice
    
    # Execute uninstall phase
    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "Starting uninstall process..." -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Magenta
    Start-Sleep -Seconds 2
    
    $configPath = Join-Path $PSScriptRoot "apps_to_remove.txt"
    $preSelected = @()

    if (Test-Path $configPath) {
        $preSelected = Get-Content $configPath
    }

    $allApps = Get-UnifiedAppList
    $guiResult = Show-UninstallGUI -AppList $allApps -PreSelectedNames $preSelected

    if (-not $guiResult) {
        Write-Host "`n[CANCELLED] Workflow aborted by user." -ForegroundColor Yellow
        Clear-WorkflowState
        Start-Sleep -Seconds 2
        return
    }

    # Save template if requested
    if ($guiResult.SaveRequested -and $guiResult.AppsToUninst) {
        $guiResult.AppsToUninst.DisplayName | Out-File $configPath -Force
        Write-Host "`n[CONFIG] Template saved." -ForegroundColor Gray
    }

    # Remove Windows Widgets if requested
    if ($guiResult.RemoveWidgets) { 
        Disable-WindowsWidgets 
    }

    # Execute uninstalls
    if ($guiResult.AppsToUninst) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Starting uninstallation process..." -ForegroundColor White
        Write-Host "========================================" -ForegroundColor Cyan
        
        foreach ($item in $guiResult.AppsToUninst) {
            Invoke-AppRemoval -App $item
        }
        
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Uninstallation completed." -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Cyan
    }
    
    # MANDATORY: Configure post-restart auto-run (no user choice)
    Write-Host "`n[WORKFLOW] Configuring automatic restart and installation..." -ForegroundColor Cyan
    Set-WorkflowAutoRun
    
    # Confirm and initiate restart
    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "READY TO RESTART" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "`nThe computer will restart now." -ForegroundColor White
    Write-Host "After restart, the installation will begin automatically." -ForegroundColor Cyan
    Write-Host "`nPress any key to restart now, or close this window to cancel..." -ForegroundColor Yellow
    
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
    Write-Host "`nRestarting in 5 seconds..." -ForegroundColor Red
    Start-Sleep -Seconds 5
    
    Restart-Computer -Force
}

# ============================================================================
# MAIN MENU
# ============================================================================

function Show-MainMenu {
    <#
    .SYNOPSIS
    Displays primary navigation menu
    #>
    while ($true) {
        Clear-Host
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "    SOFTWARE MANAGEMENT TOOL" -ForegroundColor White
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "`n1. UNINSTALL Bloatware/Applications" -ForegroundColor Gray
        Write-Host "2. INSTALL Software (Adobe, LibreOffice)" -ForegroundColor Gray
        Write-Host "3. FULL WORKFLOW (Uninstall -> Restart -> Install)" -ForegroundColor Gray
        Write-Host "4. Exit" -ForegroundColor Gray
        Write-Host "`n========================================" -ForegroundColor Cyan
        
        $choice = Read-Host "`nSelect option [1-4]"
        
        switch ($choice) {
            "1" { Start-UninstallProcess }
            "2" { Show-InstallMenu }
            "3" { Start-FullWorkflow }
            "4" { 
                Write-Host "`nExiting..." -ForegroundColor Yellow
                Exit 
            }
            default { 
                Write-Host "`n[ERROR] Invalid selection. Please choose 1-4." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    }
}

# ============================================================================
# SCRIPT ENTRY POINT
# ============================================================================

# Ensure administrative privileges
Test-Admin

# Check for workflow continuation after restart
$workflowState = Get-WorkflowState

if ($workflowState -and $workflowState.Stage -eq "POST_UNINSTALL") {
    # Resume workflow: install software
    $absolutePath = "C:\setup\installation"
    
    # Set working directory for installer access
    if (Test-Path $absolutePath) {
        Set-Location -Path $absolutePath
        Write-Host "[WORKFLOW] Working directory set to: $absolutePath" -ForegroundColor Gray
    }

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "    WORKFLOW CONTINUATION DETECTED" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "`nResuming installation phase..." -ForegroundColor Cyan
    Start-Sleep -Seconds 3
    
    # Define installer paths
    $adobeInstaller = "AcroRdrDCx642500121111_MUI.exe"
    $libreInstaller = "LibreOffice_25.8.4_Win_x86-64.msi"
    
    $adobePath = Join-Path $absolutePath $adobeInstaller
    $librePath = Join-Path $absolutePath $libreInstaller
    
    # Execute installation based on saved choice
    switch ($workflowState.InstallChoice) {
        "1" { 
            Write-Host "`n[WORKFLOW] Installing Adobe Acrobat Reader..." -ForegroundColor Cyan
            Install-Adobe -InstallerPath $adobePath 
        }
        "2" { 
            Write-Host "`n[WORKFLOW] Installing LibreOffice..." -ForegroundColor Cyan
            Install-LibreOffice -InstallerPath $librePath 
        }
        "3" { 
            Write-Host "`n[WORKFLOW] Installing both applications..." -ForegroundColor Cyan
            Install-BothApps -AdobePath $adobePath -LibrePath $librePath 
        }
    }
    
    # Clean up workflow state
    Clear-WorkflowState
    
    Write-Host "`n========================================" -ForegroundColor White
    Write-Host "WORKFLOW COMPLETED SUCCESSFULLY!" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor White
    Write-Host "`nAll tasks finished. Press any key to exit..." -ForegroundColor White
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Exit
}

# Normal operation: display main menu
Show-MainMenu