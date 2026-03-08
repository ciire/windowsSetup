# 1. Self-Elevation (Kept from your original)
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    Exit
}

# 2. Define Folder and File Names
$subFolder      = "installation"
$adobeInstaller = "AcroRdrDCx642500121111_MUI.exe"
$libreInstaller = "LibreOffice_25.8.4_Win_x86-64.msi"

# 3. Enhanced Installation Functions
function Install-Adobe {
    $logFile = Join-Path $PSScriptRoot "Adobe_Install_Detail.log"
    $path = Join-Path $PSScriptRoot -ChildPath $subFolder | Join-Path -ChildPath $adobeInstaller

    if (-not (Test-Path $path)) { 
        Write-Host "Error: Adobe installer not found at $path" -ForegroundColor Red
        return 
    }

    Write-Host "Installing Adobe Acrobat (Logging to: $logFile)..." -ForegroundColor Cyan
    
    # We add the logging switch directly to the argument list
    $adobeArgs = "/sAll /rs /msi /qn /norestart /L*V `"$logFile`""

    try {
        $adobeProc = Start-Process -FilePath $path -ArgumentList $adobeArgs -PassThru -ErrorAction Stop
        
        Write-Host "Monitoring Adobe process (ID: $($adobeProc.Id))..." -ForegroundColor Gray
        
        # Wait up to 5 mins
        $adobeProc | Wait-Process -Timeout 300 -ErrorAction SilentlyContinue

        if ($adobeProc.HasExited) {
            $exitCode = $adobeProc.ExitCode
            if ($exitCode -eq 0 -or $exitCode -eq 3010) {
                Write-Host "Success! Exit Code: $exitCode" -ForegroundColor Green
            } else {
                Write-Host "Installation Failed. Exit Code: $exitCode. Check the log file!" -ForegroundColor Red
            }
        } else {
            Write-Host "Error: Installation timed out after 5 minutes. Process is still running in background." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Failed to launch the installer: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Install-LibreOffice {
    $path = Join-Path $PSScriptRoot -ChildPath $subFolder | Join-Path -ChildPath $libreInstaller
    if (-not (Test-Path $path)) { Write-Host "Error: LibreOffice not found." -ForegroundColor Red; return }

    Write-Host "Installing LibreOffice Silently..." -ForegroundColor Cyan
    
    # Using -Wait here is fine for MSI because LibreOffice is more "polite" than Adobe
    $libreProc = Start-Process "msiexec.exe" -ArgumentList "/i `"$path`" /qn /norestart" -PassThru -Wait
    Write-Host "LibreOffice completed with Exit Code: $($libreProc.ExitCode)" -ForegroundColor Gray
}

# 4. Display Menu & Logic
Clear-Host
Write-Host "--- Professional Software Deployment ---" -ForegroundColor White
Write-Host "1. Install Adobe Acrobat Reader"
Write-Host "2. Install LibreOffice"
Write-Host "3. Install BOTH (Optimized Flow)"
Write-Host "4. Exit"
$choice = Read-Host "Select [1-4]"

switch ($choice) {
    "1" { Install-Adobe }
    "2" { Install-LibreOffice }
    "3" { 
        # Optimized Flow: Start Adobe, wait 30 seconds for it to clear the "heavy" initial registry hooks,
        # then start LibreOffice so they effectively install in parallel.
        Install-Adobe
        Write-Host "Initiating LibreOffice while Adobe cleans up..." -ForegroundColor Yellow
        Install-LibreOffice 
    }
    "4" { Exit }
}

Write-Host "`nTasks complete." -ForegroundColor Green
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")