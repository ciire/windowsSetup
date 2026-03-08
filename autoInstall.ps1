# 1. Self-Elevation: Request Admin credentials
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Admin rights required. Requesting elevation..." -ForegroundColor Yellow
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    try {
        Start-Process powershell -Verb RunAs -ArgumentList $arguments
    } catch {
        Write-Host "Error: You must provide an Admin password to install." -ForegroundColor Red
    }
    Exit
}

# 2. Set the installer filename (Matches your 64-bit MUI file)
$installerName = "AcroRdrDCx642500121111_MUI.exe"
$installerPath = Join-Path -Path $PSScriptRoot -ChildPath $installerName

# Check if the file actually exists in the folder
if (-not (Test-Path $installerPath)) {
    Write-Host "Error: Could not find $installerName in this folder!" -ForegroundColor Red
    Pause
    Exit
}

# 3. Execute Silent Install
Write-Host "Installing Adobe Acrobat Reader 64-bit Silently..." -ForegroundColor Cyan

# Adobe Enterprise Switches:
# /sAll = Silent install for all components
# /rs   = Reboot Suppress (prevents the PC from restarting automatically)
# /msi  = Passes additional parameters to the internal MSI
# /qn   = Quiet, No UI
$silentArgs = "/sAll /rs /msi /qn"

$process = Start-Process -FilePath $installerPath -ArgumentList $silentArgs -Wait -PassThru

# 4. Verification
if ($process.ExitCode -eq 0) {
    Write-Host "SUCCESS: Adobe Acrobat has been installed." -ForegroundColor Green
} else {
    Write-Host "Installation failed with Exit Code: $($process.ExitCode)" -ForegroundColor Red
}

Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")