# windowsSetup

A unified PowerShell utility designed to streamline Windows deployment by removing bloatware, configuring system settings, and automating software installations via CSV.

## 🚀 Features

* **Bloatware Removal:** Scans both Win32 and UWP (Windows Store) applications with a GUI for selective uninstallation.
* **System Tweaks:** Quickly toggle Windows Widgets, disable Superfetch (SysMain), manage Windows Search, and disable Xbox background services.
* **CSV-Driven Installer:** Bulk install applications using a simple CSV manifest.
* **Uninstall and Install** A "Set and Forget" mode that uninstalls bloatware, reboots the PC, and automatically resumes software installation after login.
* **Registry Cleanup:** Automatically mounts user hives to ensure deep cleaning of uninstalled application traces for all users on the machine.

---

## 🛠 Prerequisites

* **Windows 10 or 11**
* **Administrator Privileges:** The script will automatically prompt for elevation if not run as Admin.
* **Execution Policy:** Ensure you can run scripts by running the following command in an elevated PowerShell window:
    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
    ```

---

## 📂 Project Structure

* `SoftwareManager.ps1`: The main script containing all logic and GUI.
* `installers.csv`: Your custom list of software to install (template below).
* `apps_to_remove.txt`: (Auto-generated) Stores your preferred bloatware removal list for future runs.

---

## 📥 Installation Setup (CSV Configuration)

To use the **Install** or **Uninstall and Install** features, you must provide a CSV file. For best results, keep your installers in a dedicated folder and ensure the `InstallerPath` is accurate.

### CSV Format Requirements
The script expects the following headers:
`Name, InstallerPath, InstallerType, SilentArgs`

| Column | Description | Example |
| :--- | :--- | :--- |
| **Name** | Display name of the software | `Mozilla Firefox` |
| **InstallerPath** | Full local path to the `.exe` or `.msi` | `C:\Installers\firefox.exe` |
| **InstallerType** | Must be either `exe` or `msi` | `exe` |
| **SilentArgs** | The switches for a "silent" install (can be left empty if you want to manually install) | `/S` or `/quiet` |

### Example `installers.csv`
```csv
Name,InstallerPath,InstallerType,SilentArgs
Mozilla Firefox,C:\Users\ec\Documents\GitHub\Installers\Firefox Installer.exe,exe,-ms
7-Zip,C:\Users\ec\Documents\GitHub\Installers\7z2600-x64.exe,exe,
VLC Media Player,C:\Users\ec\Documents\GitHub\Installers\vlc-3.0.23-win32.msi,msi,
LibreOffice,C:\Users\ec\Documents\GitHub\Installers\LibreOffice.msi,msi,/quiet /norestart
Adobe Acrobat Reader,C:\Users\ec\Documents\GitHub\Installers\AcroRdrDC.exe,exe,/sAll /rs