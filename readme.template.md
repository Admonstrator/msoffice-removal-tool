<div align="center">

## Microsoft Office Removal Tool

**Remove Office installations with ease!**

[![Latest Release](https://img.shields.io/github/v/release/Admonstrator/msoffice-removal-tool?style=for-the-badge&logo=github&color=blue)](https://github.com/Admonstrator/msoffice-removal-tool/releases/latest) [![License](https://img.shields.io/github/license/Admonstrator/msoffice-removal-tool?style=for-the-badge)](LICENSE) [![Stars](https://img.shields.io/github/stars/Admonstrator/msoffice-removal-tool?style=for-the-badge)](https://github.com/Admonstrator/msoffice-removal-tool/stargazers)

---

## 💖 Support the Project

If you find this tool helpful, consider supporting its development:

[![GitHub Sponsors](https://img.shields.io/badge/GitHub-Sponsors-EA4AAA?style=for-the-badge&logo=github)](https://github.com/sponsors/admonstrator) [![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://buymeacoffee.com/admon) [![Ko-fi](https://img.shields.io/badge/Ko--fi-FF5E5B?style=for-the-badge&logo=ko-fi&logoColor=white)](https://ko-fi.com/admon) [![PayPal](https://img.shields.io/badge/PayPal-00457C?style=for-the-badge&logo=paypal&logoColor=white)](https://paypal.me/aaronviehl)

</div>

---

## 📖 About

This script downloads the current Office uninstaller from Microsoft and automatically removes all Office installations on your computer.

The script features two removal methods and can optionally install the newest Office 365 build after removal.

Created by [Aaron Viehl](https://github.com/Admonstrator) for the community.

> ⚠️ **Disclaimer** – This is an independent project, not officially affiliated with Microsoft

---

## ✨ Features

- 🚀 **Automatic Removal** – Removes all Office installations from your computer
- 🔄 **Flexible Methods** – Choose between SaRA or Office365 setup removal
- 📦 **Optional Installation** – Install Office365 after removal if desired
- 🛡️ **Safe Processing** – Built-in stage mechanism to track progress
- ⚡ **Flexible Options** – Multiple parameters for customized removal
- 🔄 **Automatic Reboot** – Optional automatic reboot after completion

---

## 📋 Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | Windows (Windows 10 or later recommended) |
| **PowerShell** | PowerShell 5.0 or later |
| **Permissions** | Administrator privileges required |
| **Internet** | Required for downloading Office uninstaller |

---

## 🚀 Quick Start

Run the script without cloning the repository:

```powershell
iwr https://get.admon.me/remove-msoffice -OutFile msoffice-removal-tool.ps1; powershell -ExecutionPolicy Bypass .\msoffice-removal-tool.ps1
```

---

## 🎛️ Parameters

The `msoffice-removal-tool.ps1` script supports the following parameters:

| Parameter | Description |
|-----------|-------------|
| `-InstallOffice365` | The script will try to install the newest Office365 build after removal |
| `-SuppressReboot` | No reboot will be executed after script is done |
| `-UseSetupRemoval` | Will use the official Office365 setup instead of SaRA for uninstalling |
| `-RunAgain` | Will skip the stage mechanism - whole script will run again |
| `-Force` | Non-interactive mode - No user interaction required |
| `-SecondsToReboot [int]` | Seconds until the computer will reboot (default: 120) |

---

## 📚 Usage Examples

### Standard Removal

Remove Office with the default SaRA method:

```powershell
.\msoffice-removal-tool.ps1
```

### Removal with Reinstallation

Remove Office and install the latest Office365 build:

```powershell
.\msoffice-removal-tool.ps1 -InstallOffice365
```

### Force Removal Without Reboot

Skip all prompts and prevent automatic reboot:

```powershell
.\msoffice-removal-tool.ps1 -Force -SuppressReboot
```

### Using Office Setup Removal

Use the Office365 setup method instead of SaRA:

```powershell
.\msoffice-removal-tool.ps1 -UseSetupRemoval -Force
```

### Complete Unattended Removal with Reinstallation

For fully unattended operation with reinstallation:

```powershell
.\msoffice-removal-tool.ps1 -InstallOffice365 -Force -SecondsToReboot 60
```

### Restarting the Entire Process

To skip the stage mechanism and run the entire script again:

```powershell
.\msoffice-removal-tool.ps1 -RunAgain -Force
```

---

## 🔍 Key Features Explained

### 🎯 Stage Mechanism

The script uses a built-in stage mechanism to ensure it only performs necessary steps and doesn't repeat work unnecessarily.

After every stage, a registry value is written to:
```
HKLM:\Software\OEM\Singleton-Factory-GmbH\M365\Install\CurrentStage
```

The mechanism is:
- ✅ Applied automatically
- ✅ Persistent across script restarts
- ✅ Can be overridden with `-RunAgain` parameter

### 🔧 Removal Methods

**Default Method (SaRA):**
- Uses the [Microsoft Support and Recovery Assistant](https://docs.microsoft.com/en-us/office365/troubleshoot/administration/sara-command-line-version)
- Recommended for most users
- Less intrusive than setup method

**Setup Method:**
- Uses the official Office365 setup
- More thorough removal
- Enabled with `-UseSetupRemoval` parameter

### 📦 Office365 Installation

When using `-InstallOffice365`, the script will:
- ✅ Automatically download the latest Office365 build
- ✅ Install Office365 after removal completes
- ✅ Configure the XML installation files from the `office365-installer/` directory

---

## 💡 Getting Help

Need assistance or have questions?

- 🐛 [Report issues on GitHub](https://github.com/Admonstrator/msoffice-removal-tool/issues) – Bug reports and feature requests
- 📧 Contact via GitHub – For private inquiries

---

## ⚠️ Disclaimer

This script is provided **as-is** without any warranty. Use it at your own risk.

It may potentially:
- 🔥 Remove important files or configurations
- 🔥 Cause system instability
- 🔥 Require manual recovery steps

**You have been warned!**

Always read the documentation carefully and understand what a script does before running it. Create a system restore point before executing this script.

---

## 📜 License

This project is licensed under the **MIT License** – see the [LICENSE](LICENSE) file for details.

---

<div align="center">

**Made with ❤️ by [Admonstrator](https://github.com/Admonstrator)**

⭐ If you find this useful, please star the repository!

</div>
