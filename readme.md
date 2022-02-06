# Microsoft Office Removal Tool

```plain
 ___   ___ ___      _____ _____ _____
|_  | |   |_  |    |   __|   __|     |
 _| |_| | |_| |_   |   __|   __| | | | 
|_____|___|_____|  |__|  |__|  |_|_|_| 

Microsoft Office Removal Tool
by Aaron Viehl (101 Frankfurt)
einsnulleins.de
```

## Synopsis

This script downloads the current Office uninstaller from Microsoft and tries to remove all Office installations on this computer.

If you wish it tries to install the newest Office365 build as well.

You can choose between 2 methods of uninstalling:\
Default method will use the [Microsoft Support and Recovery Assistant (SaRA)](https://docs.microsoft.com/en-us/office365/troubleshoot/administration/sara-command-line-version) for uninstalling.\
By using `-UseSetupRemoval` the Office365 setup method will be used.

## Parameter

| Parameter              | Usage                                                                   |
|------------------------|-------------------------------------------------------------------------|
| -InstallOffice365      | The script will try to install the newest Office365 build after removal |
| -SuppressReboot        | No reboot will be executed after script is done                         |
| -UseSetupRemoval       | Will use the official Office365 setup instead of SaRA                   |
| -RunAgain              | Will skip the stage mechanism - whole script will run again             |
| -Force                 | Non-interactive - No user interaction required                          |
| -SecondsToReboot [int] | Seconds until the computer will reboot                                  |

## Example

  ``.\msoffice-removal-tool.ps1 -InstallOffice365 -SuppressReboot -Force``

## Stage mechanism

To make sure that the program will only do the necessary parts a _stage mechanism_ is builtin. After every stage a registry value will be written to ``HKLM:\Software\OEM\101\M365\Install\CurrentStage``

To overwrite all stages and restart the whole script use ``-RunAgain``.

## Usage

You may use this script without explicit download:

```powershell
iwr https://raw.githubusercontent.com/Admonstrator/msoffice-removal-tool/main/msoffice-removal-tool.ps1 -OutFile msoffice-removal-tool.ps1; powershell -ExecutionPolicy Bypass .\msoffice-removal-tool.ps1
```
