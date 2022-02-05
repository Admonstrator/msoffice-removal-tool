# Microsoft Office Removal Tool

## Synopsis
This script downloads the current Office uninstaller from Microsoft and tries to remove all Office installations on this computer. If you wish it tries to install the newest Office365 build as well.

## Parameter

| Parameter | Usage  |
|---|---|
|  -InstallOffice365 | The script will try to install the newest Office365 build after removal  |
| -SuppressReboot | No reboot will be executed after script is done|
| -UseSetupRemoval | Will use the official Office365 setup instead of SaRA|
|-Force | Non-interactive - No user interaction required|

### Notes
  Version:  1.0
  
  Author: aaron.viehl@frankfurt.einsnulleins.de
  
  Creation Date: 2022-02-04

### Example
  ``.\msoffice-removal-tool.ps1 -InstallOffice365 -SuppressReboot -Force``