#requires -version 2
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  This script downloads the current Office uninstaller from Microsoft and tries to remove all Office installations on this computer.
  By choosing -InstallOffice365 it tries to install Office 365 as well.
.DESCRIPTION
  By choosing -InstallOffice365 it tries to install Office 365 as well.
.PARAMETER InstallOffice365
  Will install Office365 after removal.
.PARAMETER SuppressReboot
  Will supress the reboot after finishing the script.
.PARAMETER UseSetupRemoval
  Will use the setup method to remove current Office installations instead of SaRA.
.PARAMETER Force
  Skip user-input.
.INPUTS
  None
.OUTPUTS
  Just output on screen
.NOTES
  Version:        1.0
  Author:         aaron.viehl@frankfurt.einsnulleins.de
  Creation Date:  2022-02-04
  Purpose/Change: Initial script development
.EXAMPLE
  .\msoffice-removal-tool.ps1 -InstallOffice365 -SupressReboot
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Param (
    [switch]$InstallOffice365 = $False,
    [switch]$SuppressReboot = $False,
    [switch]$UseSetupRemoval = $False,
    [Switch]$Force = $False
)

#----------------------------------------------------------[Declarations]----------------------------------------------------------
$SaRA_URL = "https://aka.ms/SaRA_CommandLineVersionFiles"
$SaRA_ZIP = "$env:TEMP\SaRA.zip"
$SaRA_DIR = "$env:TEMP\SaRA"
$SaRA_EXE = "$SaRA_DIR\SaRAcmd.exe"
$Office365Setup_URL = "https://github.com/Admonstrator/msoffice-removal-tool/raw/main/office365-installer"
#-----------------------------------------------------------[Functions]------------------------------------------------------------
Function Invoke-OfficeUninstall {
    if (-Not (Test-Path "$SaRA_DIR")) {
        New-Item "$SaRA_DIR" -ItemType Directory
    }
    if ($UseSetupRemoval) {
        Write-Host "Invoking default setup method ..."
        Invoke-SetupOffice365 "$Office365Setup_URL/purge.xml"
    }
    else {
        Write-Host "Invoking SaRA method ..."
        Remove-SaRA
        Write-Host "Downloading most recent SaRA build ..."
        Invoke-SaRADownload
        Write-Host "Removing Office installations ..."
        Invoke-SaRA 
    }
}
Function Invoke-SaRADownload {    
    Start-BitsTransfer -Source "$SaRA_URL" -Destination "$SaRA_ZIP" 
    if (Test-Path "$SaRA_ZIP") {
        Write-Host "Unzipping ..."
        Expand-Archive -Path "$SaRA_ZIP" -DestinationPath "$SaRA_DIR" -Force
        if (Test-Path "$SaRA_DIR\DONE") {
            Move-Item "$SaRA_DIR\DONE\*" "$SaRA_DIR" -Force
            if (Test-Path "$SaRA_EXE") {
                Write-Host "SaRA downloaded successfully."
            }
            else {
                Write-Error "Download failed. Program terminated."
                Exit 1
            }
        }
    }
}

Function Remove-SaRA {  
    if (Test-Path "$SaRA_ZIP") {
        Remove-Item "$SaRA_ZIP" -Force
    }

    if (Test-Path "$SaRA_DIR") {
        Remove-Item "$SaRA_DIR" -Recurse -Force
    }
}
 
Function Stop-OfficeProcess {
    Write-Host "Stopping running Office applications ..."
    $OfficeProcessesArray = "lync", "winword", "excel", "msaccess", "mstore", "infopath", "setlang", "msouc", "ois", "onenote", "outlook", "powerpnt", "mspub", "groove", "visio", "winproj", "graph", "teams"
    foreach ($ProcessName in $OfficeProcessesArray) {
        if (get-process -Name $ProcessName -ErrorAction SilentlyContinue) {
            if (Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue) {
                Write-Output "Process $ProcessName was stopped."
            }
            else {
                Write-Warning "Process $ProcessName could not be stopped."
            }
        } 
    }
}

Function Invoke-SaRA {
    $SaRAProcess = Start-Process -FilePath "$SaRA_EXE" -ArgumentList "-S OfficeScrubScenario -AcceptEula" -Wait -PassThru -NoNewWindow
    switch ($SaRAProcess.ExitCode) {
        0 {
            Write-Host "Uninstall successful!"
            Break
        }
    
        7 {
            Write-Host "No office installations found."
            Break
        }

        8 {
            Write-Error "Multiple office installations found. Program need to be run in GUI mode."
            Exit 2
        }

        9 {
            Write-Error "Uninstall failed! Program need to be run in GUI mode."
            Exit 3
        }
    }
}

Function Invoke-SetupOffice365($Office365ConfigFile) {
    if ($InstallOffice365) {
        Write-Host "Downloading Office365 Installer ..."
        Start-BitsTransfer -Source "$Office365Setup_URL/setup.exe" -Destination "$SaRA_DIR\setup.exe"
        Start-BitsTransfer -Source "$Office365ConfigFile" -Destination "$SaRA_DIR\config.xml"
        Write-Host "Executing Office365 Setup ..."
        $OfficeSetup = Start-Process -FilePath "$SaRA_DIR\setup.exe" -ArgumentList "/configure $SaRA_DIR\config.xml" -Wait -PassThru -NoNewWindow 
        switch ($OfficeSetup.ExitCode) {
            0 {
                Write-Host "Install successful!"
                Break
            }

            1 {
                Write-Error "Install failed!"
                Break
            }
        }
    }
}

Function Invoke-Reboot {
    if (-not $SuppressReboot) {
        Start-Process -FilePath "$env:SystemRoot\system32\shutdown.exe" -ArgumentList "/r /c `"Reboot needed. System will reboot in 60 seconds.`" /t 60 /f /d p:4:1"
    }
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------
Write-Host " ___   ___ ___      _____ _____ _____ "
Write-Host "|_  | |   |_  |    |   __|   __|     |" 
Write-Host " _| |_| | |_| |_   |   __|   __| | | |" 
Write-Host "|_____|___|_____|  |__|  |__|  |_|_|_|" 
Write-Host ""
Write-Host "Microsoft Office Removal Tool"
Write-Host "by Aaron Viehl (101 Frankfurt)"
Write-Host "einsnulleins.de"
Write-Host ""

if (-Not $Force) {
    do {
        $YesOrNo = Read-Host "Are you sure you want to remove Office from this PC? (y/n)"
    } while ("y", "n" -notcontains $YesOrNo)

    if ($YesOrNo -eq "n") {
    exit 1
    }
}

Stop-OfficeProcess
Invoke-OfficeUninstall 
Invoke-SetupOffice365 "$Office365Setup_URL/upgrade.xml"
Invoke-Reboot