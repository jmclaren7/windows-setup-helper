# Windows Setup Helper
This project provides a GUI to use along side the traditional Windows installer interface you see on a normal Windows installation ISO/USB. You'll have access to whatever scripts and tools you choose to add along with some great included tools. Automated installs will skip everything but desk selection, once the install completes and Windows boots any scripts you selected will automatically run.

To use Windows Setup Helper you'll need to add the project files to a Windows installer image along with your custom scripts and tools. I've created a script (Build.bat) to help get this done quickly, just follow the instructions below.

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Screenshot1.png?raw=true">
</p>

## Features
- Boot Windows install media and use the custom GUI to run to tools or install Windows
- Select "Normal Install"
  - No modifications, automation or anything is made to the install process or the final Windows installation
- Select "Automated Install"
  - Skip all install steps (except for disk selection)
  - Run any of the selected "logon scripts" after install completes
  - Automatic login to the administrator account after install (Default password is 1234, be sure to disable the administrator account when done)
- Only open source or highly trusted free programs are included
- Tools and scripts are added to the interface from any available drive with folders matching the specific folder path/names
- This project focuses on only modifying the the boot image (boot.wim) and not the Windows image that gets installed, this has multiple benefits.
  - The installer image can be better trusted to be bug free and unmodified
  - Change install.wim with new or customized versions without effecting the customizations you make with this project
  - Have multiple install.wim images you can choose from (future version)
  - Use the boot image with PXE booting and still have all the customizations available
- Starts a VNC server so you can remotely connect from another computer
- Automatically add drivers to WinPE that you provide (useful for storage drivers when the installer can't detect drives)
- Has a basic taskbar for switching between open windows

<br>

<div align="center">
Building, Booting, Installing

https://www.youtube.com/watch?v=NjPxmrIeGhw<br>
[![](https://img.youtube.com/vi/NjPxmrIeGhw/maxresdefault.jpg)](https://www.youtube.com/watch?v=NjPxmrIeGhw)
</div>

## Included 3rd Party Tools
Some of the tools I normally use aren't included here because the licensing doesn't allows it, free closed source software that I've included is noted as such.

- [AutoIt.exe](https://www.autoitscript.com/) (Closed source, required) For Running The Helper Script and It's Components
- [7-Zip](https://www.7-zip.org/) For Working With Compressed Files
- [Putty](https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html) For SSH and Telnet Connections
- [Explorer++](https://github.com/derceg/explorerplusplus) For Browsing Files/Folders
- [Nirsoft Tools](https://www.nirsoft.net/) (Closed source) Including DevManView, FullEventLogView and SearchMyFiles
- [Sysinternal Tools](https://learn.microsoft.com/en-us/sysinternals/) (Closed source) Including Disk2VHD and Autoruns
- [SeaMonkey](https://www.seamonkey-project.org/) Web Browser
- [Crystal DiskInfo & DiskMark](https://github.com/hiyohiyo) For Disk Benchmark and Viewing SMART Data
- [GSmartControl](https://gsmartcontrol.shaduri.dev/) For Viewing SMART Data
- [ReactOS](https://reactos.org/) Paint For Viewing Images
- [TightVNC](https://www.tightvnc.com/) Server For Remote Access To WinPE
- [NTPWEdit](https://github.com/jmclaren7/ntpwedit) Offline Password Reset

## Adding Tools/Scripts/Programs/Drivers
- Add file to the "Tools" folder to list them in the GUI so you can run them later
- Add files to the "PEAutoRun" folder to have them automatically run when WinPE starts
  - A script to install drivers to WinPE is located at PEAutoRun\Drivers, add driver files to this folder
- Add files to the "Logon" folder to make them selectable before install and then executed after install completes
- Folders (Tools, Logon, PEAutoRun) can be used as a prefix for other folder names and they will also be processed:
  - X:\Helper\Logon
  - X:\Helper\LogonCustom
- Folders on different drives matching the path will also be processed, this allows you to add files to a bootable USB directly or a supplementary drive.
  - X:\Helper\Tools
  - F:\Helper\Tools

## Prepare Using DISM (Build.bat)
Using DISM is the recommended way to update Windows images (wim), it's more advanced but can be faster for repeatedly creating the ISO. I've created a script (Build.bat) to help automate the process.

### Prerequisites
- Download a Windows installer ISO (https://www.microsoft.com/software-download/windows11) 
- Download and install Windows ADK and the PE add-on, [Read this](https://github.com/jmclaren7/windows-setup-helper/blob/master/Extra/ADK-Versions.md) for links and information on available versions. 


### Running Build.bat
1. Edit Build.bat to configure various paths  
    - Set "sourceiso" to the path of your downloaded Windows ISO
    - Set "mediapath" to the directory where you want to extract the ISO files
    - Set "outputiso" to the path for the new ISO generated by the script
    - (Optional) "extrafiles" is the directory of files that will be added along with the project files (this is so you can keep your custom files separate from the project files)
2. Run Build.bat with administrator privileges
3. Either select the letter of the individual step you want to run or select F to go through all steps marked with a "*" (toggle if a step is included using the # of that step)

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Build1.png?raw=true">
</p>

## Prepare Using Other Methods
Any method of modifying a WIM image should work by copying the project files to the root of the boot.wim image.

* [NTLite](https://www.ntlite.com/) is one such tool that let's you work with wim images and also lets you create the final ISO. In NTLite, you "load" the "Microsoft Windows Setup (amd64)" image and then copy the project files to %TEMP%\NLTmpMnt (usually).
* [7Zip](https://www.7-zip.org/) will also let you make modifications to a wim file but you'll need to use another tool to create the ISO. If you create a bootable Windows install USB and then modify boot.wim on the USB that could also work since you won't need to create an ISO.

## Create Bootable USB From ISO
1. Download and open Rufus (https://rufus.ie/)
2. Select your target USB device and source ISO image
2. (Optional) Enable the hidden "dual UEFI/BIOS mode" by pressing ALT + E (a message in the bottom left will briefly tell you if you turned it on or off) then set the target system to "BIOS or UEFI" and file system to FAT32
5. Click start

## Other Customizations & Features
I've created a number of features which may not be clearly documented but I've tried to include examples for each of these in the project, as time goes on and if interest in the project increases I will begin to document more of these.

### TightVNC
Work in progress: Use VNC to remote access the WinPE instance from another machine on the network. This can be useful if you had a user at a remote site boot to a USB and then you VNC into it to do recovery, diagnostics or Windows installation.

- A VNC server will start automatically with WinPE (PEAutoRun Folder)
- The port and password are configured in "PEAutoRun\vncserver\settings.ini" (Defaults to port 5950, password "vncwatch")
- When the VNC server is active the main Helper window will have a message in the status bar that says "VNC Running"
- The IP of the machine is shown in the status bar or use your preferred method of locating the machine on your network

### NetBird
NetBird is an overlay/mesh network tool and the NetBird client happens to work well in WinPE, implemented correctly you can have your WinPE boot and be remotely accessible with VNC over the NetBird network automatically. 

- An example startup script is provided (PEAutoRun folder), rename the folder so it does not start with "."
- Configure the script with a setup key, be sure you understand the security implications 
- Add netbird.exe and wintun.dll from https://github.com/netbirdio/netbird/releases (netbird_x.xx.x_windows_amd64_signed.tar.gz)

### Special Files & Folders (Tools/Logon/PEAutoRun)
- Files or folders with a "." at the start of their name are treated as hidden and won't be listed or autorun
- If ".Explorer++.exe" is present in Tools, it's treated as hidden but also causes an "explorer" button to show in the GUI
- ".Options.txt" is a text file that can be used for special treatment of files/folders
    - Listing the name of a file will check it be default
    - Including the text "CheckAll" will check everything in that folder by default
    - Including the text "CollapseTree" will cause that section of the list to be collapsed by default

### Misc
- Windows Pro is used by default for automated installs, you can switch to home using the advanced menu