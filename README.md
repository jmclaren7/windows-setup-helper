# Windows Setup Helper

This project provides a GUI to use along side the traditional Windows installer interface you see on a normal Windows installation ISO/USB. You'll have access to whatever scripts and tools you choose to add along with options to start an automated Windows installation. Automated installs will use an Autounattend.xml file to skip all but partitioning steps, once the install completes any scripts you selected will automatically run.

To use Windows Setup Helper you'll need to add the project files to a Windows installer image along with your custom scripts and tools. I've created a script (Build.bat) to help get this done quickly, just follow the instructions below.

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Screenshot1.png?raw=true">
</p>

## Features

- Boot Windows install media and use the custom GUI to run to tools or install Windows
- Select "Automated Install" or "Normal Install"
  - Normal install will run the Windows installation without any modifications or automation
  - Automated install will skip all install steps (except for partitioning) and run any of the selected logon scripts after install completes
- Automatic login to the administrator account after install (Default password is 1234, be sure to disable the administrator account when done)
- Only open source or highly trusted free programs are included
- Tools and scripts are added to the interface from any available drive with folders matching the specific folder path/names
- This project focuses on only modifying the the boot image (boot.wim) and not the Windows image that gets installed, this has multiple benefits.
  - The installer image can be better trusted to be bug free and unmodified
  - Change install.wim with new or customized versions without effecting the customizations you make with this project
  - Have multiple install.wim images you can choose from (future version)
  - Use the boot image with PXE booting and still have all the customizations available
- Starts a VNC server so you can remotely connect from another computer
- Automatically add drivers that you add (useful for storage drivers when the installer can't detect drives)
- Has a basic taskbar for switching between open windows

<br>

<div align="center">
Building, Booting, Installing

https://www.youtube.com/watch?v=NjPxmrIeGhw<br>
[![](https://img.youtube.com/vi/NjPxmrIeGhw/maxresdefault.jpg)](https://www.youtube.com/watch?v=NjPxmrIeGhw)
</div>

## Included 3rd Party Tools

Some of the tools I normally use aren't included here because the licensing doesn't allows it, free closed source software that I've included is noted as such.

- [AutoIt.exe](https://www.autoitscript.com/) (Closed source, required)
- [7-Zip](https://www.7-zip.org/)
- [Putty](https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html)
- [Explorer++](https://github.com/derceg/explorerplusplus)
- Some [Nirsoft Tools](https://www.nirsoft.net/) (Closed source)
- Some [Sysinternal Tools](https://learn.microsoft.com/en-us/sysinternals/) (Closed source)
- [SeaMonkey](https://www.seamonkey-project.org/) Browser
- [Crystal DiskInfo & DiskMark](https://github.com/hiyohiyo)
- [GSmartControl](https://gsmartcontrol.shaduri.dev/)
- [ReactOS](https://reactos.org/) Components
- [SeaMonkey Web Browser](https://www.seamonkey-project.org/)
- [TightVNC](https://www.tightvnc.com/)

## Adding Tools/Scripts/Programs/Drivers

- Add file to the "Tools" folder to list them in the GUI so you can run them later
- Add files to the "PEAutoRun" folder to have them automatically run when WinPE starts
  - Add driver files under PEAutoRun\Drivers
- Add files to the "Logon" folder to make them selectable before install and then executed after install completes
- Folders (Tools, Logon, PEAutoRun) can be used as a prefix for other folder names and they will also be processed:
  - \Helper\Logon
  - \Helper\LogonCustom
- Folders on different drives matching the path will also be processed, this allows you to add files to a bootable USB after the fact, a supplementary drive or even a network drive
  - X:\Helper\Tools
  - F:\Helper\Tools

## Other Customizations & Features

I've created a number of features which may not be clearly documented but I've tried to include examples for each of these in the project, as time goes on and if interest in the project increases I will begin to document more of these.

## Prepare Using DSIM (Build.bat)

Using DSIM is the recommended way to update Windows images (wim), it's more advanced but can be faster for repeatedly creating the ISO. I've created a script (Build.bat) to help automate the process.

### Prerequisites

- Download and extract a Windows installer ISO (https://www.microsoft.com/software-download/windows11) 
- Download and install Windows ADK and the PE add-on, [Read this](https://github.com/jmclaren7/windows-setup-helper/blob/master/Extra/ADK-Versions.md) for links and information on available versions. 
- Configure Build.bat with the correct path to the directory containing the extracted files, choose an ISO output path and review the additional options noted in the script

### Running the script

1. Run Build.bat with administrator privileges
2. Either select the individual step you want or select F to go through the entire process

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Build1.png?raw=true">
</p>

## Prepare Using NTLite

[NTLite](https://www.ntlite.com/) lets you modify Windows install media and create an ISO. I no longer use this software but will leave these instructions to get you started.

1. Extract a Windows ISO to a folder and open that folder as a source in NTLite
2. Right click on and load (mount) the boot.wim image "Microsoft Windows Setup (amd64)"
3. Copy the "Helper" and "Windows" folders to the mount directory (could be either %TEMP%\NLTmpMnt or %TEMP%\NLTmpMnt01)
   - You can modify Build.bat to help copy files to the mount directory
4. Apply image
   - The Apply options allow you to "remove nonessential editions", removing all but your preferred image is recommended (Windows 11 Pro)
5. Create ISO

## Create Bootable USB From ISO

1. Download and open Rufus (https://rufus.ie/)
2. (Optional) Enable the hidden "dual UEFI/BIOS mode" by pressing ALT + E (a message in the bottom left will briefly tell you if you turned it on or off)
3. Select your USB device and ISO image
4. (Optional) Set the additional options as shown in the image (Target system: "BIOS or UEFI", File system: FAT32)
5. Click start

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Rufus1.png?raw=true">
</p>

## TightVNC
This feature is a work in progress, the idea is that if you have remote access to another machine on a network, you could have someone boot to a USB and then you can VNC into it to do recovery, diagnostics or Windows installation.

- A VNC server will start automatically with WinPE (PEAutoRun Folder)
- The port and password are configured in "PEAutoRun\vncserver\settings.ini" (Defaults to port 5950, password "vncwatch")
- When the VNC server is active the main Helper window will have a message in the status bar that says "VNC Running"
- The IP of the machine is shown in the status bar or use your preferred method of locating the machine on your network
- Incomplete: "vncviewer-helper" will scan the network for you and let you quickly connect