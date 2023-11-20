**Work in Progress:** This project has mostly been for my own use and I can't make any promises that the instructions and code available here is complete or reliable but I'll do my best to make it work for anyone who finds this project.

# Windows Setup Helper

This project provides an interface to replace the Windows Installer on normal Windows install media. The new interface gives you access to whatever scripts and tools you choose to add. The interface provides options to start a normal Windows Install or one that is automated. Automated installs will use an Autounattend.xml file to skip all but partitioning steps, once the install completes any scripts you selected will automatically run.

The objective is to leave the Windows install image (install.wim) completely unmodified and the WinPE image (boot.wim) nearly unmodified so that it can be replaced or updated easily while still having your tools, customizations and automatons available. This flexibility can be most useful when you routinely work in different environments with very different requirements.

While many other solutions for custom WindowsPE exist, very few integrate the Windows installer, all of them use tools and GUIs that are closed source and difficult to trust.

To use Windows Setup Helper you'll need to add the project files to a WinPE image (sources\boot.wim) along with your custom scripts and tools. It's recommended to us the "Prepare Using NTLite" instructions below.

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Screenshot1.png?raw=true">
</p>

## Features

- Boot Windows install media and use the custom GUI to run to tools or install Windows
- Select "Automated Install" or "Normal Install"
  - Normal install will run the Windows installation without any modifications or automation
  - Automated install will skip all install steps except for partitioning and run any selected scripts after install completes
- Automatic login to the administrator account after install (Default password is 1234, be sure to disable the administrator account when done)
- This program is integrated into the boot.wim image, this means you can pxe boot this image and use the same tools
- Tools and scripts are listed from folders in the main program's folder or specific folders from any available drive

## Included 3rd Party Tools

Many of the tools I normally add aren't included here because I'm uncertain if the licensing allows it, I _have_ included AutoIT3.exe and a couple other tools.

- TightVNC (more details below)
- 7-Zip
- Explorer++
- Nirsoft's SearchMyFiles, DevManView, FullEventLogView
- Sysinternal's disk2vhd64

## Other Useful Tools

Note: Full 64 bit is required in a 64 bit WindowsPE Environment

- Password Refixer (Commercial) - Copy the program files from their WinPE image
- Macrium Reflect (Commercial) - Copy the program files from their WinPE image
- Crystal DiskMark & DiskInfo
- Various Sysinternals & Nirsoft Tools

## TightVNC

- A VNC server will start automatically with WinPE (PEAutoRun Folder)
- The port and password are configured in "PEAutoRun\vncserver\settings.ini" (Defaults to port 5950, password "vncwatch")
- When the VNC server is active the main Helper window will have a message in the status bar that says "VNC Running"
- The IP of the machine is shown in the status bar or use your preferred method of locating the machine on your network
- Incomplete: "vncviewer-helper" will scan the network for you and let you quickly connect

## Adding Tools/Scripts/Programs

- Files in the "Tools" folder will be shown for use in the WindowsPE boot environment
- Files in the "PEAutoRun" folder will automatically run when WindowsPE boots
- Files in the "Logon" folder are selectable before install and then executed after install completes
- All the above folders can be used as a prefix for other folder names and they will be treated the same as described
  - LogonACME
  - ToolsJohn
- Folders on different drives matching the naming convention will also be listed/processed, this allows you to add files to a bootable USB after the fact
  - USB:\Helper\Tools
  - USB:\Helper\ToolsAV
  - USB:\Helper\LogonMisc

## Prepare Using NTLite

Using NTLite can be a convenient way to modify Windows install media and then create an ISO

1. Extract a Windows ISO to a folder and open that folder as a source in NTLite
2. Right click on and load (mount) the boot.wim image "Microsoft Windows Setup (amd64)"
3. Copy the "Helper" and "Windows" folders to the mount directory (could be either %TEMP%\NLTmpMnt or %TEMP%\NLTmpMnt01)
   - You can modify Update-Image.bat to help copy files to the mount directory
4. Apply image
   - The Apply options allow you to "remove nonessential editions", removing all but your preferred image is recommended (Windows 11 Pro)
5. Create ISO

## Prepare Using DSIM

Using DSIM is more advanced but can be faster for repeatedly creating the ISO by using the I've created (Update-Image.bat)

### Prerequisites

- Download and extract a Windows installer ISO (https://www.microsoft.com/software-download/windows11)
- Download and install Windows ADK and the PE add-on (https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install)
- Configure Update-Image.bat with the correct directory path to the extracted files and choose a ISO output path

### Running the script

1. Run Update-Image.bat with administrator privileges
2. Either select the individual step you want or select F to go through the entire process

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/update-image1.png?raw=true">
</p>

## Create Bootable USB From ISO

1. Download and open Rufus (https://rufus.ie/)
2. (Optional) Enable the hidden "dual UEFI/BIOS mode" by pressing ALT + E (a message in the bottom left will briefly tell you if you turned it on or off)
3. Select your USB device and ISO image
4. (Optional) Set the additional options as shown in the image (Target system: "BIOS or UEFI", File system: FAT32)
5. Click start

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Rufus1.png?raw=true">
</p>
