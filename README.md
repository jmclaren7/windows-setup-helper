# Work in Progress
This project has primarily been for my own use and I can't make any promises that the instructions and code available here is complete or reliable but I'll do my best to make it work for anyone who finds this project. Most of the tools I normally add aren't included here because they are either commercial products or I'm uncertain if the open source licensing allows it, I *have* included AutoIT3.exe (64 bit) which is required.

# Windows Setup Helper
This project provides an interface to replace the Windows Installer on a normal Windows install media. The new interface gives you access to whatever scripts and tools you choose to copy to the into the Windows install media. The interface provides options to run a normal Windows Install or one that is automated. Automated installs will use an Autounattend.xml file to skip all but partitioning steps, once the install completes any scripts you selected will automatically run.

The objective is to leave the Windows install image (install.wim) completely unmodified and the WinPE image (boot.wim) nearly unmodified so that it can be replaced or updated easily and so the option of using an unmodified install.wim is always available while still having your tools, customizations and automatons available. This flexibility can be most useful when you routinely work in different environments with very different requirements.

To use Windows Setup Helper you'll need to add the project files to a WinPE image (sources\boot.wim) along with your custom scripts and tools. It's recommended to us the "Prepare Using NTLite" instructions below.

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Screenshot1.png?raw=true">
</p>

## Features
* Boot Windows install media and use the custom GUI to run to tools or install Windows
* Select "Automated Install" or "Normal Install"
  * Normal install will run the Windows installation without any modifications or automation
  * Automated install will skip all install steps except for partitioning and run any selected scripts after install completes
* Automatic login to the administrator account after install (Default password is 1234, be sure to disable the administrator account when done)
* This program is integrated into the boot.wim image, this means you can pxe boot this image and use the same tools
* Tools and scripts are listed from folders in the main program's folder or specific folders from any available drive

## Included 3rd Party Tools For WinPE
* TightVNC (**WiP:** starts a VNC server when WinPE starts)
* 7-Zip
* Explorer++

## Other Useful 3rd Party Tools For WinPE That Have Been Tested  
Note: Full 64 bit is required in a 64 bit WindowsPE Environment
* Password Refixer (Commercial) - Copy the program files from their WinPE image
* Macrium Reflect (Commercial) - Copy the program files from their WinPE image
* Crystal DiskMark & DiskInfo
* Various Sysinternals & Nirsoft Tools

## Adding Tools/Scripts/Programs
* Files in the "Tools" folder will be shown for use in the WindowsPE boot environment
* Files in the "PEAutoRun" folder will automatically run when WindowsPE boots
* Files in the "Logon" folder are selectable before install and then executed after install completes
* All the above folders can be used as a prefix for other folder names and they will be treated the same as described
	* LogonACME
	* ToolsJohn
* Folders on different drives matching the naming convention will also be listed/processed, this allows you to add files to a bootable USB after the fact
	* USB:\Helper\Tools
	* USB:\Helper\ToolsAV
	* USB:\Helper\LogonMisc

## Prepare Using NTLite
Using NTLite can be a convenient way to modify Windows install media and then create an ISO
1. Extract a Windows ISO to a folder and open that folder as a source in NTLite
2. Right click on and load (mount) the boot.wim image "Microsoft Windows Setup (amd64)"
3. Copy the "Helper" and "Windows" folders to the mount directory (could be either %TEMP%\NLTmpMnt or %TEMP%\NLTmpMnt01)
	* You can modify Update-Image.bat to help copy files to the mount directory
4. Apply image
	* The Apply options allow you to "remove nonessential editions", removing all but your preferred image is recommended (Windows 11 Pro)
5. Create ISO

## Prepare Using DSIM
Using DSIM is more advanced and doesn't offer a way to create a bootable ISO when done, these instructions are incomplete and untested
1. Extract a Windows ISO to a folder
2. Mount sources\boot.wim
	* DISM /Mount-image /imagefile:"<path>\sources\boot.wim" /Index:1 /MountDir:%TEMP%\WIM-Mount /optimize
3. Copy the "Helper" and "Windows" folders to the mount directory (%TEMP%\WIM-Mount)
	* You can modify Update-Image.bat to help copy files to the mount directory
4. Commit changes and unmount
	* DISM /Unmount-Image /MountDir:%TEMP%\WIM-Mount /commit
	* Or unmount without committing changes: DISM /Unmount-Image /MountDir:%TEMP%\WIM-Mount /discard
5. Use your preferred tools to created a bootable USB or ISO
	* Makewinpemedia should work and is available in the Windows ADK

## Create Bootable USB From ISO
1. Download and open Rufus (https://rufus.ie/)
2. (Optional) Enable the hidden "dual UEFI/BIOS mode" by pressing ALT + E (a message in the bottom left will briefly tell you if you turned it on or off)
3. Select your USB device and ISO image
4. (Optional) Set the additional options as shown in the image (Target system: "BIOS or UEFI", File system: FAT32)
5. Click start

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/windows-setup-helper/master/Extra/Rufus1.png?raw=true">
</p>