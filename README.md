# Work in Progress
This project has primarily been for my own use with my clients, I can't make any promises that the instructions and code available here is complete or reliable but I'll do my best to make it work for anyone who find this project. Most of the tools I normally add aren't included here because they are either commercial products or I'm uncertain if the open source licensing allows it.

# Windows Setup Helper
This project integrates a flexible set of scripts and tools into the Windows install media to make Windows installations easier while also adding tools and features to the Windows PE environment. 

The objective is to leave the Windows install image (install.wim) unmodified so that it can be replaced or updated easily and so the option of using an unmodified install.wim is always available while still having your tools, customizations and automations available. The flexibility of have Windows installation media with these options can be most useful when you are routinely working in different environments with very different requirements.

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/restic-simple-backup/main/Extra/Screenshot1.png?raw=true">
</p>

## Features
* Boot Windows install media and use the custom GUI to run to tools or install Windows
* Select "Automated Install" or "Normal Install"
  * Normal install will run the Windows installation without any modifications or automation
  * Automated install will skip all install steps except for partitioning and run any selected scripts after install completes
* Automatic login to the administrator account after install (Default password is 1234, be sure to disable the administrator account when done)
* This program is integrated into the boot.wim image, this means you can pxe boot this image and use the same tools
* Tools and scripts are listed from folders in the main program's folder or specific folders from any available drive

## Required External Components
* AutoIT3.exe executable (64bit version)

## Suggested 3rd Party Tools For The WindowsPE Boot Environment
Note that 64 bit versions of programs are required in a 64 bit WindowsPE Environment
* TightVNC
* Explorer++
* Password Refixer (Commercial)
* Macrium Rescue (Commercial)
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
Using NTLite can be convenient way to modify Windows install media and then create an ISO
1. Extract a Windows ISO to a folder and open that folder as a source in NTLite
2. Right click on and load (mount) the boot.wim image "Microsoft Windows Setup (amd64)"
3. Copy the "Helper" and "Windows" folders to the mount directory (%TEMP%\NLTmpMnt)
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
5. Use your preferred tools to created a bootable USB or ISO

## Create Bootable USB From ISO
1. Download and open Rufus (https://rufus.ie/)
2. (Optional) Enable the hidden "dual UEFI/BIOS mode" by pressing ALT + E (a message in the bottom left will briefly tell you if you turned it on or off)
3. Select your USB device and ISO image
4. (Optional) Set the additional options as shown in the image (Target system: "BIOS or UEFI", File system: FAT32)
5. Click start

<p align="center">
  <img src="https://raw.githubusercontent.com/jmclaren7/restic-simple-backup/main/Extra/Rufus1.png?raw=true">
</p>