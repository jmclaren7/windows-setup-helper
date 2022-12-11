# Work in Progress
This project integrates a flexible set of scripts and tools into the Windows install media to make deployments easier while also adding tools and features to the Windows PE environment.

## Features
* Only interaction during Windows installation will be partition setup unless you select "Normal Install"
* All files are integrated into the boot.wim image, this means you can pxe boot this image and use the same tools
* Automatic login to administrator account (default password is 1234, you will need to disable the administrator account when you're done)
* Automatically Runs scripts that were selected before the setup process was started

## Required External Components
* AutoIT3.exe executable (64bit version)

## Suggested Tools To Add For The WindowsPE Boot Environment
64 bit programs are required for use with the normal 64 bit WindowsPE used with 64 bit installers provided by Microsoft
* TightVNC
* Explorer++
* Password Refixer (Commercial)
* Macrium Rescue (Commercial)
* Sysinternals & Nirsoft Tools

## Adding Tools/Scripts/Programs
* Files in the "Tools" folder are used from within the WindowsPE boot environment
* Files in the "Logon" folder are selectable before install and then executed after install completes

## Preparing Using NTLite
* Extract a Windows ISO to a folder and open that folder as a source in NTLite
* Load the install image desired (Pro) and customize it
	* Add updates
	* Add drivers (Intel chipset drivers hint: setupchipset.exe -extract "c:\temp")
	* Add "AnswersReg.reg" for answers file to work with WDS
	* Trim editions
	* Apply changes
* Load the boot image
	* Customize and run .update-image.bat to copy files to the WinPE image
	* Apply image
* Create iso from source

## USB Setup (From ISO)
1. Download and open Rufus (https://rufus.ie/)
2. Enable the hidden "dual UEFI/BIOS mode" by pressing ALT + E (a message in the bottom left will briefly tell you if you turned it on or off)
3. Select your USB device and ISO image
4. Set the additional options as shown in the image (Target system: "BIOS or UEFI", File system: FAT32)
5. Click start

![image](https://user-images.githubusercontent.com/3019173/130369524-0f8de223-60f7-4bd5-8a38-0bb15c621c5b.png)
