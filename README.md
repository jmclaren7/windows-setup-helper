# itdeployhelper
This project integrates a flexible set of scripts and tools into the Windows install media to make workstation deployments easier while also adding tools and features to the Windows PE enviroment.

## Features
* Only interaction during Windows installation will be partition setup unless you select "Normal Install"
* Uses $OEM$ folder on install media which is copied to install drive where scripts are run post-install
* Automatic login to administrator account (default password is 1234, you will need disable the administrator account when you're done)
* Runs scripts automatically
* Runs a GUI on first login to run additional routines and list additional scripts
* Can be updated
* Scripts and tools can be added to an existing USB media

## Required External Components
* AutoIT3 executables (64bit required for WinPE)
* AutoIT3 built-in UDFs (Now included)

## Suggested Tools To Add For The WindowsPE Boot Enviroment
* TightVNC
* Explorer++
* Password Refixer (Comercial)
* Paragon Hard Disk Manager (Comercial)

## Adding Tools/Scripts/Programs
* Once a USB drive is created you can update the available tools by adding them to \sources\$OEM$\$$\IT\
* Folders that start with "Opt" contain items that needs to by manualy launched
* Folders that start with "Auto" will launch automaticly
* Folders with the word "Setup" are used from within the WindowsPE boot enviroment
* Folders with the word "Login" are used from within the Windows install when it first boots

## Preparing Install

## Integrating 3rd Party Tools Into WinPE

## Making The ISO

## USB Setup (From ISO)
1. Download and open Rufus (https://rufus.ie/)
2. Enable the hidden "dual UEFI/BIOS mode" by pressing ALT + E (a message in the bottom left will breifly tell you if you turned it on or off)
3. Select your USB device and ISO image
4. Set the addiotinal options as shown in the image (Target system: "BIOS or UEFI", File system: FAT32)
5. Click start
![image](https://user-images.githubusercontent.com/3019173/130369524-0f8de223-60f7-4bd5-8a38-0bb15c621c5b.png)
