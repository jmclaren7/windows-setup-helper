# itdeployhelper
This project integrates a flexible set of scripts and tools into the Windows install media to make workstation deployments easier while also adding tools and features to the Windows PE enviroment.

## Required External Components
* AutoIt3 executables (64bit required for WinPE)

## Suggested Tools To Add
* (WinPE) TightVNC
* (WinPE) Explorer++
* (WinPE) Password Refixer (Comercial)
* (WinPE) Paragon Hard Disk Manager (Comercial)
* 
* (Post-Install/WinPE) 

## Features
* Only interaction during installation can be inital install type selection and partition setup
* Uses $OEM$ folder on install media which is copied to install drive where scripts are run post-install
* Automatic login to administrator account (default password is 1234, you'll disable the administrator account before you're done)
* Runs scripts automatically in correct folders
* Runs autoit script with GUI on first login of administrator to run additional routines and list additional scripts (OptLogin)
* Can be updated from github by clicking update button 
* Scripts and files can be added to $OEM$ folder directly on existing USB
