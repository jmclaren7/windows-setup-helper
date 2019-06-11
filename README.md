# itdeployhelper
This projects purpose is to have a flexible tool that can be integrated into Windows install media to make workstation deployments easier. The parts of the project here are limited to just the scripting/post installation steps but here are the key features of the project as a whole:

* Only interaction during installation is partition setup [Autounattend.xml] 
* Has newest intel network driver and some windows updates integrated [NTLite]
* Copies files from install media to windows folder (\sources\\$OEM$\$$\IT > C:\Windows\IT) [$OEM$]
* Automatic login to administrator account (password 1234 but you wont need it) [Autounattend.xml]
* Runs scripts automatically in correct folders (AutoSystem and AutoLogin)
* Runs autoit script with GUI on first login of administrator to run additional routines and list additional scripts (OptLogin)
* Can be updated from github by clicking update button 
* Scripts and files can be added to USB without making a new ISO (\sources\\$OEM$\$$\IT)
