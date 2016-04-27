# Office 365 (2016) Enterprise Installation Script
Full Deployment of Office 365 Script for Enterprise

This Script is fully for implementing Office 365 into an Enterprise Environment. This will not work unless you edit it for your own environment. This comes AS-IS and I provide no support. I am only posting this here in hopes it doesn't take people months to make this installation work as it did for me. 

Included Steps:
============
1. Clean Office off the computer aside from MSProject, Visio, and other singlar custom installations. This removes all Standard installs with 'Office 20** Standard'. If Access or other products are included it will remove them as well. 
2. Installs the custom 'configuration.xml' file. The included one only excludes OneDrive, Access, Project, and Visio.
3. Modify the Office Install to remove shortcuts we do not want the Users to see. We also implemented a GPO for App Locker to block the Local Outlook and Skype Clients from launching since they are not fully operational (wink, wink) and are outside the scope of our current implementation. I left them as part of the install so that we can easily deploy them later by removing that policy and adding the shortcut back to the Start-Menu.
4. Uninstalls OneDrive completely from the computer. There is no solid option to turn OneDrive installation off. This will uninstall it if it is still existing.
5. Copies Outlook Web App shortcut to the Desktop if it doesn't exist. Custom folder included. Then it sweeps the custom folders with a custom batch file to delete files in those folders. Depending on the length of the folder and number of spaces it uses you will need to change the Batch file.
6. In our enterprise we have a custom 'First time log on' Script which pins icons to the TaskBar. This script relaces one that is already there with one that is updated for Office 2016
7. Removes GroupWise shortcuts (Old Mail App)
8. Replaces shortcuts that were removed during step 1. Also creates a directory and copies the shortcuts into it since previous had that feature and this ap does not.


Suggested tools to download and implement:
============
1. Windows 10 Admin Templates - https://www.microsoft.com/en-us/download/details.aspx?id=48257
2. Office 365 Admin Templates - https://www.microsoft.com/en-us/download/details.aspx?id=49030
3. Office 365 Deployment Toolkit - https://www.microsoft.com/en-us/download/details.aspx?id=49117
