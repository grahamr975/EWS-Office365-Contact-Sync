EWS Contact Sync
===================
Utilizes both Exchange Web Services and Office 365 Remote PowerShell Services to sync Global Address List to any user in the directory.

Why would I want to use this?
iPhones in particular don't sync Office 365 contacts from the Global Address List. 

Thanks to gscales for his work on the EWSContacts powershell module. This script uses a modified version of his module.
https://github.com/gscales/Powershell-Scripts/tree/master/EWSContacts

Features
--------
- Automatically generates a list of contacts using the Office 365 Directory
- You can run the sync for any number of users
- Specify a custom contact folder that the script will create

Prerequisites
------------
-EWS API 2.2 https://www.microsoft.com/en-us/download/details.aspx?id=42951
-O365 Global Admin Account with Application Impersonation Permissions
-Powershell Version 3.0+

Installation
------------

1. Extract all resources to your local computer
2. Create a secure credential using the included Create-SecureCredential.ps1 script in the Tools folder
3. Make a .bat file calling Multi-Import.ps1 using the requried parameters: CredentialPath, FolderName, and  LogFile
4. Run the file with administrator rights

License
-------

The project is licensed under the MIT license.