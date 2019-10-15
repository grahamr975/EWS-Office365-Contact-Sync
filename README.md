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
- Imports the list of contacts into a specified user's Office 365 mailbox
- You can run the sync for any number of users
- Specify a custom contact folder that will automatically be created if it dosen't exist
- Uses Application Impersonation so this can all be done from a single admin account (that has Application Impersonation permissions)

Prerequisites
------------
- EWS API 2.2 https://www.microsoft.com/en-us/download/details.aspx?id=42951
- O365 Global Admin Account with Application Impersonation permissions
- Powershell Version 3.0+

Installation
------------
1. Validate that all prerequisites are met
2. Extract all resources to your local computer
3. Export your credentials to a CliXml credential file
4. Run the script (I'd suggest you start by testing it on a single mailbox!)

License
-------

The project is licensed under the MIT license.