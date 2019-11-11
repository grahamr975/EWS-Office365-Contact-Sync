EWS Contact Sync
===================
Utilizes both Exchange Web Services and Office 365 Remote PowerShell Services to sync your Global Address List to any/every user in the directory.

**Why would I want to use this?** iPhone/Android devices don't currently support offline Global Address List synchronization. By loading the Global Address List contacts into a folder within user's mailbox, you can circumvent this limitation.

Features
--------
- Fetch a list of contacts using the Office 365 Directory
- Import the list of contacts into a specified user's Office 365 mailbox
- You can run the sync for any number of users
- Specify a custom contact folder
- Uses Application Impersonation so this can all be done from a single admin account (that has Application Impersonation permissions)

Prerequisites
------------
- EWS API 2.2 https://www.microsoft.com/en-us/download/details.aspx?id=42951
- O365 Global Admin Account with Application Impersonation permissions
- Powershell Version 3.0+
- Think of a unique folder name (Any contacts not in the Global Address List will be deleted from the folder, so don't use 'Contacts' as the name.)

Installation
------------
1. Download the latest version here: https://github.com/grahamr975/EWS-Office365-Contact-Sync
2. Extract all resouces to a folder on your computer
3. Export your credentials to a CliXml credential file
4. Run it

Credits
-------
Thanks to gscales for his work on the EWSContacts powershell module. This script uses a modified version of their module.
https://github.com/gscales/Powershell-Scripts/tree/master/EWSContacts

License
-------
The project is licensed under the MIT license.
