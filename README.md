O365 Contact Import
===================
O365 Contact Import utilizes both Exchange Web Services and Office 365 Remote PowerShell Services to generate and import a list of contacts for every user in the organization.
The most common use for this script would be to sync your organization�s Global Access List with the phones of your users.



Features
--------
-Pseudo Multi-Threading Using Powershell Jobs Reduces Run-time
-Use a CSV for the Contact List or Fetch Directly From your Organisation�s Global Access List
-Connects Directly to Office 365, Does not Require any Connections Other than Internet Access

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