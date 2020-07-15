# EWS Contact Sync
Utilizes both Exchange Web Services and Office 365 Remote PowerShell Services to sync your Global Address List to any/every user in the directory.

**Why would I want to use this?** iPhone/Android devices don't currently support offline Global Address List synchronization. By loading the Global Address List contacts into a folder within user's mailbox, you can circumvent this limitation.

**Features**
- Fetch a list of contacts using the Office 365 Directory
- Import the list of contacts into a specified user's Office 365 mailbox
- You can run the sync for any number of users
- Specify a custom contact folder
- Uses Application Impersonation so this can all be done from a single admin account (that has Application Impersonation permissions)

## Getting Started

1. Install the EWS API here: https://www.microsoft.com/en-us/download/details.aspx?id=42951
2. Download the latest version of the script here: https://github.com/grahamr975/EWS-Office365-Contact-Sync
3. Export your Office 365 administrator credentials to a CliXml credential file. See below for an example on how to do this.
```
$Credxmlpath = "C:\MyCredentials\"
$Credential | Export-Clixml $Credxmlpath
```
4. To test the script, run for a single mailbox in your directory. See below for an example (batch file)
```
@echo off
cd "%~dp0EWS-Office365-Contact-Sync"

PowerShell.exe -ExecutionPolicy Bypass ^
-File "%CD%\EWSContactSync.ps1" ^
-CredentialPath "C:\Encrypted Credentials\SecureCredential.cred" ^
-FolderName "Directory Contacts" ^
-LogPath "%~dp0Logs" ^
-MailboxList john.doe@mycompany.com ^
-ModernAuth
pause
```
5. Once you're ready, specify DIRECTORY for MailboxList. This will sync the contacts for all users in your directory. See below for an example (batch file)
```
@echo off
cd "%~dp0EWS-Office365-Contact-Sync"

PowerShell.exe -ExecutionPolicy Bypass ^
-File "%CD%\EWSContactSync.ps1" ^
-CredentialPath "C:\Encrypted Credentials\SecureCredential.cred" ^
-FolderName "Directory Contacts" ^
-LogPath "%~dp0Logs" ^
-MailboxList DIRECTORY ^
-ModernAuth
pause
```

### Prerequisites

- EWS API 2.2 https://www.microsoft.com/en-us/download/details.aspx?id=42951
- O365 Global Admin Account with Application Impersonation permissions
- Powershell Version 3.0+
- Think of a unique folder name (Any contacts not in the Global Address List will be deleted from the folder, so don't use 'Contacts' as the name.)

## Deployment

See EWSContactSync.ps1 for documentation on optional parameters for filtering conatcts, mailboxes, etc...

## Built With

* [Powershell 5.0](https://github.com/PowerShell/PowerShell) - The main language used
* [EWS](https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/ews-reference-for-exchange) - API for reading and writing contacts
* [Exchange Online Powershell](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps) - Used to fetch contact and user mailbox data

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Ryan Graham** - *Initial work* - [grahamr975](https://github.com/grahamr975)

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Thanks to gscales for his work on the EWSContacts powershell module. This script uses a modified version of their module. https://github.com/gscales/Powershell-Scripts/tree/master/EWSContacts

