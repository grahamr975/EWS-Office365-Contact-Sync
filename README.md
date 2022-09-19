# EWS Contact Sync
Utilizes both Exchange Web Services and Office 365 Remote PowerShell Services to sync your Global Address List to any/every user in the directory.

**Why would I want to use this?** iPhone/Android devices don't currently support offline Global Address List synchronization. By loading the Global Address List contacts into a folder within user's mailbox, you can circumvent this limitation.

**Features**
- Fetch a list of contacts using the Office 365 Directory
- Import the list of contacts into a specified user's Office 365 mailbox
- You can run the sync for any number of users
- Specify a custom contact folder
- ~~Uses Application Impersonation so this can all be done from a single admin account (that has Application Impersonation permissions)~~  
    Uses a single AzureApp & certificate based authenication (See guide below)

## Modern Certificate-Based Authenication Requirements (OCT 2022 Update)
### Note to Legacy Users: This update requires you to complete the below steps in order to use the script.
1. Install the Exchange Online Powershell V2
```
Install-Module -Name ExchangeOnlineManagement -RequiredVersion 2.0.5
```
2. Create an Azure app & certificate file using [the tutorial here](https://github.com/MicrosoftDocs/office-docs-powershell/blob/main/exchange/docs-conceptual/app-only-auth-powershell-v2.md), taking note of the addendums below.
    * The app will require **Global Reader** permission (Referenced in tutorial).
    * Take a record of the Azure app's **Application (client) ID** as you'll need this later.
    * Enable Public Client Flows in the Azure App (**Authenication** -> **Allow public client flows**)
    * Specify a redirect URI (**Authenication** -> **Platform Configurations** -> **Add a platform** -> **Mobile and desktop applications** -> Enable 'https://login.microsoftonline.com/common/oauth2/nativeclient' as a redirect URI.)
    * When updating the app's Manifest, insert the below code for **requiredResourceAccess** instead of following what the tutorial suggests. The below version also includes permissions for acting as an EWS Application.
        ```
            "requiredResourceAccess": [
            {
                "resourceAppId": "00000002-0000-0ff1-ce00-000000000000",
                "resourceAccess": [
                    {
                        "id": "dc50a0fb-09a3-484d-be87-e023b12c6440",
                        "type": "Role"
                    },
                    {
                        "id": "dc890d15-9560-4a4c-9b7f-a736ec74ec40",
                        "type": "Role"
                    }
                ]
            }
        ]
        ```
3. Export your certificate password to a CliXml SecureString file. See **Create-SecureCertificatePassword.ps1** in the **Getting Started** folder for an example on how to do this.
4. You'll also need your Office 365 organization URL (Ends in .onmicrosoft.com). Do find this, navigate to the **Office 365 Admin Center** -> **Setup** -> **Domains**
### Example updated batch file with certificate authenication
```
@echo off
cd "%~dp0EWS-Office365-Contact-Sync"

PowerShell.exe -ExecutionPolicy Bypass ^
-File "%CD%\EWSContactSync.ps1" ^
-CertificatePath "C:\Users\johndoe\Desktop\automation-cert.pfx" ^
-CertificatePasswordPath "C:\Users\johndoe\Desktop\SecureCertificatePassword.cred" ^
-ClientID "36ee4c6c-0812-40a2-b820-b22ebd02bce3" ^
-FolderName "Directory Contacts" ^
-LogPath "%~dp0Logs" ^
-MailboxList john.doe@mycompany.com ^
-ExchangeOrg "mycompany.onmicrosoft.com" ^
-ModernAuth
pause
```

## Getting Started

1. ~~Install the EWS API here: https://www.microsoft.com/en-us/download/details.aspx?id=42951~~<br/>
    **Now included as a .DLL in the script (Bin folder)**
3. Download the [latest version of the script here.](https://github.com/grahamr975/EWS-Office365-Contact-Sync)
4. ~~Export your Office 365 administrator credentials to a CliXml credential file. See below for an example on how to do this.~~<br/>
    **Replaced by certificate authenication**
5. You may need to unblock the included .dll files. To do this, navigate to EWSContacts\Module\bin -> For each .dll file, right click on the file -> Check 'Unblock'

4. To test the script, run for a single mailbox in your directory. See below for an example (batch file)
```
@echo off
cd "%~dp0EWS-Office365-Contact-Sync"

PowerShell.exe -ExecutionPolicy Bypass ^
-File "%CD%\EWSContactSync.ps1" ^
-CertificatePath "C:\Users\johndoe\Desktop\automation-cert.pfx" ^
-CertificatePasswordPath "C:\Users\johndoe\Desktop\SecureCertificatePassword.cred" ^
-ClientID "36ee4c6c-0812-40a2-b820-b22ebd02bce3" ^
-FolderName "Directory Contacts" ^
-LogPath "%~dp0Logs" ^
-MailboxList john.doe@mycompany.com ^
-ExchangeOrg "mycompany.onmicrosoft.com" ^
-ModernAuth
pause
```
5. Once you're ready, specify DIRECTORY for MailboxList. This will sync the contacts for all users in your directory. See below for an example (batch file)
```
@echo off
cd "%~dp0EWS-Office365-Contact-Sync"

PowerShell.exe -ExecutionPolicy Bypass ^
-File "%CD%\EWSContactSync.ps1" ^
-CertificatePath "C:\Users\johndoe\Desktop\automation-cert.pfx" ^
-CertificatePasswordPath "C:\Users\johndoe\Desktop\SecureCertificatePassword.cred" ^
-ClientID "36ee4c6c-0812-40a2-b820-b22ebd02bce3" ^
-FolderName "Directory Contacts" ^
-LogPath "%~dp0Logs" ^
-MailboxList DIRECTORY ^
-ExchangeOrg "mycompany.onmicrosoft.com" ^
-ModernAuth
pause
```

### Prerequisites

- ~~EWS API 2.2 https://www.microsoft.com/en-us/download/details.aspx?id=42951~~

    - Now included as a .DLL in the script (Bin folder)
- ~~O365 Global Admin Account with **Application Impersonation permissions** (You MUST set this seperately)~~
    
    - Requirement replaced by Azure app (See above guide on how to set this up.)
- Verify the neccessary Office 365 URLs are whitelisted in your environment. [All Microsoft 365 Common URLs with ID#56 on this page should be allowed.](https://docs.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide)
- Powershell Version 5.0+
- Think of a unique folder name (Any contacts not in the Global Address List will be deleted from the folder, so I don't recommend using 'Contacts' as the name.)

## Deployment

See **EWSContactSync.ps1** for documentation on optional parameters for filtering conatcts, mailboxes, etc...

## Built With

* [Powershell 5.0](https://github.com/PowerShell/PowerShell) - The main language used
* [EWS](https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/ews-reference-for-exchange) - API for reading and writing contacts
* [ExchangeOnline Powershell](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/2.0.5) - Used to fetch contact and user mailbox data

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Ryan Graham** - *Initial work* - [grahamr975](https://github.com/grahamr975)
* **Glenn Scales** - *EWSContacts Powershell Module* - [gscales](https://github.com/gscales)

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Thanks to gscales for his work on the EWSContacts powershell module. This script uses a modified version of their module. https://github.com/gscales/Powershell-Scripts/tree/master/EWSContacts
* Thanks to [alexisc182](https://github.com/alexisc182) for their work on documenting the needed Office 365 URLs for whitelisting
