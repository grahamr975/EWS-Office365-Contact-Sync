<#
.SYNOPSIS
Import a Global Access List into Office 365 using Exchange Web Services

.DESCRIPTION
Uses the supplied administrator account along with Exchange Web Services to export a O365 directory's Global Address List, then imports the Global Address List into the contacts for all Mailboxes.
The main purpose for why one would want to do this, is to sync the GAL with one's iPhone/Android contacts.

Note: The administrator credentials MUST be stored as a secure credential using Export-Clixml

.PARAMETER MailboxList
Specify an email address list that recieves the contacts. Set this value to 'DIRECTORY' to specify every mailbox in your directory.

Example: $MailboxList = @("myemail1@domain.com", "myemail2@domain.com", "myemail3@domain.com")

.PARAMETER CredentialPath
Specifies the path of the Office 365 Administrator Credentials

.PARAMETER FolderName
Name of the folder that you want to import the contact into, if the folder does not exist a new one is created.
NOTE: To prevent duplicates, the specified folder is wiped before importing. For this reason, you should create a dedicated folder for the GAL.

.PARAMETER LogPath
Optional, Specifies the path of where the Log files are stored, along with the naming pattern of the log files.

.EXAMPLE

Command Prompt

C:\> PowerShell.exe -ExecutionPolicy Bypass ^
-File "%CD%\Multi-Import.ps1" ^
-CredentialPath "%CD%\Tools\SecureCredential.cred" ^
-FolderName "My Contact Folder" ^
-LogPath "%CD%\Logs\%mydate%_%mytime%.log" ^
-MailboxList "testemail@mycompany.com"

.LINK

https://www.microsoft.com/en-us/download/details.aspx?id=42951

#>

Param (
    [Parameter(Mandatory=$True)]
	[System.IO.FileInfo]$CredentialPath,
	[String]$FolderName,
	[String[]]$MailboxList,
	[System.IO.FileInfo]$LogPath
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

# Dot Source required Function Libraries
.".\Functions\library.ps1"

# Import Exchange Contacts module
Import-Module .\EWSContacts\Module\ExchangeContacts.psm1 -Force

# Import Office 365 Administrator credentials
$Credential = Import-CliXml -Path $CredentialPath

#-----------------------------------------------------------[Fetch Contact Lists]------------------------------------------------------------

# Fetch list of Global Address List contacts using Office 365 Powershell
try {
    $GALContacts = Get-GALContacts -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credentials $Credential
} catch {
    Write-Log -Level "FATAL" -Message "Failed to fetch Office 365 Global Address List contacts" -logfile $LogPath
}

# If 'DIRECTORY' is used for $MailboxList, fetch all Mailboxes from the administrator account's Office 365 directory
if ($MailboxList -eq "DIRECTORY") {
    try {
        # TO DO: ADD MAILBOX FETCH FEATURE (IT WILL REPLACE THE BELOW LINE)
        $MailboxList = $null
    } catch {
        Write-Log -Level "FATAL" -Message "Failed to fetch Office 365 mailboxes" -logfile $LogPath
    }
    
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

foreach ($Mailbox in $MailboxList) {
    Write-Log -Message "Beginning contact sync for $($Mailbox)'s mailbox" -logfile $LogPath

    # Check if a contacts folder exists with $FolderName. If not, create it.
    try {
        New-EXCContactFolder -MailboxName $Mailbox -FolderName "$FolderName" -Credential $Credential
    } catch {
        Write-Log -Level "FATAL" -Message "Failed verify that $($FolderName) exists for $($Mailbox)" -logfile $LogPath
    }

    # Fetch contacts from the user's mailbox
    $MailboxContacts = Get-EXCContacts -MailboxName $Mailbox -Credentials $Credential -Folder "Contacts\$FolderName" | Where-Object {$_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address -ne $null}
    $MailboxEmailList = Get-EmailAddressFromContact -Contact $MailboxContacts

    # Determine if each contact needs to be deleted, created, or updated
    $MailboxContactsToBeDeleted = $MailboxContacts | Where-Object {!$GALContacts.WindowsEmailAddress.ToLower().Contains($_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower())}
    $MailboxContactsToBeCreated = $GALContacts | Where-Object {!$MailboxEmailList.Contains($_.WindowsEmailAddress)}
    $MailboxContactsToBeUpdated = $GALContacts | Where-Object {$MailboxEmailList.Contains($_.WindowsEmailAddress)}

    # ------[DELETE]------
    # Delete any obsolete contacts (No longer found in the Global Address List)
    # NOTE: This cannot yet remove contacts with no email address!
    try {
        foreach ($MailboxContactToDelete in $MailboxContactsToBeDeleted) {
            Write-Verbose "Deleting Contact: $($MailboxContactToDelete.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower())"
            $MailboxContactToDelete.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
        }
        # Write-Log -Message "Removed all obsolete contacts from $($Mailbox)'s mailbox" -logfile $LogPath
    } catch {
        Write-Log -Level "ERROR" -Message "Failed to remove all obsolete contacts from $($Mailbox)'s mailbox"-logfile $LogPath -exception $_.Exception.Message
    }
    #   ------[UPDATE]------
    foreach ($GALContact in $MailboxContactsToBeUpdated) {
        # Search for a matching contact. If it already exists in the user's mailbox, don't made any changes to it since they aren't needed.
        if ($null -eq $($MailboxContacts | Where-Object {(($_.GivenName -eq $GALContact.FirstName) -or ("" -eq $GALContact.FirstName)) -and (($_.Surname -eq $GALContact.LastName) -or ("" -eq $GALContact.LastName)) -and ($_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].address -eq $GALContact.WindowsEmailAddress) -and ($GALContact.Company -eq $_.CompanyName -or $GALContact.Company -eq "") -and ($GALContact.Department -eq $_.Department -or $GALContact.Department -eq "") -and (($_.DisplayName -eq $GALContact.DisplayName) -or ($GALContact.DisplayName -eq "")) -and ($GALContact.Title -eq $_.JobTitle -or $GALContact.Title -eq "") -and ($GALContact.Phone -eq $_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] -or $GALContact.Phone -eq "") -and ($GALContact.MobilePhone -eq $_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] -or $GALContact.MobilePhone -eq "")})) {
            if ($null -ne $GALContact.WindowsEmailAddress) {
                Write-Verbose $GALContact
                Write-Verbose "Updating Contact: $($GALContact.WindowsEmailAddress)"
                try {   # Try to update the contact if it already exists
                    Set-EXCContact -MailboxName $Mailbox -DisplayName $GALContact.DisplayName -FirstName $GALContact.FirstName -LastName $GALContact.LastName -EmailAddress $GALContact.WindowsEmailAddress -CompanyName $GALContact.Company -Credentials $Credential -Department $GALContact.Department -BusinssPhone $GALContact.Phone -MobilePhone $GALContact.MobilePhone -JobTitle $GALContact.Title -Folder "Contacts\$FolderName" -useImpersonation -force
                } catch {
                    Write-Log -Level "ERROR" -Message "Failed to sync $($GALContact.WindowsEmailAddress) contact to $($Mailbox)'s mailbox" -logfile $LogPath -exception $_.Exception.Message
                }
            }
        }
    }
    #   ------[CREATE]------
    foreach ($GALContact in $MailboxContactsToBeCreated) {
        if ($null -ne $GALContact.WindowsEmailAddress) {
            Write-Verbose "Adding Contact: $($GALContact.WindowsEmailAddress)"
            try {
                New-EXCContact -MailboxName $Mailbox -DisplayName $GALContact.DisplayName -FirstName $GALContact.FirstName -LastName $GALContact.LastName -EmailAddress $GALContact.WindowsEmailAddress -CompanyName $GALContact.Company -Credentials $Credential -Department $GALContact.Department -BusinssPhone $GALContact.Phone -MobilePhone $GALContact.MobilePhone -JobTitle $GALContact.Title -Folder "Contacts\$FolderName" -useImpersonation
            } catch {
                Write-Log -Level "ERROR" -Message "Failed to sync $($GALContact.WindowsEmailAddress) contact to $($Mailbox)'s mailbox" -logfile $LogPath -exception $_.Exception.Message
            }
        }
    }
}