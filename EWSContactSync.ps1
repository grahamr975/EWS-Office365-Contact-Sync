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

.PARAMETER ClientID
Optional string; Used for ModernAuth/OAuth

.PARAMETER CredentialPath
Specifies the path of the Office 365 Administrator Credentials

.PARAMETER FolderName
Name of the folder that you want to import the contact into, if the folder does not exist a new one is created.
NOTE: To prevent duplicates, the specified folder is wiped before importing. For this reason, you should create a dedicated folder for the GAL.

.PARAMETER LogPath
Optional, Specifies the path of where the Log files are stored, along with the naming pattern of the log files.

.PARAMETER ExcludeContactsWithoutPhoneNumber
Optional Switch; Only sync contacts that have a phone or mobile number

.PARAMETER ExcludeSharedMailboxContacts
Optional Switch; Don't sync contacts that are a shared mailbox, or are a mailbox without a liscense

.PARAMETER IncludeNonUserContacts
Optional Switch; Also sync contacts that aren't users/mailboxes in your directory. These contacts must still have an email address.

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
    [System.IO.FileInfo]
    $CredentialPath,

    [Parameter(Mandatory=$True)]
    [String]
    $FolderName,

    [Parameter(Mandatory=$False)]
    [String]
    $LogPath,

    [Parameter(Mandatory=$True)]
    [String[]]
    $MailboxList,

    [Parameter(Mandatory=$false)]
    [String]
    $ClientID,

    [Parameter(Mandatory=$false)]
    [Switch]
    $ExcludeContactsWithoutPhoneNumber,

    [Parameter(Mandatory=$false)]
    [Switch]
    $ExcludeSharedMailboxContacts,
    
    [Parameter(Mandatory=$false)]
    [Switch]$IncludeNonUserContacts
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

Start-Transcript -OutputDirectory $LogPath -NoClobber

# Import Exchange Contacts module
Import-Module .\EWSContacts\Module\ExchangeContacts.psm1 -Force

# Import Office 365 Administrator credentials
$Credential = Import-CliXml -Path $CredentialPath

#-----------------------------------------------------------[Fetch Global Address List & Mailbox List]------------------------------------------------------------

# Fetch list of Global Address List contacts using Office 365 Powershell
$GALContacts = Get-GALContacts -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credentials $Credential -ExcludeContactsWithoutPhoneNumber $ExcludeContactsWithoutPhoneNumber -ExcludeSharedMailboxContacts $ExcludeSharedMailboxContacts -IncludeNonUserContacts $IncludeNonUserContacts

# If 'DIRECTORY' is used for $MailboxList, fetch all Mailboxes from the administrator account's Office 365 directory
if ($MailboxList -eq "DIRECTORY") {
    $MailboxList = Get-Mailboxes -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credentials $Credential
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

foreach ($Mailbox in $MailboxList) {
    try {
        Sync-ContactList -Mailbox $Mailbox -Credential $Credential -FolderName $FolderName -ContactList $GALContacts -ClientID $ClientID
    } catch {
        Write-Log -Level "ERROR" -Message "Failed to Sync-ContactList for $Mailbox" -exception $_.Exception.Message
    }
}

Stop-Transcript