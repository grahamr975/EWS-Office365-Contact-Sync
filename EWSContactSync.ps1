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

.PARAMETER MaxThreads
Optional, Specifies how many import jobs can be run at once.

.PARAMETER SleepTimer
Optional, Amount of time in miliseconds between the execution of each thread.

.EXAMPLE

Command Prompt

C:\> PowerShell.exe -ExecutionPolicy Bypass ^
-File "%CD%\Multi-Import.ps1" ^
-CredentialPath "%CD%\Tools\SecureCredential.cred" ^
-FolderName "L&L Contacts" ^
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

$ErrorActionPreference = "Stop"

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Import functions from library.ps1
try {
    .".\Functions\library.ps1"
    Write-Log -Message "Imported library.ps1" -logfile $LogPath
} catch {
    Write-Log -Level "FATAL" -Message "Failed to import library.ps1" -logfile $LogPath
}

# Import the ExchangeContacts Powershell module
try {
    Import-Module .\EWSContacts\Module\ExchangeContacts.psm1 -Force
    Write-Log -Message "Imported ExchangeContacts.psm1" -logfile $LogPath
} catch {
    Write-Log -Level "FATAL" -Message "Failed to import ExchangeContacts.psm1" -logfile $LogPath
}

# Import Office 365 administriator credentials (This account also needs impersonation permissions)
try {
    $Credential = Import-CliXml -Path $CredentialPath
    Write-Log -Message "Imported Office 365 Credentials from $($CredentialPath)" -logfile $LogPath
} catch {
    Write-Log -Level "FATAL" -Message "Failed to load CliXml credentials from $($CredentialPath)" -logfile $LogPath
}

# Fetch list of Global Address List contacts using Office 365 Powershell
try {
    $GALContacts = Get-GALContacts -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credentials $Credential
    Write-Log -Message "Fetched Office 365 Global Address List contacts" -logfile $LogPath
} catch {
    Write-Log -Level "FATAL" -Message "Failed to fetch Office 365 Global Address List contacts" -logfile $LogPath
}

# If 'DIRECTORY' is used for $MailboxList, fetch all Mailboxes from Office 365 directory
if ($MailboxList -eq "DIRECTORY") {
    try {
        # TO DO: ADD MAILBOX FETCH FEATURE (IT WILL REPLACE THE BELOW LINE)
        $MailboxList = $null
        Write-Log -Message "Fetched Office 365 mailboxes" -logfile $LogPath
    } catch {
        Write-Log -Level "FATAL" -Message "Failed to fetch Office 365 mailboxes" -logfile $LogPath
    }
    
}

foreach ($Mailbox in $MailboxList) {
    # Check if the contacts folder exists with $FolderName. If not, create it.
    try {
        New-EXCContactFolder -MailboxName $Mailbox -FolderName "$FolderName" -Credential $Credential
        Write-Log -Message "Verified $($FolderName) exists for $($Mailbox)" -logfile $LogPath
    } catch {
        Write-Log -Level "FATAL" -Message "Failed verify that $($FolderName) exists for $($Mailbox)" -logfile $LogPath
    }

    Write-Log -Message "Beginning contact sync for $($Mailbox)'s mailbox" -logfile $LogPath
    try {
        foreach ($GALContact in $GALContacts) {
            if ($null -ne $GALContact.WindowsEmailAddress) {
                Write-Output $GALContact.WindowsEmailAddress
                try {
                    if ($null -eq $GALContact.FirstName) {
                        # Try to update the contact if it already exists
                        $isContactFound = $(Set-EXCContact -MailboxName $Mailbox -DisplayName $GALContact.DisplayName -EmailAddress $GALContact.WindowsEmailAddress -CompanyName $GALContact.Company -Credentials $Credential -Department $GALContact.Department -BusinssPhone $GALContact.Phone -MobilePhone $GALContact.MobilePhone -JobTitle $GALContact.Title -Folder "Contacts\$FolderName" -useImpersonation -force)
                        # If the contact does not yet exist, create a new contact
                        if ($isContactFound -eq $false) {
                            New-EXCContact -MailboxName $Mailbox -DisplayName $GALContact.DisplayName -EmailAddress $GALContact.WindowsEmailAddress -CompanyName $GALContact.Company -Credentials $Credential -Department $GALContact.Department -BusinssPhone $GALContact.Phone -MobilePhone $GALContact.MobilePhone -JobTitle $GALContact.Title -Folder "Contacts\$FolderName" -useImpersonation
                        }
                    } else {
                        # Try to update the contact if it already exists
                        $isContactFound = $(Set-EXCContact -MailboxName $Mailbox -DisplayName $GALContact.DisplayName -FirstName $GALContact.FirstName -LastName $GALContact.LastName -EmailAddress $GALContact.WindowsEmailAddress -CompanyName $GALContact.Company -Credentials $Credential -Department $GALContact.Department -BusinssPhone $GALContact.Phone -MobilePhone $GALContact.MobilePhone -JobTitle $GALContact.Title -Folder "Contacts\$FolderName" -useImpersonation -force)
                        # If the contact does not yet exist, create a new contact
                        if ($isContactFound -eq $false) {
                            New-EXCContact -MailboxName $Mailbox -DisplayName $GALContact.DisplayName -FirstName $GALContact.FirstName -LastName $GALContact.LastName -EmailAddress $GALContact.WindowsEmailAddress -CompanyName $GALContact.Company -Credentials $Credential -Department $GALContact.Department -BusinssPhone $GALContact.Phone -MobilePhone $GALContact.MobilePhone -JobTitle $GALContact.Title -Folder "Contacts\$FolderName" -useImpersonation
                        }
                    }
                } catch {
                    Write-Log -Level "ERROR" -Message "Failed to sync $($GALContact.WindowsEmailAddress) contact to $($Mailbox)'s mailbox" -logfile $LogPath -exception $_.Exception.Message
                }
            }
        }
        Write-Log -Message "Sucessfully synced $($Mailbox)'s mailbox" -logfile $LogPath
    } catch {
        Write-Log -Level "FATAL" "Failed to sync $($Mailbox)'s mailbox" -logfile $LogPath
    }
}