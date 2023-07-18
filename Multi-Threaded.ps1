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
    $CertificatePath,

    [Parameter(Mandatory=$True)]
	[System.IO.FileInfo]
    $CertificatePasswordPath,

    [Parameter(Mandatory=$True)]
    [String]
    $ExchangeOrg,

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
    $ModernAuth,

    [Parameter(Mandatory=$false)]
    [Switch]
    $ExcludeContactsWithoutPhoneNumber,

    [Parameter(Mandatory=$false)]
    [Switch]
    $ExcludeSharedMailboxContacts,
    
    [Parameter(Mandatory=$false)]
    [Switch]$IncludeNonUserContacts,

    [Parameter(Mandatory=$false)]
    $ModulePath = "$((Get-Location).path)\EWSContacts\Module\ExchangeContacts.psm1",

    [Parameter(Mandatory=$false)]
    [Int]$MaxThreads = 10,

    [Parameter(Mandatory=$false)]
	[Int]$SleepTimer = 1000
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Import Exchange Contacts module
Import-Module $ModulePath -Force
Import-Module ExchangeOnlineManagement -RequiredVersion 3.2.0 -Force

# Import the Exchange Certificate Password
[Security.SecureString]$CertificatePassword = Import-CliXml -Path "$CertificatePasswordPath"

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Stop on Error
# $VerbosePreference = "Continue"
$ErrorActionPreference = "Stop"

Start-Transcript -OutputDirectory $LogPath -NoClobber
#-----------------------------------------------------------[Fetch Global Address List & Mailbox List]------------------------------------------------------------
Write-Log -Message "Fetching Contact List..."
# Fetch list of Global Address List contacts using Office 365 Powershell
$GALContacts = Get-GALContacts -ConnectionUri https://outlook.office365.com/powershell-liveid/ -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -ExchangeOrg $ExchangeOrg -ClientID $ClientID -ExcludeContactsWithoutPhoneNumber $ExcludeContactsWithoutPhoneNumber -ExcludeSharedMailboxContacts $ExcludeSharedMailboxContacts -IncludeNonUserContacts $IncludeNonUserContacts

Write-Log -Message "Fetching Mailbox List..."
# If 'DIRECTORY' is used for $MailboxList, fetch all Mailboxes from the administrator account's Office 365 directory
if ($MailboxList -eq "DIRECTORY") {
    $MailboxList = Get-Mailboxes -ConnectionUri https://outlook.office365.com/powershell-liveid/ -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -ExchangeOrg $ExchangeOrg -ClientID $ClientID
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#################
# START THREADS #
#################
#$i = 0

Write-Log -Message "Beginning Multi-Threaded Sync..."

foreach ($Mailbox in $MailboxList) {
    While ($(Get-Job -state running).count -ge $MaxThreads){
        #Write-Progress  -Activity "Importing Contacts into Office 365 Mailbox" -Status "Waiting for Contact Import Jobs to Finish..." -CurrentOperation "$i Contact Import Jobs created - $($(Get-Job -state running).count) jobs running" -PercentComplete ($i / $MailboxList.count * 100)
        Start-Sleep -Milliseconds $SleepTimer
    }
    try {
        #$i = $i + 1
        $Job = Start-Job -Name $Mailbox -ScriptBlock ${function:Sync-ContactList} -ArgumentList $Mailbox, $CertificatePath, $CertificatePassword, $ClientID, $ModernAuth, $FolderName, $GALContacts, $ModulePath
        Write-Log -Message "Started Mailbox Sync PowerShell Job for $($Job.Name)"
        #Write-Progress -Activity "Importing Contacts into Office 365 Mailbox" -Status "Starting Threads" -CurrentOperation "$i Contact Imports started - $($(Get-Job -state running).count) jobs running" -PercentComplete ($i / $MailboxList.count * 100)
    } catch {
        Write-Log -Level "ERROR" -Message "Failed to Sync-ContactList for $Mailbox" -exception $_.Exception.Message
    }
}

# Wait for remaining jobs to finish...
While ($(Get-Job -State Running).count -gt 0){
    $ThreadsStillRunning = ""
    ForEach ($System  in $(Get-Job -state running)){$ThreadsStillRunning += ", $($System.name)"}
    $ThreadsStillRunning = $ThreadsStillRunning.Substring(2)
    Write-Progress  -Activity "Importing Contacts into Office 365 Mailbox" -Status "$($(Get-Job -State Running).count) Import(s) remaining" -CurrentOperation "$ThreadsStillRunning" -PercentComplete ($(Get-Job -State Completed).count / $(Get-Job).count * 100)
    Start-Sleep -Milliseconds $SleepTimer
}

Write-Log -Message "All jobs complete: retrieving results..."

# Print results of all jobs
ForEach($Job in Get-Job) {
    try {
        Write-Output $(Receive-Job $Job -ErrorAction Stop)
    } catch {
        Write-Log -Level "ERROR" -Message "Failed to Sync-ContactList for $($Job.Name)" -exception $_.Exception.Message
    }
}

# Clean up sync jobs, stop the transcript, and exit
Stop-Job *
Remove-Job *
Stop-Transcript
Exit 0