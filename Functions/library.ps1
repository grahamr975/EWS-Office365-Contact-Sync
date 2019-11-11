function Get-GALContacts {
		<#
		.SYNOPSIS
			Uses Office 365 services to generate a list of contacts 
		
		.PARAMETER ConnectionUri
			Used to connect to Office 365, by default this is https://outlook.office365.com/powershell-liveid/.
		
		.PARAMETER Credentials
			Office 365 Admin Credentials
		
		.EXAMPLE
			PS C:\> Get-GALContacts -ConnectionUri 'https://outlook.office365.com/powershell-liveid/' -Credentials $Credentials
		#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$ConnectionUri,

		[Parameter(Position = 1, Mandatory = $true)]
		[System.Management.Automation.PSCredential]
		$Credentials
	)
	process {
		try {
			# Connect to Office 365 Exchange Server using a Remote Session
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $Credentials -Authentication Basic -AllowRedirection
		Import-PSSession $Session -DisableNameChecking
		
			$DirectoryList = $(Get-Mailbox -ResultSize unlimited | Where-Object {$_.HiddenFromAddressListsEnabled -Match "False"})
			$EmailAddressList = $DirectoryList.PrimarySMTPAddress
		
			# Import Global Address List into Powershell from Office 365 Exchange as an array
			$ContactList = Get-User -ResultSize unlimited | Where-Object {$null -ne $_.WindowsEmailAddress}
			$ContactList = $ContactList | Select-Object DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone | Where-Object {$EmailAddressList.Contains($_.WindowsEmailAddress)}
		Remove-PSSession $Session
		return $ContactList
		} catch {
			Write-Log -Level "FATAL" -Message "Failed to fetch Global Address List Contacts from Office 365 Directory" -exception $_.Exception.Message
		}
	}
}

function Get-Mailboxes {
	<#
	.SYNOPSIS
		Uses Office 365 services to generate a list of contacts 
	
	.PARAMETER ConnectionUri
		Used to connect to Office 365, by default this is https://outlook.office365.com/powershell-liveid/.
	
	.PARAMETER Credentials
		Office 365 Admin Credentials
	
	.EXAMPLE
		PS C:\> Get-GALContacts -ConnectionUri 'https://outlook.office365.com/powershell-liveid/' -Credentials $Credentials
	#>
[CmdletBinding()]
param (
	[Parameter(Position = 0, Mandatory = $true)]
	[string]
	$ConnectionUri,

	[Parameter(Position = 1, Mandatory = $true)]
	[System.Management.Automation.PSCredential]
	$Credentials
)
process {
	try {
		# Connect to Office 365 Exchange Server using a Remote Session
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $Credentials -Authentication Basic -AllowRedirection
	Import-PSSession $Session -DisableNameChecking
	
		$DirectoryList = $(Get-Mailbox -ResultSize unlimited | Where-Object {$_.HiddenFromAddressListsEnabled -Match "False"})
		$EmailAddressList = $DirectoryList.PrimarySMTPAddress
	Remove-PSSession $Session
	return $EmailAddressList
	} catch {
		Write-Log -Level "FATAL" -Message "Failed to fetch user mailbox list from Office 365 directory" -exception $_.Exception.Message
	}
}
}

function Get-EmailAddressFromContact {
	<#
	.SYNOPSIS
		Return the email address of an EWS Contact Object
	
	.PARAMETER Contact
		Used to connect to Office 365, by default this is https://outlook.office365.com/powershell-liveid/.
	
	#>
[CmdletBinding()]
param (
	[Parameter(Position = 0, Mandatory = $true)]$Contact
)
process {
		foreach ($ContactItem in $Contact) {
			$EmailList += $ContactItem.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address
		}
		return $EmailList
	}
}

Function Write-Log {
	<#
	.SYNOPSIS
		Writes the specified message to the screen. If $logfile is specified, write to the file instead.
	
	.PARAMETER Level
		Indicates the type of message. The options are INFO, WARN, ERROR, FATAL, and DEBUG. Uses INFO by default.
	
	.PARAMETER Message
		Message to write to the screen or the $logfile (if specified)

	.PARAMETER logfile
		Optional; File path of the file to write the message to.

	.PARAMETER exception
		Optional; Additional information to include when the script encounters an error. Very useful for troubleshooting.
	
	.EXAMPLE
		PS C:\> Write-Log -Level "INFO" -Message "Successfully executed the script" -logfile $LogFilePath
#>
[CmdletBinding()]
Param(
[Parameter(Mandatory=$False)]
[ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
[String]
$Level = "INFO",

[Parameter(Mandatory=$True)]
[string]
$Message,

[Parameter(Mandatory=$False)]
[string]
$logfile,

[Parameter(Mandatory=$False)]
[string]
$exception
)

if ($logfile -and !(Test-Path $logfile)) {
	New-Item -ItemType "file" -Path "$logfile" -Force | Out-Null
  }

$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
$Line = "$Stamp $Level $Message $exception"
If($logfile) {
	Add-Content $logfile -Value $Line -Force
}
Else {
	Write-Output $Line
}
if ($Level -eq "FATAL") {
	Write-Error $Message
}
}

function Sync-ContactList {
	<#
	.SYNOPSIS
		Uses Exchange Web Services API to syncronise a list of contacts to an Office 365 Mailbox
	
	.PARAMETER Mailbox
		Email address of an Office 365 user
	
	.PARAMETER Credential
		Office 365 Admin Credentials

	.PARAMETER FolderName
		Name of the contact folder that will be synced to

	.PARAMETER ContactList
		An array with DisplayName, FirstName, LastName, Title, Company, Department, WindowsEmailAddress, Phone, and MobilePhone properties
	
	.EXAMPLE
		PS C:\> Sync-ContactList -Mailbox "john.doe@example.com" -Credential $Credentials -FolderName "myFolder" -ContactList $Contacts
	#>
[CmdletBinding()]
param (
	[Parameter(Position = 0, Mandatory = $true)]
	[string]
	$Mailbox,

	[Parameter(Position = 1, Mandatory = $true)]
	[System.Management.Automation.PSCredential]
	$Credential,

	[Parameter(Position = 2, Mandatory = $true)]
	[string]
	$FolderName,

	[Parameter(Position = 3, Mandatory = $true)]
	$ContactList
)
process {
	Write-Log -Message "Beginning contact sync for $($Mailbox)'s mailbox"

	# Check if a contacts folder exists with $FolderName. If not, create it.
	New-EXCContactFolder -MailboxName $Mailbox -FolderName "$FolderName" -Credential $Credential


	# Fetch contacts from the user's mailbox
	$MailboxContacts = Get-EXCContacts -MailboxName $Mailbox -Credentials $Credential -Folder "Contacts\$FolderName" | Where-Object {$null -ne $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address}
	$MailboxEmailList = Get-EmailAddressFromContact -Contact $MailboxContacts

	# For each contact, determine if it needs to be deleted, updated, or created
	# Delete = Any contact in the user's mailbox that does not have an email address matching any contact in the Global Address List
	# Update = Any contact in the user's mailbox that has a matching email address with a contact in the Global Address List
	# Create = Any contact in the Global Address List does not does not have a matching email with a contact in the user's mailbox
	$MailboxContactsToBeDeleted = $MailboxContacts | Where-Object {!$ContactList.WindowsEmailAddress.ToLower().Contains($_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower())}
	$MailboxContactsToBeUpdated = $ContactList | Where-Object {$MailboxEmailList.Contains($_.WindowsEmailAddress)}
	$MailboxContactsToBeCreated = $ContactList | Where-Object {!$MailboxEmailList.Contains($_.WindowsEmailAddress)}

	#
	# DELETE
	#
	# NOTE: This cannot yet remove contacts with no email address!
	try {
		foreach ($Contact in $MailboxContactsToBeDeleted) {
			Write-Verbose "Deleting Contact: $($Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower())"
			$Contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
		}
	} catch {
		Write-Log -Level "ERROR" -Message "Failed to remove all obsolete contacts from $($Mailbox)'s mailbox" -exception $_.Exception.Message
	}
	#
	# UPDATE
	#
	foreach ($Contact in $MailboxContactsToBeUpdated) {
		# Search for a identical contact. If the identical contact already exists in the user's mailbox, don't made any changes to it since they aren't needed.
		if ($null -eq $($MailboxContacts | Where-Object {(($_.GivenName -eq $Contact.FirstName) -or ("" -eq $Contact.FirstName)) -and (($_.Surname -eq $Contact.LastName) -or ("" -eq $Contact.LastName)) -and ($_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].address -eq $Contact.WindowsEmailAddress) -and ($Contact.Company -eq $_.CompanyName -or $Contact.Company -eq "") -and ($Contact.Department -eq $_.Department -or $Contact.Department -eq "") -and (($_.DisplayName -eq $Contact.DisplayName) -or ($Contact.DisplayName -eq "")) -and ($Contact.Title -eq $_.JobTitle -or $Contact.Title -eq "") -and ($Contact.Phone -eq $_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] -or $Contact.Phone -eq "") -and ($Contact.MobilePhone -eq $_.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] -or $Contact.MobilePhone -eq "")})) {
			if ($null -ne $Contact.WindowsEmailAddress) {
				Write-Verbose "Updating Contact: $($Contact.WindowsEmailAddress)"
				try {
					Set-EXCContact -MailboxName $Mailbox -DisplayName $Contact.DisplayName -FirstName $Contact.FirstName -LastName $Contact.LastName -EmailAddress $Contact.WindowsEmailAddress -CompanyName $Contact.Company -Credentials $Credential -Department $Contact.Department -BusinssPhone $Contact.Phone -MobilePhone $Contact.MobilePhone -JobTitle $Contact.Title -Folder "Contacts\$FolderName" -useImpersonation -force
				} catch {
					Write-Log -Level "ERROR" -Message "Failed to sync $($Contact.WindowsEmailAddress) contact to $($Mailbox)'s mailbox" -exception $_.Exception.Message
				}
			}
		}
	}
	#
	# CREATE
	#
	foreach ($Contact in $MailboxContactsToBeCreated) {
		if ($null -ne $Contact.WindowsEmailAddress) {
			Write-Verbose "Creating Contact: $($Contact.WindowsEmailAddress)"
			try {
				New-EXCContact -MailboxName $Mailbox -DisplayName $Contact.DisplayName -FirstName $Contact.FirstName -LastName $Contact.LastName -EmailAddress $Contact.WindowsEmailAddress -CompanyName $Contact.Company -Credentials $Credential -Department $Contact.Department -BusinssPhone $Contact.Phone -MobilePhone $Contact.MobilePhone -JobTitle $Contact.Title -Folder "Contacts\$FolderName" -useImpersonation
			} catch {
				Write-Log -Level "ERROR" -Message "Failed to create $($Contact.WindowsEmailAddress) contact in $($Mailbox)'s mailbox" -exception $_.Exception.Message
			}
		}
	}
}
}