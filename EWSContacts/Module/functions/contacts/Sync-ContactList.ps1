function Sync-ContactList {
	<#
	.SYNOPSIS
		Uses Exchange Web Services API to syncronise a list of contacts to an Office 365 Mailbox
	
	.PARAMETER Mailbox
		Target mailbox that the ContactList will be synced to; Email address of an Office 365 business user who is in the same directory as the admin credentials
	
	.PARAMETER Credential
		Office 365 administrator credentials -- This account needs to have application impersonation permission

	.PARAMETER ClientID
		Used for ModernAuth/OAuth

	.PARAMETER ModernAuth
		Enables Modern Authenication, See https://docs.microsoft.com/en-us/office365/enterprise/office-365-client-support-modern-authentication 

	.PARAMETER FolderName
		Name of the contact folder that the ContactList will be synced to

	.PARAMETER ContactList
		An array with DisplayName, FirstName, LastName, Title, Company, Department, WindowsEmailAddress, Phone, and MobilePhone fields
	
	.EXAMPLE
		PS C:\> Sync-ContactList -Mailbox "john.doe@example.com" -Credential $Credentials -FolderName "myFolder" -ContactList $Contacts
	#>
[CmdletBinding()]
param (
	[Parameter(Position = 0, Mandatory = $true)]
	[string]
	$Mailbox,

	[Parameter(Position = 1, Mandatory=$True)]
    [System.IO.FileInfo]
    $CertificatePath,

    [Parameter(Position = 2, Mandatory=$True)]
	[Security.SecureString]
    $CertificatePassword,

	[Parameter(Position = 3, Mandatory = $false)]
	[string]
	$ClientID,

	[Parameter(Position = 4, Mandatory = $false)]
	[bool]
	$ModernAuth,

	[Parameter(Position = 5, Mandatory = $true)]
	[string]
	$FolderName,

	[Parameter(Position = 6, Mandatory = $true)]
	$ContactList
)
process {
	Write-Log -Message "Beginning contact sync for $($Mailbox)'s mailbox"

	# Create EWS Service object
	$service = Connect-EXCExchange -MailboxName $Mailbox -CertificateFilePath $CertificatePath -CertificatePassword $CertificatePassword -ClientID $ClientID -ModernAuth $ModernAuth
	$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox);

	# Check if a contacts folder exists with $FolderName. If not, create it.
	New-EXCContactFolder -MailboxName $Mailbox -FolderName "$FolderName" -Service $service

	# Fetch folder & contacts from the user's mailbox
	$ContactsFolderObject = Get-EXCContactFolder -Service $service -FolderPath "Contacts\$FolderName" -SmptAddress $Mailbox
	$MailboxContacts = Get-EXCContactsObject -MailboxName $Mailbox -Folder $ContactsFolderObject -service $service | Where-Object {$null -ne $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address}
	
	# If the user has no contacts, add them all in.
	if ($null -eq $MailboxContacts) {
		$MailboxContactsToBeDeleted = $null
		$MailboxContactsToBeUpdated = $null
		$MailboxContactsToBeCreated = $ContactList
	} else {
		$MailboxEmailList = Get-EmailAddressFromContact -Contact $MailboxContacts
		# For each contact, determine if it needs to be deleted, updated, or created
		# Delete = Any contact in the user's mailbox that does not have an email address matching any contact in the Global Address List; Contacts with no email address are not deleted.
		# Update = Any contact in the user's mailbox that has a matching email address with a contact in the Global Address List
		# Create = Any contact in the Global Address List does not does not have a matching email with a contact in the user's mailbox
		$MailboxContactsToBeDeleted = $MailboxContacts | Where-Object {$ContactList.WindowsEmailAddress.ToLower() -NotContains $_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower()}
		$MailboxContactsToBeUpdated = $ContactList | Where-Object {$MailboxEmailList -Contains $_.WindowsEmailAddress.ToLower()}
		$MailboxContactsToBeCreated = $ContactList | Where-Object {$MailboxEmailList -NotContains $_.WindowsEmailAddress.ToLower()}
	}

	#
	# DELETE
	#
	# NOTE: This cannot yet remove contacts with no email address!
	try {
		foreach ($Contact in $MailboxContactsToBeDeleted) {
			Write-Log -Message "Deleting Contact: $($Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower())"
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
				Write-Log -Message "Updating Contact: $($Contact.WindowsEmailAddress)"
				$ContactObject = $MailboxContacts | Where-Object {$_.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].address -eq $Contact.WindowsEmailAddress}
				# If more than one contact is found with the same email, sync the first contact and delete the reset
				if (($ContactObject | Measure-Object).Count -gt 1) {
					$ContactToSync = $ContactObject | Select-Object -First 1
					$MailboxContactDuplicates = $ContactObject | Select-Object -Skip 1
					foreach ($ContactDuplicate in $MailboxContactDuplicates) {
						Write-Log -Message "Deleting Contact With Duplicate Email: $($ContactDuplicate.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower())"
						$ContactDuplicate.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
					}
				} else {
					$ContactToSync = $ContactObject
				}

				try {
					Set-EXCContactObject -MailboxName $Mailbox -DisplayName $Contact.DisplayName -FirstName $Contact.FirstName -LastName $Contact.LastName -EmailAddress $Contact.WindowsEmailAddress -CompanyName $Contact.Company -Department $Contact.Department -BusinssPhone $Contact.Phone -MobilePhone $Contact.MobilePhone -JobTitle $Contact.Title -Folder "Contacts\$FolderName" -useImpersonation -force -Contact $ContactToSync
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
			Write-Log -Message "Creating Contact: $($Contact.WindowsEmailAddress)"
			try {
				New-EXCContactObject -MailboxName $Mailbox -DisplayName $Contact.DisplayName -FirstName $Contact.FirstName -LastName $Contact.LastName -EmailAddress $Contact.WindowsEmailAddress -CompanyName $Contact.Company -Department $Contact.Department -BusinssPhone $Contact.Phone -MobilePhone $Contact.MobilePhone -JobTitle $Contact.Title -Folder $ContactsFolderObject -useImpersonation -service $service
			} catch {
				Write-Log -Level "ERROR" -Message "Failed to create $($Contact.WindowsEmailAddress) contact in $($Mailbox)'s mailbox" -exception $_.Exception.Message
			}
		}
	}
}
}