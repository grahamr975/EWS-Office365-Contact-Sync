function Get-GALContacts {
	<#
	.SYNOPSIS
		Uses Office 365 services to generate a list of contacts. Only includes contacts with an email address.
	
	.PARAMETER ConnectionUri
		Used to connect to Office 365, by default this is https://outlook.office365.com/powershell-liveid/.
	
	.PARAMETER CertificatePath
		Office 365 Azure App Certificate File Path (See README for details)

	.PARAMETER ExcludeContactsWithoutPhoneNumber
		Switch; Only return contacts that have a phone or mobile number

	.PARAMETER ExcludeSharedMailboxContacts
		Switch; Excludes contacts that are a shared mailbox, or a mailbox without a liscense

	.PARAMETER IncludeNonUserContacts
		Switch; Also return contacts that aren't users/mailboxes in your directory. These contacts must still have an email address.
	
	.EXAMPLE
		PS C:\> Get-GALContacts -ConnectionUri 'https://outlook.office365.com/powershell-liveid/' -Credentials $Credentials
	#>
[CmdletBinding()]
param (
	[Parameter(Position = 0, Mandatory = $true)]
	[string]
	$ConnectionUri,

	[System.IO.FileInfo]
    $CertificatePath,

	[Security.SecureString]
    $CertificatePassword,

	[String]
	$ExchangeOrg,

	[String]
	$ClientID,

	[Parameter(Position = 2, Mandatory = $false)]
	[bool]
	$ExcludeContactsWithoutPhoneNumber,

	[Parameter(Position = 3, Mandatory = $false)]
	[bool]
	$ExcludeSharedMailboxContacts,

	[Parameter(Position = 4, Mandatory = $false)]
	[bool]
	$IncludeNonUserContacts
)
process {
	try {
		# Connect to Office 365 Exchange Server using a Remote Session
        Connect-ExchangeOnline -ConnectionUri $ConnectionUri -CertificateFilePath $CertificatePath -CertificatePassword $CertificatePassword -AppId $ClientID -Organization $ExchangeOrg
		
		# Import Global Address List into Powershell from Office 365 Exchange as an array
		$ContactList = Get-User -ResultSize unlimited 

		# If the ExcludeSharedMailboxContacts switch is enabled, exclude contacts that are a shared mailbox or mailbox with no liscense
		if ($ExcludeSharedMailboxContacts) {
			$DirectoryList = $(Get-EXOMailbox -ResultSize unlimited -PropertySets Minimum,AddressList | Where-Object {$_.HiddenFromAddressListsEnabled -Match "False"})
			$EmailAddressList = $DirectoryList.PrimarySMTPAddress
			$ContactList = $ContactList | Select-Object DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone | Where-Object {$EmailAddressList -Contains $_.WindowsEmailAddress}
			} else {
			$ContactList = $ContactList | Select-Object DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone
		}
		
		# If the IncludeNonUserContacts switch is enabled, also include contacts that aren't actual users in the directory
		if ($IncludeNonUserContacts) {
			$ContactList += Get-Contact -ResultSize unlimited | Select-Object DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone
		}

		# If the ExcludeContactsWithoutPhoneNumber switch is enabled, exclude contacts that don't have a phone or mobile number
		if ($ExcludeContactsWithoutPhoneNumber) {
			$ContactList = $ContactList | Where-Object {$_.Phone -or $_.MobilePhone}
		}
        Disconnect-ExchangeOnline -Confirm:$false

	# Only return contacts with email addresses
	return $ContactList | Where-Object {$null -ne $_.WindowsEmailAddress -and "" -ne $_.WindowsEmailAddress}
	} catch {
		Write-Log -Level "FATAL" -Message "Failed to fetch Global Address List Contacts from Office 365 Directory" -exception $_.Exception.Message
	}
}
}