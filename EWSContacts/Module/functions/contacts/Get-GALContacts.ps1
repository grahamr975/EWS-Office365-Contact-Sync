function Get-GALContacts {
	<#
	.SYNOPSIS
		Uses Office 365 services to generate a list of contacts. Only includes contacts with an email address.
	
	.PARAMETER ConnectionUri
		Used to connect to Office 365, by default this is https://outlook.office365.com/powershell-liveid/.
	
	.PARAMETER Credentials
		Office 365 Admin Credentials

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

	[Parameter(Position = 1, Mandatory = $true)]
	[System.Management.Automation.PSCredential]
	$Credentials,

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
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $Credentials -Authentication Basic -AllowRedirection
	Import-PSSession $Session -DisableNameChecking -AllowClobber
	
		# Import Global Address List into Powershell from Office 365 Exchange as an array
		$ContactList = Get-User -ResultSize unlimited 

		# If the ExcludeSharedMailboxContacts switch is enabled, exclude contacts that are a shared mailbox or mailbox with no liscense
		if ($ExcludeSharedMailboxContacts) {
			$DirectoryList = $(Get-Mailbox -ResultSize unlimited | Where-Object {$_.HiddenFromAddressListsEnabled -Match "False"})
			$EmailAddressList = $($DirectoryList.PrimarySMTPAddress).ToLower()
			$ContactList = $ContactList | Select-Object DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone | Where-Object {$EmailAddressList.Contains($_.WindowsEmailAddress.ToLower())}
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
	Remove-PSSession $Session
	# Only return contacts with email addresses
	return $ContactList | Where-Object {$null -ne $_.WindowsEmailAddress -and "" -ne $_.WindowsEmailAddress}
	} catch {
		Write-Log -Level "FATAL" -Message "Failed to fetch Global Address List Contacts from Office 365 Directory" -exception $_.Exception.Message
	}
}
}