function Get-GALContacts {
	<#
	.SYNOPSIS
		Uses Office 365 services to generate a list of contacts 
	
	.PARAMETER ConnectionUri
		Used to connect to Office 365, by default this is https://outlook.office365.com/powershell-liveid/.
	
	.PARAMETER Credentials
		Office 365 Admin Credentials

	.PARAMETER RequirePhoneNumber
		Switch; Only return contacts that have a phone or mobile number

	.PARAMETER IncludeNonMailboxContacts
		Switch; Also include directory users that don't have an actual mailbox
	
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
	$RequirePhoneNumber,

	[Parameter(Position = 3, Mandatory = $false)]
	[bool]
	$IncludeNonMailboxContacts
)
process {
	try {
		# Connect to Office 365 Exchange Server using a Remote Session
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $Credentials -Authentication Basic -AllowRedirection
	Import-PSSession $Session -DisableNameChecking -AllowClobber
	
		# Import Global Address List into Powershell from Office 365 Exchange as an array
		$ContactList = Get-User -ResultSize unlimited | Where-Object {$null -ne $_.WindowsEmailAddress}

		# If the IncludeNonMailboxContacts switch is enabled, also include contacts that don't have a mailbox in your directory.
		if ($IncludeNonMailboxContacts) {
			$ContactList = $ContactList | Select-Object DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone
		} else {
			$DirectoryList = $(Get-Mailbox -ResultSize unlimited | Where-Object {$_.HiddenFromAddressListsEnabled -Match "False"})
			$EmailAddressList = $DirectoryList.PrimarySMTPAddress
			$ContactList = $ContactList | Select-Object DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone | Where-Object {$EmailAddressList.Contains($_.WindowsEmailAddress)}
		}
		if ($RequirePhoneNumber) {
			$ContactList = $ContactList | Where-Object {$_.Phone -or $_.MobilePhone}
		}
	Remove-PSSession $Session
	return $ContactList
	} catch {
		Write-Log -Level "FATAL" -Message "Failed to fetch Global Address List Contacts from Office 365 Directory" -exception $_.Exception.Message
	}
}
}