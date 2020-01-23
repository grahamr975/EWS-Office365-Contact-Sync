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
	Import-PSSession $Session -DisableNameChecking -AllowClobber
	
		$DirectoryList = $(Get-Mailbox -ResultSize unlimited | Where-Object {$_.HiddenFromAddressListsEnabled -Match "False"})
	Remove-PSSession $Session

	$EmailAddressList = $DirectoryList.PrimarySMTPAddress

	} catch {
		Write-Log -Level "FATAL" -Message "Failed to fetch user mailbox list from Office 365 directory" -exception $_.Exception.Message
	}
	return $EmailAddressList
}
}