function Get-Mailboxes {
	<#
	.SYNOPSIS
		Uses Office 365 services to generate a list of user mailboxes 
	
	.PARAMETER ConnectionUri
		Used to connect to Office 365, by default this is https://outlook.office365.com/powershell-liveid/.
	
	.PARAMETER CertificatePath
		Office 365 Azure App Certificate File Path (See README for details)
	
	.EXAMPLE
		PS C:\> Get-GALContacts -ConnectionUri 'https://outlook.office365.com/powershell-liveid/' -Credentials $Credentials
	#>
[CmdletBinding()]
param (
	[Parameter(Position = 0, Mandatory = $true)]
	[string]
	$ConnectionUri,

	[Parameter(Position = 1, Mandatory = $true)]
	[System.IO.FileInfo]
	$CertificatePath,

	[Parameter(Position = 2, Mandatory = $true)]
	[Security.SecureString]
    $CertificatePassword,
	
	[Parameter(Position = 3, Mandatory = $true)]
	[String]
	$ExchangeOrg,

	[Parameter(Position = 4, Mandatory = $true)]
	[String]
	$ClientID
)
process {
	try {
	# $Null = @() is a workaround for this function returning a random filename such as "tmp_z1ci55dv.kke" at the start of the output....
	$Null = @(
		# Connect to Office 365 Exchange Server using a Remote Session
	Connect-ExchangeOnline -ConnectionUri $ConnectionUri -CertificateFilePath $CertificatePath -CertificatePassword $CertificatePassword -AppId $ClientID -Organization $ExchangeOrg
			$DirectoryList = $(Get-EXOMailbox -ResultSize unlimited | Where-Object {$_.HiddenFromAddressListsEnabled -Match "False"}).PrimarySMTPAddress
    Disconnect-ExchangeOnline -Confirm:$false
	)

	} catch {
		Write-Log -Level "FATAL" -Message "Failed to fetch user mailbox list from Office 365 directory" -exception $_.Exception.Message
	}
	return $DirectoryList | Where-Object {$null -ne $_ -and "" -ne $_}
}
}