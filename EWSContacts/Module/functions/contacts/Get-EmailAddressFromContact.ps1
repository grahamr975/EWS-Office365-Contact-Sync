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
		$EmailList = @()

		foreach ($ContactItem in $Contact) {
			$EmailList += "$($ContactItem.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address.ToLower())"
		}
		return $EmailList
	}
}