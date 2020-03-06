function Get-EXCContactsObject
{
<#
	.SYNOPSIS
		Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER Folder
		A description of the Folder parameter.
	
	.PARAMETER useImpersonation
		A description of the useImpersonation parameter.
	
	.EXAMPLE
		Example 1 To get a Contact from a Mailbox's default contacts folder
		Get-EXCContacts -MailboxName mailbox@domain.com
		
	.EXAMPLE
		Example 2 To get all the Contacts from subfolder of the Mailbox's default contacts folder
		Get-EXCContacts -MailboxName mailbox@domain.com -Folder \Contact\test
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 2, Mandatory = $true)]
		$Folder,

		[Parameter(Position = 3, Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeService]
		$Service,
		
		[switch]
		$useImpersonation
	)
	Begin
	{
		#Connect
		$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox);

		if ($service.URL)
		{
			$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Contact")
			#Define ItemView to retrive just 2000 Items    
			$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(2000)
			$fiItems = $null
			do
			{
				$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
				$fiItems = $service.FindItems($Folder.Id, $SfSearchFilter, $ivItemView)
				if ($fiItems.Items.Count -gt 0)
				{
					$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
					$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
					$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
					[Void]$service.LoadPropertiesForItems($fiItems, $psPropset)
					foreach ($Item in $fiItems.Items)
					{
						Write-Output $Item
					}
				}
				$ivItemView.Offset += $fiItems.Items.Count
			}
			while ($fiItems.MoreAvailable -eq $true)
			
		}
	}
}