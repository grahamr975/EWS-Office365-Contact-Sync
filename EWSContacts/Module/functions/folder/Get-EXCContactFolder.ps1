function Get-EXCContactFolder
{
	<#
		.SYNOPSIS
			Returns the folder object of the specified path.
		
		.DESCRIPTION
			Returns the folder object of the specified path.
		
		.PARAMETER FolderPath
			The path to the folder, relative to the message folder base.
		
		.PARAMETER SmptAddress
			The email address of the mailbox to access
		
		.PARAMETER Service
			The established Service connection to use for this connection.
			Use 'Connect-EXCExchange' in order to establish a connection and obtain such an object.
		
		.EXAMPLE
			PS C:\> Get-EXCContactFolder -FolderPath 'Contacts\Private' -SmptAddress 'peter@example.com' -Service $Service
	
			Returns the 'Private' folder within the contacts folder for the mailbox peter@example.com
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$true)]
		[String]$FolderPath,

		[Parameter(Mandatory=$true)]
		[String]$SmptAddress,
		
		[Parameter(Mandatory=$true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeService]$Service
	)
	process
	{
		$MailboxName = $SmptAddress
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName)
		$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $folderid)
		#Split the Search path into an array  
		$fldArray = $FolderPath.Split("\")
		#Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++)
		{
			#Perform search based on the displayname of each folder level 
			$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
			$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $fldArray[$lint])
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
			$findFolderResults = $Service.FindFolders($tfTargetFolder.Id, $SfSearchFilter, $fvFolderView)
			if ($findFolderResults.TotalCount -gt 0)
			{
				foreach ($folder in $findFolderResults.Folders)
				{
					$tfTargetFolder = $folder
				}
			}
			else
			{
				Write-host "Folder Not Found"
				$tfTargetFolder = $null
				break
			}
		}
		if ($null-ne $tfTargetFolder)
		{
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
			return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $tfTargetFolder.Id)
		}
		else
		{
			throw "Folder Not found"
		}
	}
}