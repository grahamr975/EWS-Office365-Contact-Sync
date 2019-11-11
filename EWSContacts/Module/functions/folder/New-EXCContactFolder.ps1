function New-EXCContactFolder
{
	<#
		.SYNOPSIS
			Creates a folder object with the specified path if the folder does not already exist.

			*Requries the EWCContacts module: https://github.com/gscales/Powershell-Scripts/tree/master/EWSContacts
		
		.PARAMETER MailboxName
			The email address of the mailbox that the contact foloder will be created in
		
		.PARAMETER FolderName
			Name of the contact folder that will be created
		
		.PARAMETER Service
			Office 365 Credentials (Admin account with application impersonatoin permissions)
		
		.EXAMPLE
			PS C:\> New-EXCContactFolder -MailboxName 'john.doe@example.com' -MailboxName "MyFolder" -Credential $Credentials
		#>
	[CmdletBinding()]
	param (
		[Parameter()]
		[string]
		$MailboxName,

		[Parameter()]
		[string]
		$FolderName,
		
		[Parameter()]
		[System.Management.Automation.PSCredential]
		$Credential
	)
	process
	{
		try {
			# Try to read from the specified folder
			$service = Connect-EXCExchange -MailboxName $MailboxName -Credential $Credential
			$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName);
			try {
				Get-EXCContactFolder -SmptAddress $MailboxName -FolderPath "Contacts\$FolderName" -Service $Service | Out-Null
			} catch {
			# # If the read fails, create the folder.
				$ContactsFolder = New-Object Microsoft.Exchange.WebServices.Data.ContactsFolder($service);
				$ContactsFolder.DisplayName = $FolderName
				$ContactsFolder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
				$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
				$RootFolder.Load()
				$FolderView = new-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
				$ContactsFolderSearch = $RootFolder.FindFolders($FolderView) | Where-Object {$_.DisplayName -eq $FolderName}
				$ContactsFolder = [Microsoft.Exchange.WebServices.Data.ContactsFolder]::Bind($service,$ContactsFolderSearch.Id);
			}
		} catch {
			Write-Log -Level "FATAL" -Message "Failed verify that $($FolderName) exists for $($Mailbox)" -exception $_.Exception.Message
		}
    }
}