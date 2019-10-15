function New-EXCContactFolder
{
	<#
		.SYNOPSIS
			Creates a folder object with the specified path.
		
		.DESCRIPTION
			Creates a folder object with the specified path.
		
		.PARAMETER FolderPath
			The path to the folder, relative to the message folder base.
		
		.PARAMETER SmptAddress
			The email address of the mailbox to access
		
		.PARAMETER Service
			The established Service connection to use for this connection.
			Use 'Connect-EXCExchange' in order to establish a connection and obtain such an object.
		
		.EXAMPLE
			PS C:\> New-EXCContactFolder -FolderName "MyFolder" -SmptAddress 'peter@example.com' -Service $Service
	
			Returns the 'Private' folder within the contacts folder for the mailbox peter@example.com
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
		# Try to read from the specified folder
		$service = Connect-EXCExchange -MailboxName $MailboxName -Credential $Credential
		$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName);
		 try {
			Get-EXCContactFolder -SmptAddress $MailboxName -FolderPath "Contacts\$FolderName" -Service $Service
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
    }
}

function Get-GALContacts {
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
		# Connect to Office 365 Exchange Server using a Remote Session
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $Credentials -Authentication Basic -AllowRedirection
		Import-PSSession $Session -DisableNameChecking
		
			$DirectoryList = $(Get-Mailbox -ResultSize unlimited | Where-Object {$_.HiddenFromAddressListsEnabled -Match "False"})
			$EmailAddressList = $DirectoryList.PrimarySMTPAddress
		
			# Import Global Address List into Powershell from Office 365 Exchange as an array
			$GALContacts = Get-User -ResultSize unlimited | Where-Object {$null -ne $_.WindowsEmailAddress}
			$GALContacts = $GALContacts | Select-Object DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone | Where-Object {$EmailAddressList.Contains($_.WindowsEmailAddress)}
		Remove-PSSession $Session
		return $GALContacts
	}
}

Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
	$logfile,
	
	[Parameter(Mandatory=$False)]
    [string]
	$exception
	)
	
	if (!(Test-Path $logfile)) {
		New-Item -ItemType "file" -Path "$logfile" -Force
	  }

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message $exception"
    If($logfile) {
		Add-Content $logfile -Value $Line -Force
	}
    Else {
        Write-Output $Line
	}
	if ($Level -eq "FATAL") {
		Write-Error $Message
	}
}