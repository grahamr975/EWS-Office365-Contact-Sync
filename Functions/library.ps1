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
    }
}


function Get-GALContacts {
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
		<#
		.SYNOPSIS
			Writes the specified message to the screen. If $logfile is specified, write to the file instead.
		
		.PARAMETER Level
			Indicates the type of message. The options are INFO, WARN, ERROR, FATAL, and DEBUG. Uses INFO by default.
		
		.PARAMETER Message
			Message to write to the screen or the $logfile (if specified)

		.PARAMETER logfile
			Optional; File path of the file to write the message to.

		.PARAMETER exception
			Optional; Additional information to include when the script encounters an error. Very useful for troubleshooting.
		
		.EXAMPLE
			PS C:\> Write-Log -Level "INFO" -Message "Successfully executed the script" -logfile $LogFilePath
	#>
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
		New-Item -ItemType "file" -Path "$logfile" -Force | Out-Null
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