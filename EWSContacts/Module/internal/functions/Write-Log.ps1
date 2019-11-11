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

if ($logfile -and !(Test-Path $logfile)) {
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