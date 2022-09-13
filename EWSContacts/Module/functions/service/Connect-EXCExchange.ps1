﻿function Connect-EXCExchange {
	<#
	.SYNOPSIS
		A brief description of the Connect-EXCExchange function.
	
	.DESCRIPTION
		A detailed description of the Connect-EXCExchange function.
	
	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.EXAMPLE
		PS C:\> Connect-EXCExchange -MailboxName 'value1' -Credentials $Credentials
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $False)]
		$Credentials,

		[Parameter(Position = 2, Mandatory = $False)]
		[switch]
		$ModernAuth,
		
		[Parameter(Position = 3, Mandatory = $False)]
		[String]
		$ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c",

		[Parameter(Position = 4, Mandatory = $False)]
		[String]
		$redirectUri= "urn:ietf:wg:oauth:2.0:oob",

		[Parameter(Position = 5, Mandatory = $False)]
		[String]
		$CertificateFilePath,
		
		[Parameter(Position = 6, Mandatory = $False)]
		[Security.SecureString]
		$CertificatePassword
	)
	Begin {
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		if (Test-Path ($script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll")) {
			Import-Module ($script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll")
			$Script:EWSDLL = $script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll"
			write-verbose ("Using EWS dll from Local Directory")
		}
		else {

			
			## Load Managed API dll  
			###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
			$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
			if (Test-Path $EWSDLL) {
				Import-Module $EWSDLL
				$Script:EWSDLL = $EWSDLL 
			}
			else {
				"$(get-date -format yyyyMMddHHmmss):"
				"This script requires the EWS Managed API 1.2 or later."
				"Please download and install the current version of the EWS Managed API from"
				"http://go.microsoft.com/fwlink/?LinkId=255472"
				""
				"Exiting Script."
				exit


			} 
		}
		
		# Force TLS 1.2
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
		
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
		
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		if ($ModernAuth.IsPresent) {
			Write-Verbose("Using Modern Auth")
			if ([String]::IsNullOrEmpty($ClientId)) {
				$ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
			}		
			Import-Module ($script:ModuleRoot + "/bin/Microsoft.Identity.Client.dll") -Force	
			if([string]::IsNullOrEmpty($CertificateFilePath)){
				if ($null -eq $Credentials) {				
					$scope = "https://outlook.office.com/EWS.AccessAsUser.All";
					$Scopes = New-Object System.Collections.Generic.List[string]
					$Scopes.Add($Scope)				
					$pcaConfig = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithRedirectUri($redirectUri)
					$token = $pcaConfig.Build().AcquireTokenInteractive($Scopes).WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount).WithLoginHint($MailboxName).ExecuteAsync().Result
					$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
				}else{
					$scope = "https://outlook.office.com/EWS.AccessAsUser.All";
					$Scopes = New-Object System.Collections.Generic.List[string]
					$Scopes.Add($Scope)				
					$pcaConfig = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithAuthority([Microsoft.Identity.Client.AadAuthorityAudience]::AzureAdMultipleOrgs)				
					$token = $pcaConfig.Build().AcquireTokenByUsernamePassword($Scopes,$Credentials.UserName,$Credentials.Password).ExecuteAsync().Result;
					$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
				}
			}
			else{
				$exVal = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
				$certificateObject = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertificateFilePath, $CertificatePassword , $exVal
				$domain = $MailboxName.Split("@")[1]
				$Scope = "https://outlook.office365.com/.default"
				$TenantId = (Invoke-WebRequest https://login.windows.net/$domain/v2.0/.well-known/openid-configuration | ConvertFrom-Json).token_endpoint.Split('/')[3]
				$app =  [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ClientId).WithCertificate($certificateObject).WithTenantId($TenantId).Build()
				$Scopes = New-Object System.Collections.Generic.List[string]
				$Scopes.Add($Scope)
				$token = $app.AcquireTokenForClient($Scopes).ExecuteAsync().Result
				$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
				$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)			
			}
		}
		else {
			Write-Verbose("Using Negotiate Auth")
			if(!$Credentials){$Credentials = Get-Credential}
			$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())
			$service.Credentials = $creds
		}

		#Credentials Option 2  
		#service.UseDefaultCredentials = $true  
		#$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		
		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
		$Compiler = $Provider.CreateCompiler()
		$Params = New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable = $False
		$Params.GenerateInMemory = $True
		$Params.IncludeDebugInformation = $False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
		
		$TASource = @'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@
		$TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
		$TAAssembly = $TAResults.CompiledAssembly
		
		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
		
		## end code from http://poshcode.org/624
		
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		
		#CAS URL Option 1 Autodiscover  
		if([string]::IsNullOrEmpty($CertificateFilePath)){
			$service.AutodiscoverUrl($MailboxName, { $true })
		}else{
			$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx" 
			$service.Url = $uri 
		}
		#Write-host ("Using CAS Server : " + $Service.url)
		
		#CAS URL Option 2 Hardcoded  
		
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
		#$service.Url = $uri    
		
		## Optional section for Exchange Impersonation  
		
		$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		if (!$service.URL) {
			throw "Error connecting to EWS"
		}
		else {
			return $service
		}
	}
}
