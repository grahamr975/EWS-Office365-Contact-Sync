# interactively prompt for the credential
$SecureString = Read-Host -AsSecureString

# export the credential object to a file
$SecureString | Export-Clixml .\SecureCertificatePassword.xml