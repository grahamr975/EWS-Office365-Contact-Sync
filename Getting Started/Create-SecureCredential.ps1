# interactively prompt for the credential
$cred = get-credential

# export the credential object to a file
$cred | Export-Clixml .\SecureCredential.cred