<#
.SYNOPSIS
Synchronizes Microsoft Entra directory users to dedicated Outlook contact folders by using Microsoft Graph.

.DESCRIPTION
Uses app-only certificate authentication.  State is retained locally so subsequent
runs use /users/delta and write only contacts whose normalized values changed.
Required application permissions: Contacts.ReadWrite and User.Read.All.
OrgContact.Read.All is required with -IncludeNonUserContacts.
#>
[CmdletBinding()]
param (
    # Microsoft Entra directory ID (usually a GUID). It identifies the tenant to sign in to.
    [Parameter(Mandatory)] [string] $TenantId,
    # Application (client) ID of the Entra app registration.
    [Parameter(Mandatory)] [string] $ClientId,
    # Local PFX file containing the certificate's private key.
    [Parameter(Mandatory)] [System.IO.FileInfo] $CertificatePath,
    # CliXml file that contains the PFX password as a Windows secure string.
    [Parameter(Mandatory)] [System.IO.FileInfo] $CertificatePasswordPath,
    # Dedicated Outlook contact folder that this script is allowed to manage.
    [Parameter(Mandatory)] [string] $FolderName,
    # Category name used to identify duplicate managed contacts outside the
    # dedicated folder. When omitted, the contact-folder name is used.
    # One or more target mailbox email addresses or Entra user IDs.
    [string[]] $MailboxList,
    # Optional CSV alternative. It needs a Mailbox or UserPrincipalName column.
    [System.IO.FileInfo] $MailboxCsvPath,
    # Optional directory for timestamped text logs.
    [string] $LogPath,
    # Local SQLite database created by Initialize-GraphContactSyncDatabase.ps1.
    [string] $DatabasePath = 'C:\ContactSync\GraphContactSync.db',
    # Force a complete directory read and mailbox-cache reconciliation after
    # this many days. Normal runs use Graph delta and SQLite mappings.
    [ValidateRange(1, 3650)] [int] $FullDirectoryRefreshDays = 30,
    # Maximum Graph operations in one JSON batch. Graph permits at most 20.
    [ValidateRange(1, 20)] [int] $BatchSize = 20,
    # Number of times to retry a throttled or temporary Graph error.
    [ValidateRange(0, 5)] [int] $MaxBatchRetries = 3,
    # Optional source-contact filters.
    [switch] $ExcludeContactsWithoutPhoneNumber,
    [switch] $ExcludeSharedMailboxContacts,
    [switch] $IncludeNonUserContacts
)

# Strict mode turns programming mistakes, such as a misspelled variable, into errors.
Set-StrictMode -Version Latest
# Stop immediately for unexpected PowerShell errors; mailbox-level errors are handled later.
$ErrorActionPreference = 'Stop'
# A caller must choose exactly one way to identify target mailboxes.
if ((@($MailboxList).Count -eq 0) -and (-not $MailboxCsvPath)) { throw 'Specify -MailboxList or -MailboxCsvPath.' }
if ((@($MailboxList).Count -gt 0) -and $MailboxCsvPath) { throw 'Use either -MailboxList or -MailboxCsvPath, not both.' }
$directoryMode = (@($MailboxList).Count -eq 1 -and $MailboxList[0].ToUpperInvariant() -eq 'DIRECTORY')

function Write-SyncLog {
    # Write the same message to the console and, when configured, to a log file.
    param([string] $Level = 'INFO', [Parameter(Mandatory)] [string] $Message)
    $line = '{0:u} [{1}] {2}' -f (Get-Date), $Level, $Message
    Write-Host $line
    $logFileVariable = Get-Variable -Scope Script -Name LogFile -ErrorAction SilentlyContinue
    if ($logFileVariable -and $logFileVariable.Value) { Add-Content -LiteralPath $logFileVariable.Value -Value $line }
}

function Get-OptionalProperty {
    # Graph often omits empty fields instead of returning a null value. This helper
    # safely reads a property without strict mode raising an error when it is absent.
    param($Object, [Parameter(Mandatory)] [string] $Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    $property.Value
}

function Format-GraphRequestError {
    # Invoke-MgGraphRequest's exception message often contains only the HTTP
    # reason phrase. Recover the JSON error details when the SDK supplies them
    # so logs retain Graph's useful code and request identifiers.
    param([Parameter(Mandatory)] $ErrorRecord)
    $exception = Get-OptionalProperty -Object $ErrorRecord -Name 'Exception'
    $message = Get-OptionalProperty -Object $exception -Name 'Message'
    if (-not $message) { $message = [string]$ErrorRecord }
    $response = Get-OptionalProperty -Object $exception -Name 'Response'
    $status = Get-OptionalProperty -Object $response -Name 'StatusCode'
    $errorDetails = Get-OptionalProperty -Object $ErrorRecord -Name 'ErrorDetails'
    $rawDetails = Get-OptionalProperty -Object $errorDetails -Name 'Message'
    $graphError = $null
    if ($rawDetails) {
        try {
            $detailObject = $rawDetails | ConvertFrom-Json -ErrorAction Stop
            $graphError = Get-OptionalProperty -Object $detailObject -Name 'error'
        } catch {
            # Some SDK versions return plain text rather than a Graph JSON body.
            if ($rawDetails -ne $message) { $message = "$message $rawDetails" }
        }
    }
    $innerError = Get-OptionalProperty -Object $graphError -Name 'innerError'
    $graphMessage = Get-OptionalProperty -Object $graphError -Name 'message'
    if ($graphMessage) { $message = $graphMessage }
    $diagnostics = @()
    $errorCode = Get-OptionalProperty -Object $graphError -Name 'code'
    $innerErrorCode = Get-OptionalProperty -Object $innerError -Name 'code'
    $requestId = Get-OptionalProperty -Object $innerError -Name 'request-id'
    $clientRequestId = Get-OptionalProperty -Object $innerError -Name 'client-request-id'
    $errorDate = Get-OptionalProperty -Object $innerError -Name 'date'
    if ($errorCode) { $diagnostics += "code=$errorCode" }
    if ($innerErrorCode -and $innerErrorCode -ne $errorCode) { $diagnostics += "inner-code=$innerErrorCode" }
    if ($requestId) { $diagnostics += "request-id=$requestId" }
    if ($clientRequestId) { $diagnostics += "client-request-id=$clientRequestId" }
    if ($errorDate) { $diagnostics += "date=$errorDate" }
    $statusText = if ($status) { "HTTP $([int]$status) $status" } else { $null }
    $diagnosticText = if ($diagnostics.Count -gt 0) { "[$($diagnostics -join '; ')]" } else { $null }
    (@($statusText, $diagnosticText, $message) | Where-Object { $_ }) -join ': '
}

function Get-GraphPages {
    # Graph lists can span multiple pages. Keep following @odata.nextLink until
    # Graph says there are no more pages, then return one combined list.
    param([Parameter(Mandatory)] [string] $Uri, [switch] $ImmutableIds)
    $items = @()
    do {
        $request = @{ Method = 'GET'; Uri = $Uri; OutputType = 'PSObject' }
        # Outlook item IDs normally change when an item is moved. Ask Graph for
        # immutable IDs on every page when the caller intends to cache item IDs.
        if ($ImmutableIds) { $request.Headers = @{ Prefer = 'IdType="ImmutableId"' } }
        $response = Invoke-MgGraphRequest @request
        if ($null -ne $response.value) { $items += @($response.value) }
        $Uri = Get-OptionalProperty -Object $response -Name '@odata.nextLink'
    } while ($Uri)
    $items
}

function Get-BetaGraphPages {
    param([Parameter(Mandatory)] [string] $Uri, [switch] $ImmutableIds)
    $items = @()
    do {
        $request = @{ Method = 'GET'; Uri = $Uri; OutputType = 'PSObject' }
        if ($ImmutableIds) { $request.Headers = @{ Prefer = 'IdType="ImmutableId"' } }
        $response = Invoke-MgGraphRequest @request
        if ($null -ne $response.value) { $items += @($response.value) }
        $Uri = Get-OptionalProperty -Object $response -Name '@odata.nextLink'
    } while ($Uri)
    $items
}

function ConvertTo-GraphPath {
    # Email addresses contain @ and other characters that must be URL encoded.
    param([Parameter(Mandatory)] [string] $Id)
    [uri]::EscapeDataString($Id)
}

function Get-StringValue {
    # Convert a possibly empty value to a safe string for fingerprint comparisons.
    param($Value)
    if ($null -eq $Value) { return '' }
    [string]$Value
}

function Get-ContactFingerprint {
    # A fingerprint is a SHA-256 hash of the fields this script synchronizes.
    # It lets later runs skip PATCH requests when nothing meaningful changed.
    param([Parameter(Mandatory)] $Contact)
    # Normalize whitespace, email case, and phone-number order before hashing.
    $canonical = @(
        (Get-StringValue $Contact.DisplayName).Trim(), (Get-StringValue $Contact.FirstName).Trim(),
        (Get-StringValue $Contact.LastName).Trim(), (Get-StringValue $Contact.Email).Trim().ToLowerInvariant(),
        (Get-StringValue $Contact.JobTitle).Trim(), (Get-StringValue $Contact.CompanyName).Trim(),
        (Get-StringValue $Contact.Department).Trim(), (@($Contact.BusinessPhones | ForEach-Object { (Get-StringValue $_).Trim() } | Sort-Object) -join '|'),
        (Get-StringValue $Contact.MobilePhone).Trim()
    ) -join "`n"
    # Hash the normalized text and return the hash as a readable hexadecimal string.
    $bytes = [Text.Encoding]::UTF8.GetBytes($canonical)
    ([Security.Cryptography.SHA256]::Create().ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join ''
}

function New-ContactModel {
    # Graph user, orgContact, and Outlook contact objects use slightly different
    # phone fields. Convert them into one common shape for the rest of the script.
    param([Parameter(Mandatory)] $Object, [string] $SourceId, [switch] $OrganizationContact)
    if ($OrganizationContact) {
        # Entra organizational contacts keep all phones in a typed phones array.
        $phones = @(Get-OptionalProperty -Object $Object -Name 'phones')
        $business = @($phones | Where-Object { $_.type -match 'business' } | ForEach-Object { $_.number })
        $mobile = ($phones | Where-Object { $_.type -eq 'mobile' } | Select-Object -First 1).number
    } else {
        # Entra users expose business phones and mobile phone as separate fields.
        $business = @(Get-OptionalProperty -Object $Object -Name 'businessPhones')
        $mobile = Get-OptionalProperty -Object $Object -Name 'mobilePhone'
    }
    # Keep only the values that this sync manages; personal Outlook fields are untouched.
    $contact = [pscustomobject]@{
        SourceId = $SourceId; DisplayName = Get-OptionalProperty $Object 'displayName'; FirstName = Get-OptionalProperty $Object 'givenName'; LastName = Get-OptionalProperty $Object 'surname'
        Email = Get-OptionalProperty $Object 'mail'; JobTitle = Get-OptionalProperty $Object 'jobTitle'; CompanyName = Get-OptionalProperty $Object 'companyName'; Department = Get-OptionalProperty $Object 'department'
        BusinessPhones = @($business); MobilePhone = $mobile
    }
    # Save the computed hash on the model so callers do not recalculate it.
    $contact | Add-Member -NotePropertyName Fingerprint -NotePropertyValue (Get-ContactFingerprint $contact)
    $contact
}

function Test-EligibleUser {
    # Return True only when a directory user should become a synchronized contact.
    param($User)
    # Exclude disabled users, guests, and objects that do not have an email address.
    if ((Get-OptionalProperty $User 'accountEnabled') -ne $true -or (Get-OptionalProperty $User 'userType') -ne 'Member' -or [string]::IsNullOrWhiteSpace((Get-OptionalProperty $User 'mail'))) { return $false }
    if ($ExcludeSharedMailboxContacts -and @(Get-OptionalProperty $User 'assignedLicenses').Count -eq 0) { return $false }
    $model = New-ContactModel -Object $User -SourceId (Get-OptionalProperty $User 'id')
    if ($ExcludeContactsWithoutPhoneNumber -and @($model.BusinessPhones).Count -eq 0 -and [string]::IsNullOrWhiteSpace($model.MobilePhone)) { return $false }
    $true
}

function Get-FilterSignature {
    # Save the selected source filters in state. If they change between runs, the
    # cached source list is no longer valid and must be rebuilt.
    @(
        "ExcludeContactsWithoutPhoneNumber=$($ExcludeContactsWithoutPhoneNumber.IsPresent)",
        "ExcludeSharedMailboxContacts=$($ExcludeSharedMailboxContacts.IsPresent)",
        "IncludeNonUserContacts=$($IncludeNonUserContacts.IsPresent)"
    ) -join ';'
}

function New-SqliteWriteCommand {
    # Build a parameterized command once so row loops only replace values and
    # execute it. The transaction makes the complete logical save atomic.
    param($Connection, $Transaction, [Parameter(Mandatory)] [string] $Query, [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]] $ParameterNames)
    $command = $Connection.CreateCommand()
    $command.CommandText = $Query
    $command.Transaction = $Transaction
    foreach ($name in $ParameterNames) { [void] $command.Parameters.AddWithValue("@$name", [DBNull]::Value) }
    $command.Prepare()
    $command
}

function Invoke-SqliteWriteCommand {
    # Reuse a prepared command with a new parameter set for one row.
    param($Command, [Parameter(Mandatory)] [hashtable] $Parameters)
    foreach ($parameter in $Command.Parameters) {
        $name = $parameter.ParameterName.TrimStart('@')
        $value = $Parameters[$name]
        $parameter.Value = if ($null -eq $value) { [DBNull]::Value } else { $value }
    }
    [void] $Command.ExecuteNonQuery()
}

function Invoke-SqliteWriteTransaction {
    # Opening one connection and committing once avoids a separate connection
    # and durable SQLite commit for every row written by a save operation.
    param([Parameter(Mandatory)] [scriptblock] $Action)
    $connection = $null
    $transaction = $null
    try {
        $connection = New-SQLiteConnection -DataSource $DatabasePath
        $transaction = $connection.BeginTransaction()
        & $Action $connection $transaction
        $transaction.Commit()
    } catch {
        $transactionError = $_
        if ($null -ne $transaction) {
            try { $transaction.Rollback() } catch { }
        }
        throw $transactionError
    } finally {
        if ($null -ne $transaction) { $transaction.Dispose() }
        if ($null -ne $connection) { $connection.Dispose() }
    }
}

function Get-State {
    # Read the small sync-wide checkpoint and source contacts from SQLite.
    $metadata = @{}
    foreach ($row in Invoke-SqliteQuery -DataSource $DatabasePath -Query 'SELECT MetadataKey, MetadataValue FROM SyncMetadata') { $metadata[$row.MetadataKey] = $row.MetadataValue }
    $source = @(Invoke-SqliteQuery -DataSource $DatabasePath -Query 'SELECT ContactJson FROM SourceContact' | ForEach-Object { $_.ContactJson | ConvertFrom-Json })
    # Changes is empty after a database read. Get-SourceState adds only the
    # users returned by this run's Graph delta feed.
    [pscustomobject]@{ FilterSignature = $metadata['FilterSignature']; UserDeltaLink = $metadata['UserDeltaLink']; LastFullDirectoryRefreshUtc = $metadata['LastFullDirectoryRefreshUtc']; SourceContacts = $source; DesiredContacts = @(); SourceChanges = @(); RebuildSourceCache = $false; RebuildMailboxCache = $false }
}

function Save-State { param($State)
    $now = (Get-Date).ToUniversalTime().ToString('o')
    # Record the successful full refresh only after every target mailbox has
    # completed. This prevents a failed run from postponing the next refresh.
    if ($State.RebuildSourceCache) { $State.LastFullDirectoryRefreshUtc = $now }
    Invoke-SqliteWriteTransaction {
        param($connection, $transaction)
        $commands = @()
        try {
            $sourceInsert = New-SqliteWriteCommand $connection $transaction 'INSERT INTO SourceContact (SourceId,Email,Fingerprint,ContactJson,UpdatedUtc) VALUES (@id,@email,@fingerprint,@json,@now)' @('id','email','fingerprint','json','now'); $commands += $sourceInsert
            $sourceInsertIgnore = New-SqliteWriteCommand $connection $transaction 'INSERT OR IGNORE INTO SourceContact (SourceId,Email,Fingerprint,ContactJson,UpdatedUtc) VALUES (@id,@email,@fingerprint,@json,@now)' @('id','email','fingerprint','json','now'); $commands += $sourceInsertIgnore
            $sourceUpdate = New-SqliteWriteCommand $connection $transaction 'UPDATE SourceContact SET Email=@email,Fingerprint=@fingerprint,ContactJson=@json,UpdatedUtc=@now WHERE SourceId=@id' @('id','email','fingerprint','json','now'); $commands += $sourceUpdate
            $sourceDeleteAll = New-SqliteWriteCommand $connection $transaction 'DELETE FROM SourceContact' @(); $commands += $sourceDeleteAll
            $sourceDelete = New-SqliteWriteCommand $connection $transaction 'DELETE FROM SourceContact WHERE SourceId=@id' @('id'); $commands += $sourceDelete
            $metadataInsert = New-SqliteWriteCommand $connection $transaction 'INSERT OR IGNORE INTO SyncMetadata (MetadataKey,MetadataValue,UpdatedUtc) VALUES (@key,@value,@now)' @('key','value','now'); $commands += $metadataInsert
            $metadataUpdate = New-SqliteWriteCommand $connection $transaction 'UPDATE SyncMetadata SET MetadataValue=@value,UpdatedUtc=@now WHERE MetadataKey=@key' @('key','value','now'); $commands += $metadataUpdate

            # A new database, expired delta token, or changed filters requires a
            # full rebuild. Normal delta runs write only the changed rows.
            if ($State.RebuildSourceCache) {
                Invoke-SqliteWriteCommand $sourceDeleteAll @{}
                foreach ($contact in @($State.SourceContacts)) {
                    Invoke-SqliteWriteCommand $sourceInsert @{ id=$contact.SourceId; email=$contact.Email; fingerprint=$contact.Fingerprint; json=($contact | ConvertTo-Json -Depth 8 -Compress); now=$now }
                }
            } elseif (@($State.SourceChanges).Count -gt 0) {
                foreach ($change in @($State.SourceChanges)) {
                    if ($change.Action -eq 'Delete') {
                        Invoke-SqliteWriteCommand $sourceDelete @{ id=$change.SourceId }
                    } else {
                        $contact = $change.Contact
                        $parameters = @{ id=$contact.SourceId; email=$contact.Email; fingerprint=$contact.Fingerprint; json=($contact | ConvertTo-Json -Depth 8 -Compress); now=$now }
                        # Retain compatibility with SQLite versions before UPSERT.
                        Invoke-SqliteWriteCommand $sourceInsertIgnore $parameters
                        Invoke-SqliteWriteCommand $sourceUpdate $parameters
                    }
                }
            } else {
                Write-SyncLog -Message 'Directory cache is unchanged; skipped source-cache database rewrite.'
            }

            # Always retain the newest Graph delta token, even when no users changed.
            foreach ($pair in @{ FilterSignature=$State.FilterSignature; UserDeltaLink=$State.UserDeltaLink; LastFullDirectoryRefreshUtc=$State.LastFullDirectoryRefreshUtc }.GetEnumerator()) {
                $parameters = @{ key=$pair.Key; value=$pair.Value; now=$now }
                Invoke-SqliteWriteCommand $metadataInsert $parameters
                Invoke-SqliteWriteCommand $metadataUpdate $parameters
            }
        } finally {
            foreach ($command in $commands) { $command.Dispose() }
        }
    }
}

function Get-SourceState {
    # Use the saved Graph delta link to retrieve only directory users changed
    # since the last successful run. The first run has no delta link and gets all users.
    param($State)
    # With no previous delta token, Graph will return the full directory. Write
    # that first source list as a complete set instead of row-by-row updates.
    if (-not $State.UserDeltaLink) { $State.RebuildSourceCache = $true }
    $byId = @{}
    # Build a fast lookup of the previously known source contacts by Entra object ID.
    foreach ($contact in @($State.SourceContacts)) { $byId[$contact.SourceId] = $contact }
    $properties = 'id,displayName,givenName,surname,jobTitle,companyName,department,mail,businessPhones,mobilePhone,accountEnabled,assignedLicenses,userType'
    $uri = if ($State.UserDeltaLink) { $State.UserDeltaLink } else { "/v1.0/users/delta?`$select=$properties" }
    $deltaLink = $null
    try {
        do {
            $page = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
            # Each delta page contains changed users and may contain a deletion marker.
            foreach ($user in @($page.value)) {
                $sourceId = Get-OptionalProperty -Object $user -Name 'id'
                if ((Get-OptionalProperty -Object $user -Name '@removed') -or -not (Test-EligibleUser $user)) {
                    $byId.Remove($sourceId) | Out-Null
                    if (-not $State.RebuildSourceCache) { $State.SourceChanges += [pscustomobject]@{ Action='Delete'; SourceId=$sourceId; Contact=$null } }
                } else {
                    $contact = New-ContactModel -Object $user -SourceId $sourceId
                    $byId[$sourceId] = $contact
                    if (-not $State.RebuildSourceCache) { $State.SourceChanges += [pscustomobject]@{ Action='Upsert'; SourceId=$sourceId; Contact=$contact } }
                }
            }
            # Graph returns nextLink until the current delta run is complete; the
            # final page returns deltaLink, which becomes the next run's checkpoint.
            $uri = Get-OptionalProperty -Object $page -Name '@odata.nextLink'
            $deltaLink = Get-OptionalProperty -Object $page -Name '@odata.deltaLink'
        } while ($uri)
    } catch {
        if ($State.UserDeltaLink) {
            # Delta links expire occasionally. Clearing it safely triggers one full rebuild.
            Write-SyncLog -Level WARN -Message 'Saved user delta token was rejected; rebuilding the directory cache.'
            $State.UserDeltaLink = $null; $State.SourceContacts = @(); $State.SourceChanges = @(); $State.RebuildSourceCache = $true; $State.RebuildMailboxCache = $true; return Get-SourceState $State
        }
        throw
    }
    if ($IncludeNonUserContacts) {
        # orgContact delta support differs by tenant; rebuild this optional source safely.
        $State.RebuildSourceCache = $true
        foreach ($key in @($byId.Keys | Where-Object { $_ -like 'org:*' })) { $byId.Remove($key) | Out-Null }
        $orgContacts = Get-GraphPages -Uri '/v1.0/contacts?$select=id,displayName,givenName,surname,jobTitle,companyName,department,mail,phones'
        foreach ($org in $orgContacts) {
            if ($org.mail) {
                $model = New-ContactModel -Object $org -SourceId "org:$($org.id)" -OrganizationContact
                if (-not $ExcludeContactsWithoutPhoneNumber -or @($model.BusinessPhones).Count -gt 0 -or $model.MobilePhone) { $byId[$model.SourceId] = $model }
            }
        }
    }
    $State.SourceContacts = @($byId.Values)
    $State.UserDeltaLink = $deltaLink
    # Deduplicate contacts by email address. If two directory objects share one
    # email address, the alphabetically first source ID wins consistently.
    $desired = @{}
    foreach ($contact in $State.SourceContacts | Sort-Object SourceId) {
        if (-not $desired.ContainsKey($contact.Email.ToLowerInvariant())) { $desired[$contact.Email.ToLowerInvariant()] = $contact }
    }
    $State.DesiredContacts = @($desired.Values)
    $State
}

function Get-MailboxState {
    # Return the locally cached information for one mailbox, if it has been synced before.
    param($State, [string] $MailboxId)
    $mailbox = Invoke-SqliteQuery -DataSource $DatabasePath -Query 'SELECT MailboxId,FolderId FROM Mailbox WHERE MailboxId=@id' -SqlParameters @{ id=$MailboxId } | Select-Object -First 1
    if (-not $mailbox) { return $null }
    $contacts = @(Invoke-SqliteQuery -DataSource $DatabasePath -Query 'SELECT Email,ContactId,Fingerprint FROM MailboxContact WHERE MailboxId=@id' -SqlParameters @{ id=$MailboxId })
    [pscustomobject]@{ MailboxId=$mailbox.MailboxId; FolderId=$mailbox.FolderId; Contacts=$contacts }
}

function Save-MailboxState {
    # Save a complete mapping only after the first scan of a mailbox folder.
    # Later syncs use Save-MailboxChanges so they change only affected rows.
    param($MailboxState)
    $now = (Get-Date).ToUniversalTime().ToString('o')
    Invoke-SqliteWriteTransaction {
        param($connection, $transaction)
        $commands = @()
        try {
            $mailboxInsert = New-SqliteWriteCommand $connection $transaction 'INSERT OR IGNORE INTO Mailbox (MailboxId,FolderId,UpdatedUtc) VALUES (@id,@folder,@now)' @('id','folder','now'); $commands += $mailboxInsert
            $mailboxUpdate = New-SqliteWriteCommand $connection $transaction 'UPDATE Mailbox SET FolderId=@folder,UpdatedUtc=@now WHERE MailboxId=@id' @('id','folder','now'); $commands += $mailboxUpdate
            $contactDelete = New-SqliteWriteCommand $connection $transaction 'DELETE FROM MailboxContact WHERE MailboxId=@id' @('id'); $commands += $contactDelete
            $contactInsert = New-SqliteWriteCommand $connection $transaction 'INSERT INTO MailboxContact (MailboxId,Email,ContactId,Fingerprint,UpdatedUtc) VALUES (@mailbox,@email,@contact,@fingerprint,@now)' @('mailbox','email','contact','fingerprint','now'); $commands += $contactInsert

            $mailboxParameters = @{ id=$MailboxState.MailboxId; folder=$MailboxState.FolderId; now=$now }
            Invoke-SqliteWriteCommand $mailboxInsert $mailboxParameters
            Invoke-SqliteWriteCommand $mailboxUpdate $mailboxParameters
            Invoke-SqliteWriteCommand $contactDelete @{ id=$MailboxState.MailboxId }
            foreach ($contact in @($MailboxState.Contacts)) {
                Invoke-SqliteWriteCommand $contactInsert @{ mailbox=$MailboxState.MailboxId; email=$contact.Email; contact=$contact.ContactId; fingerprint=$contact.Fingerprint; now=$now }
            }
        } finally {
            foreach ($command in $commands) { $command.Dispose() }
        }
    }
}

function Save-MailboxChanges {
    # Persist only the contact mappings that Graph successfully changed. This
    # avoids deleting and rebuilding thousands of rows for a one-contact update.
    param($MailboxState, [object[]] $Completed)
    $now = (Get-Date).ToUniversalTime().ToString('o')
    Invoke-SqliteWriteTransaction {
        param($connection, $transaction)
        $commands = @()
        try {
            $mailboxUpdate = New-SqliteWriteCommand $connection $transaction 'UPDATE Mailbox SET FolderId=@folder,UpdatedUtc=@now WHERE MailboxId=@id' @('id','folder','now'); $commands += $mailboxUpdate
            $contactDelete = New-SqliteWriteCommand $connection $transaction 'DELETE FROM MailboxContact WHERE MailboxId=@mailbox AND Email=@email' @('mailbox','email'); $commands += $contactDelete
            $contactInsert = New-SqliteWriteCommand $connection $transaction 'INSERT OR REPLACE INTO MailboxContact (MailboxId,Email,ContactId,Fingerprint,UpdatedUtc) VALUES (@mailbox,@email,@contact,@fingerprint,@now)' @('mailbox','email','contact','fingerprint','now'); $commands += $contactInsert
            $contactUpdate = New-SqliteWriteCommand $connection $transaction 'UPDATE MailboxContact SET Fingerprint=@fingerprint,UpdatedUtc=@now WHERE MailboxId=@mailbox AND Email=@email' @('mailbox','email','fingerprint','now'); $commands += $contactUpdate

            # Keep the mailbox timestamp meaningful without rewriting its map.
            Invoke-SqliteWriteCommand $mailboxUpdate @{ id=$MailboxState.MailboxId; folder=$MailboxState.FolderId; now=$now }
            foreach ($result in @($Completed)) {
                $operation = $result.Operation
                if ($operation.Action -eq 'Delete') {
                    Invoke-SqliteWriteCommand $contactDelete @{ mailbox=$MailboxState.MailboxId; email=$operation.Email }
                } elseif ($operation.Action -eq 'Create') {
                    Invoke-SqliteWriteCommand $contactInsert @{ mailbox=$MailboxState.MailboxId; email=$operation.Email; contact=$result.Response.body.id; fingerprint=$operation.Source.Fingerprint; now=$now }
                } elseif ($operation.Action -eq 'Update') {
                    Invoke-SqliteWriteCommand $contactUpdate @{ mailbox=$MailboxState.MailboxId; email=$operation.Email; fingerprint=$operation.Source.Fingerprint; now=$now }
                }
            }
        } finally {
            foreach ($command in $commands) { $command.Dispose() }
        }
    }
}

function New-GraphContactBody {
    # Convert the common internal model into the JSON object expected by Graph's
    # create-contact and update-contact endpoints.
    param($Contact)
    $body = @{ givenName = $Contact.FirstName; surname = $Contact.LastName; displayName = $Contact.DisplayName; fileAs = $Contact.DisplayName
        jobTitle = $Contact.JobTitle; companyName = $Contact.CompanyName; department = $Contact.Department
        emailAddresses = @(@{ address = $Contact.Email; name = $Contact.DisplayName }) }
    if (@($Contact.BusinessPhones).Count -gt 0) { $body.businessPhones = @($Contact.BusinessPhones) }
    if (-not [string]::IsNullOrWhiteSpace($Contact.MobilePhone)) { $body.mobilePhone = $Contact.MobilePhone }
    $body
}

function Get-OrCreateFolder {
    # Reuse the saved folder ID when available. On the first run, find the folder
    # by name, or create it if it does not yet exist in the mailbox.
    param([string] $MailboxId, $MailboxState, [int] $DesiredContactCount)
    $user = ConvertTo-GraphPath $MailboxId
    if ($MailboxState -and $MailboxState.FolderId) {
        try {
            Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$user/contactFolders/$($MailboxState.FolderId)?`$select=id" -OutputType PSObject | Out-Null
            return $MailboxState.FolderId
        } catch {
            Write-SyncLog -Level WARN -Message "Saved managed folder ID for $MailboxId no longer exists; locating '$FolderName' again."
        }
    }
    $candidates = @(Get-BetaGraphPages -Uri "https://graph.microsoft.com/beta/users/$user/contactFolders?`$select=id,displayName" | Where-Object { $_.displayName -eq $FolderName })
    $folder = $null
    if ($candidates.Count -gt 0) {
        $folder = $candidates | ForEach-Object {
            $count = @(Get-BetaGraphPages -Uri "https://graph.microsoft.com/beta/users/$user/contactFolders/$($_.id)/contacts?`$select=id" -ImmutableIds).Count
            [pscustomobject]@{ Folder=$_; Count=$count; Difference=[Math]::Abs($count - $DesiredContactCount) }
        } | Sort-Object Difference, Count | Select-Object -First 1
        Write-SyncLog -Message "Selected '$FolderName' folder $($folder.Folder.id) with $($folder.Count) contact(s), closest to the $DesiredContactCount directory contact(s)."
        $folder = $folder.Folder
    }
    if (-not $folder) { $folder = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/users/$user/contactFolders" -Body (@{ displayName = $FolderName } | ConvertTo-Json) -ContentType 'application/json' -OutputType PSObject }
    $folder.id
}

function Initialize-MailboxState {
    # This runs only for a mailbox without local state. It reads the existing
    # managed folder once and stores each Outlook contact ID and fingerprint.
    param([string] $MailboxId, [string] $FolderId)
    $user = ConvertTo-GraphPath $MailboxId
    $select = 'id,displayName,givenName,surname,jobTitle,companyName,department,emailAddresses,businessPhones,mobilePhone'
    $existing = Get-GraphPages -Uri "/v1.0/users/$user/contactFolders/$FolderId/contacts?`$select=$select" -ImmutableIds
    # SQLite maps a source email to one mailbox-specific contact ID. A user can
    # manually create duplicate Outlook contacts with the same email, so retain
    # one deterministic contact in the map and collect the extra copies for
    # removal during this managed-folder reconciliation.
    $contactsByEmail = @{}
    $duplicateContactIds = @()
    foreach ($item in @($existing | Sort-Object { [string](Get-OptionalProperty -Object $_ -Name 'id') })) {
        # Use the first email address as the key because the sync uses email as
        # its matching identity between the directory and Outlook contacts.
        $emailAddress = @(Get-OptionalProperty -Object $item -Name 'emailAddresses') | Select-Object -First 1
        $email = Get-OptionalProperty -Object $emailAddress -Name 'address'
        if ($email) {
            $email = [string]$email
            $model = [pscustomobject]@{ DisplayName = Get-OptionalProperty $item 'displayName'; FirstName = Get-OptionalProperty $item 'givenName'; LastName = Get-OptionalProperty $item 'surname'; Email = $email; JobTitle = Get-OptionalProperty $item 'jobTitle'; CompanyName = Get-OptionalProperty $item 'companyName'; Department = Get-OptionalProperty $item 'department'; BusinessPhones = @(Get-OptionalProperty $item 'businessPhones'); MobilePhone = Get-OptionalProperty $item 'mobilePhone' }
            $emailKey = $email.ToLowerInvariant()
            if ($contactsByEmail.ContainsKey($emailKey)) {
                $duplicateContactIds += [string](Get-OptionalProperty -Object $item -Name 'id')
                Write-SyncLog -Level WARN -Message "Duplicate contact for $emailKey was found in $MailboxId. The extra copy will be deleted."
            } else {
                $contactsByEmail[$emailKey] = [pscustomobject]@{ Email = $emailKey; ContactId = Get-OptionalProperty $item 'id'; Fingerprint = Get-ContactFingerprint $model }
            }
        }
    }
    [pscustomobject]@{ MailboxId = $MailboxId; FolderId = $FolderId; Contacts = @($contactsByEmail.Values); DuplicateContactIds = @($duplicateContactIds) }
}

function Find-DuplicateFolders {
    param([string] $MailboxId, [string] $ManagedFolderId)
    $user = ConvertTo-GraphPath $MailboxId
    foreach ($folder in Get-BetaGraphPages -Uri "https://graph.microsoft.com/beta/users/$user/contactFolders?`$select=id,displayName") {
        if ($folder.id -ne $ManagedFolderId -and $folder.displayName -eq $FolderName) {
            [pscustomobject]@{ Id=$folder.id; Name=$folder.displayName }
        }
    }
}

function Remove-DuplicateFolder {
    param([string] $MailboxId, $Folder)
    $user = ConvertTo-GraphPath $MailboxId
    $base = "https://graph.microsoft.com/beta/users/$user/contactFolders/$($Folder.Id)"
    # Legacy folders discovered only by beta can reject a v1 folder DELETE.
    # Remove their contacts individually first; then remove the empty folder.
    foreach ($contact in Get-BetaGraphPages -Uri "$base/contacts?`$select=id" -ImmutableIds) {
        Invoke-MgGraphRequest -Method DELETE -Uri "$base/contacts/$($contact.id)" -Headers @{ Prefer = 'IdType="ImmutableId"' } | Out-Null
    }
    try {
        Invoke-MgGraphRequest -Method DELETE -Uri $base | Out-Null
        Write-SyncLog -Message "Removed duplicate managed folder '$($Folder.Name)' ($($Folder.Id)) from $MailboxId."
    } catch {
        Write-SyncLog -Level WARN -Message "Removed contacts from duplicate folder '$($Folder.Name)' but Exchange kept the empty folder: $($_.Exception.Message)"
    }
}

function Invoke-GraphBatch {
    # Send a group of Graph create/update/delete operations in one HTTP request.
    # Graph supports up to 20 operations in a JSON batch.
    param([object[]] $Operations)
    $pending = @($Operations); $attempt = 0; $results = @()
    while ($pending.Count -gt 0) {
        # Give every subrequest a temporary ID so its response can be matched
        # back to the operation that created it.
        $requests = @(); $operationById = @{}
        foreach ($operation in $pending) {
            $id = [guid]::NewGuid().ToString(); $operationById[$id] = $operation
            # All operations in this batch target Outlook contacts. Request
            # immutable IDs so newly created contacts can be cached safely and
            # existing IDs remain usable if a contact is moved within a mailbox.
            $request = @{ id = $id; method = $operation.Method; url = $operation.Url; headers = @{ Prefer = 'IdType="ImmutableId"' } }
            if ($operation.Body) { $request.headers['Content-Type'] = 'application/json'; $request.body = $operation.Body }
            $requests += $request
        }
        # Post the complete batch to Graph. A successful batch envelope can still
        # contain individual subrequests that failed or were throttled.
        $response = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/$batch' -Body (@{ requests = $requests } | ConvertTo-Json -Depth 12) -ContentType 'application/json' -OutputType PSObject
        $retry = @(); $wait = 0
        foreach ($item in @($response.responses)) {
            $operation = $operationById[$item.id]
            # HTTP 2xx means this one Graph operation completed successfully.
            if ($item.status -ge 200 -and $item.status -lt 300) { $results += [pscustomobject]@{ Operation = $operation; Response = $item; RequiresMailboxReconciliation = $false } }
            elseif ($item.status -eq 404 -and $operation.Method -in @('DELETE', 'PATCH')) {
                # A cached Outlook item ID can become stale after a move or an
                # external deletion. A 404 does not prove that the logical
                # contact is absent under another ID, so make the caller rescan
                # the managed folder instead of treating the delete as complete.
                $results += [pscustomobject]@{ Operation = $operation; Response = $item; RequiresMailboxReconciliation = $true }
            }
            elseif (($item.status -eq 429 -or $item.status -ge 500) -and $attempt -lt $MaxBatchRetries) {
                # Retry throttling (429) and temporary service errors (5xx), but
                # do not retry invalid requests such as 400 or 403.
                $retry += $operation
                $retryAfter = 0
                $retryAfterValue = Get-OptionalProperty -Object (Get-OptionalProperty -Object $item -Name 'headers') -Name 'Retry-After'
                if ($retryAfterValue) { [void][int]::TryParse([string]$retryAfterValue, [ref]$retryAfter) }
                $wait = [Math]::Max($wait, $retryAfter)
            } else {
                # Include Graph's machine-readable diagnostics. The message alone
                # is often generic (for example, quota and access failures can
                # both say that properties could not be read), while the error
                # code and request ID identify the actual Exchange failure and
                # let Microsoft support trace the request.
                $errorBody = Get-OptionalProperty -Object $item -Name 'body'
                $graphError = Get-OptionalProperty -Object $errorBody -Name 'error'
                $innerError = Get-OptionalProperty -Object $graphError -Name 'innerError'
                $diagnostics = @()
                $contactEmail = Get-OptionalProperty -Object $operation -Name 'Email'
                $errorCode = Get-OptionalProperty -Object $graphError -Name 'code'
                $innerErrorCode = Get-OptionalProperty -Object $innerError -Name 'code'
                $requestId = Get-OptionalProperty -Object $innerError -Name 'request-id'
                $clientRequestId = Get-OptionalProperty -Object $innerError -Name 'client-request-id'
                $errorDate = Get-OptionalProperty -Object $innerError -Name 'date'
                if ($contactEmail) { $diagnostics += "contact=$contactEmail" }
                if ($errorCode) { $diagnostics += "code=$errorCode" }
                if ($innerErrorCode -and $innerErrorCode -ne $errorCode) { $diagnostics += "inner-code=$innerErrorCode" }
                if ($requestId) { $diagnostics += "request-id=$requestId" }
                if ($clientRequestId) { $diagnostics += "client-request-id=$clientRequestId" }
                if ($errorDate) { $diagnostics += "date=$errorDate" }
                $diagnosticText = if ($diagnostics.Count -gt 0) { " [$($diagnostics -join '; ')]" } else { '' }
                $errorMessage = Get-OptionalProperty -Object $graphError -Name 'message'
                if (-not $errorMessage) { $errorMessage = 'Graph returned no error message.' }
                throw "Graph batch operation $($operation.Method) $($operation.Url) failed with HTTP $($item.status)$diagnosticText`: $errorMessage"
            }
        }
        if ($retry.Count -gt 0) {
            # Prefer Graph's Retry-After value. When it is unavailable, wait longer
            # on each retry (exponential backoff) to avoid making throttling worse.
            $attempt++
            if ($wait -le 0) { $wait = [Math]::Min(60, [Math]::Pow(2, $attempt)) }
            Write-SyncLog -Level WARN -Message "Retrying $($retry.Count) Graph operation(s) in $wait second(s)."
            Start-Sleep -Seconds $wait
        }
        $pending = $retry
    }
    $results
}

function Sync-Mailbox {
    # Compare the directory's desired contacts with one mailbox's cached contact
    # IDs and fingerprints, then build the smallest possible set of Graph writes.
    param([string] $MailboxId, $State, $MailboxStateOverride, [int] $ReconciliationAttempt = 0)
    # A reconciliation retry passes the just-scanned state directly so duplicate
    # contact IDs discovered by that scan are not lost in the SQLite round trip.
    $mailboxState = if ($MailboxStateOverride) { $MailboxStateOverride } else { Get-MailboxState $State $MailboxId }
    try {
        $folderId = Get-OrCreateFolder $MailboxId $mailboxState @($State.DesiredContacts).Count
    } catch {
        $graphFailure = Format-GraphRequestError -ErrorRecord $_
        throw "Unable to access the managed contact folder for ${MailboxId}: $graphFailure"
    }
    if (-not $mailboxState -or $State.RebuildMailboxCache) {
        # The first sync, and each periodic full refresh, reads the actual Outlook
        # folder instead of trusting SQLite. This detects contacts users deleted
        # or changed directly in Outlook and rebuilds the mailbox-specific map.
        $mailboxState = Initialize-MailboxState $MailboxId $folderId
        Save-MailboxState $mailboxState
    }
    $mailboxState.FolderId = $folderId
    # Convert both lists into hashtables so email matching is fast.
    $current = @{}; foreach ($entry in @($mailboxState.Contacts)) { $current[$entry.Email] = $entry }
    $desired = @{}; foreach ($contact in @($State.DesiredContacts)) { $desired[$contact.Email.ToLowerInvariant()] = $contact }
    $user = ConvertTo-GraphPath $MailboxId; $operations = @()
    # A full mailbox scan can identify manual duplicate contacts. Delete only
    # the extra copies; the canonical copy remains available for an update or
    # deletion based on the current directory source below.
    foreach ($duplicateContactId in @(Get-OptionalProperty -Object $mailboxState -Name 'DuplicateContactIds' | Where-Object { $_ })) {
        # Duplicate deletes are deliberately separate from source-contact
        # deletes, so deleting an extra copy never removes the canonical SQLite
        # mapping for that email address.
        $operations += [pscustomobject]@{ Action = 'DeleteDuplicate'; Method = 'DELETE'; Url = "/users/$user/contactFolders/$folderId/contacts/$duplicateContactId"; Body = $null }
    }
    if ($State.RebuildMailboxCache) {
        # Categories can expose migrated/old copies in other contact folders.
        # Remove only copies with both the managed category and a directory email.
        $duplicateFolders = @(Find-DuplicateFolders -MailboxId $MailboxId -ManagedFolderId $folderId)
        foreach ($duplicate in $duplicateFolders) {
            Remove-DuplicateFolder -MailboxId $MailboxId -Folder $duplicate
        }
    }
    # Contacts no longer in the source directory are deleted from this managed folder.
    foreach ($email in $current.Keys) {
        if (-not $desired.ContainsKey($email)) { $operations += [pscustomobject]@{ Action = 'Delete'; Email = $email; Method = 'DELETE'; Url = "/users/$user/contactFolders/$folderId/contacts/$($current[$email].ContactId)"; Body = $null } }
    }
    # Source contacts missing in Outlook are created; matching contacts are updated
    # only when their fingerprint differs.
    foreach ($email in $desired.Keys) {
        $source = $desired[$email]; $body = New-GraphContactBody $source
        if (-not $current.ContainsKey($email)) { $operations += [pscustomobject]@{ Action = 'Create'; Email = $email; Method = 'POST'; Url = "/users/$user/contactFolders/$folderId/contacts"; Body = $body; Source = $source } }
        elseif ($current[$email].Fingerprint -ne $source.Fingerprint) { $operations += [pscustomobject]@{ Action = 'Update'; Email = $email; Method = 'PATCH'; Url = "/users/$user/contactFolders/$folderId/contacts/$($current[$email].ContactId)"; Body = $body; Source = $source } }
    }
    if ($operations.Count -eq 0) { Write-SyncLog -Message "$MailboxId is already current."; return }
    Write-SyncLog -Message "Syncing $($operations.Count) change(s) to $MailboxId."
    $completed = @()
    # Split the work into batches so no Graph request exceeds the service limit.
    for ($offset = 0; $offset -lt $operations.Count; $offset += $BatchSize) {
        $last = [Math]::Min($offset + $BatchSize - 1, $operations.Count - 1)
        $batchResults = @(Invoke-GraphBatch -Operations $operations[$offset..$last])
        $staleIdResults = @($batchResults | Where-Object { $_.RequiresMailboxReconciliation })
        if ($staleIdResults.Count -gt 0) {
            if ($ReconciliationAttempt -ge 1) {
                $staleOperation = $staleIdResults[0].Operation
                throw "Graph still could not find a contact after rescanning ${MailboxId}: $($staleOperation.Method) $($staleOperation.Url)"
            }
            # Other subrequests in this batch may already have succeeded. Read
            # the folder's actual state so the retry accounts for every result,
            # including a contact that now exists under a different ID.
            $staleMethods = @($staleIdResults | ForEach-Object { $_.Operation.Method } | Sort-Object -Unique) -join '/'
            Write-SyncLog -Level WARN -Message "Graph could not find $($staleIdResults.Count) cached contact ID(s) during $staleMethods for $MailboxId; refreshing the managed-folder cache and retrying once."
            $refreshedMailboxState = Initialize-MailboxState $MailboxId $folderId
            Save-MailboxState $refreshedMailboxState
            Sync-Mailbox -MailboxId $MailboxId -State $State -MailboxStateOverride $refreshedMailboxState -ReconciliationAttempt ($ReconciliationAttempt + 1)
            return
        }
        $completed += $batchResults
    }
    # Update the local mailbox cache only for operations Graph confirmed as successful.
    foreach ($result in $completed) {
        $operation = $result.Operation
        if ($operation.Action -eq 'Delete') { $current.Remove($operation.Email) | Out-Null }
        elseif ($operation.Action -eq 'Create') { $current[$operation.Email] = [pscustomobject]@{ Email = $operation.Email; ContactId = $result.Response.body.id; Fingerprint = $operation.Source.Fingerprint } }
        elseif ($operation.Action -eq 'Update') { $current[$operation.Email].Fingerprint = $operation.Source.Fingerprint }
    }
    # Store only rows changed by this sync. A full map is already saved during
    # first-time mailbox initialization above.
    Save-MailboxChanges -MailboxState $mailboxState -Completed $completed
}

try {
    # SQLite stores the large per-mailbox contact mapping. The initializer script
    # creates this file and its tables before the first sync.
    if (-not (Test-Path -LiteralPath $DatabasePath)) { throw "SQLite database '$DatabasePath' was not found. Run Getting Started\\Initialize-GraphContactSyncDatabase.ps1 first." }
    $DatabasePath = (Resolve-Path -LiteralPath $DatabasePath).Path
    if ($null -eq (Get-Module -ListAvailable PSSQLite | Select-Object -First 1)) { throw 'PSSQLite is required. Install it with: Install-Module PSSQLite -Scope AllUsers' }
    Import-Module PSSQLite -ErrorAction Stop
    $requiredTables = @('Mailbox', 'MailboxContact', 'SourceContact', 'SyncMetadata')
    $actualTables = @(Invoke-SqliteQuery -DataSource $DatabasePath -Query "SELECT name FROM sqlite_master WHERE type='table'" | ForEach-Object { $_.name })
    $missingTables = @($requiredTables | Where-Object { $_ -notin $actualTables })
    if ($missingTables.Count -gt 0) {
        throw "SQLite database '$DatabasePath' is not initialized (missing: $($missingTables -join ', ')). Run Getting Started\\Initialize-GraphContactSyncDatabase.ps1 with this exact -DatabasePath."
    }
    # Confirm the minimal Graph PowerShell module is installed before doing any work.
    if ($null -eq (Get-Module -ListAvailable Microsoft.Graph.Authentication | Select-Object -First 1)) { throw 'Microsoft.Graph.Authentication is required. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' }
    # Create the optional log folder and choose a unique log filename for this run.
    if ($LogPath) { New-Item -ItemType Directory -Force -Path $LogPath | Out-Null; $script:LogFile = Join-Path $LogPath ("GraphContactSync_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)) }
    # Read recipient mailboxes from a CSV when the caller chose that option.
    if ($MailboxCsvPath) { $MailboxList = Import-Csv -LiteralPath $MailboxCsvPath | ForEach-Object { if ($_.Mailbox) { $_.Mailbox } elseif ($_.UserPrincipalName) { $_.UserPrincipalName } } }
    # Decrypt the PFX password for this Windows user and load the local certificate.
    $password = Import-Clixml -LiteralPath $CertificatePasswordPath
    $certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath.FullName, $password)
    # Sign in to Graph as the application. The welcome banner is harmless.
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Certificate $certificate
    # Load the previous checkpoint, and discard only its source cache when filters changed.
    $state = Get-State
    $filterSignature = Get-FilterSignature
    if ((Get-OptionalProperty $state 'FilterSignature') -ne $filterSignature) {
        Write-SyncLog -Message 'Source filter settings changed; rebuilding the directory cache.'
        $state.UserDeltaLink = $null
        $state.SourceContacts = @()
        $state.DesiredContacts = @()
        $state.SourceChanges = @()
        $state.RebuildSourceCache = $true
        $state.RebuildMailboxCache = $true
        if ($state.PSObject.Properties['FilterSignature']) { $state.FilterSignature = $filterSignature }
        else { $state | Add-Member -NotePropertyName FilterSignature -NotePropertyValue $filterSignature }
    } else {
        # Periodically rebuild the source cache. This is useful for recovering
        # from an unexpected cache inconsistency without doing a full read every
        # scheduled run. A database created before this setting existed has no
        # timestamp, so it safely performs one full refresh on its next run.
        $lastFullRefresh = Get-OptionalProperty -Object $state -Name 'LastFullDirectoryRefreshUtc'
        $refreshDue = $true
        if ($lastFullRefresh) {
            try { $refreshDue = (((Get-Date).ToUniversalTime() - [datetime]::Parse($lastFullRefresh).ToUniversalTime()).TotalDays -ge $FullDirectoryRefreshDays) }
            catch { Write-SyncLog -Level WARN -Message 'Last full directory-refresh timestamp was invalid; rebuilding the directory cache.' }
        }
        if ($refreshDue) {
            Write-SyncLog -Message "Full directory and mailbox-cache refresh is due; rebuilding it (every $FullDirectoryRefreshDays day(s))."
            $state.UserDeltaLink = $null
            $state.SourceContacts = @()
            $state.DesiredContacts = @()
            $state.SourceChanges = @()
            $state.RebuildSourceCache = $true
            $state.RebuildMailboxCache = $true
        }
    }
    # Get current source contacts, using delta tracking when a prior checkpoint exists.
    $state = Get-SourceState $state
    Write-SyncLog -Message "Directory cache contains $(@($state.DesiredContacts).Count) contact(s)."
    if ($directoryMode) {
        # Reuse the already-loaded directory source instead of paging /users a
        # second time just to build target mailboxes. Source filters therefore
        # also determine which mailboxes DIRECTORY mode targets.
        $MailboxList = @($state.SourceContacts | Where-Object { $_.Email } | ForEach-Object { $_.Email } | Sort-Object -Unique)
        Write-SyncLog -Level WARN -Message "DIRECTORY mode selected $($MailboxList.Count) mailbox(es) from the cached directory source."
    }
    $failed = 0
    # Sync one mailbox at a time. Sequential processing is gentler on Graph throttling.
    foreach ($mailbox in @($MailboxList | Where-Object { $_ } | Select-Object -Unique)) {
        try { Sync-Mailbox $mailbox $state }
        catch { $failed++; Write-SyncLog -Level ERROR -Message "Failed to sync $mailbox : $($_.Exception.Message)" }
    }
    # Keep successful mailboxes moving even when one mailbox has a transient
    # problem. Each later run compares every mailbox's saved mapping to the
    # current source cache, so the failed mailbox is reconciled on its retry.
    Save-State $state
    if ($failed -eq 0) { Write-SyncLog -Message 'Sync completed and state was saved.' }
    else { Write-SyncLog -Level WARN -Message "$failed mailbox sync(s) failed; successful mailbox and directory state was saved. Failed mailboxes will retry next run." }
} finally {
    # Always close the Graph session, including when a failure occurs.
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
