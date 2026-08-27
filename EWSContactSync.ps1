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
    # One or more target mailbox email addresses or Entra user IDs.
    [string[]] $MailboxList,
    # Optional CSV alternative. It needs a Mailbox or UserPrincipalName column.
    [System.IO.FileInfo] $MailboxCsvPath,
    # Optional directory for timestamped text logs.
    [string] $LogPath,
    # Local SQLite database created by Initialize-GraphContactSyncDatabase.ps1.
    [string] $DatabasePath = 'C:\ContactSync\GraphContactSync.db',
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

function Get-GraphPages {
    # Graph lists can span multiple pages. Keep following @odata.nextLink until
    # Graph says there are no more pages, then return one combined list.
    param([Parameter(Mandatory)] [string] $Uri)
    $items = @()
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
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

function Get-State {
    # Read the small sync-wide checkpoint and source contacts from SQLite.
    $metadata = @{}
    foreach ($row in Invoke-SqliteQuery -DataSource $DatabasePath -Query 'SELECT MetadataKey, MetadataValue FROM SyncMetadata') { $metadata[$row.MetadataKey] = $row.MetadataValue }
    $source = @(Invoke-SqliteQuery -DataSource $DatabasePath -Query 'SELECT ContactJson FROM SourceContact' | ForEach-Object { $_.ContactJson | ConvertFrom-Json })
    # Changes is empty after a database read. Get-SourceState adds only the
    # users returned by this run's Graph delta feed.
    [pscustomobject]@{ FilterSignature = $metadata['FilterSignature']; UserDeltaLink = $metadata['UserDeltaLink']; SourceContacts = $source; DesiredContacts = @(); SourceChanges = @(); RebuildSourceCache = $false }
}

function Save-State { param($State)
    $now = (Get-Date).ToUniversalTime().ToString('o')
    # A new database, expired delta token, or changed filters requires a safe
    # full rebuild. Normal delta runs only write the individual changed rows.
    if ($State.RebuildSourceCache) {
        Invoke-SqliteQuery -DataSource $DatabasePath -Query 'DELETE FROM SourceContact' | Out-Null
        foreach ($contact in @($State.SourceContacts)) {
            Invoke-SqliteQuery -DataSource $DatabasePath -Query 'INSERT INTO SourceContact (SourceId,Email,Fingerprint,ContactJson,UpdatedUtc) VALUES (@id,@email,@fingerprint,@json,@now)' -SqlParameters @{ id=$contact.SourceId; email=$contact.Email; fingerprint=$contact.Fingerprint; json=($contact | ConvertTo-Json -Depth 8 -Compress); now=$now } | Out-Null
        }
    } elseif (@($State.SourceChanges).Count -gt 0) {
        foreach ($change in @($State.SourceChanges)) {
            if ($change.Action -eq 'Delete') {
                Invoke-SqliteQuery -DataSource $DatabasePath -Query 'DELETE FROM SourceContact WHERE SourceId=@id' -SqlParameters @{ id=$change.SourceId } | Out-Null
            } else {
                $contact = $change.Contact
                # Use older-SQLite-compatible insert/update statements in place
                # of UPSERT so this works with older PSSQLite installations.
                Invoke-SqliteQuery -DataSource $DatabasePath -Query 'INSERT OR IGNORE INTO SourceContact (SourceId,Email,Fingerprint,ContactJson,UpdatedUtc) VALUES (@id,@email,@fingerprint,@json,@now)' -SqlParameters @{ id=$contact.SourceId; email=$contact.Email; fingerprint=$contact.Fingerprint; json=($contact | ConvertTo-Json -Depth 8 -Compress); now=$now } | Out-Null
                Invoke-SqliteQuery -DataSource $DatabasePath -Query 'UPDATE SourceContact SET Email=@email,Fingerprint=@fingerprint,ContactJson=@json,UpdatedUtc=@now WHERE SourceId=@id' -SqlParameters @{ id=$contact.SourceId; email=$contact.Email; fingerprint=$contact.Fingerprint; json=($contact | ConvertTo-Json -Depth 8 -Compress); now=$now } | Out-Null
            }
        }
    } else {
        Write-SyncLog -Message 'Directory cache is unchanged; skipped source-cache database rewrite.'
    }
    # Always retain the newest Graph delta token, even when no users changed.
    foreach ($pair in @{ FilterSignature=$State.FilterSignature; UserDeltaLink=$State.UserDeltaLink }.GetEnumerator()) {
        # Use two older-SQLite-compatible statements instead of modern UPSERT syntax.
        Invoke-SqliteQuery -DataSource $DatabasePath -Query 'INSERT OR IGNORE INTO SyncMetadata (MetadataKey,MetadataValue,UpdatedUtc) VALUES (@key,@value,@now)' -SqlParameters @{ key=$pair.Key; value=$pair.Value; now=$now } | Out-Null
        Invoke-SqliteQuery -DataSource $DatabasePath -Query 'UPDATE SyncMetadata SET MetadataValue=@value,UpdatedUtc=@now WHERE MetadataKey=@key' -SqlParameters @{ key=$pair.Key; value=$pair.Value; now=$now } | Out-Null
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
            $State.UserDeltaLink = $null; $State.SourceContacts = @(); $State.SourceChanges = @(); $State.RebuildSourceCache = $true; return Get-SourceState $State
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
    # Insert a mailbox on its first sync, then update its folder ID on every run.
    Invoke-SqliteQuery -DataSource $DatabasePath -Query 'INSERT OR IGNORE INTO Mailbox (MailboxId,FolderId,UpdatedUtc) VALUES (@id,@folder,@now)' -SqlParameters @{ id=$MailboxState.MailboxId; folder=$MailboxState.FolderId; now=$now } | Out-Null
    Invoke-SqliteQuery -DataSource $DatabasePath -Query 'UPDATE Mailbox SET FolderId=@folder,UpdatedUtc=@now WHERE MailboxId=@id' -SqlParameters @{ id=$MailboxState.MailboxId; folder=$MailboxState.FolderId; now=$now } | Out-Null
    Invoke-SqliteQuery -DataSource $DatabasePath -Query 'DELETE FROM MailboxContact WHERE MailboxId=@id' -SqlParameters @{ id=$MailboxState.MailboxId } | Out-Null
    foreach ($contact in @($MailboxState.Contacts)) { Invoke-SqliteQuery -DataSource $DatabasePath -Query 'INSERT INTO MailboxContact (MailboxId,Email,ContactId,Fingerprint,UpdatedUtc) VALUES (@mailbox,@email,@contact,@fingerprint,@now)' -SqlParameters @{ mailbox=$MailboxState.MailboxId; email=$contact.Email; contact=$contact.ContactId; fingerprint=$contact.Fingerprint; now=$now } | Out-Null }
}

function Save-MailboxChanges {
    # Persist only the contact mappings that Graph successfully changed. This
    # avoids deleting and rebuilding thousands of rows for a one-contact update.
    param($MailboxState, [object[]] $Completed)
    $now = (Get-Date).ToUniversalTime().ToString('o')
    # Keep the mailbox timestamp meaningful without touching its contact rows.
    Invoke-SqliteQuery -DataSource $DatabasePath -Query 'UPDATE Mailbox SET FolderId=@folder,UpdatedUtc=@now WHERE MailboxId=@id' -SqlParameters @{ id=$MailboxState.MailboxId; folder=$MailboxState.FolderId; now=$now } | Out-Null
    foreach ($result in @($Completed)) {
        $operation = $result.Operation
        if ($operation.Action -eq 'Delete') {
            # The Graph DELETE succeeded, so remove just this contact mapping.
            Invoke-SqliteQuery -DataSource $DatabasePath -Query 'DELETE FROM MailboxContact WHERE MailboxId=@mailbox AND Email=@email' -SqlParameters @{ mailbox=$MailboxState.MailboxId; email=$operation.Email } | Out-Null
        } elseif ($operation.Action -eq 'Create') {
            # Graph returned the new, mailbox-specific Outlook contact ID.
            Invoke-SqliteQuery -DataSource $DatabasePath -Query 'INSERT OR REPLACE INTO MailboxContact (MailboxId,Email,ContactId,Fingerprint,UpdatedUtc) VALUES (@mailbox,@email,@contact,@fingerprint,@now)' -SqlParameters @{ mailbox=$MailboxState.MailboxId; email=$operation.Email; contact=$result.Response.body.id; fingerprint=$operation.Source.Fingerprint; now=$now } | Out-Null
        } else {
            # An update retains the same Outlook contact ID; only its applied
            # fingerprint needs to be advanced to the source fingerprint.
            Invoke-SqliteQuery -DataSource $DatabasePath -Query 'UPDATE MailboxContact SET Fingerprint=@fingerprint,UpdatedUtc=@now WHERE MailboxId=@mailbox AND Email=@email' -SqlParameters @{ mailbox=$MailboxState.MailboxId; email=$operation.Email; fingerprint=$operation.Source.Fingerprint; now=$now } | Out-Null
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
    param([string] $MailboxId, $MailboxState)
    if ($MailboxState -and $MailboxState.FolderId) { return $MailboxState.FolderId }
    $user = ConvertTo-GraphPath $MailboxId
    $folder = Get-GraphPages -Uri "/v1.0/users/$user/contactFolders?`$select=id,displayName" | Where-Object { $_.displayName -eq $FolderName } | Select-Object -First 1
    if (-not $folder) { $folder = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/users/$user/contactFolders" -Body (@{ displayName = $FolderName } | ConvertTo-Json) -ContentType 'application/json' -OutputType PSObject }
    $folder.id
}

function Initialize-MailboxState {
    # This runs only for a mailbox without local state. It reads the existing
    # managed folder once and stores each Outlook contact ID and fingerprint.
    param([string] $MailboxId, [string] $FolderId)
    $user = ConvertTo-GraphPath $MailboxId
    $select = 'id,displayName,givenName,surname,jobTitle,companyName,department,emailAddresses,businessPhones,mobilePhone'
    $existing = Get-GraphPages -Uri "/v1.0/users/$user/contactFolders/$FolderId/contacts?`$select=$select"
    $contacts = @()
    foreach ($item in $existing) {
        # Use the first email address as the key because the sync uses email as
        # its matching identity between the directory and Outlook contacts.
        $emailAddress = @(Get-OptionalProperty -Object $item -Name 'emailAddresses') | Select-Object -First 1
        $email = Get-OptionalProperty -Object $emailAddress -Name 'address'
        if ($email) {
            $email = [string]$email
            $model = [pscustomobject]@{ DisplayName = Get-OptionalProperty $item 'displayName'; FirstName = Get-OptionalProperty $item 'givenName'; LastName = Get-OptionalProperty $item 'surname'; Email = $email; JobTitle = Get-OptionalProperty $item 'jobTitle'; CompanyName = Get-OptionalProperty $item 'companyName'; Department = Get-OptionalProperty $item 'department'; BusinessPhones = @(Get-OptionalProperty $item 'businessPhones'); MobilePhone = Get-OptionalProperty $item 'mobilePhone' }
            $contacts += [pscustomobject]@{ Email = $email.ToLowerInvariant(); ContactId = Get-OptionalProperty $item 'id'; Fingerprint = Get-ContactFingerprint $model }
        }
    }
    [pscustomobject]@{ MailboxId = $MailboxId; FolderId = $FolderId; Contacts = @($contacts) }
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
            $request = @{ id = $id; method = $operation.Method; url = $operation.Url }
            if ($operation.Body) { $request.headers = @{ 'Content-Type' = 'application/json' }; $request.body = $operation.Body }
            $requests += $request
        }
        # Post the complete batch to Graph. A successful batch envelope can still
        # contain individual subrequests that failed or were throttled.
        $response = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/$batch' -Body (@{ requests = $requests } | ConvertTo-Json -Depth 12) -ContentType 'application/json' -OutputType PSObject
        $retry = @(); $wait = 0
        foreach ($item in @($response.responses)) {
            $operation = $operationById[$item.id]
            # HTTP 2xx means this one Graph operation completed successfully.
            if ($item.status -ge 200 -and $item.status -lt 300) { $results += [pscustomobject]@{ Operation = $operation; Response = $item } }
            elseif (($item.status -eq 429 -or $item.status -ge 500) -and $attempt -lt $MaxBatchRetries) {
                # Retry throttling (429) and temporary service errors (5xx), but
                # do not retry invalid requests such as 400 or 403.
                $retry += $operation
                $retryAfter = 0
                $retryAfterValue = Get-OptionalProperty -Object (Get-OptionalProperty -Object $item -Name 'headers') -Name 'Retry-After'
                if ($retryAfterValue) { [void][int]::TryParse([string]$retryAfterValue, [ref]$retryAfter) }
                $wait = [Math]::Max($wait, $retryAfter)
            } else { throw "Graph batch operation $($operation.Method) $($operation.Url) failed with HTTP $($item.status): $($item.body.error.message)" }
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
    param([string] $MailboxId, $State)
    $mailboxState = Get-MailboxState $State $MailboxId
    $folderId = Get-OrCreateFolder $MailboxId $mailboxState
    if (-not $mailboxState) {
        # First sync for this mailbox: create the initial local map from Outlook contacts.
        $mailboxState = Initialize-MailboxState $MailboxId $folderId
        Save-MailboxState $mailboxState
    }
    $mailboxState.FolderId = $folderId
    # Convert both lists into hashtables so email matching is fast.
    $current = @{}; foreach ($entry in @($mailboxState.Contacts)) { $current[$entry.Email] = $entry }
    $desired = @{}; foreach ($contact in @($State.DesiredContacts)) { $desired[$contact.Email.ToLowerInvariant()] = $contact }
    $user = ConvertTo-GraphPath $MailboxId; $operations = @()
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
        $completed += Invoke-GraphBatch -Operations $operations[$offset..$last]
    }
    # Update the local mailbox cache only for operations Graph confirmed as successful.
    foreach ($result in $completed) {
        $operation = $result.Operation
        if ($operation.Action -eq 'Delete') { $current.Remove($operation.Email) | Out-Null }
        elseif ($operation.Action -eq 'Create') { $current[$operation.Email] = [pscustomobject]@{ Email = $operation.Email; ContactId = $result.Response.body.id; Fingerprint = $operation.Source.Fingerprint } }
        else { $current[$operation.Email].Fingerprint = $operation.Source.Fingerprint }
    }
    # Store only rows changed by this sync. A full map is already saved during
    # first-time mailbox initialization above.
    Save-MailboxChanges -MailboxState $mailboxState -Completed $completed
}

try {
    # SQLite stores the large per-mailbox contact mapping. The initializer script
    # creates this file and its tables before the first sync.
    if (-not (Test-Path -LiteralPath $DatabasePath)) { throw "SQLite database '$DatabasePath' was not found. Run Getting Started\\Initialize-GraphContactSyncDatabase.ps1 first." }
    if ($null -eq (Get-Module -ListAvailable PSSQLite | Select-Object -First 1)) { throw 'PSSQLite is required. Install it with: Install-Module PSSQLite -Scope AllUsers' }
    Import-Module PSSQLite -ErrorAction Stop
    # Confirm the minimal Graph PowerShell module is installed before doing any work.
    if ($null -eq (Get-Module -ListAvailable Microsoft.Graph.Authentication | Select-Object -First 1)) { throw 'Microsoft.Graph.Authentication is required. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' }
    # Create the optional log folder and choose a unique log filename for this run.
    if ($LogPath) { New-Item -ItemType Directory -Force -Path $LogPath | Out-Null; $script:LogFile = Join-Path $LogPath ("GraphContactSync_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)) }
    # Read recipient mailboxes from a CSV when the caller chose that option.
    if ($MailboxCsvPath) { $MailboxList = Import-Csv -LiteralPath $MailboxCsvPath | ForEach-Object { if ($_.Mailbox) { $_.Mailbox } elseif ($_.UserPrincipalName) { $_.UserPrincipalName } } }
    if (@($MailboxList).Count -eq 1 -and $MailboxList[0].ToUpperInvariant() -eq 'DIRECTORY') {
        # DIRECTORY expands to every enabled member user with an email address.
        # Explicit recipients or a CSV are safer in production.
        Write-SyncLog -Level WARN -Message 'DIRECTORY is not recommended in production; use -MailboxCsvPath or an explicit list.'
        $MailboxList = Get-GraphPages -Uri '/v1.0/users?$select=id,mail,accountEnabled,userType' | Where-Object { $_.accountEnabled -eq $true -and $_.userType -eq 'Member' -and $_.mail } | ForEach-Object { $_.id }
    }
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
        if ($state.PSObject.Properties['FilterSignature']) { $state.FilterSignature = $filterSignature }
        else { $state | Add-Member -NotePropertyName FilterSignature -NotePropertyValue $filterSignature }
    }
    # Get current source contacts, using delta tracking when a prior checkpoint exists.
    $state = Get-SourceState $state
    Write-SyncLog -Message "Directory cache contains $(@($state.DesiredContacts).Count) contact(s)."
    $failed = 0
    # Sync one mailbox at a time. Sequential processing is gentler on Graph throttling.
    foreach ($mailbox in @($MailboxList | Where-Object { $_ } | Select-Object -Unique)) {
        try { Sync-Mailbox $mailbox $state }
        catch { $failed++; Write-SyncLog -Level ERROR -Message "Failed to sync $mailbox : $($_.Exception.Message)" }
    }
    # Only advance the delta checkpoint when every mailbox completed. This ensures
    # failed mailbox changes are retried next time instead of being silently skipped.
    if ($failed -eq 0) { Save-State $state; Write-SyncLog -Message 'Sync completed and state was saved.' }
    else { throw "$failed mailbox sync(s) failed; state was not advanced." }
} catch  {
    If ($LogPath) { 
        Write-SyncLog -Message "An initialization error has occured."
        Write-SyncLog -Message $_.ScriptStackTrace
    }
} finally {
    # Always close the Graph session, including when a failure occurs.
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
