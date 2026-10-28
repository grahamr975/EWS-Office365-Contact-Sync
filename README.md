# Microsoft Graph Contact Sync

Synchronize Microsoft Entra directory contacts into a dedicated Outlook contact
folder for one or more Exchange Online mailboxes. This is a Microsoft Graph-only
rewrite of the former EWS sync script.

After the first sync, uses SQL Lite to cache the contact state of each mailbox. This way, the script won't have to re-read each's mailboxes contact list on future runs. This caching significantly improves the sync run-time after the first run. After caching, 2,000 contacts x 2,000 mailboxes will sync in 1-2 hours.

**Why would I want to use this?** iPhone/Android devices don't currently support offline Global Address List synchronization. By loading the Global Address List contacts into a folder within user's mailbox, you can circumvent this limitation.

## Prerequisites

1. PowerShell 5.1+ or PowerShell 7+.
2. Install the required Graph authentication module for the account running the job:

   ```powershell
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
   ```

3. Create a Microsoft Entra app registration & certificate using [the tutorial here](https://github.com/MicrosoftDocs/office-docs-powershell/blob/main/exchange/docs-conceptual/app-only-auth-powershell-v2.md), upload the certificate to it, and
   grant tenant-wide admin consent for these **application** permissions:

   | Permission | Why |
   | --- | --- |
   | `Contacts.ReadWrite` | Create, update, and delete contacts in target mailboxes. |
   | `User.Read.All` | Read directory users and resolve `DIRECTORY` mailboxes. |
   | `OrgContact.Read.All` | Required only with `-IncludeNonUserContacts`. |

   The app can be restricted to an approved mailbox scope by using Exchange
   Online application RBAC. Do that before using `DIRECTORY` in production.

4. Export the PFX password as a CliXml secure string on the same Windows account
   that will run the scheduled task. The existing
   `Getting Started/Create-SecureCertificatePassword.ps1` helper can be used.

5. Install the SQL Lite Powershell Module & initialize the local sync database.
```powershell
Install-Module PSSQLite -Scope AllUsers

.\Getting Started\Initialize-GraphContactSyncDatabase.ps1 `
  -DatabasePath 'C:\ContactSync\GraphContactSync.db'
```

6. Create a unique folder name for the script to use. The named folder is managed as a whole. Contacts no longer present in the chosen
directory source are deleted from that folder, so do not use a personal contacts
folder as the target.


## Run

```powershell
.\EWSContactSync.ps1 `
  -TenantId 'contoso.onmicrosoft.com' `
  -ClientId '00000000-0000-0000-0000-000000000000' `
  -CertificatePath 'C:\Certs\contact-sync.pfx' `
  -CertificatePasswordPath 'C:\Certs\contact-sync-password.cred' `
  -FolderName 'Directory Contacts' `
  -MailboxList 'person@contoso.com' `
  -DatabasePath 'C:\ContactSync\GraphContactSync.db' `
  -LogPath 'C:\ContactSync\Logs'
```

When starting the script interactively, supply `-MailboxList 'person@contoso.com'`
or `-MailboxCsvPath 'C:\ContactSync\Mailboxes.csv'` on the command line. PowerShell
prompts for mandatory parameters through `FolderName`, but does not prompt for
either mailbox-selection parameter. Omitting both produces
`Specify -MailboxList or -MailboxCsvPath.`

## Performance and state

The first run reads the directory and each managed contact folder. It saves
source contact fingerprints, mailbox contact IDs, and a Microsoft Graph user
delta link to SQLite. Later runs read only directory
changes and create, update, or delete only affected contacts. Keep this state
database in a protected, persistent folder and do not delete it unless you intend
to perform a full reconciliation.

Changing any source filter (`-ExcludeSharedMailboxContacts`,
`-ExcludeContactsWithoutPhoneNumber`, or `-IncludeNonUserContacts`) automatically
rebuilds the source cache on the next run. Contacts excluded by the new filter
are removed from the managed folder.

`-ExcludeContactsWithoutPhoneNumber` requires at least one nonblank business or
mobile phone value. Null, empty, and whitespace-only values do not count, for
either directory users or organizational contacts. The first run after this
phone-filter update rebuilds the source cache when the switch is enabled, so
previously included contacts without phone values are removed from the managed
folder without resetting the database.

For directory users whose first and last names are both empty, the sync puts the
full directory display name in the contact's first-name field. This gives mobile
clients a personal name to display while retaining the company field. Existing
first or last names are preserved; organizational contacts remain company-only
when their source has no personal name. If the directory display name is empty,
the sync uses the available first/last name, then the email address.

The first run after this name-mapping update automatically rebuilds the directory
cache and rescans managed mailbox folders, updating affected contacts without a
database reset. Allow the phone to sync afterward. If a contact already has
correct first/last names in Outlook but still shows a company on the phone,
check for linked or duplicate phone contacts and stale device contact data.

Graph write calls are sent as batches of up to 20 operations. The script honors
`Retry-After` for throttled and transient batch responses. Adjust the batch size
with `-BatchSize` (1–20) only if your tenant needs a lower rate.

Contact IDs are requested in Microsoft Graph's immutable-ID format so moving a
contact within the same mailbox does not invalidate its cached ID. If Graph
still returns `404` for a cached contact during an update or deletion, the
script rescans that mailbox's managed folder and retries once. The rescan
matches contacts by normalized email address and also accounts for any other
operations that already succeeded in the same Graph batch.

For scheduled runs, use an explicit `-MailboxList` or a CSV file:

```csv
Mailbox
person1@contoso.com
person2@contoso.com
```

```powershell
.\EWSContactSync.ps1 ... -MailboxCsvPath 'C:\ContactSync\Mailboxes.csv'
```

`DIRECTORY` is still available as an option you can pass as a string to the MailBoxList paramter. It
attempts every enabled member user with an email address. `-ExcludeSharedMailboxContacts`
omits unlicensed directory users; Graph does not expose Exchange's
`HiddenFromAddressListsEnabled` property, so that legacy filter cannot be reproduced exactly.

Optional switches retained from the EWS version:

- `-ExcludeContactsWithoutPhoneNumber`
- `-ExcludeSharedMailboxContacts`
- `-IncludeNonUserContacts` (Microsoft Entra organizational contacts)

`Multi-Threaded.ps1` was removed as an option due to graph-API throttling limits. EWS was a bit more forgiving.

## Graph API Documentation

In `DIRECTORY` mode, Graph requests use the cached Entra user object ID, so a
user's email address does not need to match their sign-in name. Existing SQLite
mailbox keys and log labels remain email addresses. Organizational contacts are
source contacts only, not target mailboxes. Explicit `MailboxList` and
`MailboxCsvPath` targets should use Entra object IDs or user principal names.

### Recovery and run status

The script automatically adds a `MailboxHealth` table to existing databases; no
database reset is required. On the first run after this upgrade, each target
mailbox is rescanned to establish its recovery state, so that run can take longer.
Later runs rescan mailboxes after failures or interruptions, when their folder
changes, or when their individual reconciliation interval expires.

Successful Graph operations are saved after each batch, even if another operation
fails. An uncertain contact creation is reconciled before a later attempt rather
than blindly retried. Cleared directory phone numbers now clear the corresponding
Outlook fields.

Duplicate-folder cleanup requires a contact category matching `FolderName` and a
matching email already present in the canonical folder. Unique or untagged
contacts are preserved with a warning; folders containing them remain for review.

When `LogPath` is configured, each run writes a text log. Mailbox failure entries
include consecutive failures and the last successful sync time. Log timestamps
are UTC; recovery state remains stored in the database.

The script exits with code `1` when any mailbox fails and `0` after a successful
run. Unexpected terminating errors also fail the process. A `.sync.lock` file
beside the SQLite database prevents overlapping runs against that database; the
operating system releases its lock when the process exits. The file itself stays
in place and does not need to be deleted. Schedule the script in its own
PowerShell process, and propagate its exit code if using a wrapper script.

Offline regression checks are documented in [Tests/README.md](Tests/README.md).

Microsoft documents the contact API and its required application permission in
[Create contact](https://learn.microsoft.com/en-us/graph/api/user-post-contacts?view=graph-rest-1.0), and documents the optional organizational-contact source in
[List orgContacts](https://learn.microsoft.com/en-us/graph/api/orgcontact-list?view=graph-rest-1.0).

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 


## Authors

* **Ryan Graham** - *Initial work* - [grahamr975](https://github.com/grahamr975)

## Acknowledgments

* Thanks to gscales for his work on the EWSContacts powershell module (now obsoleted). This script uses a modified version of their module. https://github.com/gscales/Powershell-Scripts/tree/master/EWSContacts

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
