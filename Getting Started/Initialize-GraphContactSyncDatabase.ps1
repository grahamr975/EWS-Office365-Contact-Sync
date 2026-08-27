<#
.SYNOPSIS
Creates the SQLite database used by a large Microsoft Graph contact-sync deployment.

.DESCRIPTION
This creates the schema needed to replace a large JSON state file when syncing
many mailboxes. SQLite is a single local file, so run the sync from one Windows
server and do not place the database on a network share.

The script requires the PSSQLite PowerShell module. Install it once with:
    Install-Module PSSQLite -Scope AllUsers
#>
[CmdletBinding()]
param(
    # Store this on the sync server's local drive. The scheduled-task account
    # must have Modify permission to this folder.
    [string] $DatabasePath = 'C:\ContactSync\GraphContactSync.db'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Verify that the SQLite PowerShell provider is installed before doing any work.
if ($null -eq (Get-Module -ListAvailable PSSQLite | Select-Object -First 1)) {
    throw 'PSSQLite is required. Install it in an elevated PowerShell window with: Install-Module PSSQLite -Scope AllUsers'
}

Import-Module PSSQLite -ErrorAction Stop

# Make the target directory when it does not already exist.
$databaseDirectory = Split-Path -Parent $DatabasePath
New-Item -ItemType Directory -Force -Path $databaseDirectory | Out-Null

# SQLite uses a file lock. WAL mode lets one writer and readers coexist more
# efficiently, while busy_timeout makes a second accidental process wait briefly
# instead of failing immediately.
$schema = @'
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
PRAGMA busy_timeout = 30000;
PRAGMA foreign_keys = ON;

-- Stores small sync-wide values such as the Graph user delta link and selected filters.
CREATE TABLE IF NOT EXISTS SyncMetadata (
    MetadataKey   TEXT PRIMARY KEY,
    MetadataValue TEXT NULL,
    UpdatedUtc    TEXT NOT NULL
);

-- One row for every Entra user or organizational contact selected as a source.
CREATE TABLE IF NOT EXISTS SourceContact (
    SourceId      TEXT PRIMARY KEY,
    Email          TEXT NOT NULL COLLATE NOCASE,
    Fingerprint    TEXT NOT NULL,
    ContactJson    TEXT NOT NULL,
    UpdatedUtc     TEXT NOT NULL
);

-- One row per target mailbox. FolderId is the Graph ID of the dedicated folder.
CREATE TABLE IF NOT EXISTS Mailbox (
    MailboxId      TEXT PRIMARY KEY COLLATE NOCASE,
    FolderId       TEXT NOT NULL,
    UpdatedUtc     TEXT NOT NULL
);

-- One row per synchronized contact in each mailbox. This is the large table:
-- approximately 1,805 rows for each mailbox in the user's deployment.
CREATE TABLE IF NOT EXISTS MailboxContact (
    MailboxId      TEXT NOT NULL COLLATE NOCASE,
    Email          TEXT NOT NULL COLLATE NOCASE,
    ContactId      TEXT NOT NULL,
    Fingerprint    TEXT NOT NULL,
    UpdatedUtc     TEXT NOT NULL,
    PRIMARY KEY (MailboxId, Email),
    FOREIGN KEY (MailboxId) REFERENCES Mailbox(MailboxId) ON DELETE CASCADE
);

-- These indexes make the two normal lookups fast: contacts for a mailbox and
-- all mailbox copies of a changed source email.
CREATE INDEX IF NOT EXISTS IX_SourceContact_Email ON SourceContact(Email);
CREATE INDEX IF NOT EXISTS IX_MailboxContact_Email ON MailboxContact(Email);
'@

# Create the database file and all tables. Running this script again is safe;
# CREATE ... IF NOT EXISTS leaves existing data alone.
Invoke-SqliteQuery -DataSource $DatabasePath -Query $schema | Out-Null

# Save a schema version so a later script version can safely apply migrations.
$parameters = @{ key = 'SchemaVersion'; value = '1'; updated = (Get-Date).ToUniversalTime().ToString('o') }
# INSERT OR IGNORE is supported by older SQLite versions included with some PSSQLite installs.
$schemaVersionSql = 'INSERT OR IGNORE INTO SyncMetadata (MetadataKey, MetadataValue, UpdatedUtc) VALUES (@key, @value, @updated)'
Invoke-SqliteQuery -DataSource $DatabasePath -Query $schemaVersionSql -SqlParameters $parameters | Out-Null

Write-Host "SQLite database is ready: $DatabasePath"
Write-Host 'The main sync script can now use this database with its -DatabasePath parameter.'
