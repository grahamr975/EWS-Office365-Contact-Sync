# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased] (To do list)

## [2.0.1] - 8/21/2020
## Removed
- Deleted unused/obsolete functions from from the EWS Contacts module

## [2.0.1] - 7/15/2020
## New
- Added expermental support for Modern Authenication... See here for some backround information: https://techcommunity.microsoft.com/t5/exchange-team-blog/upcoming-changes-to-exchange-web-services-ews-api-for-office-365/ba-p/608055

## [2.0.0] - 2/5/2020 (Note: This update may break previous installations...)
## New
- Added IncludeNonUserContacts switch to Get-GALContacts and EWSContactSync
- Added ExcludeSharedMailboxContacts switch to Get-GALContacts and EWSContactSync
## Changed
- Shared mailboxes are now automatically included as contacts by default. To exclude shared mailboxes, use the ExcludeSharedMailboxContacts switch. This change is intented to improve clarity.
- The RequirePhoneNumber parameter has been changed to ExcludeContactsWithoutPhoneNumber in order to improve clarity

## [1.0.4] - 2019-11-15
## New
- New function: SetEXCContactObject; Updates an EWS Contact object that is passed into the function
- New function: NewEXCContactObject; Creates a new EWS Contact. The EWS Service and Contact Folder objects are passed into this function
## Changed
- Sync-ContactList function now uses SetEXCContactObject and NewEXCContactObject to improve the speed of the script

## [1.0.3] - 2019-11-15
## Fixed
- Fatal error when there are no contacts in the user's mailbox

## [1.0.2] - 2019-11-11
## Changed
- Changed the logging method to transcript
## Fixed
- Re-did the previous changes to fix an unknown parameter error

## [1.0.1] - 2019-11-8
## Added
- When "DIRECTORY" is specified for the MailboxList, now every user in the directory will be included
## Changed
- Moved the main functionality of the script into a function called Sync-ContactList
- Integrated all custom functions (library.ps1) into the EWSContacts Module

## [1.0.0] - 2019-11-2019
## Changed
- Read the user's mailbox once for all contacts rather than for every contact when determining if a contact needs to be deleted, updated, or added.
- Only update a contact if it needs to be updated. If both the new and old contact are exact matches, skip to the next contact.

## [0.0.3] - 2019-10-15
## Added
- Removes contacts from the target folder that are no longer in the Global Address List. (NOTE: Does not currently delete contacts with no email address)

## [0.0.2] - 2019-10-15
## Added
- Parameters for CredentialPath, FolderName, MailboxList, & LogPath
- Log writing functionality (See Write-Log function in library.ps1)
- Error handling

## [0.0.1] - 2019-10-14
### Alpha
- Ported from previous version of the the Multi-Contact Update script, this fork looks to overwrite contacts rather than delete them.
