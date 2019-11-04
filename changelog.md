# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased] (To do list)
- Automagically import a list of mailboxes from Active Directory
- Multi threading
- Create switch parameter for wether to import all contacts, or only user contacts

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