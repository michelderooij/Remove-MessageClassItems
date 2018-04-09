# Remove-MessageClassItems

## Getting Started

This script will remove items of a certain class from a mailbox, traversing through mail 
item folders (IPF.Note). You can also specify how the items should be deleted. Example usages 
are cleaning up mailboxes of stubs, by removing the shortcut messages left behind by archiving 
products such as Enterprise Vault.

### Requirements

* PowerShell 3.0 or later
* EWS Managed API 1.2 or later

### Usage

Syntax:
```
Remove-MessageClassItems.ps1 [-Identity] <String> [-MessageClass] <String> [-Server <String>] [-Impersonation] [-Credentials <PSCredential>] [-DeleteMode <String>] [-ScanAllFolders] [-Before <DateTime>] [-MailboxOnly] [-ArchiveOnly] [-IncludeFolders <String[]>] [-ExcludeFolders <String[]>] [-NoProgressBar] [-Report] [-WhatIf] [-Confirm] [<CommonParameters>]
```

Examples:
```
.\Remove-MessageClassItems.ps1 -Identity user1 -Impersonation -Verbose -DeleteMode MoveToDeletedItems -MessageClass IPM.Note.EnterpriseVault.Shortcut
```
Process mailbox of user1, moving "IPM.Note.EnterpriseVault.Shortcut" message class items to the
DeletedItems folder, using impersonation and verbose output.

```
.\Remove-MessageClassItems.ps1 -Identity user1 -Impersonation -Verbose -DeleteMode SoftDelete -MessageClass EnterpriseVault -PartialMatching -ScanAllFolders -Before ((Get-Date).AddDays(-90))
```
Process mailbox of user1, scanning all folders and soft-deleting items that contain the string 'EnterpriseVault' in their
message class items and are older than 90 days.

```
$Credentials= Get-Credential
.\Remove-MessageClassItems.ps1 -Identity olrik@office365tenant.com -Credentials $Credentials -MessageClass IPM.Note.EnterpriseVault.Shortcut -MailboxOnly
```
Get credentials and process only the mailbox of olrik@office365tenant.com, removing "IPM.Note.EnterpriseVault.Shortcut" message class items.

```
$Credentials= Get-Credential
.\Remove-MessageClassItems.ps1 -Identity michel@contoso.com -MessageClass IPM* -Credentials $Credentials -WhatIf:$true -Verbose -IncludeFolders Archive\* -ExcludeFolders ArchiveX
```
Scan mailbox of michel@contoso looking for items of classes starting with IPM using provided credentials, limited to folders named Archive and their subfolders, but
excluding folders named ArchiveX, showing what it would do (WhatIf) in verbose mode.

```
Import-CSV users.csv1 | .\Remove-MessageClassItems.ps1 -Impersonation -DeleteMode HardDelete -MessageClass IPM.ixos-archive
```
Uses a CSV file to fix specified mailboxes (containing Identity column), removing "IPM.ixos-archive" items permanently, using impersonation.

### About

For more information on this script, as well as usage and examples, see
the related blog article, [Removing Message Class Items from a Mailbox](http://eightwone.com/2013/05/16/removing-messages-by-message-class-from-mailbox/).

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

 