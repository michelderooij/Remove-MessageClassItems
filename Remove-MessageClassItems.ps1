<#
    .SYNOPSIS
    Remove-MessageClassItems

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 2.10, December 17th, 2021

    .DESCRIPTION
    This script will remove items of a certain class from a mailbox, traversing through
    mail item folders (IPF.Note). You can also specify how the items should be deleted.
    Example usages are cleaning up mailboxes of stubs, by removing the shortcut messages
    left behind by archiving products such as Enterprise Vault.
    Note that usage of the Verbose, Confirm and WhatIf parameters is supported. When
    using Confirm, you will be prompted per removal batch.

    .LINK
    http://eightwone.com

    .NOTES
    Microsoft Exchange Web Services (EWS) Managed API 1.2 or up is required. Recommended to use package 
    EWS.WebServices.Managed.Api, see https://eightwone.com/2020/10/05/ews-webservices-managed-api
    For OAuth, Microsoft Authentication Authentication Libraries (MSAL) is  required.
    
    Search order for DLL's is script Folder then installed packages.

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial release
    1.01    Fixed example (IPM.ixos-archive instead of IBM.ixos-archive)
            Removed requirement for EMS
    1.1     Added switch -ScanAllFolders to scan all mail folders, not only IPF.Note folders
            Added parameter -Before to only remove items received before specified date.
    1.2     Fixed issue with PowerShell v3 (System.Collections.Generic.List`1)
    1.3     Changed parameter Mailbox, you can now use an e-mail address as well
            Added parameter Credentials
            Added parameter PartialMatching for partial class name matching
            Changed item removal process. Remove items after, not while processing
            folder. Avoids asynchronous deletion issues.
            Works against Office 365
            Deleted Items folder will be processed, unless MoveToDeletedItems is used
            Changed EWS DLL loading, can now be in current folder as well
    1.31    Fixed bug when more than $MaxBatchSize entries were found
    1.4     Added personal archive support
    1.41    Fixed typo preventing script from working on Ex2007
    1.5     Added IncludeFolder parameter
            Added ExcludeFolder parameter
            Removed PartialMatching parameter, replaced with MessageClass wildcard option
    1.51    Fixed using 2+ Exclude folders
    1.52    Identity parameter replaces Mailbox
            Made "can't access information store" more verbose.
            Fixed bug in non-wildcard matching
    1.53    Fixed (another) non-wildcard matching bug
    1.6     Added EWS throttling handling
            Added progress bar for progress indication
            Added NoProgressBar switch
            Added more statistics, e.g. items/minute summary
    1.61    Fixed bug when store (archive) is inaccessible
    1.7     Renamed IncludeFolder/ExcludeFolder to IncludeFolders/ExcludeFolders
            Changed IncludeFolders and ExcludeFolders to add path matching
            Added #JunkEmail# and #DeletedItems# to IncludeFolders/ExcludeFolders
            Fixed Well-Known Folder processing to use current mailbox folder name
            Optimalizations when running against multiple mailboxes
            Some code rewriting
    1.71    Fixed partial folder name matching
    1.8     Added ReplaceClass parameter
            Added Report switch
            Added MessageClass size limits check
            Added EWS Managed API DLL version reporting (Verbose)
    1.81    Added X-AnchorMailbox for impersonation requests
    1.82    Fixed issue with processing delegate mailboxes using Full Access permissions
    1.83    Fixed bug in folder selection process
    1.84    Added code to leverage installed package EWS.WebServices.Managed.Api
    2.00    Added OAuth authentication options
            Changed DLL loading routing (EWS Managed API + MSAL)
            Not trusting self-signed certs by default; added -TrustAll switch to trust all certs
            Added pipeline proper processing with begin/process/end
            Replaced all strings with var-subsitution with -f 
            Added certificate authentication example
            Determine DeletedItems once per mailbox, not for every folder to process
            Replaced ScanAllFolders with Type
    2.01    Fixed loading of module when using installed NuGet packages
    2.02    Changed PropertySet constructors to prevent possible initialization issues
    2.10    Added DefaultAuth for usage on-premises (using current security context)

    .PARAMETER Identity
    Identity of the Mailbox. Can be CN/SAMAccountName (for on-premises) or e-mail format (on-prem & Office 365)

    .PARAMETER MessageClass
    Specifies the PR_MESSAGE_CLASS to remove, for example IPM.ixos-archive (Opentext/IXOS LiveLink E-Mail Archive)
    or IPM.Note.EnterpriseVault.Shortcut (EnterpriseVault). You can use wildcards around or at the end to
    include folders containing or starting with this string, e.g. 'IPM.ixos*' or '*EnterpriseVault*'.
    Matching is always case-insensitive.

    .PARAMETER ReplaceClass
    Specifies that instead of removing the item, its PR_MESSAGE_CLASS class property will be modified to this value.
    For example, can be used in conjunction with MessageClass to modify any IPM.Note items pending Evault archival back
    to regular items: -MessageClass IPM.Note.EnterpriseVault.PendingArchive -ReplaceClass IPM.Note

    .PARAMETER DeleteMode
    Determines how to remove messages. Options are:
    - HardDelete:         Items will be permanently deleted.
    - SoftDelete:         Items will be moved to the dumpster (default).
    - MoveToDeletedItems: Items will be moved to the Deleted Items folder.
    When using MoveToDeletedItems, the Deleted Items folder will not be processed.

    .PARAMETER Type
    Determines what kind of folderclass to check for items.
    Options: Mail, Calendar, Contacts, Tasks, Notes or All (Default).

    .PARAMETER Before
    Allows you to remove items of a certain age by filtering on their DateTimeReceived attribute. Default is all items.

    .PARAMETER MailboxOnly
    Only process primary mailbox of specified users. You als need to use this parameter when
    running against mailboxes on Exchange Server 2007.

    .PARAMETER ArchiveOnly
    Only process personal archives of specified users.

    .PARAMETER IncludeFolders
    Specify one or more names of folder(s) to include, e.g. 'Projects'. You can use wildcards
    around or at the end to include folders containing or starting with this string, e.g.
    'Projects*' or '*Project*'. To match folders and subfolders, add a trailing \*,
    e.g. Projects\*. This will include folders named Projects and all subfolders.
    To match from the top of the structure, prepend using '\'. Matching is case-insensitive.

    Some examples, using the following folder structure:

    + TopFolderA
        + FolderA
            + SubFolderA
            + SubFolderB
        + FolderB
    + TopFolderB
        + FolderA

    Filter              Match(es)
    --------------------------------------------------------------------------------------------------------------------
    FolderA             \TopFolderA\FolderA, \TopFolderB\FolderA
    Folder*             \TopFolderA\FolderA, \TopFolderA\FolderB, \TopFolderA\FolderA\SubFolderA, \TopFolderA\FolderA\SubFolderB
    FolderA\*Folder*    \TopFolderA\FolderA\SubFolderA, \TopFolderA\FolderA\SubFolderB
    \*FolderA\*         \TopFolderA, \TopFolderA\FolderA, \TopFolderA\FolderB, \TopFolderA\FolderA\SubFolderA, \TopFolderA\FolderA\SubFolderB, \TopFolderB\FolderA
    \*\FolderA          \TopFolderA\FolderA, \TopFolderB\FolderA

    You can also use well-known folders, by using this format: #WellKnownFolderName#, e.g. #Inbox#.
    Supported are #Calendar#, #Contacts#, #Inbox#, #Notes#, #SentItems#, #Tasks#, #JunkEmail# and #DeletedItems#.
    The script uses the currently configured Well-Known Folder of the mailbox to be processed.

    .PARAMETER ExcludeFolders
    Specify one or more folder(s) to exclude. Usage of wildcards and well-known folders identical to IncludeFolders.
    Note that ExcludeFolders criteria overrule IncludeFolders when matching folders.

    .PARAMETER NoProgressBar
    Use this switch to prevent displaying of a progress bar as folders and items are being processed.

    .PARAMETER Report
    Reports individual items detected as duplicate. Can be used together with WhatIf to perform pre-analysis.

    .PARAMETER Server
    Exchange Client Access Server to use for Exchange Web Services. When ommited, script will attempt to
    use Autodiscover.

    .PARAMETER TrustAll
    Specifies if all certificates should be accepted, including self-signed certificates.

    .PARAMETER TenantId
    Specifies the identity of the Tenant.

    .PARAMETER ClientId
    Specifies the identity of the application configured in Azure Active Directory.

    .PARAMETER Credentials
    Specify credentials to use with Basic Authentication. Credentials can be set using $Credentials= Get-Credential
    This parameter is mutually exclusive with CertificateFile, CertificateThumbprint and Secret. 

    .PARAMETER CertificateThumbprint
    Specify the thumbprint of the certificate to use with OAuth authentication. The certificate needs
    to reside in the personal store. When using OAuth, providing TenantId and ClientId is mandatory.
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificateFile
    Specify the .pfx file containing the certificate to use with OAuth authentication. When a password is required,
    you will be prompted or you can provide it using CertificatePassword.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificatePassword
    Sets the password to use with the specified .pfx file. The provided password needs to be a secure string, 
    eg. -CertificatePassword (ConvertToSecureString -String 'P@ssword' -Force -AsPlainText)

    .PARAMETER Secret
    Specifies the client secret to use with OAuth authentication. The secret needs to be provided as a secure string.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and CertificateThumbprint. 

    .EXAMPLE
    .\Remove-MessageClassItems.ps1 -Identity user1 -Impersonation -Verbose -DeleteMode MoveToDeletedItems -MessageClass IPM.Note.EnterpriseVault.Shortcut

    Process mailbox of user1, moving "IPM.Note.EnterpriseVault.Shortcut" message class items to the
    DeletedItems folder, using impersonation and verbose output.

    .EXAMPLE
    .\Remove-MessageClassItems.ps1 -Identity user1 -Impersonation -Verbose -DeleteMode SoftDelete -MessageClass EnterpriseVault -PartialMatching -Type All -Before ((Get-Date).AddDays(-90))

    Process mailbox of user1, scanning all folders and soft-deleting items that contain the string 'EnterpriseVault' in their
    message class items and are older than 90 days.

    .EXAMPLE
    $Credentials= Get-Credential
    .\Remove-MessageClassItems.ps1 -Identity olrik@office365tenant.com -Credentials $Credentials -MessageClass IPM.Note.EnterpriseVault.Shortcut -MailboxOnly

    Get credentials and process only the mailbox of olrik@office365tenant.com, removing "IPM.Note.EnterpriseVault.Shortcut" message class items.

    .EXAMPLE
    $Credentials= Get-Credential
    .\Remove-MessageClassItems.ps1 -Identity michel@contoso.com -MessageClass IPM* -Credentials $Credentials -WhatIf:$true -Verbose -IncludeFolders Archive\* -ExcludeFolders ArchiveX

    Scan mailbox of michel@contoso looking for items of classes starting with IPM using provided credentials, limited to folders named Archive and their subfolders, but
    excluding folders named ArchiveX, showing what it would do (WhatIf) in verbose mode.

    .EXAMPLE
    $Secret= Read-Host 'Secret' -AsSecureString
    Import-CSV users.csv1 | .\Remove-MessageClassItems.ps1 -Server outlook.office365.com -DeleteMode HardDelete -MessageClass IPM.ixos-archive -TenantId '1ab81a53-2c16-4f28-98f3-fd251f0459f3' -ClientId 'ea76025c-592d-43f1-91f4-2dec7161cc59' -Secret $Secret

    Permanently remove items of type IPM.ixos-archive from mailboxes identified by CSV file in Office365 bypassing AutoDiscover. 
    OAuth authentication is performed against indicated tenant <TenantID> using registered App <ClientID> and App secret entered.
#>

[cmdletbinding(
    DefaultParameterSetName = 'DefaultAuth',
    SupportsShouldProcess= $true,
    ConfirmImpact= 'High'
)]
param(
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'DefaultAuth')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuth')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [alias('Mailbox')]
    [string]$Identity,
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [ValidateLength(1, 255)]
    [string]$MessageClass,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [ValidateLength(1, 255)]
    [string]$ReplaceClass,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [ValidateSet( 'Mail', 'Calendar', 'Contacts', 'Tasks', 'Notes', 'All')]
    [string]$Type= 'All',
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [string]$Server,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [switch]$Impersonation,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [ValidateSet( 'HardDelete', 'SoftDelete', 'MoveToDeletedItems')]
    [string]$DeleteMode= 'SoftDelete',
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [switch]$ScanAllFolders,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [datetime]$Before,
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [switch]$MailboxOnly,
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [switch]$ArchiveOnly,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [string[]]$IncludeFolders,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [string[]]$ExcludeFolders,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [switch]$NoProgressBar,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [switch]$Report,
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [Switch]$DefaultAuth,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [System.Security.SecureString]$Secret,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [String]$CertificateThumbprint,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [String]$CertificateFile,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [System.Security.SecureString]$CertificatePassword,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [string]$TenantId,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [string]$ClientId,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [switch]$TrustAll
)
#Requires -Version 3.0

begin {

write-host ($PSCommandSet)

    # Process folders these batches
    $script:MaxFolderBatchSize= 100
    # Process items in these page sizes
    $script:MaxItemBatchSize= 1000
    # Max of concurrent item deletes
    $script:MaxDeleteBatchSize= 100

    # Initial sleep timer (ms) and treshold before lowering
    $script:SleepTimerMax= 300000               # Maximum delay (5min)
    $script:SleepTimerMin= 100                  # Minimum delay
    $script:SleepAdjustmentFactor= 2.0          # When tuning, use this factor
    $script:SleepTimer= $script:SleepTimerMin   # Initial sleep timer value

    # Errors
    $ERR_DLLNOTFOUND= 1000
    $ERR_DLLLOADING= 1001
    $ERR_MAILBOXNOTFOUND= 1002
    $ERR_AUTODISCOVERFAILED= 1003
    $ERR_CANTACCESSMAILBOXSTORE= 1004
    $ERR_PROCESSINGMAILBOX= 1005
    $ERR_PROCESSINGARCHIVE= 1006
    $ERR_INVALIDCREDENTIALS= 1007
    $ERR_PROBLEMIMPORTINGCERT= 1008
    $ERR_CERTNOTFOUND= 1009

    ### HELPER FUNCTIONS ###

    Function Import-ModuleDLL {
        param(
            [string]$Name,
            [string]$FileName,
            [string]$Package,
            [string]$ValidateObjName
        )

        $AbsoluteFileName= Join-Path -Path $PSScriptRoot -ChildPath $FileName
        If ( Test-Path $AbsoluteFileName) {
            # OK
        }
        Else {
            If( $Package) {
                If( Get-Command -Name Get-Package -ErrorAction SilentlyContinue) {
                    If( Get-Package -Name $Package -ErrorAction SilentlyContinue) {
                        $AbsoluteFileName= (Get-ChildItem -ErrorAction SilentlyContinue -Path (Split-Path -Parent (get-Package -Name $Package | Sort-Object -Property Version -Descending | Select-Object -First 1).Source) -Filter $FileName -Recurse).FullName
                    }
                }
            }
        }

        If( $absoluteFileName) {
            $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
            If( $ModLoaded) {
                Write-Verbose ('Module {0} v{1} already loaded' -f $ModLoaded.Name, $ModLoaded.Version)
            }
            Else {
                Write-Verbose ('Loading module {0}' -f $absoluteFileName)
                try {
                    Import-Module -Name $absoluteFileName -Global -Force
                    Start-Sleep 1
                }
                catch {
                    Write-Error ('Problem loading module {0}: {1}' -f $Name, $error[0])
                    Exit $ERR_DLLLOADING
                }
                $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
                If( $ModLoaded) {
                    Write-Verbose ('Module {0} v{1} loaded' -f $ModLoaded.Name, $ModLoaded.Version)
                }
                If( $validateObjName) {
                    # Try initializing test object to validate proper loading of module
                    Try {
                        $null= New-Object -TypeName $validateObjName
                    }
                    Catch {
                        Write-Error ('Problem initializing test-object from module {0}: {1}' -f $Name, $_.Exception.Message)
                        Exit $ERR_DLLLOADING
                    }
                }
            }
        }
        Else {
            Write-Verbose ('Required module {0} could not be located' -f $FileName)
            Exit $ERR_DLLNOTFOUND
        }
    }

    Function Set-SSLVerification {
        param(
            [switch]$Enable,
            [switch]$Disable
        )

        Add-Type -TypeDefinition  @"
            using System.Net.Security;
            using System.Security.Cryptography.X509Certificates;
            public static class TrustEverything
            {
                private static bool ValidationCallback(object sender, X509Certificate certificate, X509Chain chain,
                    SslPolicyErrors sslPolicyErrors) { return true; }
                public static void SetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback= ValidationCallback; }
                public static void UnsetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback= null; }
        }
"@
        If($Enable) {
            Write-Verbose ('Enabling SSL certificate verification')
            [TrustEverything]::UnsetCallback()
        }
        Else {
            Write-Verbose ('Disabling SSL certificate verification')
            [TrustEverything]::SetCallback()
        }
    }

    Function Get-EmailAddress {
        param(
            [string]$Identity
        )
        $address= [regex]::Match([string]$Identity, ".*@.*\..*", "IgnoreCase")
        if ( $address.Success ) {
            return $address.value.ToString()
        }
        Else {
            # Use local AD to look up e-mail address using $Identity as SamAccountName
            $ADSearch= New-Object DirectoryServices.DirectorySearcher( [ADSI]"")
            $ADSearch.Filter= "(|(cn=$Identity)(samAccountName=$Identity)(mail=$Identity))"
            $Result= $ADSearch.FindOne()
            If ( $Result) {
                $objUser= $Result.getDirectoryEntry()
                return $objUser.mail.toString()
            }
            else {
                return $null
            }
        }
    }

    Function iif( $eval, $tv= '', $fv= '') {
        If ( $eval) { return $tv } else { return $fv}
    }

    Function Construct-FolderFilter {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string[]]$Folders,
            [string]$emailAddress
        )
        If ( $Folders) {
            $FolderFilterSet= [System.Collections.ArrayList]@()
            ForEach ( $Folder in $Folders) {
                # Convert simple filter to (simple) regexp
                $Parts= $Folder -match '^(?<root>\\)?(?<keywords>.*?)?(?<sub>\\\*)?$'
                If ( !$Parts) {
                    Write-Error ('Invalid regular expression matching against {0}' -f $Folder)
                }
                Else {
                    $Keywords= Search-ReplaceWellKnownFolderNames -EwsService $EwsService -Criteria ($Matches.keywords) -EmailAddress $emailAddress
                    $EscKeywords= [Regex]::Escape( $Keywords) -replace '\\\*', '.*'
                    $Pattern= iif -eval $Matches.Root -tv '^\\' -fv '^\\(.*\\)*'
                    $Pattern += iif -eval $EscKeywords -tv $EscKeywords -fv ''
                    $Pattern += iif -eval $Matches.sub -tv '(\\.*)?$' -fv '$'
                    $Obj= [pscustomobject]@{
                        'Pattern'    = [string]$Pattern
                        'IncludeSubs'= [bool]$Matches.Sub
                        'OrigFilter' = [string]$Folder
                    }
                    $null= $FolderFilterSet.Add( $Obj)
                    Write-Debug ($Obj -join ',')
                }
            }
        }
        Else {
            $FolderFilterSet= $null
        }
        return $FolderFilterSet
    }

    Function Search-ReplaceWellKnownFolderNames {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$criteria= '',
            [string]$emailAddress
        )
        $AllowedWKF= 'Inbox', 'Calendar', 'Contacts', 'Notes', 'SentItems', 'Tasks', 'JunkEmail', 'DeletedItems'
        # Construct regexp to see if allowed WKF is part of criteria string
        ForEach ( $ThisWKF in $AllowedWKF) {
            If ( $criteria -match '#{0}#') {
                $criteria= $criteria -replace ('#{0}#' -f $ThisWKF), (myEWSBind-WellKnownFolder $EwsService $ThisWKF $emailAddress).DisplayName
            }
        }
        return $criteria
    }
    Function Tune-SleepTimer {
        param(
            [bool]$previousResultSuccess= $false
        )
        if ( $previousResultSuccess) {
            If ( $script:SleepTimer -gt $script:SleepTimerMin) {
                $script:SleepTimer= [int]([math]::Max( [int]($script:SleepTimer / $script:SleepAdjustmentFactor), $script:SleepTimerMin))
                Write-Warning ('Previous EWS operation successful, adjusted sleep timer to {0}ms' -f $script:SleepTimer)
            }
        }
        Else {
            $script:SleepTimer= [int]([math]::Min( ($script:SleepTimer * $script:SleepAdjustmentFactor) + 100, $script:SleepTimerMax))
            If ( $script:SleepTimer -eq 0) {
                $script:SleepTimer= 5000
            }
            Write-Warning ('Previous EWS operation failed, adjusted sleep timer to {0}ms' -f $script:SleepTimer)
        }
        Start-Sleep -Milliseconds $script:SleepTimer
    }

    Function myEWSFind-Folders {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,
            [Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]$FolderSearchCollection,
            [Microsoft.Exchange.WebServices.Data.FolderView]$FolderView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $EwsService.FindFolders( $FolderId, $FolderSearchCollection, $FolderView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindFolders with Search options in {0}. Error: {1}' -f $FolderId.FolderName, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSFind-FoldersNoSearch {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,
            [Microsoft.Exchange.WebServices.Data.FolderView]$FolderView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $EwsService.FindFolders( $FolderId, $FolderView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindFolders without Search options in {0}. Error: {1}' -f $FolderId.FolderName, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSFind-Items {
        param(
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]$ItemSearchFilterCollection,
            [Microsoft.Exchange.WebServices.Data.ItemView]$ItemView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $Folder.FindItems( $ItemSearchFilterCollection, $ItemView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindItems with Search options in {0}. Error: {1}' -f $Folder.DisplayName, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSFind-ItemsNoSearch {
        param(
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [Microsoft.Exchange.WebServices.Data.ItemView]$ItemView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $Folder.FindItems( $ItemView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindItems without Search options in {0}. Error {1}' -f $Folder.DisplayName, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSRemove-Items {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            $ItemIds,
            [Microsoft.Exchange.WebServices.Data.DeleteMode]$DeleteMode,
            [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]$SendCancellationsMode,
            [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]$AffectedTaskOccurrences,
            [bool]$SuppressReadReceipt
        )
        $OpSuccess= $false
        $critErr= $false
        Do {
            Try {
                If ( @([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013, [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1) -contains $EwsService.RequestedServerVersion) {
                    $res= $EwsService.DeleteItems( $ItemIds, $DeleteMode, $SendCancellationsMode, $AffectedTaskOccurrences, $SuppressReadReceipt)
                }
                Else {
                    $res= $EwsService.DeleteItems( $ItemIds, $DeleteMode, $SendCancellationsMode, $AffectedTaskOccurrences)
                }
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation RemoveItems with {0}. Error: {1}' -f $RemoveItems, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSBind-WellKnownFolder {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$WellKnownFolderName,
            [string]$emailAddress
        )
        $OpSuccess= $false
        $critErr= $false
        Do {
            Try {
                $explicitFolder= New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName, $emailAddress)  
                $res= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $explicitFolder)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Cannot bind to {0}: {1}' -f $WellKnownFolderName, $_.Exception.Message)
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function Get-SubFolders {
        param(
            $Folder,
            $CurrentPath,
            $IncludeFilter,
            $ExcludeFilter
        )
        $FoldersToProcess= [System.Collections.ArrayList]@()
        $FolderView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( $MaxFolderBatchSize)
        $FolderView.Traversal= [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
        $FolderView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $FolderView.PropertySet.Add( [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
        $FolderView.PropertySet.Add( [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass)
        $FolderSearchCollection= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection( [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        If ( $Type -ne 'All') {
            $FolderSearchClass= (@{Mail= 'IPF.Note'; Calendar= 'IPF.Appointment'; Contacts= 'IPF.Contact'; Tasks= 'IPF.Task'; Notes= 'IPF.StickyNotes'})[$Type]
            $FolderSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, $FolderSearchClass)
            $FolderSearchCollection.Add( $FolderSearchFilter)
        }
        Do {
            If ( $FolderSearchCollection.Count -ge 1) {
                $FolderSearchResults= myEWSFind-Folders $EwsService $Folder.Id $FolderSearchCollection $FolderView
            }
            Else {
                $FolderSearchResults= myEWSFind-FoldersNoSearch $EwsService $Folder.Id $FolderView
            }
            ForEach ( $FolderItem in $FolderSearchResults) {
                $FolderPath= '{0}\{1}' -f $CurrentPath, $FolderItem.DisplayName
                If ( $IncludeFilter) {
                    $Add= $false
                    # Defaults to true, unless include does not specifically include subfolders
                    $Subs= $true
                    ForEach ( $Filter in $IncludeFilter) {
                        If ( $FolderPath -match $Filter.Pattern) {
                            $Add= $true
                            # When multiple criteria match, one with and one without subfolder processing, subfolders will be processed.
                            $Subs= $Filter.IncludeSubs
                        }
                    }
                }
                Else {
                    # If no includeFolders specified, include all (unless excluded)
                    $Add= $true
                    $Subs= $true
                }
                If ( $ExcludeFilter) {
                    # Excludes can overrule includes
                    ForEach ( $Filter in $ExcludeFilter) {
                        If ( $FolderPath -match $Filter.Pattern) {
                            $Add= $false
                            # When multiple criteria match, one with and one without subfolder processing, subfolders will be processed.
                            $Subs= $Filter.IncludeSubs
                        }
                    }
                }
                If ( $Add) {
                    Write-Verbose ( 'Adding folder {0}' -f $FolderPath)

                    $Obj= New-Object -TypeName PSObject -Property @{
                        'Name'    = $FolderPath;
                        'Folder'  = $FolderItem
                    }
                    $null= $FoldersToProcess.Add( $Obj)
                }
                If ( $Subs) {
                    # Could be that specific folder is to be excluded, but subfolders needs evaluation
                    ForEach ( $AddFolder in (Get-SubFolders -Folder $FolderItem -CurrentPath $FolderPath -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter)) {
                        $null= $FoldersToProcess.Add( $AddFolder)
                    }
                }
            }
            $FolderView.Offset += $FolderSearchResults.Folders.Count
        } While ($FolderSearchResults.MoreAvailable)
        Write-Output -NoEnumerate $FoldersToProcess
    }

    Function Process-Mailbox {
        [CmdletBinding(SupportsShouldProcess=$true)]
        Param(
            [string]$Identity,
            $Folder,
            $IncludeFilter,
            $ExcludeFilter,
            $emailAddress,
            $DeletedItemsFolder
        )

        $ProcessingOK= $True
        $TotalMatch= 0
        $TotalRemoved= 0
        $FoldersFound= 0
        $FoldersProcessed= 0
        $TimeProcessingStart= Get-Date

        # Build list of folders to process
        Write-Verbose (iif $ScanAllFolders -fv 'Collecting folders containing e-mail items to process' -tv 'Collecting folders to process')
        $FoldersToProcess= Get-SubFolders -Folder $Folder -CurrentPath '' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -ScanAllFolders $ScanAllFolders

        $FoldersFound= $FoldersToProcess.Count
        Write-Verbose ('Found {0} folders matching folder search criteria' -f $FoldersFound)

        ForEach ( $SubFolder in $FoldersToProcess) {
            If (!$NoProgressBar) {
                Write-Progress -Id 1 -Activity ('Processing {0}' -f $Identity) -Status ('Processed folder {0} of {1}' -f $FoldersProcessed, $FoldersFound) -PercentComplete ( $FoldersProcessed / $FoldersFound * 100)
            }
            If ( ! ( $DeleteMode -eq 'MoveToDeletedItems' -and $SubFolder.Id -eq $DeletedItemsFolder.Id)) {
                If ( $Report.IsPresent) {
                    Write-Host ('Processing folder {0}' -f $SubFolder.Name)
                }
                Else {
                    Write-Verbose ('Processing folder {0}' -f $SubFolder.Name)
                }
                $ItemView= New-Object Microsoft.Exchange.WebServices.Data.ItemView( $script:MaxItemBatchSize, 0, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
                $ItemView.Traversal= [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
                $ItemView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                $ItemView.PropertySet.Add( [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
                $ItemView.PropertySet.Add( [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
                $ItemView.PropertySet.Add( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass)

                $ItemSearchFilterCollection= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
                If ($MessageClass -match '^\*(?<substring>.*?)\*$') {
                    $ItemSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
                        $matches['substring'], [Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
                }
                Else {
                    If ($MessageClass -match '^(?<prefix>.*?)\*$') {
                        $ItemSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
                            $matches['prefix'], [Microsoft.Exchange.WebServices.Data.ContainmentMode]::Prefixed, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
                    }
                    Else {
                        $ItemSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, $MessageClass)
                    }
                }
                $ItemSearchFilterCollection.add( $ItemSearchFilter)

                If ( $Before) {
                    $ItemSearchFilterCollection.add( (New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan( [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $Before)))
                }

                $ProcessList= [System.Collections.ArrayList]@()
                If ( $psversiontable.psversion.major -lt 3) {
                    $ItemIds= [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.ItemId]))
                }
                Else {
                    $type= ("System.Collections.Generic.List" + '`' + "1") -as 'Type'
                    $type= $type.MakeGenericType([Microsoft.Exchange.WebServices.Data.ItemId] -as 'Type')
                    $ItemIds= [Activator]::CreateInstance($type)
                }

                Do {
                    $ItemSearchResults= MyEWSFind-Items $SubFolder.Folder $ItemSearchFilterCollection $ItemView
                    If (!$NoProgressBar) {
                        Write-Progress -Id 2 -Activity ('Processing folder {0}' -f $SubFolder.Name) -Status ('Found {0} matching items' -f $ProcessList.Count)
                    }
                    If ( $ItemSearchResults.Items.Count -gt 0) {
                        ForEach ( $Item in $ItemSearchResults.Items) {
                            If ( $Report.IsPresent) {
                                Write-Host ('Item: {0} of {1} ({2})' -f $Item.Subject, $Item.DateTimeReceived, $Item.ItemClass)
                            }
                            $ProcessList.Add( $Item.Id)
                        }
                    }
                    Else {
                        Write-Debug "No matching items found"
                    }
                    $ItemView.Offset += $ItemSearchResults.Items.Count
                } While ( $ItemSearchResults.MoreAvailable)
            }
            Else {
                Write-Debug "Skipping DeletedItems folder"
            }
            $TotalMatch += $ItemSearchResults.TotalCount
            If ( ($ProcessList.Count -gt 0) -and ($Force -or $PSCmdlet.ShouldProcess( ('{0} item(s) from {1}' -f $ProcessList.Count, $SubFolder.Name)))) {
                If ( $ReplaceClass) {
                    Write-Verbose ('Modifying {0} items from {1}' -f $ProcessList.Count, $SubFolder.Name)
                    $ItemsChanged= 0
                    $ItemsRemaining= $ProcessList.Count
                    ForEach ( $ItemID in $ProcessList) {
                        If (!$NoProgressBar) {
                            Write-Progress -Id 2 -Activity ('Processing folder {0}' -f $SubFolder.DisplayName) -Status ('Items processed {0} - remaining {1}' -f $ItemsChanged, $ItemsRemaining) -PercentComplete ( $ItemsRemoved / $ProcessList.Count * 100)
                        }
                        Try {
                            $ItemObj= [Microsoft.Exchange.WebServices.Data.Item]::Bind( $EwsService, $ItemID)
                            $ItemObj.ItemClass= $ReplaceClass
                            $ItemObj.Update( [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
                        }
                        Catch {
                            Write-Error ('Problem modifying item: {0}' -f $error[0])
                            $ProcessingOK= $False
                        }
                        $ItemsChanged++
                        $ItemsRemaining--
                    }
                }
                Else {
                    Write-Verbose ('Removing {0} items from {1}' -f $ProcessList.Count, $SubFolder.Name)

                    # How to deal with removing calendar items
                    $SendCancellationsMode= [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone
                    $AffectedTaskOccurrences= [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::SpecifiedOccurrenceOnly
                    $SuppressReadReceipt= $true # Only works using EWS with Exchange2013+ mode

                    $ItemsRemoved= 0
                    $ItemsRemaining= $ProcessList.Count

                    # Remove ItemIDs in batches
                    ForEach ( $ItemID in $ProcessList) {
                        $ItemIds.Add( $ItemID)
                        If ( $ItemIds.Count -eq $script:MaxDeleteBatchSize) {
                            $ItemsRemoved += $ItemIds.Count
                            $ItemsRemaining -= $ItemIds.Count
                            If (!$NoProgressBar) {
                                Write-Progress -Id 2 -Activity ('Processing folder {0}' -f $SubFolder.DisplayName) -Status ('Items processed {0} - remaining {1}' -f $ItemsRemoved, $ItemsRemaining) -PercentComplete ( $ItemsRemoved / $ProcessList.Count * 100)
                            }
                            $null= myEWSRemove-Items $EwsService $ItemIds $DeleteMode $SendCancellationsMode $AffectedTaskOccurrences $SuppressReadReceipt
                            $ItemIds.Clear()
                        }
                    }
                    # .. also remove last ItemIDs
                    If ( $ItemIds.Count -gt 0) {
                        $ItemsRemoved += $ItemIds.Count
                        $ItemsRemaining= 0
                        $res= myEWSRemove-Items $EwsService $ItemIds $DeleteMode $SendCancellationsMode $AffectedTaskOccurrences $SuppressReadReceipt
                        $ItemIds.Clear()
                    }
                    $TotalRemoved+= $ProcessList.Count
                }
                If (!$NoProgressBar) {
                    Write-Progress -Id 2 -Activity ('Processing folder {0}' -f $SubFolder.DisplayName) -Status 'Finished processing.' -Completed
                }
            }
            Else {
                # No matches
            }
            $FoldersProcessed++
        } # ForEach SubFolder
        If (!$NoProgressBar) {
            Write-Progress -Id 1 -Activity ('Processing {0}' -f $Identity) -Status 'Finished processing.' -Completed
        }
        If ( $ProcessingOK) {
            $TimeProcessingDiff= New-TimeSpan -Start $TimeProcessingStart -End (Get-Date)
            $Speed= [int]( $TotalMatch / $TimeProcessingDiff.TotalSeconds * 60)
            Write-Verbose ('{0} item(s) processed in {1:hh}:{1:mm}:{1:ss} - average {2} items/min' -f $TotalMatch, $TimeProcessingDiff, $Speed)
        }
        Return $ProcessingOK
    }

    Import-ModuleDLL -Name 'Microsoft.Exchange.WebServices' -FileName 'Microsoft.Exchange.WebServices.dll' -Package 'Exchange.WebServices.Managed.Api' -validateObjName 'Microsoft.Exchange.WebServices.Data.ExchangeVersion'
    Import-ModuleDLL -Name 'Microsoft.Identity.Client' -FileName 'Microsoft.Identity.Client.dll' -Package 'Microsoft.Identity.Client' -validateObjName 'Microsoft.Identity.Client.TokenCache'

    If ( $MailboxOnly) {
        $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
    }
    Else {
        $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    }
    $EwsService= [Microsoft.Exchange.WebServices.Data.ExchangeService]::new( $ExchangeVersion)

    If( $Credentials -or $DefaultAuth) {
        If( $Credentials) {
            try {
                Write-Verbose ('Using credentials {0}' -f $Credentials.UserName)
                $EwsService.Credentials= [System.Net.NetworkCredential]::new( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
            }
            catch {
                Write-Error ('Invalid credentials provided: {0}' -f $_.Exception.Message)
                Exit $ERR_INVALIDCREDENTIALS
            }
        }
        Else {
            Write-Verbose ('DefaultAuth specified, using current security context')
        }
    }
    Else {
        # Use OAuth (and impersonation/X-AnchorMailbox always set)
        $Impersonation= $true

        If( $CertificateThumbprint -or $CertificateFile) {
            If( $CertificateFile) {
                
                # Use certificate from file using absolute path to authenticate
                $CertificateFile= (Resolve-Path -Path $CertificateFile).Path
                
                Try {
                    If( $CertificatePassword) {
                        $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $CertificatePassword)))
                    }
                    Else {
                        $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile)
                    }
                }
                Catch {
                    Write-Error ('Problem importing PFX: {0}' -f $_.Exception.Message)
                    Exit $ERR_PROBLEMIMPORTINGCERT
                }
            }
            Else {
                # Use provided certificateThumbprint to retrieve certificate from My store, and authenticate with that
                $CertStore= [System.Security.Cryptography.X509Certificates.X509Store]::new( [Security.Cryptography.X509Certificates.StoreName]::My, [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
                $CertStore.Open( [System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly )
                $X509Certificate2= $CertStore.Certificates.Find( [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $CertificateThumbprint, $False) | Select-Object -First 1
                If(!( $X509Certificate2)) {
                    Write-Error ('Problem locating certificate in My store: {0}' -f $error[0])
                    Exit $ERR_CERTNOTFOUND
                }
            }
            Write-Verbose ('Will use certificate {0}, issued by {1} and expiring {2}' -f $X509Certificate2.Thumbprint, $X509Certificate2.Issuer, $X509Certificate2.NotAfter)
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithCertificate( $X509Certificate2).withTenantId( $TenantId).Build()
               
        }
        Else {
            # Use provided secret to authenticate
            Write-Verbose ('Will use provided secret to authenticate')
            $PlainSecret= [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Secret))
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithClientSecret( $PlainSecret).withTenantId( $TenantId).Build()
        }
        $Scopes= New-Object System.Collections.Generic.List[string]
        $Scopes.Add( 'https://outlook.office365.com/.default')
        Try {
            $Response=$App.AcquireTokenForClient( $Scopes).executeAsync()
            $Token= $Response.Result
            $EwsService.Credentials= [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$Token.AccessToken
            Write-Verbose ('Authentication token acquired')
        }
        Catch {
            Write-Error ('Problem acquiring token: {0}' -f $error[0])
            Exit $ERR_INVALIDCREDENTIALS
        }
    }

    If ( $ReplaceClass) {
        Write-Verbose ('Changing messages of class {0} to {1}' -f $MessageClass, $ReplaceClass)
    }
    Else {
        Write-Verbose ('DeleteMode is {0}' -f $DeleteMode)
        Write-Verbose ('Removing messages of class {0}' -f $MessageClass)
    }
    If ( $Before) {
        Write-Verbose "Removing messages older than $Before"
    }

    If( $TrustAll) {
        Set-SSLVerification -Disable
    }

}


process {

    ForEach ( $CurrentIdentity in $Identity) {

        $EmailAddress= get-EmailAddress -Identity $CurrentIdentity
        If ( !$EmailAddress) {
            Write-Error ('Specified mailbox {0} not found' -f $EmailAddress)
            Exit $ERR_MAILBOXNOTFOUND
        }

        Write-Host ('Processing mailbox {0} ({1})' -f $EmailAddress, $CurrentIdentity)

        If( $Impersonation) {
            Write-Verbose ('Using {0} for impersonation' -f $EmailAddress)
            $EwsService.ImpersonatedUserId= [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new( [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
            $EwsService.HttpHeaders.Clear()
            $EwsService.HttpHeaders.Add( 'X-AnchorMailbox', $EmailAddress)
        }
            
        If ($Server) {
            $EwsUrl= 'https://{0}/EWS/Exchange.asmx' -f $Server
            Write-Verbose ('Using Exchange Web Services URL {0}' -f $EwsUrl)
            $EwsService.Url= $EwsUrl
        }
        Else {
            Write-Verbose ('Looking up EWS URL using Autodiscover for {0}' -f $EmailAddress)
            try {
                # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
                $ErrorActionPreference= 'Stop'
                $EwsService.autodiscoverUrl( $EmailAddress, {$true})
            }
            catch {
                Write-Error ('Autodiscover failed: {0}' -f $_.Exception.Message)
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference= 'Continue'
            Write-Verbose 'Using EWS endpoint {0}' -f $EwsService.Url
        } 

        # Construct search filters
        Write-Verbose 'Constructing folder matching rules'
        $IncludeFilter= Construct-FolderFilter $EwsService $IncludeFolders $EmailAddress
        $ExcludeFilter= Construct-FolderFilter $EwsService $ExcludeFolders $EmailAddress

        If ( -not $ArchiveOnly.IsPresent) {
            try {
                $RootFolder= myEWSBind-WellKnownFolder $EwsService 'MsgFolderRoot' $emailAddress
                If ( $RootFolder) {
                    Write-Verbose ('Processing primary mailbox {0}' -f $emailAddress)
                    $DeletedItemsFolder= myEWSBind-WellKnownFolder $EwsService 'DeletedItems' $emailAddress
                    If (! ( Process-Mailbox -Identity $emailAddress -Folder $RootFolder -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -emailAddress $emailAddress -DeletedItemsFolder $DeletedItemsFolder)) {
                        Write-Error ('Problem processing primary mailbox of {0} ({1})' -f $CurrentIdentity, $EmailAddress)
                        Exit $ERR_PROCESSINGMAILBOX
                    }
                }
            }
            catch {
                Write-Error ('Cannot access mailbox information store for {0}: {1}' -f $EmailAddress, $_.Exception.Message)
                Exit $ERR_CANTACCESSMAILBOXSTORE
            }
        }

        If ( -not $MailboxOnly.IsPresent) {
            try {
                $ArchiveRootFolder= myEWSBind-WellKnownFolder $EwsService 'ArchiveMsgFolderRoot' $emailAddress
                If ( $ArchiveRootFolder) {
                    Write-Verbose ('Processing archive mailbox {0}' -f $Identity)
                    $DeletedItemsFolder= myEWSBind-WellKnownFolder $EwsService 'DeletedItems' $emailAddress
                    If (! ( Process-Mailbox -Identity $emailAddress -Folder $ArchiveRootFolder -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -emailAddress $emailAddress -DeletedItemsFolder $DeletedItemsFolder)) {
                        Write-Error ('Problem processing archive mailbox of {0} ({1})' -f $emailAddress, $CurrentIdentity)
                        Exit $ERR_PROCESSINGARCHIVE
                    }
                }
            }
            catch {
                Write-Debug 'No archive configured or cannot access archive'
            }
        }
        Write-Verbose ('Processing {0} finished' -f $emailAddress)
    }
}

End {
    If( $TrustAll) {
        Set-SSLVerification -Enable
    }
}