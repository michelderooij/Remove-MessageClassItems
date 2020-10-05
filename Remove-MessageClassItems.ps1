<#
    .SYNOPSIS
    Remove-MessageClassItems

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.84, October 6th, 2020

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
    Microsoft Exchange Web Services (EWS) Managed API 1.2 or up is required.
    Recommended EWS.WebServices.Managed.Api (see https://eightwone.com/2020/10/05/ews-webservices-managed-api)

    Search order for Microsoft.Exchange.WebServices.dll: 
    Script Folder, EWS.WebServices.Managed.Api (package), EWS Managed API (install)

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

    .PARAMETER Server
    Exchange Client Access Server to use for Exchange Web Services. When ommited, script will attempt to
    use Autodiscover.

    .PARAMETER Credentials
    Specify credentials to use. When not specified, current credentials are used.
    Credentials can be set using $Credentials= Get-Credential

    .PARAMETER Impersonation
    When specified, uses impersonation for mailbox access, otherwise current logged on user is used.
    For details on how to configure impersonation access for Exchange 2010 using RBAC, see this article:
    http://msdn.microsoft.com/en-us/library/exchange/bb204095(v=exchg.140).aspx
    For details on how to configure impersonation for Exchange 2007, see KB article:
    http://msdn.microsoft.com/en-us/library/exchange/bb204095%28v=exchg.80%29.aspx

    .PARAMETER DeleteMode
    Determines how to remove messages. Options are:
    - HardDelete:         Items will be permanently deleted.
    - SoftDelete:         Items will be moved to the dumpster (default).
    - MoveToDeletedItems: Items will be moved to the Deleted Items folder.
    When using MoveToDeletedItems, the Deleted Items folder will not be processed.

    .PARAMETER ScanAllFolders
    Specifies if you want to scan all folders or only IPF.Note folders (default).
    Useful when there are folders in the mailbox of which Folder Class isn't set.

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

    .EXAMPLE
    .\Remove-MessageClassItems.ps1 -Identity user1 -Impersonation -Verbose -DeleteMode MoveToDeletedItems -MessageClass IPM.Note.EnterpriseVault.Shortcut

    Process mailbox of user1, moving "IPM.Note.EnterpriseVault.Shortcut" message class items to the
    DeletedItems folder, using impersonation and verbose output.

    .EXAMPLE
    .\Remove-MessageClassItems.ps1 -Identity user1 -Impersonation -Verbose -DeleteMode SoftDelete -MessageClass EnterpriseVault -PartialMatching -ScanAllFolders -Before ((Get-Date).AddDays(-90))

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
    Import-CSV users.csv1 | .\Remove-MessageClassItems.ps1 -Impersonation -DeleteMode HardDelete -MessageClass IPM.ixos-archive

    Uses a CSV file to fix specified mailboxes (containing Identity column), removing "IPM.ixos-archive" items permanently, using impersonation.

#>

[cmdletbinding(
    SupportsShouldProcess = $true,
    ConfirmImpact = "High"
)]
param(
    [parameter( Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "All")]
    [alias('Mailbox')]
    [string]$Identity,
    [parameter( Position = 1, Mandatory = $true, ParameterSetName = "All")]
    [ValidateLength(1, 255)]
    [string]$MessageClass,
    [parameter( Position = 2, ParameterSetName = "All")]
    [ValidateLength(1, 255)]
    [string]$ReplaceClass,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [string]$Server,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$Impersonation,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [ValidateSet("HardDelete", "SoftDelete", "MoveToDeletedItems")]
    [string]$DeleteMode = 'SoftDelete',
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$ScanAllFolders,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [datetime]$Before,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [parameter( Mandatory = $false, ParameterSetName = "MailboxOnly")]
    [switch]$MailboxOnly,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [parameter( Mandatory = $false, ParameterSetName = "ArchiveOnly")]
    [switch]$ArchiveOnly,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [string[]]$IncludeFolders,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [string[]]$ExcludeFolders,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$NoProgressBar,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$Report
)

process {

    # Process folders these batches
    $script:MaxFolderBatchSize = 100
    # Process items in these page sizes
    $script:MaxItemBatchSize = 1000
    # Max of concurrent item deletes
    $script:MaxDeleteBatchSize = 100

    # Initial sleep timer (ms) and treshold before lowering
    $script:SleepTimerMax = 300000               # Maximum delay (5min)
    $script:SleepTimerMin = 100                  # Minimum delay
    $script:SleepAdjustmentFactor = 2.0          # When tuning, use this factor
    $script:SleepTimer = $script:SleepTimerMin   # Initial sleep timer value

    # Errors
    $ERR_EWSDLLNOTFOUND = 1000
    $ERR_EWSLOADING = 1001
    $ERR_MAILBOXNOTFOUND = 1002
    $ERR_AUTODISCOVERFAILED = 1003
    $ERR_CANTACCESSMAILBOXSTORE = 1004
    $ERR_PROCESSINGMAILBOX = 1005
    $ERR_PROCESSINGARCHIVE = 1006
    $ERR_INVALIDCREDENTIALS = 1007

    Function Get-EmailAddress( $Identity) {
        $address = [regex]::Match([string]$Identity, ".*@.*\..*", "IgnoreCase")
        if ( $address.Success ) {
            return $address.value.ToString()
        }
        Else {
            # Use local AD to look up e-mail address using $Identity as CN or SamAccountName
            $ADSearch = New-Object DirectoryServices.DirectorySearcher( [ADSI]"")
            $ADSearch.Filter = "(|(cn=$Identity)(samAccountName=$Identity)(mail=$Identity))"
            $Result = $ADSearch.FindOne()
            If ( $Result) {
                $objUser = $Result.getDirectoryEntry()
                return $objUser.mail.toString()
            }
            else {
                return $null
            }
        }
    }

    Function Load-EWSManagedAPIDLL {
        $EWSDLL = 'Microsoft.Exchange.WebServices.dll'
        If ( Test-Path "$pwd\$EWSDLL") {
            $EWSDLLPath = "$pwd"
        }
        Else {

            If( Get-Command -Name Get-Package -ErrorAction SilentlyContinue) {
                If( Get-Package -Name Exchange.WebServices.Managed.Api -ErrorAction SilentlyContinue) {
                    EWSDLLPath= (Get-ChildItem -ErrorAction SilentlyContinue -Path (Split-Path -Parent (get-Package Exchange.WebServices.Managed.Api).Source) -Filter $EWSDLL -Recurse | Sort -Descending -Property LastWriteTime | Select -First 1).DirectoryName
                }    
            }
            If($null -eq $EWSDLLPath) {
                $EWSDLLPath = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory'))
                if (!( Test-Path "$EWSDLLPath\$EWSDLL")) {
                    Write-Error "This script requires EWS Managed API 1.2 or later to be installed, or the Microsoft.Exchange.WebServices.DLL in the current folder."
                    Write-Error "You can download and install EWS Managed API from http://go.microsoft.com/fwlink/?LinkId=255472"
                    Exit $ERR_EWSDLLNOTFOUND
                }
            }
        }

        Write-Verbose "Loading $EWSDLLPath\$EWSDLL"
        try {
            # EX2010
            If (!( Get-Module Microsoft.Exchange.WebServices)) {
                Import-Module "$EWSDLLPATH\$EWSDLL"
            }
        }
        catch {
            #<= EX2010
            [void][Reflection.Assembly]::LoadFile( "$EWSDLLPath\$EWSDLL")
        }
        try {
            $Temp = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
        }
        catch {
            Write-Error "Problem loading $EWSDLL"
            Exit $ERR_EWSLOADING
        }
        $DLLObj = Get-ChildItem -Path "$EWSDLLPATH\$EWSDLL" -ErrorAction SilentlyContinue
        If ( $DLLObj) {
            Write-Verbose ('Loaded EWS Managed API v{0}' -f $DLLObj.VersionInfo.FileVersion)
        }
    }

    # After calling this any SSL Warning issues caused by Self Signed Certificates will be ignored
    # Source: http://poshcode.org/624
    Function set-TrustAllWeb() {
        Write-Verbose "Set to trust all certificates"
        $Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
        $Compiler = $Provider.CreateCompiler()
        $Params = New-Object System.CodeDom.Compiler.CompilerParameters
        $Params.GenerateExecutable = $False
        $Params.GenerateInMemory = $True
        $Params.IncludeDebugInformation = $False
        $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

        $TASource = @'
            namespace Local.ToolkitExtensions.Net.CertificatePolicy {
                public class TrustAll : System.Net.ICertificatePolicy {
                    public TrustAll() {
                    }
                    public bool CheckValidationResult(System.Net.ServicePoint sp, System.Security.Cryptography.X509Certificates.X509Certificate cert,   System.Net.WebRequest req, int problem) {
                        return true;
                    }
                }
            }
'@

        $TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
        $TAAssembly = $TAResults.CompiledAssembly
        $TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
        [System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
    }

    Function iif( $eval, $tv = '', $fv = '') {
        If ( $eval) { return $tv } else { return $fv}
    }

    Function Construct-FolderFilter {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string[]]$Folders,
            [string]$emailAddress
        )
        If ( $Folders) {
            $FolderFilterSet = @()
            ForEach ( $Folder in $Folders) {
                # Convert simple filter to (simple) regexp
                $Parts = $Folder -match '^(?<root>\\)?(?<keywords>.*?)?(?<sub>\\\*)?$'
                If ( !$Parts) {
                    Write-Error ('Invalid regular expression matching against {0}' -f $Folder)
                }
                Else {
                    $Keywords = Search-ReplaceWellKnownFolderNames $EwsService ($Matches.keywords) $emailAddress
                    $EscKeywords = [Regex]::Escape( $Keywords) -replace '\\\*', '.*'
                    $Pattern = iif -eval $Matches.Root -tv '^\\' -fv '^\\(.*\\)*'
                    $Pattern += iif -eval $EscKeywords -tv $EscKeywords -fv ''
                    $Pattern += iif -eval $Matches.sub -tv '(\\.*)?$' -fv '$'
                    $Obj = New-Object -TypeName PSObject -Prop @{
                        'Pattern'     = $Pattern;
                        'IncludeSubs' = -not [string]::IsNullOrEmpty( $Matches.Sub)
                        'OrigFilter'  = $Folder
                    }
                    $FolderFilterSet += $Obj
                    Write-Debug ($Obj -join ',')
                }
            }
        }
        Else {
            $FolderFilterSet = $null
        }
        return $FolderFilterSet
    }

    Function Search-ReplaceWellKnownFolderNames {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$criteria = '',
            [string]$emailAddress
        )
        $AllowedWKF = 'Inbox', 'Calendar', 'Contacts', 'Notes', 'SentItems', 'Tasks', 'JunkEmail', 'DeletedItems'
        # Construct regexp to see if allowed WKF is part of criteria string
        ForEach ( $ThisWKF in $AllowedWKF) {
            If ( $criteria -match '#{0}#') {
                $criteria = $criteria -replace ('#{0}#' -f $ThisWKF), (myEWSBind-WellKnownFolder $EwsService $ThisWKF $emailAddress).DisplayName
            }
        }
        return $criteria
    }
    Function Tune-SleepTimer {
        param(
            [bool]$previousResultSuccess = $false
        )
        if ( $previousResultSuccess) {
            If ( $script:SleepTimer -gt $script:SleepTimerMin) {
                $script:SleepTimer = [int]([math]::Max( [int]($script:SleepTimer / $script:SleepAdjustmentFactor), $script:SleepTimerMin))
                Write-Warning ('Previous EWS operation successful, adjusted sleep timer to {0}ms' -f $script:SleepTimer)
            }
        }
        Else {
            $script:SleepTimer = [int]([math]::Min( ($script:SleepTimer * $script:SleepAdjustmentFactor) + 100, $script:SleepTimerMax))
            If ( $script:SleepTimer -eq 0) {
                $script:SleepTimer = 5000
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
        $OpSuccess = $false
        $CritErr = $false
        Do {
            Try {
                $res = $EwsService.FindFolders( $FolderId, $FolderSearchCollection, $FolderView)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $CritErr = $false
        Do {
            Try {
                $res = $EwsService.FindFolders( $FolderId, $FolderView)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $CritErr = $false
        Do {
            Try {
                $res = $Folder.FindItems( $ItemSearchFilterCollection, $ItemView)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $CritErr = $false
        Do {
            Try {
                $res = $Folder.FindItems( $ItemView)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $critErr = $false
        Do {
            Try {
                If ( @([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013, [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1) -contains $EwsService.RequestedServerVersion) {
                    $res = $EwsService.DeleteItems( $ItemIds, $DeleteMode, $SendCancellationsMode, $AffectedTaskOccurrences, $SuppressReadReceipt)
                }
                Else {
                    $res = $EwsService.DeleteItems( $ItemIds, $DeleteMode, $SendCancellationsMode, $AffectedTaskOccurrences)
                }
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $critErr = $false
        Do {
            Try {
                $explicitFolder= New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName, $emailAddress)  
                $res = [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $explicitFolder)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
                Write-Warning ('Cannot bind to {0} - skipping. Error: {1}' -f $WellKnownFolderName, $Error[0])
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
            $ExcludeFilter,
            $ScanAllFolders
        )
        $FoldersToProcess = [System.Collections.ArrayList]@()
        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView( $MaxFolderBatchSize)
        $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
        $FolderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass)
        $FolderSearchCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection( [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        If ( -not $ScanAllFolders) {
            $FolderSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Note")
            $FolderSearchCollection.Add( $FolderSearchFilter)
        }

        Do {
            If ( $FolderSearchCollection.Count -ge 1) {
                $FolderSearchResults = myEWSFind-Folders $EwsService $Folder.Id $FolderSearchCollection $FolderView
            }
            Else {
                $FolderSearchResults = myEWSFind-FoldersNoSearch $EwsService $Folder.Id $FolderView
            }
            ForEach ( $FolderItem in $FolderSearchResults) {
                $FolderPath = '{0}\{1}' -f $CurrentPath, $FolderItem.DisplayName
                If ( $IncludeFilter) {
                    $Add = $false
                    # Defaults to true, unless include does not specifically include subfolders
                    $Subs = $true
                    ForEach ( $Filter in $IncludeFilter) {
                        If ( $FolderPath -match $Filter.Pattern) {
                            $Add = $true
                            # When multiple criteria match, one with and one without subfolder processing, subfolders will be processed.
                            $Subs = $Filter.IncludeSubs
                        }
                    }
                }
                Else {
                    # If no includeFolders specified, include all (unless excluded later on)
                    $Add = $true
                    $Subs = $true
                }
                If ( $ExcludeFilter) {
                    # Excludes can overrule includes
                    ForEach ( $Filter in $ExcludeFilter) {
                        If ( $FolderPath -match $Filter.Pattern) {
                            $Add = $false
                            # When multiple criteria match, one with and one without subfolder processing, subfolders will be processed.
                            $Subs = $Filter.IncludeSubs
                        }
                    }
                }
                If ( $Add) {
                    Write-Verbose ( 'Adding folder {0}' -f $FolderPath, $Prio)

                    $Obj = New-Object -TypeName PSObject -Property @{
                        'Name'   = $FolderPath;
                        'Folder' = $FolderItem
                    }
                    $FoldersToProcess.Add( $Obj) | Out-Null
                }
                If ( $Subs) {
                    # Could be that specific folder is to be excluded, but subfolders needs evaluation, or specific folder without subfolders excluded
                    $SubFolders = Get-SubFolders -Folder $FolderItem -CurrentPath $FolderPath -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter -ScanAllFolders $ScanAllFolders
                    ForEach ( $AddFolder in $Subfolders) {
                        $FoldersToProcess.Add( $AddFolder)  | Out-Null
                    }
                }
            }
            $FolderView.Offset += $FolderSearchResults.Folders.Count
        } While ($FolderSearchResults.MoreAvailable)
        Write-Output -NoEnumerate $FoldersToProcess
    }

    Function Process-Mailbox {
        param(
            [string]$Identity,
            $Folder,
            $IncludeFilter,
            $ExcludeFilter,
            $emailAddress
        )

        $ProcessingOK = $True
        $TotalMatch = 0
        $TotalRemoved = 0
        $FoldersFound = 0
        $FoldersProcessed = 0
        $TimeProcessingStart = Get-Date
        $DeletedItemsFolder = myEWSBind-WellKnownFolder $EwsService 'DeletedItems' $emailAddress

        # Build list of folders to process
        Write-Verbose (iif $ScanAllFolders -fv 'Collecting folders containing e-mail items to process' -tv 'Collecting folders to process')
        $FoldersToProcess = Get-SubFolders -Folder $Folder -CurrentPath '' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -ScanAllFolders $ScanAllFolders

        $FoldersFound = $FoldersToProcess.Count
        Write-Verbose ('Found {0} folders matching folder search criteria' -f $FoldersFound)

        ForEach ( $SubFolder in $FoldersToProcess) {
            If (!$NoProgressBar) {
                Write-Progress -Id 1 -Activity "Processing $Identity" -Status "Processed folder $FoldersProcessed of $FoldersFound" -PercentComplete ( $FoldersProcessed / $FoldersFound * 100)
            }
            If ( ! ( $DeleteMode -eq 'MoveToDeletedItems' -and $SubFolder.Id -eq $DeletedItemsFolder.Id)) {
                If ( $Report.IsPresent) {
                    Write-Host ('Processing folder {0}' -f $SubFolder.Name)
                }
                Else {
                    Write-Verbose ('Processing folder {0}' -f $SubFolder.Name)
                }
                $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView( $script:MaxItemBatchSize, 0, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
                $ItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
                $ItemView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
                    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
                    [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass)

                $ItemSearchFilterCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
                If ($MessageClass -match '^\*(?<substring>.*?)\*$') {
                    $ItemSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
                        $matches['substring'], [Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
                }
                Else {
                    If ($MessageClass -match '^(?<prefix>.*?)\*$') {
                        $ItemSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
                            $matches['prefix'], [Microsoft.Exchange.WebServices.Data.ContainmentMode]::Prefixed, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
                    }
                    Else {
                        $ItemSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, $MessageClass)
                    }
                }
                $ItemSearchFilterCollection.add( $ItemSearchFilter)

                If ( $Before) {
                    $ItemSearchFilterCollection.add( (New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan( [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $Before)))
                }

                $ProcessList = [System.Collections.ArrayList]@()
                If ( $psversiontable.psversion.major -lt 3) {
                    $ItemIds = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.ItemId]))
                }
                Else {
                    $type = ("System.Collections.Generic.List" + '`' + "1") -as 'Type'
                    $type = $type.MakeGenericType([Microsoft.Exchange.WebServices.Data.ItemId] -as 'Type')
                    $ItemIds = [Activator]::CreateInstance($type)
                }

                Do {
                    $ItemSearchResults = MyEWSFind-Items $SubFolder.Folder $ItemSearchFilterCollection $ItemView
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
            If ( ($ProcessList.Count -gt 0) -and $PSCmdlet.ShouldProcess( ('{0} item(s) from {1}' -f $ProcessList.Count, $SubFolder.Name))) {
                If ( $ReplaceClass) {
                    Write-Verbose ('Modifying {0} items from {1}' -f $ProcessList.Count, $SubFolder.Name)
                    $ItemsChanged = 0
                    $ItemsRemaining = $ProcessList.Count
                    ForEach ( $ItemID in $ProcessList) {
                        If (!$NoProgressBar) {
                            Write-Progress -Id 2 -Activity "Processing folder $($SubFolder.DisplayName)" -Status "Items processed $ItemsChanged - remaining $ItemsRemaining" -PercentComplete ( $ItemsRemoved / $ProcessList.Count * 100)
                        }
                        Try {
                            $ItemObj = [Microsoft.Exchange.WebServices.Data.Item]::Bind( $EwsService, $ItemID)
                            $ItemObj.ItemClass = $ReplaceClass
                            $ItemObj.Update( [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
                        }
                        Catch {
                            Write-Error ('Problem modifying item: {0}' -f $error[0])
                            $ProcessingOK = $False
                        }
                        $ItemsChanged++
                        $ItemsRemaining--
                    }
                }
                Else {
                    Write-Verbose ('Removing {0} items from {1}' -f $ProcessList.Count, $SubFolder.Name)

                    $SendCancellationsMode = [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone
                    $AffectedTaskOccurrences = [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::SpecifiedOccurrenceOnly
                    $SuppressReadReceipt = $true # Only works using EWS with Exchange2013+ mode

                    $ItemsRemoved = 0
                    $ItemsRemaining = $ProcessList.Count

                    # Remove ItemIDs in batches
                    ForEach ( $ItemID in $ProcessList) {
                        $ItemIds.Add( $ItemID)
                        If ( $ItemIds.Count -eq $script:MaxDeleteBatchSize) {
                            $ItemsRemoved += $ItemIds.Count
                            $ItemsRemaining -= $ItemIds.Count
                            If (!$NoProgressBar) {
                                Write-Progress -Id 2 -Activity "Processing folder $($SubFolder.DisplayName)" -Status "Items processed $ItemsRemoved - remaining $ItemsRemaining" -PercentComplete ( $ItemsRemoved / $ProcessList.Count * 100)
                            }
                            $res = myEWSRemove-Items $EwsService $ItemIds $DeleteMode $SendCancellationsMode $AffectedTaskOccurrences $SuppressReadReceipt
                            $ItemIds.Clear()
                        }
                    }
                    # .. also remove last ItemIDs
                    If ( $ItemIds.Count -gt 0) {
                        $ItemsRemoved += $ItemIds.Count
                        $ItemsRemaining = 0
                        $res = myEWSRemove-Items $EwsService $ItemIds $DeleteMode $SendCancellationsMode $AffectedTaskOccurrences $SuppressReadReceipt
                        $ItemIds.Clear()
                    }
                }
                If (!$NoProgressBar) {
                    Write-Progress -Id 2 -Activity "Processing folder $($SubFolder.DisplayName)" -Status 'Finished processing.' -Completed
                }
            }
            Else {
                # No matches
            }
            $FoldersProcessed++
        } # ForEach SubFolder
        If (!$NoProgressBar) {
            Write-Progress -Id 1 -Activity "Processing $Identity" -Status "Finished processing." -Completed
        }
        If ( $ProcessingOK) {
            $TimeProcessingDiff = (Get-Date) - $TimeProcessingStart
            $Speed = [int]( $TotalMatch / $TimeProcessingDiff.TotalSeconds * 60)
            Write-Verbose ('{0} item(s) processed in {1:hh}:{1:mm}:{1:ss} - average {2} items/min' -f $TotalMatch, $TimeProcessingDiff, $Speed)
        }
        Return $ProcessingOK
    }

    ##################################################
    # Main
    ##################################################
    #Requires -Version 3

    Load-EWSManagedAPIDLL
    set-TrustAllWeb

    If ( $MailboxOnly) {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
    }
    Else {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    }

    $EwsService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService( $ExchangeVersion)
    If ( $Credentials) {
        try {
            Write-Verbose ('Using credentials {0}' -f $Credentials.UserName)
            $EwsService.Credentials = New-Object System.Net.NetworkCredential( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
        }
        catch {
            Write-Error ('Invalid credentials provided, error: {0}' -f $error[0])
            Exit $ERR_INVALIDCREDENTIALS
        }
    }
    Else {
        $EwsService.UseDefaultCredentials = $true
    }

    ForEach ( $CurrentIdentity in $Identity) {

        $EmailAddress = get-EmailAddress $CurrentIdentity
        If ( !$EmailAddress) {
            Write-Error ('Specified mailbox {0} not found' -f $CurrentIdentity)
            Exit $ERR_MAILBOXNOTFOUND
        }

        Write-Host ('Processing mailbox {0} ({1})' -f $CurrentIdentity, $EmailAddress)

        If ( $Impersonation) {
            Write-Verbose ('Using {0} for impersonation' -f $EmailAddress)
            $EwsService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
            $EwsService.HttpHeaders.Add("X-AnchorMailbox", $EmailAddress)
        }

        If ($Server) {
            $EwsUrl = ('https://{0}/EWS/Exchange.asmx' -f $Server)
            Write-Verbose ('Using Exchange Web Services URL {0}' -f $EwsUrl)
            $EwsService.Url = $EwsUrl
        }
        Else {
            Write-Verbose ('Looking up EWS URL using Autodiscover for {0}' -f $EmailAddress)
            try {
                # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
                $ErrorActionPreference = 'Stop'
                $EwsService.autodiscoverUrl( $EmailAddress, {$true})
            }
            catch {
                Write-Error ('Autodiscover failed, error: {0}' -f $_.Exception.Message)
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference = 'Continue'
            Write-Verbose ('Using EWS on CAS {0}' -f $EwsService.Url)
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

        # Construct search filters
        Write-Verbose 'Constructing folder matching rules'
        $IncludeFilter = Construct-FolderFilter $EwsService $IncludeFolders
        $ExcludeFilter = Construct-FolderFilter $EwsService $ExcludeFolders

        If ( -not $ArchiveOnly.IsPresent) {
            try {
                $RootFolder = myEWSBind-WellKnownFolder $EwsService 'MsgFolderRoot' $emailAddress
                If ( $RootFolder) {
                    Write-Verbose ('Processing primary mailbox {0}' -f $Identity)
                    If (! ( Process-Mailbox -Identity $Identity -Folder $RootFolder -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -emailAddress $emailAddress)) {
                        Write-Error ('Problem processing primary mailbox of {0} ({1})' -f $CurrentIdentity, $EmailAddress)
                        Exit $ERR_PROCESSINGMAILBOX
                    }
                }
            }
            catch {
                Write-Error ('Cannot access mailbox information store, error: {0}' -f $Error[0])
                Exit $ERR_CANTACCESSMAILBOXSTORE
            }
        }

        If ( -not $MailboxOnly.IsPresent) {
            try {
                $ArchiveRootFolder = myEWSBind-WellKnownFolder $EwsService 'ArchiveMsgFolderRoot' $emailAddress
                If ( $ArchiveRootFolder) {
                    Write-Verbose ('Processing archive mailbox {0}' -f $Identity)
                    If (! ( Process-Mailbox -Identity $Identity -Folder $ArchiveRootFolder -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -emailAddress $emailAddress)) {
                        Write-Error ('Problem processing archive mailbox of {0} ({1})' -f $CurrentIdentity, $EmailAddress)
                        Exit $ERR_PROCESSINGARCHIVE
                    }
                }
            }
            catch {
                Write-Debug 'No archive configured or cannot access archive'
            }
        }
        Write-Verbose ('Processing {0} finished' -f $CurrentIdentity)
    }
}
