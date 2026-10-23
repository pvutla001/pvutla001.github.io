#Requires -Version 5.1
<#
.SYNOPSIS
Diagnoses SharePoint Online file locks and repairs explicit checkouts.
.DESCRIPTION
Version 2.1. Default is read-only Diagnose. No Graph calls, device-code login,
file deletion/replacement, compliance-policy changes or invented force-unlock API.
Uses PnP.PowerShell browser MFA or a supplied current PnP connection.
Editing-session locks must end at the client or expire; persistent locks are
reported for Microsoft support. A clear snapshot does not prove edit permission.
LibraryUrl scans all visible file types recursively without content downloads.
FileUrl accepts a direct HTTPS file URL. Unlock is a preserving CheckIn alias,
not a force-release operation. See README-v2.txt for authentication and evidence.
Every run writes UTC-stamped log, JSON and CSV beside this script by default.
CheckIn/DiscardCheckout require ClientSessionsClosed and a local content backup.
ClientEnvironment/ClientSessionState/SessionHost/SessionId are operator-supplied
incident context, not discovered from SharePoint. Disconnected or unreachable
sessions block repair. No arbitrary Office lock or Citrix session force termination.
Do not dot-source this script: it sets a process exit code when it finishes.
.EXAMPLE
./Repair-SPOFileLock.ps1 -SiteUrl https://contoso.sharepoint.com/sites/QA -FilePath '/sites/QA/LockLab/Test.doc' -ClientId '<app-id>'
.EXAMPLE
./Repair-SPOFileLock.ps1 -SiteUrl https://contoso.sharepoint.com/sites/QA -FilePath '/sites/QA/LockLab/Test.doc' -Connection $conn -Action CheckIn -ClientSessionsClosed -WhatIf
.EXAMPLE
./Repair-SPOFileLock.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/QA/Documents' -AuthenticationMode LegacyWebLogin
.EXAMPLE
./Repair-SPOFileLock.ps1 -FileUrl 'https://contoso.sharepoint.com/sites/QA/Documents/Test.docx' -Action Unlock -ClientSessionsClosed -ClientId '<app-id>'
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName='Path')]
param(
    [Parameter(Mandatory,ParameterSetName='Path')][Parameter(ParameterSetName='FileUrl')][Parameter(ParameterSetName='Library')][string]$SiteUrl,
    # DECODED server-relative file path. Literal %, # and apostrophes are allowed.
    # Do not supply a sharing link, Doc.aspx URL, or URL-encoded path.
    [Parameter(Mandatory,ParameterSetName='Path')][ValidateNotNullOrEmpty()][string]$FilePath,
    [Parameter(Mandatory,ParameterSetName='FileUrl')][string]$FileUrl,
    [Parameter(Mandatory,ParameterSetName='Library')][string]$LibraryUrl,
    [ValidateSet('Interactive','LegacyWebLogin')][string]$AuthenticationMode='Interactive',
    [string]$ClientId,
    [object]$Connection,
    [ValidateSet('Diagnose','WaitForRelease','Unlock','CheckIn','DiscardCheckout')]
    [string]$Action = 'Diagnose',
    [switch]$ClientSessionsClosed,
    [switch]$AcknowledgeDiscardChanges,
    [ValidateSet('Unknown','LocalDesktop','Citrix','AzureVirtualDesktop','Windows365','OtherVDI')][string]$ClientEnvironment='Unknown',
    [ValidateSet('Unknown','Open','Disconnected','Unreachable')][string]$ClientSessionState='Unknown',
    [ValidateLength(0,256)][string]$SessionHost,
    [ValidateLength(0,128)][string]$SessionId,
    [ValidateLength(0,128)][string]$IncidentReference,
    [ValidateLength(0,2000)][string]$ReportedError,
    [ValidateRange(0,3600)][int]$WaitSeconds = 600,
    [ValidateRange(5,60)][int]$PollSeconds = 15,
    [string]$OutputDirectory,
    [ValidateLength(1,200)][string]$Comment = 'Checkout recovery after user confirmed saved changes and closed editing sessions.'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $OutputDirectory) { $OutputDirectory=Join-Path $PSScriptRoot 'SPO-Lock-Reports' }
$runId = [DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss_fff') + '_' + [guid]::NewGuid().ToString('N').Substring(0,8)
# Local audit files are written even for -WhatIf. All remote mutations are gated.
[void][System.IO.Directory]::CreateDirectory($OutputDirectory)
$runDirectory = Join-Path ([System.IO.Path]::GetFullPath($OutputDirectory)) $runId
[void][System.IO.Directory]::CreateDirectory($runDirectory)
$logPath = Join-Path $runDirectory "SPO-Lock_$runId.log"
$reportPath = Join-Path $runDirectory "SPO-Lock_$runId.json"
$csvPath = Join-Path $runDirectory "SPO-Lock_$runId.csv"
$summaryPath = Join-Path $runDirectory "SPO-Lock_$runId.txt"
$script:snapshots = [System.Collections.Generic.List[object]]::new()
$script:requests = [System.Collections.Generic.List[object]]::new()
$report = [ordered]@{
    ScriptVersion = '2.1'; RunId = $runId; StartedUtc = [DateTime]::UtcNow.ToString('o')
    SiteUrl = $SiteUrl; FilePath = $FilePath; RequestedAction = $Action
    ClientSessionsClosed = [bool]$ClientSessionsClosed; WhatIf = [bool]$WhatIfPreference
    Status = 'Starting'; ExitCode = 1; MutationAttempted = $false; MutationCompleted = $false
    Library = $null; SignedInUser = $null; PnPVersion = $null; Backup = $null
    Error = $null; FinalSnapshot = $null; FinishedUtc = $null
    LibraryUrl=$LibraryUrl; FileUrl=$FileUrl; AuthenticationMode=$AuthenticationMode
    EnumerationComplete=$false; FilesScanned=0; FilesWithSignals=0; FileReadErrors=0
    MetadataReadErrors=0; FileMetadata=$null
    ClientContext=[ordered]@{Source='OperatorSuppliedNotVerified';Environment=$ClientEnvironment;SessionState=$ClientSessionState;Host=$SessionHost;SessionId=$SessionId;IncidentReference=$IncidentReference;ReportedError=$ReportedError}
    Capabilities=[ordered]@{RepairsExplicitCheckouts=$true;ForceReleaseArbitraryEditingLock=$false;ControlsRemoteDesktopSessions=$false;RecoversUnsavedClientEdits=$false}
}
$script:libraryResults=[System.Collections.Generic.List[object]]::new()
$script:conn = $null
$script:targetEndpoint = $null
$script:expectedId = $null

function Write-RunLog([string]$Message) {
    $line = '{0} {1}' -f [DateTime]::UtcNow.ToString('o'), $Message
    [System.IO.File]::AppendAllText($logPath, $line + [Environment]::NewLine)
    Write-Host $line
}

function Read-RequiredProperty($Value, [string]$Name) {
    if ($null -eq $Value -or $null -eq $Value.PSObject.Properties[$Name]) {
        throw "Required response property '$Name' is absent; lock state is unknown."
    }
    return $Value.$Name
}

function Read-OptionalProperty($Value,[string]$Name) {
    if ($null -eq $Value) { return $null }
    try {
        if ($null -ne $Value.PSObject.Properties[$Name]) { return $Value.$Name }
    } catch { } # Optional/uninitialized diagnostic properties are not lock evidence.
    return $null
}

function ConvertTo-UtcText($Value) {
    if ($null -eq $Value) { return $null }
    try { return ([datetime]$Value).ToUniversalTime().ToString('o') } catch { return $null }
}

function ConvertTo-PersonMetadata($Value) {
    if ($null -eq $Value) { return $null }
    return [pscustomobject]@{Id=(Read-OptionalProperty $Value 'LookupId');DisplayName=(Read-OptionalProperty $Value 'LookupValue');Email=(Read-OptionalProperty $Value 'Email')}
}

function ConvertTo-ItemMetadata($Fields) {
    return [pscustomobject]@{
        Status='Read';Error=$null;ItemId=$Fields['ID'];Title=$Fields['Title']
        CreatedUtc=(ConvertTo-UtcText $Fields['Created']);ModifiedUtc=(ConvertTo-UtcText $Fields['Modified'])
        CreatedBy=(ConvertTo-PersonMetadata $Fields['Author']);ModifiedBy=(ConvertTo-PersonMetadata $Fields['Editor'])
        ContentTypeId=[string]$Fields['ContentTypeId']
    }
}

function Read-FileMetadata([string]$UniqueId) {
    try {
        $ctx=Get-PnPContext -Connection $script:conn
        $item=$ctx.Web.GetFileById([guid]$UniqueId).ListItemAllFields
        $ctx.Load($item); $ctx.ExecuteQuery()
        $script:requests.Add([pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('o');Endpoint='CSOM file item metadata';RequestId=(Read-OptionalProperty $ctx 'TraceCorrelationId');Success=$true})
        return ConvertTo-ItemMetadata $item.FieldValues
    } catch {
        $report.MetadataReadErrors++
        Write-RunLog "Supplemental metadata unavailable: $($_.Exception.Message)"
        return [pscustomobject]@{Status='Unavailable';Error=$_.Exception.Message}
    }
}

function Get-RecoveryAdvice($Snapshot) {
    $state=Read-OptionalProperty $Snapshot 'State'
    if (-not $state -or $state -in @('UnknownReadFailed','InconsistentState')) { return 'Resolve read errors or inconsistent state before repair; no unlocked state can be inferred.' }
    if ($state -eq 'NoCheckoutOrEditingLockReported') { return 'No checkout/editing signal at this observation. If opening or saving fails, capture the actual error and investigate the client, permissions, format, validation and policy restrictions.' }
    if ($ClientSessionState -in @('Disconnected','Unreachable')) {
        return 'Remote session recovery required. Ask the VDI administrator to identify the exact user/host/session, reconnect or shadow it and preserve edits. Logging off or terminating that session can lose unsaved work and requires a separate deliberate decision. SharePoint cannot save those edits. After session closure, use WaitForRelease; escalate persistent server locks with these reports.'
    }
    if ((Read-OptionalProperty $Snapshot 'CheckOutType') -ne 2) { return 'Explicit checkout reported. Preserve edits and confirm all editing sessions closed, then use CheckIn/Unlock. DiscardCheckout separately requires acknowledgement and discards checkout changes. A remaining editing lock is not cleared by claiming checkout success.' }
    if ($ClientSessionsClosed) { return 'Editing lock remains despite operator-confirmed closure. Check other devices, background Office processes and disconnected VDI sessions. Use WaitForRelease; if it persists, provide the report, closure time and actual error to Microsoft support. This is not proof that the lock is orphaned.' }
    return 'Editing lock reported. Save and close all known editing sessions, including remote desktops and other devices, then use WaitForRelease. There is no arbitrary editing-lock force release in this script.'
}

function Get-ObservationSummary([object[]]$Samples) {
    if ($Samples.Count -eq 0) { return $null }
    $signals=@($Samples | Where-Object {$_.State -in @('CheckedOut','EditingLockReported','CheckoutAndLockReported')})
    $first=$Samples[0]; $last=$Samples[-1]
    $firstSignal=$null; $lastSignal=$null
    if ($signals.Count) { $firstSignal=$signals[0].Utc; $lastSignal=$signals[-1].Utc }
    $conclusion='PointInTimeObservationsOnly'
    if ($report.RequestedAction -eq 'WaitForRelease') {
        if ($last.State -eq 'NoCheckoutOrEditingLockReported') { $conclusion='ClearAtLastObservation' }
        elseif ($ClientSessionsClosed) { $conclusion='SignalRemainsAfterOperatorConfirmedClosure_NotProofOfOrphanedLock' }
        else { $conclusion='SignalRemains_ClientClosureUnconfirmed' }
    }
    return [pscustomobject]@{FirstObservedUtc=$first.Utc;LastObservedUtc=$last.Utc;ObservationCount=$Samples.Count;ObservedIntervalSeconds=[math]::Round((([datetime]$last.Utc)-([datetime]$first.Utc)).TotalSeconds,2);RequestedWaitSeconds=$WaitSeconds;PollSeconds=$PollSeconds;SignalObservationCount=$signals.Count;FirstSignalObservedUtc=$firstSignal;LastSignalObservedUtc=$lastSignal;SignalAtEveryObservation=($signals.Count -eq $Samples.Count);ClientClosureConfirmed=[bool]$ClientSessionsClosed;Conclusion=$conclusion;Limitation='Sampled observations are not lock age, lease expiry or proof of a continuously held/orphaned lock.'}
}

function ConvertTo-DiagnosticRow($Entry,$Metadata) {
    $owner=Read-OptionalProperty $Entry 'CheckedOutBy'; $locker=Read-OptionalProperty $Entry 'LockedBy'
    $createdBy=Read-OptionalProperty $Metadata 'CreatedBy'; $modifiedBy=Read-OptionalProperty $Metadata 'ModifiedBy'
    $path=Read-OptionalProperty $Entry 'ServerRelativeUrl'; $absoluteUrl=$null
    if ($path -and $SiteUrl) { $absoluteUrl=([uri]$SiteUrl).GetLeftPart([UriPartial]::Authority)+(($path.Split('/') | ForEach-Object {[uri]::EscapeDataString($_)}) -join '/') }
    $row=[ordered]@{
        RunId=$runId;Utc=(Read-OptionalProperty $Entry 'Utc');Phase=(Read-OptionalProperty $Entry 'Phase');State=(Read-OptionalProperty $Entry 'State')
        SiteUrl=$SiteUrl;LibraryId=(Read-OptionalProperty $report.Library 'Id');LibraryTitle=(Read-OptionalProperty $report.Library 'Title')
        ItemId=(Read-OptionalProperty $Metadata 'ItemId');FilePath=$path;FileUrl=$absoluteUrl;UniqueId=(Read-OptionalProperty $Entry 'UniqueId')
        FileName=(Read-OptionalProperty $Entry 'Name');Extension=[IO.Path]::GetExtension([string]$path)
        CheckOutType=(Read-OptionalProperty $Entry 'CheckOutType');CheckOutTypeName=(Read-OptionalProperty $Entry 'CheckOutTypeName')
        CheckedOutBy=(Read-OptionalProperty $owner 'LoginName');CheckoutOwnerName=(Read-OptionalProperty $owner 'Title');CheckoutOwnerEmail=(Read-OptionalProperty $owner 'Email')
        LockedBy=(Read-OptionalProperty $locker 'LoginName');LockOwnerName=(Read-OptionalProperty $locker 'Title');LockOwnerEmail=(Read-OptionalProperty $locker 'Email')
        MajorVersion=(Read-OptionalProperty $Entry 'MajorVersion');MinorVersion=(Read-OptionalProperty $Entry 'MinorVersion');ETag=(Read-OptionalProperty $Entry 'ETag')
        LastModified=(Read-OptionalProperty $Entry 'LastModified');Length=(Read-OptionalProperty $Entry 'Length');CreatedUtc=(Read-OptionalProperty $Metadata 'CreatedUtc')
        CreatedBy=(Read-OptionalProperty $createdBy 'DisplayName');CreatedByEmail=(Read-OptionalProperty $createdBy 'Email')
        ModifiedBy=(Read-OptionalProperty $modifiedBy 'DisplayName');ModifiedByEmail=(Read-OptionalProperty $modifiedBy 'Email');Title=(Read-OptionalProperty $Metadata 'Title');ContentTypeId=(Read-OptionalProperty $Metadata 'ContentTypeId')
        IrmEnabled=(Read-OptionalProperty $Entry 'IrmEnabled');FileLevel=(Read-OptionalProperty $Entry 'FileLevel');RequestId=(Read-OptionalProperty $Entry 'RequestId')
        LibraryForceCheckout=(Read-OptionalProperty $report.Library 'ForceCheckout');LibraryVersioning=(Read-OptionalProperty $report.Library 'EnableVersioning');LibraryMinorVersions=(Read-OptionalProperty $report.Library 'EnableMinorVersions');LibraryModeration=(Read-OptionalProperty $report.Library 'EnableModeration')
        MetadataStatus=(Read-OptionalProperty $Metadata 'Status');MetadataError=(Read-OptionalProperty $Metadata 'Error');Error=(Read-OptionalProperty $Entry 'Error')
        ClientEnvironment=$ClientEnvironment;ClientSessionState=$ClientSessionState;ClientSessionsClosed=[bool]$ClientSessionsClosed;SessionHost=$SessionHost;SessionId=$SessionId;ContextSource='OperatorSuppliedNotVerified'
        IncidentReference=$IncidentReference;ReportedError=$ReportedError;NextAction=(Get-RecoveryAdvice $Entry);FinalStatus=$report.Status;ExitCode=$report.ExitCode
    }
    # Excel should display untrusted names/error text as text, not evaluate formulas.
    foreach ($key in @($row.Keys)) { if ($row[$key] -is [string] -and $row[$key] -match '^[\s]*[=+@-]') { $row[$key]="'"+$row[$key] } }
    return [pscustomobject]$row
}

function Invoke-SPRead([string]$Endpoint) {
    $headers = $null
    try {
        $readArgs=@{Url=$Endpoint;Method='Get';Connection=$script:conn;ErrorAction='Stop'}
        if ((Get-Command Invoke-PnPSPRestMethod).Parameters.ContainsKey('ResponseHeadersVariable')) { $readArgs.ResponseHeadersVariable='headers' }
        $data = Invoke-PnPSPRestMethod @readArgs
        $requestId = $null
        if ($null -ne $headers) { $requestId = [string]$headers['SPRequestGuid'] }
        $script:requests.Add([pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('o'); Endpoint=$Endpoint; RequestId=$requestId; Success=$true})
        if ($null -ne $data.PSObject.Properties['d']) { return $data.d }
        return $data
    } catch {
        $script:requests.Add([pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('o'); Endpoint=$Endpoint; RequestId=$null; Success=$false})
        throw
    }
}

function Read-UserIdentity($User) {
    if ($null -eq $User) { return $null }
    $id = Read-RequiredProperty $User 'Id'
    if ([int]$id -le 0) { throw 'Unexpected lock/checkout user response; refusing to infer unlocked state.' }
    return [pscustomobject]@{
        Id = [int]$id
        LoginName = [string](Read-RequiredProperty $User 'LoginName')
        Title = [string](Read-RequiredProperty $User 'Title')
        Email = Read-OptionalProperty $User 'Email'
    }
}

function Read-LockSnapshot([string]$Phase,[string]$ResolvedPath=$FilePath) {
    # Load navigation properties explicitly. Legacy PnP REST conversion drops
    # null-valued user properties, which must not be mistaken for a clear state.
    $ctx=Get-PnPContext -Connection $script:conn
    if ($script:expectedId) { $file=$ctx.Web.GetFileById([guid]$script:expectedId) }
    else { $file=$ctx.Web.GetFileByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl($ResolvedPath)) }
    $ctx.Load($file)
    $ctx.Load($file.CheckedOutByUser)
    $ctx.Load($file.LockedByUser)
    $ctx.ExecuteQuery()
    $coUser=$null; $lockUser=$null
    if ($null -eq $file.CheckedOutByUser.ServerObjectIsNull -or $null -eq $file.LockedByUser.ServerObjectIsNull) { throw 'User properties were not initialized; lock state unknown.' }
    if (-not $file.CheckedOutByUser.ServerObjectIsNull) { $coUser=[pscustomobject]@{Id=$file.CheckedOutByUser.Id;LoginName=$file.CheckedOutByUser.LoginName;Title=$file.CheckedOutByUser.Title;Email=(Read-OptionalProperty $file.CheckedOutByUser 'Email')} }
    if (-not $file.LockedByUser.ServerObjectIsNull) { $lockUser=[pscustomobject]@{Id=$file.LockedByUser.Id;LoginName=$file.LockedByUser.LoginName;Title=$file.LockedByUser.Title;Email=(Read-OptionalProperty $file.LockedByUser 'Email')} }
    $f=[pscustomobject]@{Exists=$file.Exists;UniqueId=$file.UniqueId;Name=$file.Name;ServerRelativeUrl=$file.ServerRelativeUrl;CheckOutType=[int]$file.CheckOutType;MajorVersion=$file.MajorVersion;MinorVersion=$file.MinorVersion;ETag=$file.ETag;TimeLastModified=$file.TimeLastModified.ToUniversalTime().ToString('o');Length=$file.Length;CheckedOutByUser=$coUser;LockedByUser=$lockUser}
    $requestId=Read-OptionalProperty $ctx 'TraceCorrelationId'
    $script:requests.Add([pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('o');Endpoint='CSOM file/checkout/lock snapshot';RequestId=$requestId;Success=$true})
    if ((Read-RequiredProperty $f 'Exists') -ne $true) { throw 'File does not exist or is not visible to this account.' }
    $id = [string](Read-RequiredProperty $f 'UniqueId')
    if ($script:expectedId -and $id -ne $script:expectedId) { throw 'File identity changed; refusing to continue.' }
    $co = [int](Read-RequiredProperty $f 'CheckOutType')
    if ($co -notin @(0,1,2)) { throw "Unknown CheckOutType '$co'." }
    $owner = Read-UserIdentity (Read-RequiredProperty $f 'CheckedOutByUser')
    $locker = Read-UserIdentity (Read-RequiredProperty $f 'LockedByUser')
    $state = 'NoCheckoutOrEditingLockReported'
    if ($co -ne 2) { $state = 'CheckedOut' }
    if ($null -ne $locker) {
        if ($co -eq 2) { $state = 'EditingLockReported' } else { $state = 'CheckoutAndLockReported' }
    }
    if ($co -eq 2 -and $null -ne $owner) { $state = 'InconsistentState' }
    if ($co -ne 2 -and $null -eq $owner) { $state = 'InconsistentState' }
    $snapshot = [pscustomobject]@{
        Utc = [DateTime]::UtcNow.ToString('o'); Phase = $Phase; State = $state
        UniqueId = $id; Name = [string](Read-RequiredProperty $f 'Name')
        ServerRelativeUrl = [string](Read-RequiredProperty $f 'ServerRelativeUrl')
        CheckOutType = $co; CheckedOutBy = $owner; LockedBy = $locker
        CheckOutTypeName = @('Online','Offline','None')[$co]
        MajorVersion = [int](Read-RequiredProperty $f 'MajorVersion')
        MinorVersion = [int](Read-RequiredProperty $f 'MinorVersion')
        ETag = [string](Read-RequiredProperty $f 'ETag')
        LastModified = [string](Read-RequiredProperty $f 'TimeLastModified')
        Length = [long](Read-RequiredProperty $f 'Length')
        RequestId=$requestId;IrmEnabled=(Read-OptionalProperty $file 'IrmEnabled');FileLevel=[string](Read-OptionalProperty $file 'Level')
    }
    $script:snapshots.Add($snapshot)
    Write-RunLog ("{0}: {1}; checkout={2}; checkout owner={3}; lock owner={4}" -f $Phase,$state,$co,($owner | ConvertTo-Json -Compress),($locker | ConvertTo-Json -Compress))
    return $snapshot
}

function Test-StateUnchanged($Before, $After) {
    # Reduces races; it cannot make the subsequent CSOM operation atomic.
    return ($Before.UniqueId -eq $After.UniqueId -and $Before.ETag -eq $After.ETag -and
        $Before.LastModified -eq $After.LastModified -and $Before.Length -eq $After.Length -and
        $Before.CheckOutType -eq $After.CheckOutType -and
        (ConvertTo-Json -InputObject $Before.CheckedOutBy -Compress) -eq (ConvertTo-Json -InputObject $After.CheckedOutBy -Compress) -and
        (ConvertTo-Json -InputObject $Before.LockedBy -Compress) -eq (ConvertTo-Json -InputObject $After.LockedBy -Compress))
}

function Save-ContentBackup($Snapshot) {
    $ctx = Get-PnPContext -Connection $script:conn
    $file = $ctx.Web.GetFileById([guid]$Snapshot.UniqueId)
    $binary = $file.OpenBinaryStream()
    $ctx.ExecuteQuery()
    if ($null -eq $binary.Value) { throw 'Content backup unavailable; no mutation will be attempted.' }
    $extension = [System.IO.Path]::GetExtension($Snapshot.Name)
    if ($extension -notmatch '^\.[a-zA-Z0-9]{1,12}$') { $extension = '.bin' }
    $destination = Join-Path $runDirectory ("BeforeChange_$runId$extension")
    $output = $null
    try {
        $output = [System.IO.File]::Open($destination,[System.IO.FileMode]::CreateNew,[System.IO.FileAccess]::Write)
        $binary.Value.CopyTo($output)
    } finally {
        if ($null -ne $output) { $output.Dispose() }
        $binary.Value.Dispose()
    }
    $bytes = (Get-Item -LiteralPath $destination).Length
    if ($bytes -ne $Snapshot.Length) { throw 'Backup length differs from the file snapshot; no mutation attempted.' }
    return [pscustomobject]@{
        Path=$destination; Bytes=$bytes; SHA256=(Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash
        Scope='Server-readable file bytes only; does not preserve unsaved client edits, full version history, or all metadata.'
    }
}

try {
    Write-RunLog "Starting $Action. Reports: $runDirectory"
    if ($ClientSessionsClosed -and $ClientSessionState -ne 'Unknown') { throw 'ClientSessionsClosed conflicts with an Open, Disconnected or Unreachable session. Confirm closure and omit ClientSessionState before repair.' }
    if ($LibraryUrl -and $Action -ne 'Diagnose') { throw 'Library scans are read-only. Supply -FileUrl for an individual repair.' }
    if ($FileUrl -or $LibraryUrl) {
        $inputUrl=$FileUrl; if ($LibraryUrl) { $inputUrl=$LibraryUrl }
        $targetUri=[uri]$inputUrl
        if (-not $targetUri.IsAbsoluteUri -or $targetUri.Scheme -ne 'https' -or $targetUri.UserInfo -or $targetUri.Fragment) { throw 'Supply a direct HTTPS URL; encode literal # in filenames as %23.' }
        $targetPath=[uri]::UnescapeDataString($targetUri.AbsolutePath).TrimEnd('/')
        if ($LibraryUrl) {
            $targetPath=$targetPath -replace '(?i)/Forms/[^/]+\.aspx$',''
            $libraryPath=$targetPath
            if ($targetUri.Query) { Write-RunLog 'Library view filters ignored; all current visible files are scanned recursively.' }
        } else {
            if ($targetUri.Query) { throw 'FileUrl must be a direct file URL without viewer or sharing query parameters.' }
            $FilePath=$targetPath
        }
        if (-not $SiteUrl) {
            $SiteUrl=$targetUri.GetLeftPart([UriPartial]::Authority)
            if ($targetPath -match '^/(sites|teams|personal)/[^/]+') { $SiteUrl+=$Matches[0] }
        }
        if (([uri]$SiteUrl).Authority -ine $targetUri.Authority) { throw 'SiteUrl and target URL must use the same host.' }
    }
    $site = [uri]$SiteUrl
    if (-not $site.IsAbsoluteUri -or $site.Scheme -ne 'https' -or $site.Query -or $site.Fragment -or $site.UserInfo) {
        throw 'SiteUrl must be an HTTPS SharePoint web URL without query, fragment or credentials.'
    }
    $SiteUrl = $site.AbsoluteUri.TrimEnd('/')
    if (-not $LibraryUrl -and (-not $FilePath.StartsWith('/') -or $FilePath.StartsWith('//') -or $FilePath.Contains('\') -or $FilePath -match '[\r\n]')) {
        throw 'FilePath must be a decoded server-relative file path starting with a single slash.'
    }
    if ($FilePath -match '(?i)/_layouts/|/Forms/AllItems\.aspx|^/:[a-z]:/') {
        throw 'Provide the actual file path, not a viewer, library view, or sharing URL.'
    }
    $webPath = [uri]::UnescapeDataString($site.AbsolutePath).TrimEnd('/')
    $scopePath=$FilePath; if ($LibraryUrl) { $scopePath=$libraryPath }
    if ($webPath -and -not $scopePath.StartsWith($webPath + '/', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'FilePath is outside the supplied SharePoint web URL.'
    }
    $report.SiteUrl=$SiteUrl; $report.FilePath=$FilePath
    if ($AuthenticationMode -eq 'LegacyWebLogin') {
        if ($PSVersionTable.PSVersion.Major -ne 5 -or [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'LegacyWebLogin requires Windows PowerShell 5.1 -STA.' }
        $dependencyRoot=Join-Path (Split-Path $PSScriptRoot -Parent) '.dependencies'
        if (Test-Path -LiteralPath $dependencyRoot) { $env:PSModulePath=$dependencyRoot+';'+$env:PSModulePath }
        Import-Module PnP.PowerShell -RequiredVersion 1.12.0 -ErrorAction Stop
        Write-RunLog 'Using legacy PnP 1.12.0 embedded MFA; tenant policy may reject this unsupported authentication method.'
    } else { Import-Module PnP.PowerShell -ErrorAction Stop }
    $report.PnPVersion = (Get-Module PnP.PowerShell).Version.ToString()
    if ($null -ne $Connection) {
        $script:conn = $Connection
    } else {
        $connect = @{Url=$SiteUrl; ReturnConnection=$true; ErrorAction='Stop'}
        if ($AuthenticationMode -eq 'LegacyWebLogin') { $connect.UseWebLogin=$true; $connect.ForceAuthentication=$true }
        else { $connect.Interactive=$true; if ($ClientId) { $connect.ClientId = $ClientId } }
        $script:conn = Connect-PnPOnline @connect
    }
    $web = Invoke-SPRead '/_api/web?$select=Url'
    if (([string](Read-RequiredProperty $web 'Url')).TrimEnd('/') -ne $SiteUrl) {
        throw 'The connection is for a different SharePoint web. Reconnect to SiteUrl.'
    }
    $report.SignedInUser = Invoke-SPRead '/_api/web/currentuser?$select=Id,Title,LoginName'
    if ($LibraryUrl) {
        $libraries=@(Get-PnPList -Includes RootFolder,BaseType,ForceCheckout,EnableVersioning,EnableMinorVersions,EnableModeration -Connection $script:conn | Where-Object { $_.BaseType.ToString() -eq 'DocumentLibrary' -and $_.RootFolder.ServerRelativeUrl.TrimEnd('/') -ieq $libraryPath })
        if ($libraries.Count -ne 1) { throw 'No unique library root matches LibraryUrl. For a subweb also supply -SiteUrl.' }
        $report.Library=[pscustomobject]@{Id=$libraries[0].Id;Title=$libraries[0].Title;Root=$libraryPath;ForceCheckout=$libraries[0].ForceCheckout;EnableVersioning=$libraries[0].EnableVersioning;EnableMinorVersions=$libraries[0].EnableMinorVersions;EnableModeration=$libraries[0].EnableModeration}
        $query="<View Scope='RecursiveAll'><Query><OrderBy><FieldRef Name='ID'/></OrderBy></Query><ViewFields><FieldRef Name='FileRef'/><FieldRef Name='FSObjType'/><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Created'/><FieldRef Name='Modified'/><FieldRef Name='Author'/><FieldRef Name='Editor'/><FieldRef Name='ContentTypeId'/></ViewFields><RowLimit Paged='TRUE'>500</RowLimit></View>"
        Get-PnPListItem -List $libraries[0].Id -Query $query -PageSize 500 -Connection $script:conn | ForEach-Object {
            if ([int]$_.FieldValues['FSObjType'] -eq 0) {
                $itemFields=$_.FieldValues; $path=[string]$itemFields['FileRef']; $script:expectedId=$null
                $script:targetEndpoint="/_api/web/GetFileByServerRelativePath(decodedUrl='$([uri]::EscapeDataString($path.Replace("'","''")))')"
                try {
                    $snapshot=Read-LockSnapshot 'LibraryDiscovery' $path
                    $snapshot | Add-Member -NotePropertyName Metadata -NotePropertyValue (ConvertTo-ItemMetadata $itemFields)
                    $script:libraryResults.Add($snapshot); $report.FilesScanned++
                    if ($snapshot.State -ne 'NoCheckoutOrEditingLockReported') { $report.FilesWithSignals++ }
                } catch {
                    $report.FileReadErrors++
                    $script:libraryResults.Add([pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('o');ServerRelativeUrl=$path;State='UnknownReadFailed';Error=$_.Exception.Message})
                    Write-RunLog "Unknown lock state for ${path}: $($_.Exception.Message)"
                }
            }
        }
        $report.EnumerationComplete=$true
        $report.Status='LibraryScanComplete'; $report.ExitCode=0
        if ($report.FilesWithSignals -gt 0) { $report.Status='LibraryLocksOrCheckoutsReported'; $report.ExitCode=2 }
        if ($report.FileReadErrors -gt 0) { $report.Status='LibraryScanIncompleteFileReads'; $report.ExitCode=1 }
    } else {
    # Unlock preserves saved changes: it maps to CheckIn, never automatic discard.
    if ($Action -eq 'Unlock') { $Action='CheckIn' }
    # Escape an OData string literal before URL encoding. Never decode FilePath.
    $encodedPath = [uri]::EscapeDataString($FilePath.Replace("'", "''"))
    $script:targetEndpoint = "/_api/web/GetFileByServerRelativePath(decodedUrl='$encodedPath')"
    $initial = Read-LockSnapshot 'Initial'
    $script:expectedId = $initial.UniqueId
    # Bind all following reads and the mutation to the observed identity.
    $script:targetEndpoint = "/_api/web/GetFileById('$($script:expectedId)')"
    $report.Library = Invoke-SPRead ($script:targetEndpoint + '/ListItemAllFields/ParentList?$select=Id,Title,EnableVersioning,EnableMinorVersions,EnableModeration,ForceCheckout')
    $report.FileMetadata = Read-FileMetadata $initial.UniqueId
    $current = $initial

    if ($Action -eq 'Diagnose') {
        $report.Status = $current.State
        if ($current.State -eq 'NoCheckoutOrEditingLockReported') { $report.ExitCode=0 } else { $report.ExitCode=2 }
    } elseif ($Action -eq 'WaitForRelease') {
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        while ($current.State -ne 'NoCheckoutOrEditingLockReported' -and $timer.Elapsed.TotalSeconds -lt $WaitSeconds) {
            $remaining = [int][Math]::Ceiling($WaitSeconds - $timer.Elapsed.TotalSeconds)
            Start-Sleep -Seconds ([Math]::Min($PollSeconds,$remaining))
            $current = Read-LockSnapshot 'Poll'
        }
        $timer.Stop()
        if ($current.State -eq 'NoCheckoutOrEditingLockReported') {
            $report.Status='NoCheckoutOrEditingLockReported'; $report.ExitCode=0
        } else { $report.Status='LockOrCheckoutStillReported'; $report.ExitCode=2 }
    } else {
        if ($current.State -eq 'InconsistentState') { throw 'Inconsistent lock state. Re-run diagnosis; do not repair automatically.' }
        if ($ClientSessionState -in @('Disconnected','Unreachable')) {
            $report.Status='BlockedClientSessionUnavailable'; $report.ExitCode=3
            Write-RunLog 'Recover the disconnected/unreachable session and preserve edits before repair. This script does not control VDI sessions or force-release their editing locks.'
        } elseif ($current.CheckOutType -eq 2) {
            if ($null -ne $current.LockedBy) {
                $report.Status='EditingLockRequiresClientClosureOrSupport'; $report.ExitCode=2
                Write-RunLog 'No explicit checkout to repair. Close editing sessions, then use WaitForRelease. No force-unlock attempted.'
            } else { $report.Status='NoCheckoutToRepair'; $report.ExitCode=0 }
        } elseif (-not $ClientSessionsClosed) {
            $report.Status='BlockedClientSessionsNotConfirmed'; $report.ExitCode=3
            Write-RunLog 'Save/reconcile client edits and close every session. Re-run with -ClientSessionsClosed only after verifying this.'
        } elseif ($Action -eq 'DiscardCheckout' -and -not $AcknowledgeDiscardChanges) {
            $report.Status='BlockedDiscardNotAcknowledged'; $report.ExitCode=3
            Write-RunLog 'Discarding checkout loses changes. Explicit -AcknowledgeDiscardChanges is required.'
        } elseif ($Action -eq 'DiscardCheckout' -and $current.MajorVersion -lt 1) {
            $report.Status='BlockedDiscardWithoutMajorVersion'; $report.ExitCode=3
            Write-RunLog 'Discard is blocked for files without a major version to protect initial uploads and unpublished drafts. Use CheckIn or review manually.'
        } else {
            $minor = [bool](Read-RequiredProperty $report.Library 'EnableMinorVersions')
            $checkInType = if ($minor) { 'MinorCheckIn' } else { 'MajorCheckIn' }
            $description = if ($Action -eq 'CheckIn') { "Back up server content and check in using $checkInType" } else { 'Back up server-readable content, then DISCARD checkout changes' }
            if ($PSCmdlet.ShouldProcess($FilePath, $description)) {
                $report.Backup = Save-ContentBackup $current
                Write-RunLog "Backup saved and SHA256 recorded: $($report.Backup.Path)"
                $fresh = Read-LockSnapshot 'BeforeMutation'
                if (-not (Test-StateUnchanged $current $fresh)) { throw 'File/lock state changed during preparation. No mutation attempted; re-run diagnosis.' }
                $ctx = Get-PnPContext -Connection $script:conn
                $file = $ctx.Web.GetFileById([guid]$script:expectedId)
                $report.MutationAttempted = $true
                if ($Action -eq 'CheckIn') {
                    $kind = [Microsoft.SharePoint.Client.CheckinType]::$checkInType
                    $file.CheckIn($Comment,$kind)
                } else { $file.UndoCheckOut() }
                # No automatic retry: after a timeout the server may have committed.
                $ctx.ExecuteQuery()
                $report.MutationCompleted = $true
                $current = Read-LockSnapshot 'AfterMutation'
                if ($current.State -eq 'NoCheckoutOrEditingLockReported') {
                    $report.Status='CheckoutRepairedNoEditingLockReported'; $report.ExitCode=0
                } else { $report.Status='MutationCompletedButLockOrCheckoutStillReported'; $report.ExitCode=2 }
            } else {
                $report.Status = if ($WhatIfPreference) { 'WhatIfNoRemoteChanges' } else { 'CancelledNoRemoteChanges' }
                $report.ExitCode=3
            }
        }
    }
    }
} catch {
    $e = $_
    $report.Status = if ($report.MutationAttempted) { 'MutationOrVerificationFailedInspectBeforeRetry' } else { 'FailedNoRemoteMutationAttempted' }
    $report.ExitCode=1
    $correlation = $null
    $exception = $e.Exception
    for ($i=0; $i -lt 8 -and $null -ne $exception; $i++) {
        if ($null -ne $exception.PSObject.Properties['ServerErrorTraceCorrelationId']) {
            $correlation = [string]$exception.ServerErrorTraceCorrelationId
        }
        $exception = $exception.InnerException
    }
    $report.Error = [pscustomobject]@{
        Message=$e.Exception.Message; ErrorId=$e.FullyQualifiedErrorId
        Category=[string]$e.CategoryInfo.Category; CorrelationId=$correlation
    }
    Write-RunLog ("ERROR: {0}; correlation={1}" -f $e.Exception.Message,$correlation)
    if ($report.MutationAttempted -and $script:targetEndpoint) {
        try { $current = Read-LockSnapshot 'AfterError' } catch { Write-RunLog "Post-error state could not be read: $($_.Exception.Message)" }
    }
} finally {
    if ($script:snapshots.Count -gt 0) { $report.FinalSnapshot = $script:snapshots[$script:snapshots.Count-1] }
    $report.FinishedUtc = [DateTime]::UtcNow.ToString('o')
    $report['Snapshots'] = $script:snapshots.ToArray()
    $report['Requests'] = $script:requests.ToArray()
    $report['LibraryResults'] = $script:libraryResults.ToArray()
    $report['Interpretation'] = 'A clear snapshot means no checkout/editing lock was reported at that instant. It does not test edit permission, retention/record restrictions, application validation, or future locks. Verify opening, saving, metadata and publishing separately.'
    $report['RecoveryNotes'] = 'For editing locks: preserve edits; close Word/browser/other-device sessions; check lingering Word processes; wait and diagnose again. For persistent locks provide these reports to Microsoft. Never retry an uncertain mutation until current state is reviewed.'
    $report['NotAvailableFromThisInterface'] = @('Exact lock acquisition time and lease expiry','Lock token or exclusive/shared lease subtype','All coauthors and their device/session identifiers','Remote client reachability and unsaved edits','Root cause of a browser/Office/Power Apps failure')
    $report['NextAction'] = if ($LibraryUrl) { 'Review each CSV row and diagnose affected files individually. Discovery does not perform bulk repair.' } else { Get-RecoveryAdvice $report.FinalSnapshot }
    $report['Observation'] = if ($LibraryUrl) { $null } else { Get-ObservationSummary $script:snapshots.ToArray() }
    $json = $report | ConvertTo-Json -Depth 12
    [System.IO.File]::WriteAllText($reportPath,$json)
    $rows = @($script:snapshots.ToArray() | ForEach-Object { ConvertTo-DiagnosticRow $_ $report.FileMetadata })
    if ($LibraryUrl) {
        $rows=@($script:libraryResults.ToArray() | ForEach-Object { ConvertTo-DiagnosticRow $_ (Read-OptionalProperty $_ 'Metadata') })
    }
    if ($rows.Count -eq 0) { $rows = @([pscustomobject]@{RunId=$runId; FilePath=$FilePath; FinalStatus=$report.Status; ExitCode=$report.ExitCode}) }
    $csv = $rows | ConvertTo-Csv -NoTypeInformation
    [System.IO.File]::WriteAllLines($csvPath,[string[]]$csv,[Text.UTF8Encoding]::new($true))
    $summary=@("SharePoint file lock diagnostic - version $($report.ScriptVersion)","Run: $runId", "UTC: $($report.StartedUtc) to $($report.FinishedUtc)","Status: $($report.Status); exit code: $($report.ExitCode)","Site: $SiteUrl", "Target: $FilePath$LibraryUrl", "Scanned: $($report.FilesScanned); signals: $($report.FilesWithSignals); read errors: $($report.FileReadErrors); supplemental metadata errors: $($report.MetadataReadErrors)", "Remote mutation attempted/completed: $($report.MutationAttempted)/$($report.MutationCompleted)","Client context (operator supplied): $ClientEnvironment / $ClientSessionState / host=$SessionHost / session=$SessionId", "Client closure confirmed: $([bool]$ClientSessionsClosed)","Incident: $IncidentReference", "Reported error: $ReportedError",'', $report.NextAction, '', 'No arbitrary Office editing-lock force release, remote session termination, or recovery of unsaved client edits.', 'Blank optional fields mean unavailable or inapplicable; LastModified is not lock age. See JSON for observations, metadata errors and request IDs.', "JSON: $reportPath", "CSV: $csvPath")
    [System.IO.File]::WriteAllLines($summaryPath,[string[]]$summary,[Text.UTF8Encoding]::new($true))
    Write-RunLog "Finished: $($report.Status); exit code=$($report.ExitCode)"
    Write-Host "JSON: $reportPath"
    Write-Host "CSV:  $csvPath"
}
[pscustomobject]@{ Status=$report.Status; ExitCode=$report.ExitCode; Report=$reportPath; Log=$logPath; Csv=$csvPath; Summary=$summaryPath }
exit ([int]$report.ExitCode)
