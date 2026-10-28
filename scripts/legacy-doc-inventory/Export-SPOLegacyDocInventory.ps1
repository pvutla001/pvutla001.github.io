#Requires -Version 5.1
<# 
.SYNOPSIS
Stream a large SharePoint library inventory with bounded parallel metadata/content workers. 
.DESCRIPTION 
Version 3.10.1. Windows PowerShell 5.1 STA; legacy PnP.PowerShell 1.12.0 MFA retained.
Default: .doc metadata only, all folder depths. CountOnly retains MaxDepth=2;
use MaxDepth=-1 for unlimited depth. Folder is a decoded library-relative path.
One paged discovery request plus up to ThrottleLimit workers; every request batch
owns a separate cloned CSOM context. No concurrent sharing of a mutable context.
ID-ordered discovery captures a starting high-watermark and uses no nonindexed
extension/path filter. Metadata is read per page/field batch, not per file.
JSONL and worksheet XML stream to disk. Excel reports rotate into bounded parts;
no ImportExcel dependency or whole-library in-memory workbook is required.
Progress is displayed, logged and saved to timestamped Progress JSON. Partial records survive
failure. This is not a transactional snapshot or a resumable synchronization engine.
DownloadFiles/CheckVBA/CheckMacro/CheckBoth retain local copies (Flat by default; -DownloadLayout Mirror preserves folders). Content
checks use native PowerShell/.NET parsing and never execute Word/VBA. Metadata/count modes
never download contents. VBA/macros overlap; NotDetected does not mean malware-free.
Files hidden by permissions, historical versions and newly added IDs above the
starting watermark are not included. File moves/changes during scanning are possible.
Keep SPOLegacyDocInventory.Engine.ps1 and SPOLegacyDocInventory.Runtime.cs beside
this script; content checks also require Inspect-LegacyWord.ps1. See the accompanying blog post.
.NOTES
Version 3.10.1 fixes legacy .NET path validation for local document I/O,
including deep Mirror downloads and the isolated scanner. No registry or host
configuration changes are needed. Keep script/module/report roots short;
these use standard PowerShell/.NET paths. Replace all four files together and
start a fresh PowerShell process when upgrading.
All reports and logs have timestamped names. ReviewAndFailures workbooks contain
Documents and Diagnostics sheets. Detailed failures are always retained; -Verbose
also logs every successful document and prints per-document results to the console.
ScanStatus describes the retained local snapshot. CorrelationStatus separately
describes its relationship to the current server version. VBAProjectCount is blank
when it could not be established, rather than a misleading zero.
.EXAMPLE
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/Example/Documents' -ThrottleLimit 4 -BatchSize 500
.EXAMPLE
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/Example/Documents' -CountOnly -MaxDepth -1
.EXAMPLE
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/Example/Documents' -CheckBoth -NoInstall
Downloads and statically inspects .doc files using the native PowerShell helper.
.EXAMPLE
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/Example/Documents' -CheckBoth -Verbose -DownloadLayout Mirror
Retain folder structure and record every document's processing result.
.EXAMPLE
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/Example/Documents' -Folder 'Policies/Finance' -CheckBoth -NoInstall -Verbose
Inspect one library-relative folder and its descendants using pre-staged dependencies.
#>
[CmdletBinding(DefaultParameterSetName='Info')]
param(
    [string]$LibraryUrl,[string]$SiteUrl,[string]$Folder,[string]$DownloadDirectory,
    [switch]$DownloadFiles,
    [ValidateSet('Flat','Mirror')][string]$DownloadLayout='Flat',
    [Parameter(ParameterSetName='Info')][switch]$FetchDocOnly,
    [Parameter(Mandatory,ParameterSetName='Counts')][switch]$CountOnly,
    [Parameter(ParameterSetName='Counts')][ValidateRange(-1,2147483647)][int]$MaxDepth=2,
    [Parameter(Mandatory,ParameterSetName='VBA')][switch]$CheckVBA,
    [Parameter(Mandatory,ParameterSetName='Macro')][switch]$CheckMacro,
    [Parameter(Mandatory,ParameterSetName='Both')][switch]$CheckBoth,
    [ValidateSet('LegacyWebLogin','SystemBrowser')][string]$AuthenticationMode='LegacyWebLogin',
    [string]$ClientId,[object]$Connection,
    [ValidateRange(1,2048)][int]$MaxScanMB=100,
    [ValidateRange(10,1800)][int]$ScanTimeoutSeconds=120,
    [ValidateRange(1,8)][int]$ThrottleLimit=4,
    [ValidateRange(10,2000)][int]$BatchSize=500,
    [ValidateRange(1,8)][int]$FieldBatchSize=8,
    [ValidateRange(0,12)][int]$RetryCount=6,
    [ValidateRange(1,60)][int]$ProgressSeconds=5,
    [ValidateRange(100,250000)][int]$ReportDocumentsPerWorkbook=10000,
    [ValidateRange(1000,1000000)][int]$ReportMetadataRowsPerWorkbook=200000,
    [string]$OutputDirectory,[switch]$NoInstall
)
$ConfiguredLibraryUrl='' # Supply your library URL with -LibraryUrl.
Set-StrictMode -Version 2
$ErrorActionPreference='Stop'
$previousPnPUpdateCheck=$env:PNPPOWERSHELL_UPDATECHECK
$env:PNPPOWERSHELL_UPDATECHECK='Off'
$enginePath=Join-Path $PSScriptRoot 'SPOLegacyDocInventory.Engine.ps1'
. $enginePath
$runId=[DateTimeOffset]::Now.ToString('yyyyMMdd_HHmmss_fff')+'_'+[guid]::NewGuid().ToString('N').Substring(0,8)
if(-not $OutputDirectory){$OutputDirectory=Join-Path $PSScriptRoot 'DocInventoryReports'}
$runDirectory=[IO.Path]::GetFullPath((Join-Path $OutputDirectory ('LegacyDocInventory_'+$runId)))
[void][IO.Directory]::CreateDirectory($runDirectory)
$logPath=Join-Path $runDirectory ('ExecutionLog_'+$runId+'.log')
$journalPath=Join-Path $runDirectory ('DocumentRecords_'+$runId+'.jsonl')
$diagnosticsPath=Join-Path $runDirectory ('DocumentDiagnostics_'+$runId+'.jsonl')
$failureLogPath=Join-Path $runDirectory ('FailureDetails_'+$runId+'.log')
$summaryPath=Join-Path $runDirectory ('RunSummary_'+$runId+'.json')
$summaryTextPath=Join-Path $runDirectory ('RunSummary_'+$runId+'.txt')
$progressPath=Join-Path $runDirectory ('Progress_'+$runId+'.json')
$journal=[IO.StreamWriter]::new($journalPath,$false,[Text.UTF8Encoding]::new($false),65536)
$diagnosticsWriter=[IO.StreamWriter]::new($diagnosticsPath,$false,[Text.UTF8Encoding]::new($false),65536)
$failureWriter=[IO.StreamWriter]::new($failureLogPath,$false,[Text.UTF8Encoding]::new($false),65536)
$failureWriter.WriteLine('PER-DOCUMENT ERRORS AND REVIEW ITEMS | Run '+$runId)
$summary=[ordered]@{
    ScriptVersion='3.10.1';RunId=$runId;Started=[DateTimeOffset]::Now.ToString('o');Finished='';Status='Starting'
    LibraryUrl='';SiteUrl='';Account='';Mode=$PSCmdlet.ParameterSetName;FolderScope='';DownloadLayout=$DownloadLayout
    ItemsEnumerated=0L;FilesEnumerated=0L;FoldersEnumerated=0L;MatchedDocFiles=0L;DocFiles=0L
    Succeeded=0L;NeedsReview=0L;Failed=0L;Errors=0L;NotProcessed=0L;Downloaded=0L;DownloadFailures=0L;ScansCompleted=0L;VbaIndicatorsDetected=0L
    FailedBatches=0;EnumerationComplete=$false;HighWatermarkId=0;LibraryItemCountEstimate=0;LastEnumeratedId=0;MetadataRows=0L;DocFilesExcludedByDepth=0L
    ThrottleLimit=$ThrottleLimit;BatchSize=$BatchSize;VerboseDocuments=[bool]$PSBoundParameters['Verbose']
    Coverage='Visible current .doc files; no historical versions; IDs at or below the initial high watermark. No transactional snapshot.'
    ScanCoverage='Local snapshot VBA indicators; no exhaustive WordBasic, embedded-file, external-template or stomped-VBA analysis.'
    ReportParts=@();Artifacts=@();DownloadRoot='';RunErrors=@();Error=''
}
$script:inventoryWorkbook=$null;$script:inventoryPart=0;$script:partDocs=0;$script:partMetadata=0
$script:outcomeBooks=@{};$script:findingCounts=[ordered]@{}
$script:artifacts=[Collections.Generic.List[object]]::new()
$script:documentHeaders=@('Outcome','FileName','LibraryRelativePath','Url','ItemId','UniqueId','FolderDepth','SizeBytes','Modified','MetadataStatus','DownloadStatus','LocalPath','ScanStatus','CorrelationStatus','Format','FormatDetails','DetectionScope','EmbeddedObjects','FIBVersion','EffectiveFIBVersion','IsTemplate','VBA','VBAProjectCount','VBASourceModules','VBASourceModulesParsed','Macro','Both','ErrorCode','ErrorStage','Error','Notes','SHA256','ScannerVersion','ScanMilliseconds','ScanInputPath','SignatureHex','DownloadedBytes','DiscoveryModifiedUtc','DownloadModifiedUtc','AfterScanModifiedUtc','ETagBefore','ETagAfter')
$script:metadataHeaders=@('ItemId','FileName','Field','DisplayName','Type','Value','Status')
$script:issueHeaders=@('ItemId','FileName','LibraryRelativePath','Severity','Code','Stage','Message','ExceptionType','ExceptionDetails','FullyQualifiedErrorId','ScriptStackTrace','ProjectPath','StreamPath','ModuleName','FileOffset','CompressedOffset','LocalPath','SHA256')
$folderCounts=@{};$schema=@();$ctx=$null;$ownsContext=$false;$runFailure=$null;$gate=$null;$clock=[Diagnostics.Stopwatch]::StartNew();$scannerExecutable='';$book=$null;$servicePoint=$null;$previousConnectionLimit=0;$download=$false;$scan=$false
function Write-RunLog([string]$Message,[string]$Level='INFO',[switch]$Quiet) {
    $line='{0} [{1}] {2}' -f [DateTimeOffset]::Now.ToString('o'),$Level,$Message
    [IO.File]::AppendAllText($logPath,$line+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    if(-not $Quiet){Write-Host $line}
}
function Save-Journal([string]$Kind,$Data) {$journal.WriteLine((@{Kind=$Kind;Timestamp=[DateTimeOffset]::Now.ToString('o');Data=$Data}|ConvertTo-Json -Depth 18 -Compress))}
function Add-Artifact([string]$Purpose,[string]$Path) {$script:artifacts.Add([pscustomobject]@{Purpose=$Purpose;Path=$Path;Created=$false})}
function Get-RowValues($Object,[string[]]$Headers) {
    $values=foreach($header in $Headers){$p=$Object.PSObject.Properties[$header];if($p){[string]$p.Value}else{''}}
    return ,([string[]]$values)
}
function Add-FolderCount([string]$Relative,[bool]$Document) {
    $depth=Get-RelativeFolderDepth $Relative;if($MaxDepth -ge 0 -and $depth -gt $MaxDepth){return}
    $cursor=$Relative;$direct=$true
    while($true){
        if(-not $folderCounts.ContainsKey($cursor)){$label=$cursor;if(-not $label){$label='(library root)'};$folderCounts[$cursor]=[pscustomobject]@{Folder=$label;Depth=(Get-RelativeFolderDepth $cursor);DirectDocCount=0L;DocsWithinDepthLimit=0L}}
        if($Document){$folderCounts[$cursor].DocsWithinDepthLimit++;if($direct){$folderCounts[$cursor].DirectDocCount++}}
        if(-not $cursor){break};$direct=$false;$slash=$cursor.LastIndexOf('/');if($slash -lt 0){$cursor=''}else{$cursor=$cursor.Substring(0,$slash)}
    }
}
function Complete-InventoryWorkbook {
    if($script:inventoryWorkbook){try{$script:inventoryWorkbook.Complete()}finally{$script:inventoryWorkbook.Dispose();$script:inventoryWorkbook=$null}}
}
function Write-DocumentReport($Document,[object[]]$Metadata) {
    if(-not $script:inventoryWorkbook -or $script:partDocs -ge $ReportDocumentsPerWorkbook -or ($script:partMetadata+$Metadata.Count) -gt $ReportMetadataRowsPerWorkbook){
        Complete-InventoryWorkbook;$script:inventoryPart++;$script:partDocs=0;$script:partMetadata=0
        $name='AllDocuments_Metadata_{0}_Part{1:D4}.xlsx' -f $runId,$script:inventoryPart
        $script:inventoryWorkbook=[SPOInventoryV310.StreamingWorkbook]::new((Join-Path $runDirectory $name))
        [void]$script:inventoryWorkbook.AddSheet('Documents',$script:documentHeaders);[void]$script:inventoryWorkbook.AddSheet('Metadata',$script:metadataHeaders)
        $fieldSheet=$script:inventoryWorkbook.AddSheet('Fields',@('InternalName','DisplayName','Type','Hidden','ReadOnly'))
        foreach($field in $schema){$script:inventoryWorkbook.AddRow($fieldSheet,[string[]]@($field.InternalName,$field.DisplayName,$field.Type,$field.Hidden,$field.ReadOnly))}
        $summary.ReportParts+=$name;Add-Artifact 'Complete document inventory and metadata' (Join-Path $runDirectory $name)
    }
    $script:inventoryWorkbook.AddRow(0,(Get-RowValues $Document $script:documentHeaders));$script:partDocs++
    foreach($m in $Metadata){$script:inventoryWorkbook.AddRow(1,(Get-RowValues $m $script:metadataHeaders));$script:partMetadata++}
}
function Get-OutcomeWorkbook([string]$Category,[int]$IssueRows=0) {
    if(-not $script:outcomeBooks.ContainsKey($Category)){$script:outcomeBooks[$Category]=@{Book=$null;Part=0;Rows=0;Issues=0}}
    $state=$script:outcomeBooks[$Category]
    if(-not $state.Book -or $state.Rows -ge $ReportDocumentsPerWorkbook -or ($state.Issues+$IssueRows) -gt $ReportMetadataRowsPerWorkbook){
        if($state.Book){try{$state.Book.Complete()}finally{$state.Book.Dispose();$state.Book=$null}}
        $state.Part++;$state.Rows=0;$state.Issues=0;$name='{0}_{1}_Part{2:D4}.xlsx' -f $Category,$runId,$state.Part
        $state.Book=[SPOInventoryV310.StreamingWorkbook]::new((Join-Path $runDirectory $name));[void]$state.Book.AddSheet('Documents',$script:documentHeaders)
        if($Category -eq 'ReviewAndFailures'){[void]$state.Book.AddSheet('Diagnostics',$script:issueHeaders)}
        $summary.ReportParts+=$name;Add-Artifact $Category (Join-Path $runDirectory $name)
    }
    return $state
}
function Write-OutcomeReport($Row) {
    $category='SuccessfulDocuments';if($Row.Outcome -ne 'Succeeded'){$category='ReviewAndFailures'}
    $state=Get-OutcomeWorkbook $category $Row.Diagnostics.Count
    $state.Book.AddRow(0,(Get-RowValues $Row $script:documentHeaders));$state.Rows++
    if($category -eq 'ReviewAndFailures'){
        foreach($issue in $Row.Diagnostics){
            $record=[ordered]@{ItemId=$Row.ItemId;FileName=$Row.FileName;LibraryRelativePath=$Row.LibraryRelativePath;LocalPath=$Row.LocalPath;SHA256=$Row.SHA256}
            foreach($p in $issue.PSObject.Properties){$record[$p.Name]=$p.Value}
            $state.Book.AddRow(1,(Get-RowValues ([pscustomobject]$record) $script:issueHeaders));$state.Issues++
        }
    }
}
function Save-DocumentDiagnostics($Row) {
    $message='ItemId={0}; File={1}; Outcome={2}; Scan={3}; VBA={4}; Macro={5}; Projects={6}; Error={7}' -f $Row.ItemId,$Row.LibraryRelativePath,$Row.Outcome,$Row.ScanStatus,$Row.VBA,$Row.Macro,$Row.VBAProjectCount,$Row.Error
    if($Row.Diagnostics.Count -or $summary.VerboseDocuments){
        $record=[ordered]@{Timestamp=[DateTimeOffset]::Now.ToString('o');RunId=$runId;Document=$Row}
        $diagnosticsWriter.WriteLine(($record|ConvertTo-Json -Depth 18 -Compress))
        Write-RunLog $message $(if($Row.Outcome -eq 'Succeeded'){'DETAIL'}else{'REVIEW'}) -Quiet
    }
    if($Row.Outcome -ne 'Succeeded'){
        $failureWriter.WriteLine([Environment]::NewLine+('[{0}] {1}' -f [DateTimeOffset]::Now.ToString('o'),$message))
        $failureWriter.WriteLine('SharePoint: '+$Row.Url);$failureWriter.WriteLine('Local snapshot: '+$Row.LocalPath)
        foreach($issue in $Row.Diagnostics){$failureWriter.WriteLine(($issue|ConvertTo-Json -Depth 12))}
    }
    Write-Verbose $message
}
function Accept-InventoryResult($Result,$Failure,$Page) {
    if($Failure){
        $summary.FailedBatches++;Write-RunLog "Batch ($($Page.After), $($Page.Through)] failed: $($Failure.Exception.Message)" 'ERROR'
        $failedDocs=[Collections.Generic.List[object]]::new()
        foreach($item in $Page.Items){$v=$item.Values;if([int]$v['FSObjType'] -eq 0 -and [IO.Path]::GetExtension($v['FileLeafRef']) -ieq '.doc' -and $v['FileRef'].StartsWith($folderPath+'/',[StringComparison]::OrdinalIgnoreCase)){
            $failed=New-InventoryDocument $item $config;$failed.MetadataStatus='Failed';Add-InventoryIssue $failed (ConvertTo-InventoryError $Failure 'BatchWorker' 'BatchFailure');Set-InventoryOutcome $failed;$failedDocs.Add($failed)
        }}
        $Result=[pscustomobject]@{Documents=$failedDocs.ToArray();Metadata=@();Errors=$failedDocs.Count;After=$Page.After;Through=$Page.Through}
    }
    $groups=@{};foreach($m in $Result.Metadata){if(-not $groups.ContainsKey($m.ItemId)){$groups[$m.ItemId]=[Collections.Generic.List[object]]::new()};$groups[$m.ItemId].Add($m)}
    foreach($row in $Result.Documents){
        Set-InventoryOutcome $row
        $values=@();if($groups.ContainsKey($row.ItemId)){$values=$groups[$row.ItemId].ToArray()}
        Save-Journal 'Document' @{Document=$row;Metadata=$values};Save-DocumentDiagnostics $row
        Write-DocumentReport $row $values;Write-OutcomeReport $row
        $summary.DocFiles++;$summary.MetadataRows+=$values.Count;$summary[$row.Outcome]++
        if($row.Outcome -ne 'Succeeded'){$summary.Errors++}
        if($row.DownloadStatus -eq 'Downloaded'){$summary.Downloaded++}elseif($row.DownloadStatus -eq 'Failed'){$summary.DownloadFailures++}
        if($row.ScanStatus -eq 'Scanned'){$summary.ScansCompleted++}
        if($row.VBA -eq 'Detected' -or $row.Macro -eq 'VBAIndicatorsDetected'){$summary.VbaIndicatorsDetected++}
        $category=$row.ScanStatus+'/'+$row.Format
        if(-not $script:findingCounts.Contains($category)){$script:findingCounts[$category]=0L};$script:findingCounts[$category]++
    }
    Save-Journal 'BatchCommitted' @{After=$Page.After;Through=$Page.Through;Documents=$Result.Documents.Count;Errors=$Result.Errors}
    $journal.Flush();$diagnosticsWriter.Flush();$failureWriter.Flush()
}
function Write-ProgressSnapshot([long]$Cursor,[int]$Active,[bool]$Discovering){
    $summary.LastEnumeratedId=$Cursor;$rate=[math]::Round($summary.DocFiles/[math]::Max(1,$clock.Elapsed.TotalSeconds),2)
    $retries=0L;$cooldown=0;if($gate){$retries=$gate.Retries;$cooldown=[math]::Ceiling($gate.RemainingSeconds)}
    $progress=[ordered]@{Utc=[DateTime]::UtcNow.ToString('o');Status=$summary.Status;EnumeratedItems=$summary.ItemsEnumerated;FilesEnumerated=$summary.FilesEnumerated;MatchedDocFiles=$summary.MatchedDocFiles;Processed=$summary.DocFiles;Succeeded=$summary.Succeeded;NeedsReview=$summary.NeedsReview;Failed=$summary.Failed;ActiveWorkers=$Active;DiscoveryRequestActive=$Discovering;FilesPerSecond=$rate;RetryCount=$retries;CooldownSeconds=$cooldown;LastEnumeratedId=$Cursor}
    $text="items=$($summary.ItemsEnumerated); files=$($summary.FilesEnumerated); .doc matched/processed=$($summary.MatchedDocFiles)/$($summary.DocFiles); success=$($summary.Succeeded); review=$($summary.NeedsReview); failed=$($summary.Failed); workers=$Active; retries=$retries; cooldown=${cooldown}s"
    Write-Progress -Activity 'SharePoint legacy document inventory' -Status $text;Write-RunLog $text
    $temp=$progressPath+'.tmp';[IO.File]::WriteAllText($temp,($progress|ConvertTo-Json),[Text.UTF8Encoding]::new($false));Move-Item -LiteralPath $temp -Destination $progressPath -Force
    $journal.Flush();$diagnosticsWriter.Flush();$failureWriter.Flush()
}
function Register-RunFailure($Record,[string]$Stage) {
    $detail=ConvertTo-InventoryError $Record $Stage 'RunFailure';$summary.RunErrors+=@($detail);$summary.Error=$Record.Exception.Message
    $failureWriter.WriteLine([Environment]::NewLine+'RUN FAILURE: '+($detail|ConvertTo-Json -Depth 12));$diagnosticsWriter.WriteLine((@{RunId=$runId;RunError=$detail}|ConvertTo-Json -Depth 12 -Compress))
    Write-RunLog ($Stage+': '+$Record.Exception.Message) 'ERROR'
}
function Get-CompletionText {
    $lines=[Collections.Generic.List[string]]::new()
    $lines.Add('==========================================================================')
    $lines.Add(' SHAREPOINT LEGACY DOCUMENT INVENTORY | v3.10.1')
    $lines.Add('==========================================================================')
    $lines.Add((' RESULT       : {0}' -f $summary.Status));$lines.Add((' RUN          : {0}' -f $runId))
    $lines.Add((' LIBRARY      : {0}' -f $summary.LibraryUrl));$lines.Add((' SCOPE / MODE : {0} / {1}' -f $summary.FolderScope,$summary.Mode))
    $lines.Add((' ACCOUNT      : {0}' -f $summary.Account));$lines.Add((' START / END  : {0} / {1}' -f $summary.Started,$summary.Finished))
    $lines.Add((' DURATION     : {0:N1} seconds' -f $summary.ElapsedSeconds));$lines.Add('')
    $lines.Add(' COVERAGE')
    $lines.Add((' Library items enumerated       : {0:N0} ({1:N0} files; {2:N0} folders)' -f $summary.ItemsEnumerated,$summary.FilesEnumerated,$summary.FoldersEnumerated))
    $lines.Add((' Matching .doc files in scope    : {0:N0}' -f $summary.MatchedDocFiles))
    $lines.Add((' Enumeration completed          : {0}' -f $summary.EnumerationComplete))
    if($CountOnly){$lines.Add((' .doc files counted within depth : {0:N0}' -f $summary.DocFiles));$lines.Add((' Excluded by depth limit         : {0:N0}' -f $summary.DocFilesExcludedByDepth));$lines.Add(' No per-document processing or downloads requested.')}
    else{
        $lines.Add((' Documents processed            : {0:N0}' -f $summary.DocFiles));$lines.Add(('   Succeeded                    : {0:N0}' -f $summary.Succeeded))
        $lines.Add(('   Needs review                 : {0:N0}' -f $summary.NeedsReview));$lines.Add(('   Failed                       : {0:N0}' -f $summary.Failed))
        $lines.Add((' Matched but not processed      : {0:N0}' -f $summary.NotProcessed));$lines.Add((' Complete content inspections   : {0:N0}' -f $summary.ScansCompleted))
        $lines.Add((' Documents with VBA indicators  : {0:N0}' -f $summary.VbaIndicatorsDetected))
    }
    $lines.Add((' Run-level failures             : {0:N0}' -f @($summary.RunErrors).Count));$lines.Add('')
    $lines.Add(' OUTPUT FILES')
    foreach($artifact in $script:artifacts){$label='CREATED';if(-not [IO.File]::Exists($artifact.Path)){$label='NOT CREATED / INCOMPLETE'};$lines.Add((' [{0}] {1}' -f $label,$artifact.Purpose));$lines.Add('   '+$artifact.Path)}
    $lines.Add('');$lines.Add(' LOCAL DOCUMENT COPIES')
    if($download){$lines.Add((' Downloaded: {0:N0}; download failures: {1:N0}; layout: {2}' -f $summary.Downloaded,$summary.DownloadFailures,$DownloadLayout));$lines.Add(' Directory: '+$summary.DownloadRoot);$lines.Add(' ReviewAndFailures maps failed/review records to retained copies through LocalPath.')}
    else{$lines.Add(' Not requested. No document contents were downloaded.')}
    $lines.Add('');$lines.Add(' INTERPRETATION')
    $lines.Add(' Success means the requested operations completed, not that a file is safe.')
    $lines.Add(' Review includes partial/unsupported inspections, missing metadata or version changes.')
    $lines.Add(' Local VBA findings describe the SHA256 snapshot; server correlation is separate.')
    if($summary.Status -ne 'Completed'){$lines.Add(' Review FailureDetails and ReviewAndFailures before relying on coverage.')}
    $lines.Add('==========================================================================')
    return ($lines -join [Environment]::NewLine)
}
Add-Artifact 'Execution log' $logPath;Add-Artifact 'Full document records (JSONL)' $journalPath
Add-Artifact 'Structured per-document diagnostics' $diagnosticsPath;Add-Artifact 'Detailed failures and review log' $failureLogPath
Add-Artifact 'Machine-readable run summary' $summaryPath;Add-Artifact 'Readable run summary' $summaryTextPath;Add-Artifact 'Progress snapshot' $progressPath
try {
    Write-RunLog "Starting v3.10.1; $ThrottleLimit batch workers; page size $BatchSize."
    if($CountOnly -and $DownloadFiles){throw '-CountOnly cannot be combined with -DownloadFiles.'}
    if($PSVersionTable.PSVersion.Major -ne 5 -or [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA'){throw 'Run with powershell.exe -NoProfile -STA -File Export-SPOLegacyDocInventory.ps1'}
    $dependencyDir=Join-Path $PSScriptRoot '.dependencies';[void][IO.Directory]::CreateDirectory($dependencyDir);$env:PSModulePath=$dependencyDir+';'+$env:PSModulePath
    Import-RequiredModule 'PnP.PowerShell' '1.12.0'
    $runtimePath=Join-Path $PSScriptRoot 'SPOLegacyDocInventory.Runtime.cs'
    if(-not (Test-Path -LiteralPath $runtimePath -PathType Leaf)){throw "Runtime source file not found: $runtimePath"}
    if(-not ('SPOInventoryV310.StreamingWorkbook' -as [type])){
        Add-Type -Path $runtimePath -ReferencedAssemblies @('System.dll','System.Core.dll','System.Xml.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll',[Microsoft.SharePoint.Client.ClientContext].Assembly.Location,[Microsoft.SharePoint.Client.ClientRuntimeContext].Assembly.Location)
    }
    if(-not ('SPOInventoryV310.LocalIO' -as [type])){throw 'SPOInventoryV310 runtime failed to load.'}
    if([SPOInventoryV310.BuildInfo]::Version -ne '3.10.1'){throw 'Runtime version mismatch. Start a fresh PowerShell process with the complete v3.10.1 package.'}
    $legacyPaths=$false;$blockLongPaths=$false
    [void][AppContext]::TryGetSwitch('Switch.System.IO.UseLegacyPathHandling',[ref]$legacyPaths)
    [void][AppContext]::TryGetSwitch('Switch.System.IO.BlockLongPaths',[ref]$blockLongPaths)
    $summary['LocalFileIO']='Win32UnicodeWithSafeFileHandle'
    $summary['HostPathSettings']=@{UseLegacyPathHandling=$legacyPaths;BlockLongPaths=$blockLongPaths}
    Write-RunLog ('Local document I/O: Windows Unicode APIs; host UseLegacyPathHandling={0}, BlockLongPaths={1} (unchanged).' -f $legacyPaths,$blockLongPaths)
    $summary['ComponentSHA256']=@('Export-SPOLegacyDocInventory.ps1','SPOLegacyDocInventory.Engine.ps1','SPOLegacyDocInventory.Runtime.cs','Inspect-LegacyWord.ps1')|ForEach-Object {$component=Join-Path $PSScriptRoot $_;if(Test-Path -LiteralPath $component){[pscustomobject]@{File=$_;SHA256=(Get-FileHash -LiteralPath $component -Algorithm SHA256).Hash}}}
    if(-not [SPOInventoryV310.LocalIO].GetMethod('CopyLimited',[Reflection.BindingFlags]'Public,Static')){throw 'SPOInventoryV310 runtime validation failed: CopyLimited() is unavailable.'}
    $gate=[SPOInventoryV310.RetryGate]::new()
    if(-not $LibraryUrl){$LibraryUrl=$ConfiguredLibraryUrl};if([string]::IsNullOrWhiteSpace($LibraryUrl)){throw 'Supply -LibraryUrl with your SharePoint document library URL.'};$uri=[uri]$LibraryUrl
    if(-not $uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or $uri.UserInfo -or $uri.Fragment){throw 'Supply a direct HTTPS library root/view URL; encode literal # as %23.'}
    $origin=$uri.GetLeftPart([UriPartial]::Authority);$libraryPath=[uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/') -replace '(?i)/Forms/[^/]+\.aspx$',''
    if($uri.Query){Write-RunLog 'View filters ignored. Use -Folder for a decoded library-relative scope.'}
    if(-not $SiteUrl){$SiteUrl=$origin;if($libraryPath -match '^/(sites|teams|personal)/[^/]+'){$SiteUrl+=$Matches[0]}}
    $site=[uri]$SiteUrl;if($site.Scheme -ne 'https' -or $site.Authority -ine $uri.Authority -or $site.UserInfo -or $site.Query -or $site.Fragment){throw 'SiteUrl must be a direct HTTPS web URL on the same host.'}
    $SiteUrl=$site.AbsoluteUri.TrimEnd('/');$summary.SiteUrl=$SiteUrl;$summary.LibraryUrl=$LibraryUrl
    $servicePoint=[Net.ServicePointManager]::FindServicePoint($site);$previousConnectionLimit=$servicePoint.ConnectionLimit
    $servicePoint.ConnectionLimit=[math]::Max($previousConnectionLimit,$ThrottleLimit+1)
    $scan=($CheckVBA -or $CheckMacro -or $CheckBoth);$download=($scan -or $DownloadFiles)
    $summary['ContentScanner']='NotRequested';if($scan){$summary.ContentScanner='PowerShell-CFB-OVBA-1.4.1'}
    if($scan){Initialize-Scanner}
    $downloadRoot='';if($download){if(-not $DownloadDirectory){$DownloadDirectory=Join-Path $PSScriptRoot 'DocInventoryDownloads'};$downloadRoot=[SPOInventoryV310.LocalIO]::FullPath($DownloadDirectory.TrimEnd('\')+'\'+$runId);[SPOInventoryV310.LocalIO]::CreateDirectory($downloadRoot)}
    $summary['DownloadRoot']=$downloadRoot
    if($download){
        # Validate local staging before potentially enumerating thousands of files.
        $probe=$downloadRoot+'\.write_test_'+[guid]::NewGuid().ToString('N')+'.tmp'
        try{$probeStream=[SPOInventoryV310.LocalIO]::CreateNew($probe);$probeStream.Dispose();[SPOInventoryV310.LocalIO]::Delete($probe)}
        catch{throw ('Download staging directory is not writable: {0}. {1}' -f $downloadRoot,$_.Exception.Message)}
    }
    if($Connection){$ctx=Get-PnPContext -Connection $Connection}else{
        $auth=@{Url=$SiteUrl;ReturnConnection=$true;ErrorAction='Stop';ForceAuthentication=$true;Verbose=$false}
        if($AuthenticationMode -eq 'SystemBrowser'){$parsed=[guid]::Empty;if(-not [guid]::TryParse($ClientId,[ref]$parsed) -or $parsed -eq [guid]::Empty){throw 'SystemBrowser needs an approved Entra app ClientId.'};$auth.Interactive=$true;$auth.LaunchBrowser=$true;$auth.ClientId=$ClientId}else{$auth.UseWebLogin=$true}
        Write-RunLog "Sign in to $SiteUrl using $AuthenticationMode."
        $Connection=Connect-PnPOnline @auth;$ctx=Get-PnPContext -Connection $Connection;$ownsContext=$true
    }
    $ctx.Load($ctx.Web);$ctx.Load($ctx.Web.CurrentUser);$gate.ExecuteRead($ctx,$RetryCount)
    if($ctx.Web.Url.TrimEnd('/') -ine $SiteUrl){throw 'Connection is for a different SharePoint web.'}
    $summary.Account=$ctx.Web.CurrentUser.LoginName
    $matches=@(Get-PnPList -Includes RootFolder,BaseType,ItemCount -Connection $Connection | Where-Object {$_.BaseType.ToString() -eq 'DocumentLibrary' -and $_.RootFolder.ServerRelativeUrl.TrimEnd('/') -ieq $libraryPath})
    if($matches.Count -ne 1){throw 'Library root not uniquely found. For subwebs supply -SiteUrl.'};$list=$matches[0]
    $folderPath=Resolve-LibraryFolder $Folder $libraryPath
    if($Folder){$f=$ctx.Web.GetFolderByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl($folderPath));$ctx.Load($f);$gate.ExecuteRead($ctx,$RetryCount);if(-not $f.Exists){throw 'Requested folder does not exist.'}}
    $summary['FolderScope']=$folderPath;$summary.LibraryItemCountEstimate=$list.ItemCount
    $q=[Microsoft.SharePoint.Client.CamlQuery]::new();$q.ViewXml="<View Scope='RecursiveAll'><Query><OrderBy Override='TRUE'><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><ViewFields><FieldRef Name='ID'/></ViewFields><RowLimit>1</RowLimit></View>"
    $highest=$ctx.Web.Lists.GetById($list.Id).GetItems($q);$ctx.Load($highest);$gate.ExecuteRead($ctx,$RetryCount)
    if($highest.Count){$summary.HighWatermarkId=$highest[0].Id}
    if($CountOnly){$summary['MaxDepthFromLibraryRoot']=$MaxDepth;$scopeRelative='';if($folderPath -ine $libraryPath){$scopeRelative=Get-LibraryRelativePath $folderPath $libraryPath};if($MaxDepth -ge 0 -and (Get-RelativeFolderDepth $scopeRelative) -gt $MaxDepth){throw 'Folder is deeper than MaxDepth. Use -MaxDepth -1 for unlimited depth.'};Add-FolderCount $scopeRelative $false}
    else{$allFields=@(Get-PnPField -List $list.Id -Includes SchemaXml,FromBaseType -Connection $Connection);$schema=@($allFields | Where-Object {Test-UserMetadataField $_} | ForEach-Object {[pscustomobject]@{InternalName=$_.InternalName;DisplayName=$_.Title;Type=$_.TypeAsString;Hidden=$_.Hidden;ReadOnly=$_.ReadOnlyField}});$summary['MetadataFieldsSelected']=$schema.Count}
    Save-Journal 'Fields' $schema
    $scannerTemp=Join-Path $runDirectory 'ScannerTemp';if($scan){[void][IO.Directory]::CreateDirectory($scannerTemp)}
    $baseFields=@('ID','FileLeafRef','FileRef','FSObjType');if(-not $CountOnly){$baseFields+=@('UniqueId','Modified','File_x0020_Size')}
    $config=@{ListId=$list.Id;HighWatermark=[long]$summary.HighWatermarkId;PageSize=$BatchSize;ThrottleLimit=$ThrottleLimit;FieldBatchSize=$FieldBatchSize;RetryCount=$RetryCount;ProgressSeconds=$ProgressSeconds;EnginePath=$enginePath;Gate=$gate;BaseFields=$baseFields;Fields=$schema;LibraryPath=$libraryPath;FolderPath=$folderPath;Origin=$origin;CountOnly=[bool]$CountOnly;Download=$download;Scan=$scan;CheckVBA=($CheckVBA -or $CheckBoth);CheckMacro=($CheckMacro -or $CheckBoth);CheckBoth=[bool]$CheckBoth;DownloadRoot=$downloadRoot;DownloadLayout=$DownloadLayout;MaxBytes=([long]$MaxScanMB*1MB);ScannerExecutable=$scannerExecutable;ScannerHelper=(Join-Path $PSScriptRoot 'Inspect-LegacyWord.ps1');ScannerTemp=$scannerTemp;ScanTimeoutSeconds=$ScanTimeoutSeconds}
    $contextFactory={ $clone=[Microsoft.SharePoint.Client.ClientContextExtensions]::Clone($ctx,[string]$SiteUrl,$null);$clone.RequestTimeout=120000;$clone.DisableReturnValueCache=$true;return $clone }
    $onPage={param($page)
        $summary.ItemsEnumerated+=$page.Items.Count
        foreach($item in $page.Items){$v=$item.Values;if([int]$v['FSObjType'] -eq 0){$summary.FilesEnumerated++}else{$summary.FoldersEnumerated++};$path=$v['FileRef'];if(-not $path.StartsWith($folderPath+'/',[StringComparison]::OrdinalIgnoreCase)){continue};$relative=Get-LibraryRelativePath $path $libraryPath
            if([int]$v['FSObjType'] -ne 0){if($CountOnly){Add-FolderCount $relative $false};continue}
            if([IO.Path]::GetExtension($v['FileLeafRef']) -ine '.doc'){continue};$summary.MatchedDocFiles++
            if($CountOnly){$slash=$relative.LastIndexOf('/');$parent='';if($slash -ge 0){$parent=$relative.Substring(0,$slash)};if($MaxDepth -lt 0 -or (Get-RelativeFolderDepth $parent) -le $MaxDepth){Add-FolderCount $parent $true;$summary.DocFiles++}else{$summary.DocFilesExcludedByDepth++}}
        }
    }
    $onResult=${function:Accept-InventoryResult}
    $summary.Status='Running'
    Invoke-BoundedInventory $config $contextFactory $onPage $onResult ${function:Write-ProgressSnapshot} | Out-Null
    $summary.EnumerationComplete=$true;$summary.Status='Completed';if($summary.Errors -gt 0 -or $summary.FailedBatches -gt 0){$summary.Status='CompletedWithIssues'}
} catch {$runFailure=$_;$summary.Status='Failed';Register-RunFailure $_ 'Run'}
finally {
    try {
        Complete-InventoryWorkbook
        if(-not $CountOnly -and ('SPOInventoryV310.StreamingWorkbook' -as [type])){
            # Header-only files make a zero-result category explicit.
            if(-not $script:outcomeBooks.ContainsKey('SuccessfulDocuments')){[void](Get-OutcomeWorkbook 'SuccessfulDocuments')}
            if(-not $script:outcomeBooks.ContainsKey('ReviewAndFailures')){[void](Get-OutcomeWorkbook 'ReviewAndFailures')}
            foreach($state in $script:outcomeBooks.Values){if($state.Book){try{$state.Book.Complete()}finally{$state.Book.Dispose();$state.Book=$null}}}
        }
        if($CountOnly -and ('SPOInventoryV310.StreamingWorkbook' -as [type])){
            $part=0;$count=0;$book=$null
            try {
                foreach($row in $folderCounts.Values){
                    if(-not $book -or $count -ge $ReportMetadataRowsPerWorkbook){if($book){$book.Complete();$book.Dispose()};$part++;$count=0;$name='FolderCounts_{0}_Part{1:D4}.xlsx' -f $runId,$part
                        $book=[SPOInventoryV310.StreamingWorkbook]::new((Join-Path $runDirectory $name));[void]$book.AddSheet('FolderCounts',@('Folder','Depth','DirectDocCount','DocsWithinDepthLimit'));$summary.ReportParts+=$name;Add-Artifact 'Folder counts' (Join-Path $runDirectory $name)}
                    Save-Journal 'FolderCount' $row;$book.AddRow(0,[string[]]@($row.Folder,$row.Depth,$row.DirectDocCount,$row.DocsWithinDepthLimit));$count++
                };if($book){$book.Complete()}
            }finally{if($book){$book.Dispose()}}
        }
    } catch {if(-not $runFailure){$runFailure=$_};$summary.Status='ReportExportFailed';Register-RunFailure $_ 'ExportReports'}
    finally {foreach($state in $script:outcomeBooks.Values){if($state.Book){$state.Book.Dispose()}}}
    $summary.NotProcessed=[math]::Max(0,$summary.MatchedDocFiles-$summary.DocFiles-$summary.DocFilesExcludedByDepth)
    $summary.Finished=[DateTimeOffset]::Now.ToString('o');$summary['ElapsedSeconds']=[math]::Round($clock.Elapsed.TotalSeconds,2)
    if($gate){$summary['Requests']=$gate.Requests;$summary['Retries']=$gate.Retries}
    $summary['Findings']=$script:findingCounts
    $summary['ReportsDirectory']=$runDirectory;$summary['ExecutionLog']=$logPath;$summary['FailureLog']=$failureLogPath;$summary['Diagnostics']=$diagnosticsPath;$summary['Journal']=$journalPath
    $summaryBookPath=Join-Path $runDirectory ('RunSummary_'+$runId+'.xlsx');Add-Artifact 'Run summary workbook' $summaryBookPath
    if('SPOInventoryV310.StreamingWorkbook' -as [type]){
        try {
            $book=[SPOInventoryV310.StreamingWorkbook]::new($summaryBookPath);[void]$book.AddSheet('RunSummary',@('Property','Value'))
            $firstKeys=@('Status','LibraryUrl','FolderScope','Mode','Account','Started','Finished','ElapsedSeconds','ItemsEnumerated','FilesEnumerated','FoldersEnumerated','MatchedDocFiles','DocFiles','Succeeded','NeedsReview','Failed','NotProcessed','EnumerationComplete','Downloaded','DownloadFailures','DownloadRoot')
            foreach($key in @($firstKeys)+@($summary.Keys|Where-Object{$_ -notin $firstKeys -and $_ -notin @('Artifacts','RunErrors','Findings','ComponentSHA256')})){$book.AddRow(0,[string[]]@($key,($summary[$key] -join '; ')))}
            [void]$book.AddSheet('OutputFiles',@('Purpose','Path'));foreach($artifact in $script:artifacts){$book.AddRow(1,[string[]]@($artifact.Purpose,$artifact.Path))}
            [void]$book.AddSheet('Findings',@('ScanStatus / Format','Documents'));foreach($key in $script:findingCounts.Keys){$book.AddRow(2,[string[]]@($key,$script:findingCounts[$key]))}
            $book.Complete()
        }catch{if(-not $runFailure){$runFailure=$_};$summary.Status='ReportExportFailed';Register-RunFailure $_ 'ExportSummary'}finally{if($book){$book.Dispose()}}
    }
    try{Write-ProgressSnapshot $summary.LastEnumeratedId 0 $false}catch{if(-not $runFailure){$runFailure=$_};$summary.Status='ReportExportFailed';Register-RunFailure $_ 'WriteProgress'}
    if($servicePoint){$servicePoint.ConnectionLimit=$previousConnectionLimit};if($ownsContext -and $ctx){$ctx.Dispose()}
    $env:PNPPOWERSHELL_UPDATECHECK=$previousPnPUpdateCheck
    $journal.Flush();$diagnosticsWriter.Flush();$failureWriter.Flush()
    foreach($artifact in $script:artifacts){$artifact.Created=[IO.File]::Exists($artifact.Path)}
    # Create both summaries before final path/existence rendering, then refresh their manifest.
    [IO.File]::WriteAllText($summaryPath,($summary|ConvertTo-Json -Depth 18),[Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText($summaryTextPath,'Preparing final summary',[Text.UTF8Encoding]::new($false))
    foreach($artifact in $script:artifacts){$artifact.Created=[IO.File]::Exists($artifact.Path)};$summary.Artifacts=$script:artifacts.ToArray()
    $completion=Get-CompletionText
    [IO.File]::WriteAllText($summaryTextPath,$completion,[Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText($summaryPath,($summary|ConvertTo-Json -Depth 18),[Text.UTF8Encoding]::new($false))
    Save-Journal 'Summary' $summary;$journal.Dispose();$diagnosticsWriter.Dispose();$failureWriter.Dispose()
    Write-Progress -Activity 'SharePoint legacy document inventory' -Completed
    Write-RunLog "Finished: $($summary.Status). Summary: $summaryTextPath" -Quiet
    Write-Host '';Write-Host $completion -ForegroundColor Cyan
}
if($runFailure){throw $runFailure}
[pscustomobject]@{Status=$summary.Status;Documents=$summary.DocFiles;Succeeded=$summary.Succeeded;NeedsReview=$summary.NeedsReview;Failed=$summary.Failed;ReportDirectory=$runDirectory;Summary=$summaryPath;Downloads=$summary.DownloadRoot}
