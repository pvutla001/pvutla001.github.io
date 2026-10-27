#Requires -Version 5.1
<#
.SYNOPSIS
Stream a large SharePoint library inventory with bounded parallel metadata/content workers.
.DESCRIPTION
Version 3.1. Windows PowerShell 5.1 STA; legacy PnP.PowerShell 1.12.0 MFA retained.
Default: .doc metadata only, all folder depths. CountOnly retains MaxDepth=2;
use MaxDepth=-1 for unlimited depth. Folder is a decoded library-relative path.
One paged discovery request plus up to ThrottleLimit workers; every request batch
owns a separate cloned CSOM context. No concurrent sharing of a mutable context.
ID-ordered discovery captures a starting high-watermark and uses no nonindexed
extension/path filter. Metadata is read per page/field batch, not per file.
JSONL and worksheet XML stream to disk. Excel reports rotate into bounded parts;
no ImportExcel dependency or whole-library in-memory workbook is required.
Progress is displayed, logged and saved to Progress.json. Partial records survive
failure. This is not a transactional snapshot or a resumable synchronization engine.
DownloadFiles/CheckVBA/CheckMacro/CheckBoth retain mirrored local copies. Content
checks use native PowerShell/.NET parsing and never execute Word/VBA. Metadata/count modes
never download contents. VBA/macros overlap; NotDetected does not mean malware-free.
Files hidden by permissions, historical versions and newly added IDs above the
starting watermark are not included. File moves/changes during scanning are possible.
Keep SPOLegacyDocInventory.Engine.ps1 and SPOLegacyDocInventory.Runtime.cs beside
this script; content checks also require Inspect-LegacyWord.ps1. See the accompanying blog post.
.EXAMPLE
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/Example/Documents' -ThrottleLimit 4 -BatchSize 500
.EXAMPLE
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/Example/Documents' -CountOnly -MaxDepth -1
.EXAMPLE
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl 'https://contoso.sharepoint.com/sites/Example/Documents' -CheckBoth -NoInstall
Downloads and statically inspects .doc files using the native PowerShell helper.
#>
[CmdletBinding(DefaultParameterSetName='Info')]
param(
    [string]$LibraryUrl,[string]$SiteUrl,[string]$Folder,[string]$DownloadDirectory,
    [switch]$DownloadFiles,
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
$enginePath=Join-Path $PSScriptRoot 'SPOLegacyDocInventory.Engine.ps1'
. $enginePath
$runId=[DateTimeOffset]::Now.ToString('yyyyMMdd_HHmmss_fff')+'_'+[guid]::NewGuid().ToString('N').Substring(0,8)
if(-not $OutputDirectory){$OutputDirectory=Join-Path $PSScriptRoot 'DocInventoryReports'}
$runDirectory=[IO.Path]::GetFullPath((Join-Path $OutputDirectory $runId));[void][IO.Directory]::CreateDirectory($runDirectory)
$logPath=Join-Path $runDirectory ('Run_'+$runId+'.log');$journalPath=Join-Path $runDirectory 'Inventory.jsonl'
$journal=[IO.StreamWriter]::new($journalPath,$false,[Text.UTF8Encoding]::new($false),65536)
$summary=[ordered]@{ScriptVersion='3.1';RunId=$runId;Started=[DateTimeOffset]::Now.ToString('o');Finished='';Status='Starting';LibraryUrl='';SiteUrl='';Account='';Mode=$PSCmdlet.ParameterSetName;ItemsEnumerated=0L;MatchedDocFiles=0L;DocFiles=0L;Errors=0L;FailedBatches=0;EnumerationComplete=$false;HighWatermarkId=0;LibraryItemCountEstimate=0;LastEnumeratedId=0;MetadataRows=0L;DocFilesExcludedByDepth=0L;ThrottleLimit=$ThrottleLimit;BatchSize=$BatchSize;Coverage='Current visible .doc files; recursive; no historical versions; IDs at or below initial high watermark';ScanCoverage='Local VBA indicators only; not exhaustive WordBasic, embedded objects, external-template or stomped-VBA detection';ReportParts=@();Error=''}
$script:inventoryWorkbook=$null;$script:inventoryPart=0;$script:partDocs=0;$script:partMetadata=0
$script:documentHeaders=@('ItemId','UniqueId','FileName','Url','LibraryRelativePath','FolderDepth','SizeBytes','Modified','MetadataStatus','DownloadStatus','LocalPath','ScanStatus','VBA','VBASourceModules','Macro','Both','Format','SHA256','Notes','Error')
$script:metadataHeaders=@('ItemId','FileName','Field','DisplayName','Type','Value','Status')
$folderCounts=@{};$schema=@();$ctx=$null;$ownsContext=$false;$runFailure=$null;$gate=$null;$clock=[Diagnostics.Stopwatch]::StartNew();$scannerExecutable='';$book=$null;$servicePoint=$null;$previousConnectionLimit=0
function Write-RunLog([string]$Message,[string]$Level='INFO') {
    $line='{0} [{1}] {2}' -f [DateTimeOffset]::Now.ToString('o'),$Level,$Message
    [IO.File]::AppendAllText($logPath,$line+[Environment]::NewLine,[Text.UTF8Encoding]::new($false));Write-Host $line
}
function Save-Journal([string]$Kind,$Data) {$journal.WriteLine((@{Kind=$Kind;Timestamp=[DateTimeOffset]::Now.ToString('o');Data=$Data}|ConvertTo-Json -Depth 14 -Compress))}
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
    if($script:inventoryWorkbook){$script:inventoryWorkbook.Complete();$script:inventoryWorkbook.Dispose();$script:inventoryWorkbook=$null}
}
function Write-DocumentReport($Document,[object[]]$Metadata) {
    if(-not $script:inventoryWorkbook -or $script:partDocs -ge $ReportDocumentsPerWorkbook -or ($script:partMetadata+$Metadata.Count) -gt $ReportMetadataRowsPerWorkbook){
        Complete-InventoryWorkbook;$script:inventoryPart++;$script:partDocs=0;$script:partMetadata=0
        $name='Inventory_{0}_{1:D4}.xlsx' -f $runId,$script:inventoryPart
        $script:inventoryWorkbook=[SPOInventoryV3.StreamingWorkbook]::new((Join-Path $runDirectory $name))
        [void]$script:inventoryWorkbook.AddSheet('Documents',$script:documentHeaders);[void]$script:inventoryWorkbook.AddSheet('Metadata',$script:metadataHeaders)
        $fieldSheet=$script:inventoryWorkbook.AddSheet('Fields',@('InternalName','DisplayName','Type','Hidden','ReadOnly'))
        foreach($field in $schema){$script:inventoryWorkbook.AddRow($fieldSheet,[string[]]@($field.InternalName,$field.DisplayName,$field.Type,$field.Hidden,$field.ReadOnly))}
        $summary.ReportParts+= $name
    }
    $values=foreach($header in $script:documentHeaders){$property=$Document.PSObject.Properties[$header];if($property){[string]$property.Value}else{''}}
    $script:inventoryWorkbook.AddRow(0,[string[]]$values);$script:partDocs++
    foreach($m in $Metadata){$script:inventoryWorkbook.AddRow(1,[string[]]@($m.ItemId,$m.FileName,$m.Field,$m.DisplayName,$m.Type,$m.Value,$m.Status));$script:partMetadata++}
}
function Write-ProgressSnapshot([long]$Cursor,[int]$Active,[bool]$Discovering){
    $summary.LastEnumeratedId=$Cursor;$rate=[math]::Round($summary.DocFiles/[math]::Max(1,$clock.Elapsed.TotalSeconds),2)
    $retries=0L;$cooldown=0;if($gate){$retries=$gate.Retries;$cooldown=[math]::Ceiling($gate.RemainingSeconds)}
    $progress=[ordered]@{Utc=[DateTime]::UtcNow.ToString('o');Status=$summary.Status;EnumeratedItems=$summary.ItemsEnumerated;LibraryItemCountEstimate=$summary.LibraryItemCountEstimate;MatchedDocFiles=$summary.MatchedDocFiles;ReportedDocFiles=$summary.DocFiles;ActiveWorkers=$Active;DiscoveryRequestActive=$Discovering;FilesPerSecond=$rate;RetryCount=$retries;CooldownSeconds=$cooldown;Errors=$summary.Errors;LastEnumeratedId=$Cursor;HighWatermarkId=$summary.HighWatermarkId}
    $text="items=$($summary.ItemsEnumerated) / ~$($summary.LibraryItemCountEstimate); .doc matched/reported=$($summary.MatchedDocFiles)/$($summary.DocFiles); workers=$Active; retries=$retries; cooldown=${cooldown}s; errors=$($summary.Errors)"
    Write-Progress -Activity 'SharePoint .doc inventory' -Status $text
    Write-RunLog $text
    $temp=Join-Path $runDirectory 'Progress.json.tmp';[IO.File]::WriteAllText($temp,($progress|ConvertTo-Json),[Text.UTF8Encoding]::new($false));Move-Item -LiteralPath $temp -Destination (Join-Path $runDirectory 'Progress.json') -Force
    $journal.Flush()
}
try {
    Write-RunLog "Starting v3.1; $ThrottleLimit batch workers; page size $BatchSize."
    if($CountOnly -and $DownloadFiles){throw '-CountOnly cannot be combined with -DownloadFiles.'}
    if($PSVersionTable.PSVersion.Major -ne 5 -or [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA'){throw 'Run with powershell.exe -NoProfile -STA -File Export-SPOLegacyDocInventory.ps1'}
    $dependencyDir=Join-Path $PSScriptRoot '.dependencies';[void][IO.Directory]::CreateDirectory($dependencyDir);$env:PSModulePath=$dependencyDir+';'+$env:PSModulePath
    Import-RequiredModule 'PnP.PowerShell' '1.12.0'
    if(-not ('SPOInventoryV3.StreamingWorkbook' -as [type])){
        Add-Type -Path (Join-Path $PSScriptRoot 'SPOLegacyDocInventory.Runtime.cs') -ReferencedAssemblies @('System.dll','System.Core.dll','System.Xml.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll',[Microsoft.SharePoint.Client.ClientContext].Assembly.Location,[Microsoft.SharePoint.Client.ClientRuntimeContext].Assembly.Location)
    }
    $gate=[SPOInventoryV3.RetryGate]::new()
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
    $summary['ContentScanner']='NotRequested';if($scan){$summary.ContentScanner='PowerShell-CFB-OVBA-1.0'}
    if($scan){Initialize-Scanner}
    $downloadRoot='';if($download){if(-not $DownloadDirectory){$DownloadDirectory=Join-Path $PSScriptRoot 'DocInventoryDownloads'};$downloadRoot=[IO.Path]::GetFullPath((Join-Path $DownloadDirectory $runId));[void][IO.Directory]::CreateDirectory($downloadRoot)}
    $summary['DownloadRoot']=$downloadRoot
    if($Connection){$ctx=Get-PnPContext -Connection $Connection}else{
        $auth=@{Url=$SiteUrl;ReturnConnection=$true;ErrorAction='Stop';ForceAuthentication=$true}
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
    $config=@{ListId=$list.Id;HighWatermark=[long]$summary.HighWatermarkId;PageSize=$BatchSize;ThrottleLimit=$ThrottleLimit;FieldBatchSize=$FieldBatchSize;RetryCount=$RetryCount;ProgressSeconds=$ProgressSeconds;EnginePath=$enginePath;Gate=$gate;BaseFields=$baseFields;Fields=$schema;LibraryPath=$libraryPath;FolderPath=$folderPath;Origin=$origin;CountOnly=[bool]$CountOnly;Download=$download;Scan=$scan;CheckVBA=($CheckVBA -or $CheckBoth);CheckMacro=($CheckMacro -or $CheckBoth);CheckBoth=[bool]$CheckBoth;DownloadRoot=$downloadRoot;MaxBytes=([long]$MaxScanMB*1MB);ScannerExecutable=$scannerExecutable;ScannerHelper=(Join-Path $PSScriptRoot 'Inspect-LegacyWord.ps1');ScannerTemp=$scannerTemp;ScanTimeoutSeconds=$ScanTimeoutSeconds}
    $contextFactory={ $clone=[Microsoft.SharePoint.Client.ClientContextExtensions]::Clone($ctx,[string]$SiteUrl,$null);$clone.RequestTimeout=120000;$clone.DisableReturnValueCache=$true;return $clone }
    $onPage={param($page)
        $summary.ItemsEnumerated+=$page.Items.Count
        foreach($item in $page.Items){$v=$item.Values;$path=$v['FileRef'];if(-not $path.StartsWith($folderPath+'/',[StringComparison]::OrdinalIgnoreCase)){continue};$relative=Get-LibraryRelativePath $path $libraryPath
            if([int]$v['FSObjType'] -ne 0){if($CountOnly){Add-FolderCount $relative $false};continue}
            if([IO.Path]::GetExtension($v['FileLeafRef']) -ine '.doc'){continue};$summary.MatchedDocFiles++
            if($CountOnly){$slash=$relative.LastIndexOf('/');$parent='';if($slash -ge 0){$parent=$relative.Substring(0,$slash)};if($MaxDepth -lt 0 -or (Get-RelativeFolderDepth $parent) -le $MaxDepth){Add-FolderCount $parent $true;$summary.DocFiles++}else{$summary.DocFilesExcludedByDepth++}}
        }
    }
    $onResult={param($result,$failure,$page)
        if($failure){
            $summary.FailedBatches++;Write-RunLog "Batch ($($page.After), $($page.Through)] failed: $($failure.Exception.Message)" 'ERROR'
            $failedDocs=[Collections.Generic.List[object]]::new()
            foreach($item in $page.Items){$v=$item.Values;if([int]$v['FSObjType'] -eq 0 -and [IO.Path]::GetExtension($v['FileLeafRef']) -ieq '.doc' -and $v['FileRef'].StartsWith($folderPath+'/',[StringComparison]::OrdinalIgnoreCase)){$failedDocs.Add([pscustomobject]@{ItemId=$item.Id;UniqueId=$v['UniqueId'];FileName=$v['FileLeafRef'];LibraryRelativePath=(Get-LibraryRelativePath $v['FileRef'] $libraryPath);MetadataStatus='Failed';Error=$failure.Exception.Message})}}
            $result=[pscustomobject]@{Documents=$failedDocs.ToArray();Metadata=@();Errors=$failedDocs.Count;After=$page.After;Through=$page.Through}
        }
        $groups=@{};foreach($m in $result.Metadata){if(-not $groups.ContainsKey($m.ItemId)){$groups[$m.ItemId]=[Collections.Generic.List[object]]::new()};$groups[$m.ItemId].Add($m)}
        foreach($row in $result.Documents){$values=@();if($groups.ContainsKey($row.ItemId)){$values=$groups[$row.ItemId].ToArray()};Save-Journal 'Document' @{Document=$row;Metadata=$values};Write-DocumentReport $row $values;$summary.DocFiles++;$summary.MetadataRows+=$values.Count}
        $summary.Errors+=$result.Errors;Save-Journal 'BatchCommitted' @{After=$page.After;Through=$page.Through;Documents=$result.Documents.Count;Errors=$result.Errors};$journal.Flush()
    }
    $summary.Status='Running'
    Invoke-BoundedInventory $config $contextFactory $onPage $onResult ${function:Write-ProgressSnapshot} | Out-Null
    $summary.EnumerationComplete=$true;$summary.Status='Completed';if($summary.Errors -gt 0 -or $summary.FailedBatches -gt 0){$summary.Status='CompletedWithErrors'}
} catch {$runFailure=$_;$summary.Status='Failed';$summary.Error=$_.Exception.Message;Write-RunLog $summary.Error 'ERROR'}
finally {
    try {
        Complete-InventoryWorkbook
        if($CountOnly -and ('SPOInventoryV3.StreamingWorkbook' -as [type])){
            $part=0;$count=0;$book=$null
            try {foreach($row in $folderCounts.Values){if(-not $book -or $count -ge $ReportMetadataRowsPerWorkbook){if($book){$book.Complete();$book.Dispose()};$part++;$count=0;$name='FolderCounts_{0}_{1:D4}.xlsx' -f $runId,$part;$book=[SPOInventoryV3.StreamingWorkbook]::new((Join-Path $runDirectory $name));[void]$book.AddSheet('FolderCounts',@('Folder','Depth','DirectDocCount','DocsWithinDepthLimit'));$summary.ReportParts+=$name};Save-Journal 'FolderCount' $row;$book.AddRow(0,[string[]]@($row.Folder,$row.Depth,$row.DirectDocCount,$row.DocsWithinDepthLimit));$count++};if($book){$book.Complete()}}finally{if($book){$book.Dispose()}}
        }
    }catch{if(-not $runFailure){$runFailure=$_};$summary.Status='ReportExportFailed';$summary.Error=$_.Exception.Message;Write-RunLog "Excel export failed; JSONL retained: $($_.Exception.Message)" 'ERROR'}
    if($script:inventoryWorkbook){$script:inventoryWorkbook.Dispose()}
    $summary.Finished=[DateTimeOffset]::Now.ToString('o');$summary['ElapsedSeconds']=[math]::Round($clock.Elapsed.TotalSeconds,2);if($gate){$summary['Requests']=$gate.Requests;$summary['Retries']=$gate.Retries}
    try{Write-ProgressSnapshot $summary.LastEnumeratedId 0 $false}catch{Write-Warning "Final progress snapshot unavailable: $($_.Exception.Message)"}
    Save-Journal 'Summary' $summary;$journal.Dispose()
    $summary|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $runDirectory 'Summary.json') -Encoding UTF8
    if('SPOInventoryV3.StreamingWorkbook' -as [type]){try{$book=[SPOInventoryV3.StreamingWorkbook]::new((Join-Path $runDirectory ('RunSummary_'+$runId+'.xlsx')));[void]$book.AddSheet('RunSummary',@('Property','Value'));foreach($key in $summary.Keys){$book.AddRow(0,[string[]]@($key,($summary[$key] -join '; ')))};$book.Complete()}catch{Write-RunLog "RunSummary workbook unavailable: $($_.Exception.Message)" 'WARN'}finally{if($book){$book.Dispose()}}}
    Write-Progress -Activity 'SharePoint .doc inventory' -Completed
    if($servicePoint){$servicePoint.ConnectionLimit=$previousConnectionLimit}
    if($ownsContext -and $ctx){$ctx.Dispose()}
    Write-RunLog "Finished: $($summary.Status); .doc reported=$($summary.DocFiles); errors=$($summary.Errors). Reports: $runDirectory"
    Write-Host "Full-value journal: $journalPath"
}
if($runFailure){throw $runFailure}
[pscustomobject]@{Status=$summary.Status;Documents=$summary.DocFiles;Errors=$summary.Errors;ReportDirectory=$runDirectory;Summary=(Join-Path $runDirectory 'Summary.json')}
