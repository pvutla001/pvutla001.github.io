# Shared pure helpers and worker functions for Export-SPOLegacyDocInventory v3.1. 
function Convert-FieldText($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return (ConvertTo-InventoryUtc $Value) }
    if ($Value -is [string] -or $Value -is [ValueType]) { return [string]$Value }
    if ($Value -is [System.Collections.IEnumerable]) { return (@($Value | ForEach-Object { Convert-FieldText $_ }) -join '; ') }
    if ($Value.PSObject.Properties['LookupValue']) {
        $email=''; if ($Value.PSObject.Properties['Email']) { $email=$Value.Email }
        return ('{0} [ID={1}; Email={2}]' -f $Value.LookupValue,$Value.LookupId,$email)
    }
    if ($Value.PSObject.Properties['TermGuid']) { return ('{0} [{1}]' -f $Value.Label,$Value.TermGuid) }
    if ($Value.PSObject.Properties['Url']) { return ('{0} ({1})' -f $Value.Url,$Value.Description) }
    return [string]$Value
}
function Test-UserMetadataField($Field) {
    # Keep useful default properties even when SharePoint marks them hidden/read-only.
    $defaults=@('FileLeafRef','Title','Created','Author','Modified','Editor','File_x0020_Type','File_x0020_Size','_UIVersionString')
    if ($Field.InternalName -in $defaults) { return $true }
    if ($Field.Hidden -or $Field.InternalName -eq 'ContentType') { return $false }
    # SourceID identifies Microsoft schema fields versus custom list/site columns.
    # Do not exclude custom calculated/read-only columns or underscore-prefixed names.
    $schemaXml=[xml]$Field.SchemaXml
    $sourceId=$schemaXml.DocumentElement.GetAttribute('SourceID')
    if ($sourceId -match '^https?://schemas\.microsoft\.com/' -or $Field.FromBaseType) { return $false }
    return $true
}
function Get-LibraryRelativePath([string]$ServerPath,[string]$Root) {
    $prefix=$Root.TrimEnd('/')+'/'
    if (-not $ServerPath.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) { throw 'File path is outside the document library.' }
    return $ServerPath.Substring($prefix.Length)
}
function Resolve-LibraryFolder([string]$Value,[string]$Root) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Root.TrimEnd('/') }
    # Accept decoded library-relative names only, avoiding ambiguous URL decoding.
    if ($Value.StartsWith('/') -or $Value.Contains('\') -or $Value.Contains(':')) { throw 'Folder must be a plain library-relative path, e.g. Policies/Finance (not a URL).' }
    $parts=$Value.TrimEnd('/').Split('/')
    if (@($parts | Where-Object { $_ -eq '.' -or $_ -eq '..' -or $_ -eq '' }).Count) { throw 'Folder contains an empty or traversal segment.' }
    return $Root.TrimEnd('/')+'/'+($parts -join '/')
}
function Get-RelativeFolderDepth([string]$RelativeFolder) {
    if (-not $RelativeFolder) { return 0 }
    return $RelativeFolder.Split('/').Count
}

function New-InventoryQueryXml([long]$After,[long]$Through,[string[]]$Fields,[int]$PageSize) {
    $refs=(@($Fields | Select-Object -Unique | ForEach-Object {"<FieldRef Name='$([Security.SecurityElement]::Escape($_))'/>"}) -join '')
    return "<View Scope='RecursiveAll'><Query><Where><And><Gt><FieldRef Name='ID'/><Value Type='Counter'>$After</Value></Gt><Leq><FieldRef Name='ID'/><Value Type='Counter'>$Through</Value></Leq></And></Where><OrderBy Override='TRUE'><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query><ViewFields>$refs</ViewFields><RowLimit Paged='TRUE'>$PageSize</RowLimit></View>"
}

function Read-InventoryRange($Context,$Config,[long]$After,[long]$Through,[string[]]$Fields) {
    $query=[Microsoft.SharePoint.Client.CamlQuery]::new()
    $query.DatesInUtc=$true
    $query.ViewXml=New-InventoryQueryXml $After $Through $Fields $Config.PageSize
    $items=$Context.Web.Lists.GetById([guid]$Config.ListId).GetItems($query)
    $Context.Load($items); $Config.Gate.ExecuteRead($Context,$Config.RetryCount)
    return ,$items
}

function Get-InventoryPage($Context,$Config,[long]$Cursor) {
    # Unfiltered server paging avoids a broad filter over a >5000-item library.
    # ID ranges are used only for metadata on the already bounded discovered page.
    $query=[Microsoft.SharePoint.Client.CamlQuery]::new()
    $query.DatesInUtc=$true
    $refs=(@($Config.BaseFields | ForEach-Object {"<FieldRef Name='$([Security.SecurityElement]::Escape($_))'/>"}) -join '')
    $query.ViewXml="<View Scope='RecursiveAll'><Query><OrderBy Override='TRUE'><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query><ViewFields>$refs</ViewFields><RowLimit Paged='TRUE'>$($Config.PageSize)</RowLimit></View>"
    if($Config.ContainsKey('PagingInfo') -and $Config.PagingInfo){$position=[Microsoft.SharePoint.Client.ListItemCollectionPosition]::new();$position.PagingInfo=$Config.PagingInfo;$query.ListItemCollectionPosition=$position}
    $items=$Context.Web.Lists.GetById([guid]$Config.ListId).GetItems($query);$Context.Load($items);$Config.Gate.ExecuteRead($Context,$Config.RetryCount)
    $nextPosition=$null;if($items.ListItemCollectionPosition){$nextPosition=$items.ListItemCollectionPosition.PagingInfo}
    $rows=[Collections.Generic.List[object]]::new(); $last=$Cursor
    $beyondWatermark=$false
    foreach($item in $items) {
        if($item.Id -gt $Config.HighWatermark){$beyondWatermark=$true;break}
        if($item.Id -le $last) { throw 'Non-monotonic ID received; enumeration stopped to prevent skipped or duplicate items.' }
        $last=$item.Id; $values=@{}
        foreach($key in $Config.BaseFields) { if($item.FieldValues.ContainsKey($key)) { $values[$key]=Convert-FieldText $item.FieldValues[$key] } }
        $rows.Add([pscustomobject]@{Id=$item.Id;Values=$values})
    }
    if($nextPosition -and $Config.ContainsKey('PagingInfo') -and $nextPosition -eq $Config.PagingInfo){throw 'SharePoint repeated a paging token; enumeration stopped.'}
    return [pscustomobject]@{After=$Cursor;Through=$last;Items=$rows.ToArray();PagingInfo=$nextPosition;Done=(-not $nextPosition -or $beyondWatermark -or $last -ge $Config.HighWatermark)}
}

function Test-NetworkFailure($ErrorRecord) {
    for($e=$ErrorRecord.Exception;$null -ne $e;$e=$e.InnerException) {
        if($e -is [Net.WebException] -or $e -is [UnauthorizedAccessException]) { return $true }
    }
    return ($ErrorRecord.Exception.Message -match '(?i)unauthorized|access denied|401|403|429|503|authentication|token.*expir')
}

function ConvertTo-InventoryUtc($Value) {
    if($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)){return ''}
    if($Value -is [datetime]){
        # SharePoint list-item Modified is UTC. Do not let an Unspecified kind
        # acquire the workstation's timezone during conversion.
        $date=$Value;if($date.Kind -eq [DateTimeKind]::Unspecified){$date=[datetime]::SpecifyKind($date,[DateTimeKind]::Utc)}
        return $date.ToUniversalTime().ToString('o',[Globalization.CultureInfo]::InvariantCulture)
    }
    $parsed=[DateTimeOffset]::MinValue
    if(-not [DateTimeOffset]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$parsed)){throw "Cannot normalize SharePoint timestamp: $Value"}
    return $parsed.UtcDateTime.ToString('o',[Globalization.CultureInfo]::InvariantCulture)
}
function New-InventoryDocument($Item,$Config) {
    $v=$Item.Values;$path=$v['FileRef'];$relative=Get-LibraryRelativePath $path $Config.LibraryPath
    $slash=$relative.LastIndexOf('/');$depth=0;if($slash -ge 0){$depth=Get-RelativeFolderDepth $relative.Substring(0,$slash)}
    [pscustomobject][ordered]@{
        ItemId=$Item.Id;UniqueId=$v['UniqueId'];FileName=$v['FileLeafRef'];Url=$Config.Origin+(($path.Split('/')|ForEach-Object{[uri]::EscapeDataString($_)}) -join '/')
        LibraryRelativePath=$relative;FolderDepth=$depth;SizeBytes=$v['File_x0020_Size'];Modified=$v['Modified']
        Outcome='Pending';MetadataStatus='Complete';DownloadStatus='NotRequested';LocalPath='';ScanStatus='NotRequested';CorrelationStatus='NotRequested'
        VBA='NotRequested';VBAProjectCount=$null;VBASourceModules=$null;VBASourceModulesParsed=$null;Macro='NotRequested';Both='NotRequested'
        Format='';FormatDetails='';DetectionScope='';EmbeddedObjects=$null;FIBVersion=$null;EffectiveFIBVersion=$null;IsTemplate=$null;SignatureHex='';SHA256='';ScannerVersion='';ScanMilliseconds=0;ScanInputPath=''
        DiscoveryModifiedUtc='';DownloadModifiedUtc='';AfterScanModifiedUtc='';ETagBefore='';ETagAfter='';DownloadedBytes=$null
        ErrorCode='';ErrorStage='';Error='';Notes='';Diagnostics=[Collections.Generic.List[object]]::new()
    }
}
function ConvertTo-InventoryError($Record,[string]$Stage,[string]$Code,[string]$Severity='Error') {
    [pscustomobject][ordered]@{Severity=$Severity;Code=$Code;Stage=$Stage;Message=$Record.Exception.Message
        ExceptionType=$Record.Exception.GetType().FullName;ExceptionDetails=$Record.Exception.ToString()
        FullyQualifiedErrorId=$Record.FullyQualifiedErrorId;ScriptStackTrace=$Record.ScriptStackTrace}
}
function Add-InventoryIssue($Row,$Issue) {
    $Row.Diagnostics.Add($Issue)
    $Row.Error=(@($Row.Diagnostics|ForEach-Object{'['+$_.Stage+'] '+$_.Message}) -join ' | ')
    $Row.ErrorCode=(@($Row.Diagnostics|ForEach-Object{$_.Code}|Select-Object -Unique) -join '; ')
    $Row.ErrorStage=(@($Row.Diagnostics|ForEach-Object{$_.Stage}|Select-Object -Unique) -join '; ')
}
function Set-InventoryOutcome($Row) {
    if($Row.MetadataStatus -eq 'Failed' -or $Row.DownloadStatus -eq 'Failed'){$Row.Outcome='Failed'}
    elseif($Row.MetadataStatus -ne 'Complete' -or $Row.Diagnostics.Count -gt 0 -or $Row.ScanStatus -notin @('Scanned','NotRequested')){$Row.Outcome='NeedsReview'}
    else{$Row.Outcome='Succeeded'}
}
function Set-InventoryScanResult($Row,$Result,$Config) {
    $Row.ScanStatus=$Result.Status;$Row.Format=$Result.Format;$Row.SignatureHex=$Result.SignatureHex
    $Row.FormatDetails=$Result.FormatDetails;$Row.DetectionScope=$Result.DetectionScope;$Row.EmbeddedObjects=$Result.EmbeddedObjects;$Row.FIBVersion=$Result.FIBVersion
    $Row.EffectiveFIBVersion=$Result.EffectiveFIBVersion;$Row.IsTemplate=$Result.IsTemplate
    $Row.ScannerVersion=$Result.Scanner;$Row.ScanMilliseconds=$Result.ElapsedMilliseconds;$Row.ScanInputPath=$Result.FilePath
    $Row.VBAProjectCount=$Result.ProjectCount;$Row.VBASourceModulesParsed=$Result.SourceModulesParsed
    if($Result.Notes){$Row.Notes+=(' '+$Result.Notes)}
    if($Config.CheckVBA){$Row.VBA=$Result.VBA;$Row.VBASourceModules=$Result.SourceModules}
    if($Config.CheckMacro){$Row.Macro=$Result.Macro}
    if($Config.CheckBoth){$Row.Both=$Result.Macro}
    foreach($issue in $Result.Diagnostics){Add-InventoryIssue $Row $issue}
}
function Set-InventoryCorrelation($Row,$BeforeModified,$AfterModified,[string]$BeforeEtag,[string]$AfterEtag) {
    $Row.DiscoveryModifiedUtc=ConvertTo-InventoryUtc $Row.Modified
    $Row.DownloadModifiedUtc=ConvertTo-InventoryUtc $BeforeModified
    $Row.AfterScanModifiedUtc=ConvertTo-InventoryUtc $AfterModified
    $Row.ETagBefore=$BeforeEtag;$Row.ETagAfter=$AfterEtag
    if(-not $BeforeEtag -or -not $AfterEtag){$Row.CorrelationStatus='Unverified';$code='MissingVersionEvidence';$message='A server ETag was not returned; the downloaded snapshot could not be correlated with the later server version.'}
    elseif($BeforeEtag -cne $AfterEtag -or $Row.DownloadModifiedUtc -ne $Row.AfterScanModifiedUtc){$Row.CorrelationStatus='ServerChangedDuringScan';$code='ServerVersionChanged';$message='Server version changed between download and post-scan verification. The VBA/format findings describe the retained local snapshot and are preserved; rerun to assess the current server version.'}
    elseif($Row.DiscoveryModifiedUtc -and $Row.DownloadModifiedUtc -and $Row.DiscoveryModifiedUtc -ne $Row.DownloadModifiedUtc){$Row.CorrelationStatus='MetadataChangedSinceDiscovery';$code='MetadataVersionChanged';$message='List-item Modified changed between discovery and download. The downloaded snapshot was stable through post-scan verification. Review metadata correlation; local content findings remain valid for the recorded SHA256.'}
    else{$Row.CorrelationStatus='Verified';return}
    Add-InventoryIssue $Row ([pscustomobject]@{Severity='Warning';Code=$code;Stage='VersionCorrelation';Message=$message;ExceptionType='';ExceptionDetails='';FullyQualifiedErrorId='';ScriptStackTrace='';DiscoveryModifiedUtc=$Row.DiscoveryModifiedUtc;DownloadModifiedUtc=$Row.DownloadModifiedUtc;AfterScanModifiedUtc=$Row.AfterScanModifiedUtc;ETagBefore=$BeforeEtag;ETagAfter=$AfterEtag})
}
function Invoke-InventoryScan([string]$Path,$Config) {
    $resultPath=Join-Path $Config.ScannerTemp ([guid]::NewGuid().ToString('N')+'.json')
    $psi=[Diagnostics.ProcessStartInfo]::new();$psi.FileName=$Config.ScannerExecutable
    $psi.Arguments='-NoLogo -NoProfile -NonInteractive -File "{0}" -Path "{1}" -ResultPath "{2}" -MaxFileMB {3} -TimeoutSeconds {4}' -f $Config.ScannerHelper,$Path,$resultPath,([long]($Config.MaxBytes/1MB)),$Config.ScanTimeoutSeconds
    $psi.UseShellExecute=$false;$psi.CreateNoWindow=$true;$psi.RedirectStandardOutput=$true;$psi.RedirectStandardError=$true
    $process=$null
    try {
        $process=[Diagnostics.Process]::Start($psi)
        $stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
        if(-not $process.WaitForExit($Config.ScanTimeoutSeconds*1000)){ $process.Kill();$process.WaitForExit();throw ('Scanner timed out after {0}s. Input={1}; helper={2}; process ID={3}.' -f $Config.ScanTimeoutSeconds,$Path,$Config.ScannerHelper,$process.Id) }
        if($process.ExitCode -ne 0 -or -not [IO.File]::Exists($resultPath)){
            throw ('Scanner process failed. ExitCode={0}; input={1}; helper={2}; stderr={3}; stdout={4}' -f $process.ExitCode,$Path,$Config.ScannerHelper,$stderr.GetAwaiter().GetResult(),$stdout.GetAwaiter().GetResult())
        }
        $result=[IO.File]::ReadAllText($resultPath) | ConvertFrom-Json
        if(-not $result.PSObject.Properties['SchemaVersion'] -or $result.SchemaVersion -ne '2.0'){throw 'Scanner result schema mismatch; use the matching v3.10 helper files.'}
        if([SPOInventoryV310.LocalIO]::FullPath($result.FilePath) -ine [SPOInventoryV310.LocalIO]::FullPath($Path)){throw 'Scanner result input path does not match the requested document.'}
        return $result
    } finally {if($process){$process.Dispose()};if([IO.File]::Exists($resultPath)){[IO.File]::Delete($resultPath)}}
}
function Invoke-InventoryDownload($Context,$Config,$Row) {
    $local=$null;$partial=$null;$stage='PrepareDownload'
    try {
        if([long]$Row.SizeBytes -gt $Config.MaxBytes){throw 'File exceeds MaxScanMB; content not downloaded.'}
        $safeId=([string]$Row.UniqueId -replace '[^0-9A-Fa-f-]','');if(-not $safeId){$safeId='no-guid'}
        $layout='Flat';if($Config.ContainsKey('DownloadLayout')){$layout=$Config.DownloadLayout}
        if($layout -eq 'Mirror'){$local=[SPOInventoryV310.LocalIO]::Mirror($Config.DownloadRoot,$Row.LibraryRelativePath)}
        else{$local=[IO.Path]::Combine($Config.DownloadRoot,('{0}_{1}.doc' -f [long]$Row.ItemId,$safeId))}
        $Row.Notes='Local snapshot uses '+$layout+' layout; original SharePoint name/path and local mapping are recorded.'
        if([SPOInventoryV310.LocalIO]::FileExists($local)){throw 'Download target already exists; refusing overwrite.'}
        $directory=[SPOInventoryV310.LocalIO]::Parent($local);[SPOInventoryV310.LocalIO]::CreateDirectory($directory)
        $partial=$directory+'\.inventory_'+[guid]::NewGuid().ToString('N')+'.partial'
        $stage='Download'
        for($attempt=0;;$attempt++) {
            $stream=$null
            try {
                $file=$Context.Web.GetFileById([guid]$Row.UniqueId);$Context.Load($file)
                # File.ListItemAllFields can return site-local wall time, whereas
                # discovery CAML returns UTC. Use the identical UTC-enabled query
                # for every Modified comparison, loaded alongside the File.
                $snapshotQuery=[Microsoft.SharePoint.Client.CamlQuery]::new();$snapshotQuery.DatesInUtc=$true
                $snapshotQuery.ViewXml=New-InventoryQueryXml ([long]$Row.ItemId-1) ([long]$Row.ItemId) @('ID','Modified') 1
                $snapshotItems=$Context.Web.Lists.GetById([guid]$Config.ListId).GetItems($snapshotQuery);$Context.Load($snapshotItems)
                $Config.Gate.ExecuteRead($Context,$Config.RetryCount)
                if($snapshotItems.Count -ne 1){throw 'The document list item was not returned for the pre-download version snapshot.'}
                if($file.Length -gt $Config.MaxBytes){throw 'Current server file exceeds MaxScanMB.'}
                $beforeEtag=$file.ETag;$beforeModified=$snapshotItems[0].FieldValues['Modified'];$expectedLength=$file.Length
                $binary=$file.OpenBinaryStream();$Config.Gate.ExecuteRead($Context,$Config.RetryCount);$stream=$binary.Value
                if($null -eq $stream){throw 'No download stream returned.'}
                [SPOInventoryV310.LocalIO]::CopyLimited($stream,$partial,$Config.MaxBytes)
                break
            } catch {
                [SPOInventoryV310.LocalIO]::Delete($partial)
                if($attempt -ge $Config.RetryCount -or -not $Config.Gate.Backoff($_.Exception,$attempt)){throw}
                $Config.Gate.Wait()
            } finally {if($stream){$stream.Dispose()}}
        }
        [SPOInventoryV310.LocalIO]::Move($partial,$local)
        $Row.DownloadStatus='Downloaded';$Row.LocalPath=$local;$Row.DownloadedBytes=[SPOInventoryV310.LocalIO]::Length($local)
        $stage='HashSnapshot';$Row.SHA256=[SPOInventoryV310.LocalIO]::Hash($local)
        if($Row.DownloadedBytes -ne $expectedLength){Add-InventoryIssue $Row ([pscustomobject]@{Severity='Warning';Code='DownloadSizeChanged';Stage='Download';Message="Expected $expectedLength bytes; received $($Row.DownloadedBytes). Server content may have changed during download."})}
        if($Config.Scan){
            $stage='LaunchScanner'
            try {$result=Invoke-InventoryScan $local $Config;Set-InventoryScanResult $Row $result $Config}
            catch {$Row.ScanStatus='Unknown';if($Config.CheckVBA){$Row.VBA='Unknown'};if($Config.CheckMacro){$Row.Macro='Unknown'};if($Config.CheckBoth){$Row.Both='Unknown'};Add-InventoryIssue $Row (ConvertTo-InventoryError $_ $stage 'ScannerProcessFailure')}
        }
        $stage='VersionCorrelation'
        $after=$Context.Web.GetFileById([guid]$Row.UniqueId);$Context.Load($after)
        $snapshotQuery=[Microsoft.SharePoint.Client.CamlQuery]::new();$snapshotQuery.DatesInUtc=$true
        $snapshotQuery.ViewXml=New-InventoryQueryXml ([long]$Row.ItemId-1) ([long]$Row.ItemId) @('ID','Modified') 1
        $snapshotItems=$Context.Web.Lists.GetById([guid]$Config.ListId).GetItems($snapshotQuery);$Context.Load($snapshotItems)
        $Config.Gate.ExecuteRead($Context,$Config.RetryCount)
        if($snapshotItems.Count -ne 1){throw 'The document list item was not returned for post-scan version verification.'}
        Set-InventoryCorrelation $Row $beforeModified $snapshotItems[0].FieldValues['Modified'] $beforeEtag $after.ETag
    } catch {
        if($Row.DownloadStatus -ne 'Downloaded'){$Row.DownloadStatus='Failed'}
        if($stage -eq 'VersionCorrelation'){$Row.CorrelationStatus='Unverified'}
        Add-InventoryIssue $Row (ConvertTo-InventoryError $_ $stage 'DocumentProcessingError')
        if($Config.Scan -and $Row.ScanStatus -eq 'NotRequested'){$Row.ScanStatus='Unknown';if($Config.CheckVBA){$Row.VBA='Unknown'};if($Config.CheckMacro){$Row.Macro='Unknown'};if($Config.CheckBoth){$Row.Both='Unknown'}}
    } finally {
        if($partial){try{[SPOInventoryV310.LocalIO]::Delete($partial)}catch{Add-InventoryIssue $Row (ConvertTo-InventoryError $_ 'CleanupDownload' 'PartialFileCleanupFailed')}}
    }
}
function Invoke-InventoryBatch($Context,$Config,$Page) {
    $documents=[Collections.Generic.List[object]]::new();$byId=@{};$maps=@{};$statuses=@{}
    foreach($item in $Page.Items) {
        $v=$item.Values;$name=$v['FileLeafRef'];$path=$v['FileRef']
        if([int]$v['FSObjType'] -ne 0 -or [IO.Path]::GetExtension($name) -ine '.doc' -or -not $path.StartsWith($Config.FolderPath+'/',[StringComparison]::OrdinalIgnoreCase)){continue}
        $relative=Get-LibraryRelativePath $path $Config.LibraryPath
        $row=New-InventoryDocument $item $Config
        $documents.Add($row);$byId[$item.Id]=$row;$maps[$item.Id]=@{};$statuses[$item.Id]=@{}
        foreach($field in $Config.Fields){if($v.ContainsKey($field.InternalName)){$maps[$item.Id][$field.InternalName]=$v[$field.InternalName];$statuses[$item.Id][$field.InternalName]='Read'}}
    }
    if(-not $documents.Count){return [pscustomobject]@{After=$Page.After;Through=$Page.Through;Documents=@();Metadata=@();Errors=0}}
    $extra=@($Config.Fields | Where-Object {$_.InternalName -notin $Config.BaseFields})
    $networkFailure=$null;$networkDiagnostic=$null;$fieldDiagnostics=@{}
    for($offset=0;$offset -lt $extra.Count;$offset+=$Config.FieldBatchSize) {
        $group=@($extra[$offset..([math]::Min($offset+$Config.FieldBatchSize-1,$extra.Count-1))]);$failure=$networkFailure
        if(-not $failure) {
            try {
                $items=Read-InventoryRange $Context $Config $Page.After $Page.Through (@('ID')+@($group.InternalName))
                foreach($item in $items){if($byId.ContainsKey($item.Id)){foreach($field in $group){if($item.FieldValues.ContainsKey($field.InternalName)){$maps[$item.Id][$field.InternalName]=Convert-FieldText $item.FieldValues[$field.InternalName];$statuses[$item.Id][$field.InternalName]='Read'}}}}
            } catch {$failure=$_.Exception.Message;if(Test-NetworkFailure $_){$networkFailure=$failure;$networkDiagnostic=ConvertTo-InventoryError $_ 'ReadMetadata' 'MetadataRequestFailed'}}
        }
        if($failure) {
            foreach($field in $group) {
                $fieldError=$networkFailure;$fieldDiagnostic=$networkDiagnostic
                if(-not $fieldError) {
                    try {
                        $items=Read-InventoryRange $Context $Config $Page.After $Page.Through @('ID',$field.InternalName)
                        foreach($item in $items){if($byId.ContainsKey($item.Id) -and $item.FieldValues.ContainsKey($field.InternalName)){$maps[$item.Id][$field.InternalName]=Convert-FieldText $item.FieldValues[$field.InternalName];$statuses[$item.Id][$field.InternalName]='Read'}}
                    } catch {$fieldError=$_.Exception.Message;$fieldDiagnostic=ConvertTo-InventoryError $_ 'ReadMetadata' 'MetadataRequestFailed';if(Test-NetworkFailure $_){$networkFailure=$fieldError;$networkDiagnostic=$fieldDiagnostic}}
                }
                if($fieldError){$fieldDiagnostics[$field.InternalName]=$fieldDiagnostic;foreach($row in $documents){$statuses[$row.ItemId][$field.InternalName]='Error: '+$fieldError}}
            }
        }
    }
    $metadata=[Collections.Generic.List[object]]::new();$errors=0
    foreach($row in $documents) {
        foreach($field in $Config.Fields){
            $status='NotReturned';$value='';if($statuses[$row.ItemId].ContainsKey($field.InternalName)){$status=$statuses[$row.ItemId][$field.InternalName]};if($maps[$row.ItemId].ContainsKey($field.InternalName)){$value=$maps[$row.ItemId][$field.InternalName]}
            if($status -ne 'Read'){
                $row.MetadataStatus='Partial'
                $details=[ordered]@{Severity='Error';Code='MetadataReadIncomplete';Stage='ReadMetadata';Message=($field.InternalName+': '+$status);Field=$field.InternalName}
                if($fieldDiagnostics.ContainsKey($field.InternalName) -and $fieldDiagnostics[$field.InternalName]){foreach($p in $fieldDiagnostics[$field.InternalName].PSObject.Properties){if($p.Name -ne 'Message'){$details[$p.Name]=$p.Value}}}
                Add-InventoryIssue $row ([pscustomobject]$details)
            }
            $metadata.Add([pscustomobject]@{ItemId=$row.ItemId;FileName=$row.FileName;Field=$field.InternalName;DisplayName=$field.DisplayName;Type=$field.Type;Value=$value;Status=$status})
        }
        if($Config.Download){Invoke-InventoryDownload $Context $Config $row}
        Set-InventoryOutcome $row;if($row.Outcome -ne 'Succeeded'){$errors++}
    }
    return [pscustomobject]@{After=$Page.After;Through=$Page.Through;Documents=$documents.ToArray();Metadata=$metadata.ToArray();Errors=$errors}
}

# Bounded producer/consumer: one asynchronous discovery request, at most N batch
# jobs. Completed results are delivered before allocating more work.
function Invoke-BoundedInventory($Config,[scriptblock]$ContextFactory,[scriptblock]$OnPage,[scriptblock]$OnResult,[scriptblock]$OnProgress,
    [scriptblock]$PageReader={param($Context,$Config,$InputPage) Get-InventoryPage $Context $Config $InputPage},
    [scriptblock]$BatchReader={param($Context,$Config,$InputPage) Invoke-InventoryBatch $Context $Config $InputPage}) {
    $initial=[Management.Automation.Runspaces.InitialSessionState]::CreateDefault();[void]$initial.StartupScripts.Add($Config.EnginePath)
    # An isolated default host keeps worker output out of the interactive host's
    # formatter. Only the coordinator publishes progress and final results.
    $pool=[runspacefactory]::CreateRunspacePool($initial);[void]$pool.SetMaxRunspaces($Config.ThrottleLimit+1);$pool.Open()
    $active=[Collections.Generic.List[object]]::new();$enumerator=$null;$cursor=0L;$done=$false;$failure=$null;$lastProgress=[datetime]::MinValue
    $wrapper={param($Worker,$Context,$Config,$InputPage) $ErrorActionPreference='Stop';Set-StrictMode -Version 2;try {& ([scriptblock]::Create($Worker)) $Context $Config $InputPage} finally {if($Context){$Context.Dispose()}}}
    try {
        while(-not $done -or $enumerator -or $active.Count) {
            for($i=$active.Count-1;$i -ge 0;$i--){
                $job=$active[$i];if(-not $job.Handle.IsCompleted){continue}
                $jobOutput=$null;$jobFailure=$null
                try {$jobOutput=@($job.PowerShell.EndInvoke($job.Handle));if($job.PowerShell.HadErrors -or $jobOutput.Count -ne 1){throw ($job.PowerShell.Streams.Error | Out-String)}}
                catch {$jobFailure=$_}
                finally {$job.PowerShell.Dispose();$active.RemoveAt($i)}
                # Report-writer failures must abort; never replay a partially written batch.
                if($jobFailure){[void](& $OnResult $null $jobFailure $job.Page)}else{[void](& $OnResult $jobOutput[0] $null $job.Page)}
            }
            if($enumerator -and $enumerator.Handle.IsCompleted){
                $page=$null
                try {$value=@($enumerator.PowerShell.EndInvoke($enumerator.Handle));if($enumerator.PowerShell.HadErrors -or $value.Count -ne 1){throw ($enumerator.PowerShell.Streams.Error | Out-String)};$page=$value[0];$cursor=$page.Through;$done=$page.Done;if($page.PSObject.Properties['PagingInfo']){$Config.PagingInfo=$page.PagingInfo};[void](& $OnPage $page)}
                catch {$failure=$_;$done=$true;$page=$null}
                finally {$enumerator.PowerShell.Dispose();$enumerator=$null}
                if($page -and -not $Config.CountOnly -and $page.Items.Count){
                    $context=& $ContextFactory
                    $ps=[powershell]::Create();$ps.RunspacePool=$pool;[void]$ps.AddScript($wrapper.ToString()).AddArgument($BatchReader.ToString()).AddArgument($context).AddArgument($Config).AddArgument($page)
                    $active.Add([pscustomobject]@{PowerShell=$ps;Handle=$ps.BeginInvoke();Page=$page})
                }
            }
            if(-not $done -and -not $enumerator -and $active.Count -lt $Config.ThrottleLimit){
                $context=& $ContextFactory
                $ps=[powershell]::Create();$ps.RunspacePool=$pool;[void]$ps.AddScript($wrapper.ToString()).AddArgument($PageReader.ToString()).AddArgument($context).AddArgument($Config).AddArgument($cursor)
                $enumerator=[pscustomobject]@{PowerShell=$ps;Handle=$ps.BeginInvoke()}
            }
            if(([datetime]::UtcNow-$lastProgress).TotalSeconds -ge $Config.ProgressSeconds){& $OnProgress $cursor $active.Count ([bool]$enumerator);$lastProgress=[datetime]::UtcNow}
            if($active.Count -or $enumerator){Start-Sleep -Milliseconds 50}
        }
        & $OnProgress $cursor 0 $false
        if($failure){throw $failure}
    } finally {
        if($enumerator){$enumerator.PowerShell.Stop();$enumerator.PowerShell.Dispose()}
        foreach($job in $active){$job.PowerShell.Stop();$job.PowerShell.Dispose()}
        $pool.Close();$pool.Dispose()
    }
}

function Import-RequiredModule([string]$Name,[string]$Version) {
    $loaded=Get-Module $Name
    if ($loaded -and $loaded.Version -ne [version]$Version) { throw "Start a fresh PowerShell window: a different $Name version is loaded." }
    if (-not (Get-Module -ListAvailable $Name | Where-Object Version -eq ([version]$Version))) {
        if ($NoInstall) { throw "Missing $Name $Version; installation disabled." }
        Write-RunLog "Installing $Name $Version from PSGallery to local .dependencies."
        [Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        if (-not (Get-PackageProvider NuGet -ListAvailable -ErrorAction SilentlyContinue)) { Install-PackageProvider NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null }
        Save-Module -Name $Name -RequiredVersion $Version -Repository PSGallery -Path $dependencyDir -Force -ErrorAction Stop
    }
    Import-Module $Name -RequiredVersion $Version -ErrorAction Stop -Verbose:$false
}
function Initialize-Scanner {
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'Inspect-LegacyWord.ps1'))) { throw 'Inspect-LegacyWord.ps1 must be beside this script.' }
    $script:scannerExecutable=Join-Path $PSHOME 'powershell.exe'
    if(-not [IO.File]::Exists($script:scannerExecutable)){throw 'The scanner requires Windows PowerShell 5.1.'}
    Write-RunLog 'Content scanner: native PowerShell CFB/OVBA parser; no Python or package installation.'
}
