# Shared pure helpers and worker functions for Export-SPOLegacyDocInventory v3.1.
function Convert-FieldText($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('o') }
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
    $query.ViewXml=New-InventoryQueryXml $After $Through $Fields $Config.PageSize
    $items=$Context.Web.Lists.GetById([guid]$Config.ListId).GetItems($query)
    $Context.Load($items); $Config.Gate.ExecuteRead($Context,$Config.RetryCount)
    return ,$items
}

function Get-InventoryPage($Context,$Config,[long]$Cursor) {
    # Unfiltered server paging avoids a broad filter over a >5000-item library.
    # ID ranges are used only for metadata on the already bounded discovered page.
    $query=[Microsoft.SharePoint.Client.CamlQuery]::new()
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

function Invoke-InventoryScan([string]$Path,$Config) {
    $resultPath=Join-Path $Config.ScannerTemp ([guid]::NewGuid().ToString('N')+'.json')
    $psi=[Diagnostics.ProcessStartInfo]::new();$psi.FileName=$Config.ScannerExecutable
    $psi.Arguments='-NoLogo -NoProfile -NonInteractive -File "{0}" -Path "{1}" -ResultPath "{2}" -MaxFileMB {3} -TimeoutSeconds {4}' -f $Config.ScannerHelper,$Path,$resultPath,([long]($Config.MaxBytes/1MB)),$Config.ScanTimeoutSeconds
    $psi.UseShellExecute=$false;$psi.CreateNoWindow=$true;$psi.RedirectStandardOutput=$true;$psi.RedirectStandardError=$true
    $process=[Diagnostics.Process]::Start($psi)
    $stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
    try {
        if(-not $process.WaitForExit($Config.ScanTimeoutSeconds*1000)){ $process.Kill();$process.WaitForExit();throw 'Scanner timed out; result Unknown.' }
        if($process.ExitCode -ne 0 -or -not [IO.File]::Exists($resultPath)){
            $detail=$stderr.GetAwaiter().GetResult();if($detail.Length -gt 600){$detail=$detail.Substring(0,600)}
            throw "PowerShell scanner did not produce a result; result Unknown. $detail"
        }
        return ([IO.File]::ReadAllText($resultPath) | ConvertFrom-Json)
    } finally {$process.Dispose();if([IO.File]::Exists($resultPath)){[IO.File]::Delete($resultPath)}}
}

function Invoke-InventoryDownload($Context,$Config,$Row) {
    $local=$null;$partial=$null
    try {
        if([long]$Row.SizeBytes -gt $Config.MaxBytes){throw 'File exceeds MaxScanMB; content not downloaded.'}
        $local=[SPOInventoryV3.LocalIO]::Mirror($Config.DownloadRoot,$Row.LibraryRelativePath)
        $native=[SPOInventoryV3.LocalIO]::Extended($local)
        if([IO.File]::Exists($native)){throw 'Download target already exists; refusing overwrite.'}
        $directory=$local.Substring(0,$local.LastIndexOf('\'))
        [void][IO.Directory]::CreateDirectory([SPOInventoryV3.LocalIO]::Extended($directory))
        $partial=$directory+'\.inventory_'+[guid]::NewGuid().ToString('N')+'.partial'
        for($attempt=0;;$attempt++) {
            $stream=$null
            try {
                $file=$Context.Web.GetFileById([guid]$Row.UniqueId);$Context.Load($file)
                $fileItem=$file.ListItemAllFields;$Context.Load($fileItem);$Config.Gate.ExecuteRead($Context,$Config.RetryCount)
                if($file.Length -gt $Config.MaxBytes){throw 'Current server file exceeds MaxScanMB.'}
                $binary=$file.OpenBinaryStream();$Config.Gate.ExecuteRead($Context,$Config.RetryCount);$stream=$binary.Value
                if($null -eq $stream){throw 'No download stream returned.'}
                # Compare the same list-item field used in enumeration; File.TimeLastModified
                # can differ from list-item Modified even for an unchanged file.
                $beforeEtag=$file.ETag;$beforeModified=Convert-FieldText $fileItem.FieldValues['Modified']
                [SPOInventoryV3.LocalIO]::CopyLimited($stream,$partial,$Config.MaxBytes)
                break
            } catch {
                if([IO.File]::Exists([SPOInventoryV3.LocalIO]::Extended($partial))){[IO.File]::Delete([SPOInventoryV3.LocalIO]::Extended($partial))}
                if($attempt -ge $Config.RetryCount -or -not $Config.Gate.Backoff($_.Exception,$attempt)){throw}
                $Config.Gate.Wait()
            } finally {if($stream){$stream.Dispose()}}
        }
        [IO.File]::Move([SPOInventoryV3.LocalIO]::Extended($partial),$native)
        $Row.DownloadStatus='Downloaded';$Row.LocalPath=$local;$Row.SHA256=[SPOInventoryV3.LocalIO]::Hash($local)
        if($Config.Scan) {
            $result=Invoke-InventoryScan $local $Config
            $Row.ScanStatus=$result.Status;$Row.Format=$result.Format;$Row.Notes=$result.Notes;$Row.Error=$result.Error
            if($Config.CheckVBA){$Row.VBA=$result.VBA;$Row.VBASourceModules=$result.SourceModules}
            if($Config.CheckMacro){$Row.Macro=$result.Macro}
            if($Config.CheckBoth){$Row.Both=$result.Macro}
        }
        $after=$Context.Web.GetFileById([guid]$Row.UniqueId);$Context.Load($after);$Config.Gate.ExecuteRead($Context,$Config.RetryCount)
        if($beforeEtag -ne $after.ETag -or $beforeModified -ne $Row.Modified){
            $Row.ScanStatus='ChangedDuringScan';$Row.Notes+=' Content/metadata changed during enumeration or scan; rerun to correlate.'
            if($Config.Scan){$Row.VBA='Unknown';$Row.Macro='Unknown';$Row.Both='Unknown'}
        }
    } catch {
        if($Row.DownloadStatus -ne 'Downloaded'){$Row.DownloadStatus='Failed'}
        $Row.Error=$_.Exception.Message
        if($Config.Scan){$Row.ScanStatus='Unknown';if($Config.CheckVBA){$Row.VBA='Unknown'};if($Config.CheckMacro){$Row.Macro='Unknown'};if($Config.CheckBoth){$Row.Both='Unknown'}}
        if($partial -and [IO.File]::Exists([SPOInventoryV3.LocalIO]::Extended($partial))){[IO.File]::Delete([SPOInventoryV3.LocalIO]::Extended($partial))}
    }
}

function Invoke-InventoryBatch($Context,$Config,$Page) {
    $documents=[Collections.Generic.List[object]]::new();$byId=@{};$maps=@{};$statuses=@{}
    foreach($item in $Page.Items) {
        $v=$item.Values;$name=$v['FileLeafRef'];$path=$v['FileRef']
        if([int]$v['FSObjType'] -ne 0 -or [IO.Path]::GetExtension($name) -ine '.doc' -or -not $path.StartsWith($Config.FolderPath+'/',[StringComparison]::OrdinalIgnoreCase)){continue}
        $relative=Get-LibraryRelativePath $path $Config.LibraryPath
        $row=[pscustomobject][ordered]@{ItemId=$item.Id;UniqueId=$v['UniqueId'];FileName=$name;Url=$Config.Origin+(($path.Split('/') | ForEach-Object {[uri]::EscapeDataString($_)}) -join '/');LibraryRelativePath=$relative;FolderDepth=0;SizeBytes=$v['File_x0020_Size'];Modified=$v['Modified'];MetadataStatus='Complete';DownloadStatus='NotRequested';LocalPath='';ScanStatus='NotRequested';VBA='NotRequested';VBASourceModules='';Macro='NotRequested';Both='NotRequested';Format='';SHA256='';Notes='';Error=''}
        $slash=$relative.LastIndexOf('/');$row.FolderDepth=0;if($slash -ge 0){$row.FolderDepth=Get-RelativeFolderDepth $relative.Substring(0,$slash)}
        $documents.Add($row);$byId[$item.Id]=$row;$maps[$item.Id]=@{};$statuses[$item.Id]=@{}
        foreach($field in $Config.Fields){if($v.ContainsKey($field.InternalName)){$maps[$item.Id][$field.InternalName]=$v[$field.InternalName];$statuses[$item.Id][$field.InternalName]='Read'}}
    }
    if(-not $documents.Count){return [pscustomobject]@{After=$Page.After;Through=$Page.Through;Documents=@();Metadata=@();Errors=0}}
    $extra=@($Config.Fields | Where-Object {$_.InternalName -notin $Config.BaseFields})
    $networkFailure=$null
    for($offset=0;$offset -lt $extra.Count;$offset+=$Config.FieldBatchSize) {
        $group=@($extra[$offset..([math]::Min($offset+$Config.FieldBatchSize-1,$extra.Count-1))]);$failure=$networkFailure
        if(-not $failure) {
            try {
                $items=Read-InventoryRange $Context $Config $Page.After $Page.Through (@('ID')+@($group.InternalName))
                foreach($item in $items){if($byId.ContainsKey($item.Id)){foreach($field in $group){if($item.FieldValues.ContainsKey($field.InternalName)){$maps[$item.Id][$field.InternalName]=Convert-FieldText $item.FieldValues[$field.InternalName];$statuses[$item.Id][$field.InternalName]='Read'}}}}
            } catch {$failure=$_.Exception.Message;if(Test-NetworkFailure $_){$networkFailure=$failure}}
        }
        if($failure) {
            foreach($field in $group) {
                $fieldError=$networkFailure
                if(-not $fieldError) {
                    try {
                        $items=Read-InventoryRange $Context $Config $Page.After $Page.Through @('ID',$field.InternalName)
                        foreach($item in $items){if($byId.ContainsKey($item.Id) -and $item.FieldValues.ContainsKey($field.InternalName)){$maps[$item.Id][$field.InternalName]=Convert-FieldText $item.FieldValues[$field.InternalName];$statuses[$item.Id][$field.InternalName]='Read'}}
                    } catch {$fieldError=$_.Exception.Message;if(Test-NetworkFailure $_){$networkFailure=$fieldError}}
                }
                if($fieldError){foreach($row in $documents){$statuses[$row.ItemId][$field.InternalName]='Error: '+$fieldError}}
            }
        }
    }
    $metadata=[Collections.Generic.List[object]]::new();$errors=0
    foreach($row in $documents) {
        foreach($field in $Config.Fields){
            $status='NotReturned';$value='';if($statuses[$row.ItemId].ContainsKey($field.InternalName)){$status=$statuses[$row.ItemId][$field.InternalName]};if($maps[$row.ItemId].ContainsKey($field.InternalName)){$value=$maps[$row.ItemId][$field.InternalName]}
            if($status -ne 'Read'){$row.MetadataStatus='Partial'}
            $metadata.Add([pscustomobject]@{ItemId=$row.ItemId;FileName=$row.FileName;Field=$field.InternalName;DisplayName=$field.DisplayName;Type=$field.Type;Value=$value;Status=$status})
        }
        if($Config.Download){Invoke-InventoryDownload $Context $Config $row}
        if($row.MetadataStatus -ne 'Complete' -or $row.Error -or $row.ScanStatus -notin @('NotRequested','Scanned')){$errors++}
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
    Import-Module $Name -RequiredVersion $Version -ErrorAction Stop
}
function Initialize-Scanner {
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'Inspect-LegacyWord.ps1'))) { throw 'Inspect-LegacyWord.ps1 must be beside this script.' }
    $script:scannerExecutable=Join-Path $PSHOME 'powershell.exe'
    if(-not [IO.File]::Exists($script:scannerExecutable)){throw 'The scanner requires Windows PowerShell 5.1.'}
    Write-RunLog 'Content scanner: native PowerShell CFB/OVBA parser; no Python or package installation.'
}
