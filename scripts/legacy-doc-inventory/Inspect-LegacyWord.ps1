#Requires -Version 5.1
<#
.SYNOPSIS
Inspect a local legacy Word binary using PowerShell and built-in .NET only.
.DESCRIPTION
Read-only CFB/OVBA parser. No Python, Word, COM automation, Add-Type, network,
third-party modules, source export or macro execution. This helper replaces
Inspect-LegacyWord.py for inventory v3.1. It is not an antivirus or Word validator.
Supported scope: Word 97+ OLE binaries and VBA project storages in the outer
compound hierarchy. WordBasic, packed embedded files, external templates and
compiled-only/stomped VBA are not exhaustively assessed. Unknown never means clean.
Resource/cycle/bounds checks reject corrupt or unusually large parser structures.
The inventory also enforces a hard child-process timeout for each inspection.
.EXAMPLE
.\Inspect-LegacyWord.ps1 -Path 'D:\Exports\Example.doc'
.EXAMPLE
.\Inspect-LegacyWord.ps1 -Path 'D:\Exports\Example.doc' -ResultPath 'D:\Reports\Example.json'
.NOTES
Implementation references: Microsoft MS-CFB section 2; MS-DOC FibBase;
MS-OVBA sections 2.3.4 and 2.4.1. No executable content is evaluated.
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/05060311-bfce-4b12-874d-71fd4ce63aea
https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc
#>
[CmdletBinding()]
param([string]$Path,[string]$ResultPath,[ValidateRange(1,2048)][int]$MaxFileMB=100,
      [ValidateRange(1,1800)][int]$TimeoutSeconds=120)

function ConvertTo-LwNativePath([string]$Value) {
    if($Value.StartsWith('\\?\')){return $Value}
    $full=[IO.Path]::GetFullPath($Value)
    if($full.StartsWith('\\')){return '\\?\UNC\'+$full.Substring(2)}
    return '\\?\'+$full
}
function Assert-LwTime($State) {
    if($State.Clock.Elapsed.TotalSeconds -gt $State.Timeout){throw 'Parser time limit exceeded.'}
}
function Read-LwAt($State,[long]$Offset,[int]$Count) {
    Assert-LwTime $State
    if($Offset -lt 0 -or $Count -lt 0 -or $Offset -gt $State.Stream.Length-$Count){throw 'Truncated or out-of-range compound-file read.'}
    $bytes=[byte[]]::new($Count);[void]$State.Stream.Seek($Offset,[IO.SeekOrigin]::Begin);$done=0
    while($done -lt $Count){$n=$State.Stream.Read($bytes,$done,$Count-$done);if($n -le 0){throw 'Unexpected end of file.'};$done+=$n}
    return ,$bytes
}
function Get-LwChain($State,[uint32]$Start,[uint32[]]$Table,[long]$Limit) {
    $chain=[Collections.Generic.List[uint32]]::new();$seen=[Collections.Generic.HashSet[uint32]]::new();$id=$Start
    while($id -ne 4294967294){
        if($id -ge $Limit -or $id -ge $Table.Length -or -not $seen.Add($id)){throw 'Invalid, cross-repeated or cyclic sector chain.'}
        $chain.Add($id);$id=$Table[$id]
        if(($chain.Count -band 1023) -eq 0){Assert-LwTime $State}
    }
    return ,$chain.ToArray()
}
function Read-LwRegular($State,[uint32]$Start,[long]$Size,[int]$MaxBytes=33554432,[int]$Prefix=0) {
    $chain=Get-LwChain $State $Start $State.Fat $State.SectorCount
    $capacity=[long]$chain.Length*$State.SectorSize
    if($Size -lt 0){$Size=$capacity}
    if($Size -gt $capacity -or ($Size -gt 0 -and $capacity-$Size -ge $State.SectorSize)){throw 'Stream size does not match its sector chain.'}
    $readSize=$Size;if($Prefix -gt 0){$readSize=[math]::Min($Size,$Prefix)}
    if($readSize -gt $MaxBytes){throw 'Stream exceeds parser allocation limit.'}
    $bytes=[byte[]]::new([int]$readSize);$done=0
    foreach($sector in $chain){
        if($done -ge $readSize){break};$count=[int][math]::Min($State.SectorSize,$readSize-$done)
        $part=Read-LwAt $State (([long]$sector+1)*$State.SectorSize) $count
        [Buffer]::BlockCopy($part,0,$bytes,$done,$count);$done+=$count
    }
    return ,$bytes
}
function Read-LwStream($State,$Entry,[int]$MaxBytes=33554432,[int]$Prefix=0) {
    if($Entry.Type -ne 2){throw 'Expected a stream directory entry.'}
    if($Entry.Size -ge 4096){return ,(Read-LwRegular $State $Entry.Start $Entry.Size $MaxBytes $Prefix)}
    if($Entry.Size -eq 0){return ,([byte[]]::new(0))}
    $chain=Get-LwChain $State $Entry.Start $State.MiniFat ([math]::Floor($State.RootSize/64))
    if([long]$chain.Length*64 -lt $Entry.Size -or [long]$chain.Length*64-$Entry.Size -ge 64){throw 'Mini-stream size does not match its chain.'}
    $readSize=[int]$Entry.Size;if($Prefix -gt 0){$readSize=[math]::Min($readSize,$Prefix)}
    if($readSize -gt $MaxBytes){throw 'Mini stream exceeds parser allocation limit.'}
    $bytes=[byte[]]::new($readSize);$done=0
    foreach($mini in $chain){
        if($done -ge $readSize){break};$offset=[long]$mini*64;$rootIndex=[int][math]::Floor($offset/$State.SectorSize)
        if($rootIndex -ge $State.RootChain.Length){throw 'Mini stream is outside the root stream.'}
        $count=[math]::Min(64,$readSize-$done)
        $part=Read-LwAt $State (([long]$State.RootChain[$rootIndex]+1)*$State.SectorSize+($offset%$State.SectorSize)) $count
        [Buffer]::BlockCopy($part,0,$bytes,$done,$count);$done+=$count
    }
    return ,$bytes
}
function Initialize-LwCompound($State) {
    $header=Read-LwAt $State 0 512
    if([BitConverter]::ToString($header,0,8) -ne 'D0-CF-11-E0-A1-B1-1A-E1'){throw 'Not a compound file.'}
    $version=[BitConverter]::ToUInt16($header,26);$shift=[BitConverter]::ToUInt16($header,30)
    if([BitConverter]::ToUInt16($header,28) -ne 65534 -or ($version -eq 3 -and $shift -ne 9) -or ($version -eq 4 -and $shift -ne 12) -or $version -notin @(3,4) -or [BitConverter]::ToUInt16($header,32) -ne 6 -or [BitConverter]::ToUInt32($header,56) -ne 4096){throw 'Unsupported or malformed CFB header.'}
    $State.SectorSize=1 -shl $shift;$State.SectorCount=[long][math]::Floor($State.Stream.Length/$State.SectorSize)-1
    if($State.SectorCount -lt 2 -or $State.Stream.Length%$State.SectorSize -ne 0){throw 'Truncated compound file.'}
    $fatCount=[BitConverter]::ToUInt32($header,44);$difCount=[BitConverter]::ToUInt32($header,72)
    if($fatCount -lt 1 -or $fatCount -gt [math]::Ceiling(($State.SectorCount+1)/($State.SectorSize/4))+1 -or $difCount -gt $State.SectorCount){throw 'Invalid FAT allocation counts.'}
    $fatIds=[Collections.Generic.List[uint32]]::new();$reserved=[Collections.Generic.HashSet[uint32]]::new()
    for($i=0;$i -lt 109;$i++){$id=[BitConverter]::ToUInt32($header,76+4*$i);if($id -ne 4294967295){$fatIds.Add($id)}}
    $dif=[BitConverter]::ToUInt32($header,68)
    for($i=0;$i -lt $difCount;$i++){
        if($dif -ge $State.SectorCount -or -not $reserved.Add($dif)){throw 'Invalid/cyclic DIFAT.'}
        $part=Read-LwAt $State (([long]$dif+1)*$State.SectorSize) $State.SectorSize
        for($j=0;$j -lt $State.SectorSize-4;$j+=4){$id=[BitConverter]::ToUInt32($part,$j);if($id -ne 4294967295){$fatIds.Add($id)}}
        $dif=[BitConverter]::ToUInt32($part,$State.SectorSize-4)
    }
    if($difCount -gt 0 -and $dif -ne 4294967294){throw 'DIFAT is not terminated.'}
    if($fatIds.Count -ne $fatCount){throw 'FAT sector count mismatch.'}
    $State.Fat=[uint32[]]::new([int]($fatCount*$State.SectorSize/4));$at=0
    foreach($id in $fatIds){
        if($id -ge $State.SectorCount -or -not $reserved.Add($id)){throw 'Invalid/duplicate FAT sector.'}
        $part=Read-LwAt $State (([long]$id+1)*$State.SectorSize) $State.SectorSize
        [Buffer]::BlockCopy($part,0,$State.Fat,$at,$part.Length);$at+=$part.Length
    }
    if($State.Fat.Length -lt $State.SectorCount){throw 'FAT does not cover the file.'}
    foreach($id in $fatIds){if($State.Fat[$id] -ne 4294967293){throw 'FAT sector marker invalid.'}}
    $miniCount=[BitConverter]::ToUInt32($header,64)
    if($miniCount -gt $State.SectorCount){throw 'Invalid mini FAT count.'}
    $miniBytes=Read-LwRegular $State ([BitConverter]::ToUInt32($header,60)) ([long]$miniCount*$State.SectorSize) 16777216
    $State.MiniFat=[uint32[]]::new($miniBytes.Length/4);[Buffer]::BlockCopy($miniBytes,0,$State.MiniFat,0,$miniBytes.Length)
    $directory=Read-LwRegular $State ([BitConverter]::ToUInt32($header,48)) -1 8388608
    if($directory.Length%128 -ne 0){throw 'Directory entry size mismatch.'}
    $entries=[Collections.Generic.List[object]]::new()
    for($at=0;$at -lt $directory.Length;$at+=128){
        $type=$directory[$at+66];$name='';$size=0L
        if($type -ne 0){
            $length=[BitConverter]::ToUInt16($directory,$at+64)
            if($type -notin @(1,2,5) -or $length -lt 2 -or $length -gt 64 -or $length%2 -ne 0 -or [BitConverter]::ToUInt16($directory,$at+$length-2) -ne 0){throw 'Malformed directory entry.'}
            $name=[Text.Encoding]::Unicode.GetString($directory,$at,$length-2)
            if($name.IndexOfAny([char[]]@([char]0,'/','\')) -ge 0){throw 'Invalid directory entry name.'}
            $wide=[BitConverter]::ToUInt64($directory,$at+120)
            if($type -ne 1 -and ($wide -gt $State.Stream.Length -or ($version -eq 3 -and $wide -gt 2147483647))){throw 'Invalid stream length.'}
            $size=[long]$wide
        }
        $entries.Add([pscustomobject]@{Id=($at/128);Name=$name;Type=$type;Left=[BitConverter]::ToUInt32($directory,$at+68);Right=[BitConverter]::ToUInt32($directory,$at+72);Child=[BitConverter]::ToUInt32($directory,$at+76);Start=[BitConverter]::ToUInt32($directory,$at+116);Size=$size;Path='';Parent='';Linked=$false})
    }
    if(-not $entries.Count -or $entries[0].Type -ne 5){throw 'Root storage missing.'}
    $root=$entries[0];$root.Linked=$true;$State.RootSize=$root.Size
    $State.RootChain=Get-LwChain $State $root.Start $State.Fat $State.SectorCount
    if([long]$State.RootChain.Length*$State.SectorSize -lt $root.Size -or ($root.Size -gt 0 -and [long]$State.RootChain.Length*$State.SectorSize-$root.Size -ge $State.SectorSize)){throw 'Root mini stream size mismatch.'}
    $stack=[Collections.Generic.Stack[object]]::new();$stack.Push(@($root.Child,''));$map=@{}
    while($stack.Count){
        Assert-LwTime $State;$task=$stack.Pop();$id=[uint32]$task[0];$parent=[string]$task[1]
        if($id -eq 4294967295){continue}
        if($id -eq 0 -or $id -ge $entries.Count){throw 'Invalid directory tree reference.'}
        $e=$entries[[int]$id];if($e.Linked -or $e.Type -notin @(1,2)){throw 'Cyclic or multiply-linked directory tree.'}
        $e.Linked=$true;$e.Parent=$parent;$e.Path=$e.Name;if($parent){$e.Path=$parent+'/'+$e.Name}
        if($e.Path.Length -gt 32768 -or $map.ContainsKey($e.Path)){throw 'Ambiguous or excessive directory path.'}
        $map[$e.Path]=$e;$stack.Push(@($e.Left,$parent));$stack.Push(@($e.Right,$parent))
        if($e.Type -eq 1){$stack.Push(@($e.Child,$e.Path))}elseif($e.Child -ne 4294967295){throw 'Stream has child directory entries.'}
    }
    foreach($e in $entries){if($e.Type -ne 0 -and -not $e.Linked){throw 'Unreachable allocated directory entry; inspection incomplete.'}}
    $State.Entries=$entries;$State.Map=$map
}

function Expand-LwVba([byte[]]$Bytes,[int]$Offset,$State,[int]$Maximum=16777216) {
    if($Offset -lt 0 -or $Offset -ge $Bytes.Length -or $Bytes[$Offset] -ne 1){throw 'Invalid VBA compression signature/offset.'}
    $output=[IO.MemoryStream]::new();$pos=$Offset+1;$chunk=[byte[]]::new(4096)
    try {
        while($pos -lt $Bytes.Length){
            Assert-LwTime $State
            if($pos+2 -gt $Bytes.Length){throw 'Truncated VBA chunk header.'}
            $start=$pos;$header=[BitConverter]::ToUInt16($Bytes,$pos);$pos+=2
            $size=($header -band 4095)+3;$end=$start+$size;$count=0
            if(($header -band 28672) -ne 12288 -or $end -gt $Bytes.Length){throw 'Invalid/truncated VBA chunk.'}
            if(($header -band 32768) -eq 0){
                if($size -ne 4098){throw 'Invalid raw VBA chunk size.'}
                [Buffer]::BlockCopy($Bytes,$pos,$chunk,0,4096);$count=4096;$pos=$end
            }else{
                while($pos -lt $end){
                    $flags=$Bytes[$pos];$pos++
                    if($pos -eq $end){throw 'VBA flag byte has no token.'}
                    for($bit=0;$bit -lt 8 -and $pos -lt $end;$bit++){
                        if(($flags -band (1 -shl $bit)) -eq 0){
                            if($count -ge 4096){throw 'VBA chunk output overflow.'};$chunk[$count]=$Bytes[$pos];$count++;$pos++
                        }else{
                            if($pos+2 -gt $end -or $count -eq 0){throw 'Invalid VBA copy token.'}
                            $token=[BitConverter]::ToUInt16($Bytes,$pos);$pos+=2
                            $bits=4;while((1 -shl $bits) -lt $count){$bits++}
                            $lengthMask=65535 -shr $bits;$length=($token -band $lengthMask)+3;$distance=($token -shr (16-$bits))+1
                            if($distance -gt $count -or $count+$length -gt 4096){throw 'VBA copy token is outside the current chunk.'}
                            for($j=0;$j -lt $length;$j++){$chunk[$count]=$chunk[$count-$distance];$count++}
                        }
                    }
                }
            }
            if($output.Length+$count -gt $Maximum -or $State.ExpandedBytes+$count -gt 67108864){throw 'VBA decompression limit exceeded.'}
            $output.Write($chunk,0,$count);$State.ExpandedBytes+=$count
        }
        return ,$output.ToArray()
    }finally{$output.Dispose()}
}

function Read-LwNumber($Reader,[int]$Width=2) {
    if($Reader.BaseStream.Length-$Reader.BaseStream.Position -lt $Width){throw 'Truncated VBA directory record.'}
    if($Width -eq 2){return $Reader.ReadUInt16()};return $Reader.ReadUInt32()
}
function Skip-LwBytes($Reader,[long]$Count) {
    if($Count -lt 0 -or $Count -gt $Reader.BaseStream.Length-$Reader.BaseStream.Position){throw 'VBA directory length exceeds available bytes.'}
    [void]$Reader.BaseStream.Seek($Count,[IO.SeekOrigin]::Current)
}
function Peek-LwId($Reader) {$id=Read-LwNumber $Reader;[void]$Reader.BaseStream.Seek(-2,[IO.SeekOrigin]::Current);return $id}
function Read-LwRecord($Reader,[int]$Expected,[int]$Fixed=-1) {
    $id=Read-LwNumber $Reader;$size=Read-LwNumber $Reader 4
    if($id -ne $Expected -or ($Fixed -ge 0 -and $size -ne $Fixed) -or $size -gt $Reader.BaseStream.Length-$Reader.BaseStream.Position){throw ('Unexpected VBA record/size: expected 0x{0:X4}, found 0x{1:X4}.' -f $Expected,$id)}
    return ,$Reader.ReadBytes([int]$size)
}
function Skip-LwReferenceName($Reader) {
    [void](Read-LwRecord $Reader 0x16)
    # Older Office writers can omit the Unicode companion.
    if((Peek-LwId $Reader) -eq 0x3e){[void](Read-LwRecord $Reader 0x3e)}
}
function Get-LwModules([byte[]]$Bytes,$State) {
    $ms=[IO.MemoryStream]::new($Bytes,$false);$reader=[IO.BinaryReader]::new($ms)
    try {
        [void](Read-LwRecord $reader 1 4)
        if((Peek-LwId $reader) -eq 0x4a){[void](Read-LwRecord $reader 0x4a 4)}
        [void](Read-LwRecord $reader 2 4);[void](Read-LwRecord $reader 0x14 4)
        $codepage=[BitConverter]::ToUInt16((Read-LwRecord $reader 3 2),0);$encoding=[Text.Encoding]::GetEncoding($codepage,[Text.EncoderFallback]::ExceptionFallback,[Text.DecoderFallback]::ExceptionFallback)
        [void](Read-LwRecord $reader 4);[void](Read-LwRecord $reader 5);[void](Read-LwRecord $reader 0x40)
        [void](Read-LwRecord $reader 6);[void](Read-LwRecord $reader 0x3d)
        [void](Read-LwRecord $reader 7 4);[void](Read-LwRecord $reader 8 4)
        [void](Read-LwRecord $reader 9 4);[void](Read-LwNumber $reader)
        if((Peek-LwId $reader) -eq 0x0c){[void](Read-LwRecord $reader 0x0c);[void](Read-LwRecord $reader 0x3c)}
        while((Peek-LwId $reader) -ne 0x0f){
            Assert-LwTime $State;$id=Read-LwNumber $reader
            switch($id){
                0x16 {[void]$ms.Seek(-2,[IO.SeekOrigin]::Current);Skip-LwReferenceName $reader}
                0x33 {Skip-LwBytes $reader (Read-LwNumber $reader 4)}
                0x0d {[void](Read-LwNumber $reader 4);Skip-LwBytes $reader (Read-LwNumber $reader 4);Skip-LwBytes $reader 6}
                0x0e {[void](Read-LwNumber $reader 4);Skip-LwBytes $reader (Read-LwNumber $reader 4);Skip-LwBytes $reader (Read-LwNumber $reader 4);Skip-LwBytes $reader 6}
                0x2f {
                    [void](Read-LwNumber $reader 4);Skip-LwBytes $reader (Read-LwNumber $reader 4);Skip-LwBytes $reader 6
                    if((Peek-LwId $reader) -eq 0x16){Skip-LwReferenceName $reader}
                    if((Read-LwNumber $reader) -ne 0x30){throw 'Invalid reference-control record.'}
                    [void](Read-LwNumber $reader 4);Skip-LwBytes $reader (Read-LwNumber $reader 4);Skip-LwBytes $reader 26
                }
                default {throw ('Unsupported VBA reference record 0x{0:X4}.' -f $id)}
            }
        }
        $count=[BitConverter]::ToUInt16((Read-LwRecord $reader 0x0f 2),0)
        if($count -gt 4096){throw 'VBA module count exceeds parser limit.'}
        [void](Read-LwRecord $reader 0x13 2);$modules=[Collections.Generic.List[object]]::new();$seen=@{}
        for($i=0;$i -lt $count;$i++){
            Assert-LwTime $State
            [void](Read-LwRecord $reader 0x19)
            if((Peek-LwId $reader) -eq 0x47){[void](Read-LwRecord $reader 0x47)}
            $stream=$encoding.GetString((Read-LwRecord $reader 0x1a));$unicode=Read-LwRecord $reader 0x32
            if($unicode.Length%2 -ne 0){throw 'Odd UTF16 stream-name length.'}
            $unicodeName=[Text.Encoding]::Unicode.GetString($unicode);if($unicodeName){$stream=$unicodeName}
            if(-not $stream -or $stream.IndexOfAny([char[]]@([char]0,'/','\')) -ge 0 -or $seen.ContainsKey($stream)){throw 'Invalid/duplicate VBA module stream name.'};$seen[$stream]=$true
            [void](Read-LwRecord $reader 0x1c);[void](Read-LwRecord $reader 0x48)
            $offset=[BitConverter]::ToUInt32((Read-LwRecord $reader 0x31 4),0)
            [void](Read-LwRecord $reader 0x1e 4);[void](Read-LwRecord $reader 0x2c 2)
            $type=Read-LwNumber $reader;if($type -notin @(0x21,0x22)){throw 'Invalid module type.'};[void](Read-LwNumber $reader 4)
            if((Peek-LwId $reader) -eq 0x25){[void](Read-LwRecord $reader 0x25 0)}
            if((Peek-LwId $reader) -eq 0x28){[void](Read-LwRecord $reader 0x28 0)}
            [void](Read-LwRecord $reader 0x2b 0)
            $modules.Add([pscustomobject]@{Stream=$stream;Offset=$offset})
        }
        [void](Read-LwRecord $reader 0x10 0)
        if($ms.Position -ne $ms.Length){throw 'Unexpected trailing VBA directory data.'}
        return [pscustomobject]@{Modules=$modules.ToArray();Encoding=$encoding}
    }finally{$reader.Dispose();$ms.Dispose()}
}

function Get-LegacyWordInspection([string]$FilePath,[long]$MaximumBytes=104857600,[int]$Seconds=120) {
    $result=[ordered]@{Status='Unknown';Format='';VBA='Unknown';Macro='Unknown';SourceModules=$null;ProjectCount=0;Scanner='PowerShell-CFB-OVBA-1.0';Error='';Notes='Static VBA project/source indicators in the outer OLE hierarchy. Not exhaustive WordBasic, packed embedded objects, external-template or compiled-only/stomped-code detection. NotDetected is not a safety verdict.'}
    $state=@{Stream=$null;Clock=[Diagnostics.Stopwatch]::StartNew();Timeout=$Seconds;ExpandedBytes=0L}
    try {
        $state.Stream=[IO.FileStream]::new((ConvertTo-LwNativePath $FilePath),[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        if($state.Stream.Length -gt $MaximumBytes){throw 'File exceeds the configured size limit.'}
        if($state.Stream.Length -lt 8 -or [BitConverter]::ToString((Read-LwAt $state 0 8)) -ne 'D0-CF-11-E0-A1-B1-1A-E1'){$result.Format='NotOLE';throw 'Not a legacy OLE Word binary; unsupported or mislabeled .doc.'}
        $result.Format='MalformedOLE';Initialize-LwCompound $state
        if($state.Map.ContainsKey('EncryptedPackage') -or $state.Map.ContainsKey('EncryptionInfo')){$result.Format='EncryptedOLE';throw 'Encrypted package; not decrypted or inspected.'}
        if(-not $state.Map.ContainsKey('WordDocument')){$result.Format='OtherOLE';throw 'OLE container has no root WordDocument stream.'}
        $fib=Read-LwStream $state $state.Map['WordDocument'] 32 32
        if($fib.Length -lt 32 -or [BitConverter]::ToUInt16($fib,0) -ne 0xa5ec -or [BitConverter]::ToUInt16($fib,2) -lt 0xc1){$result.Format='UnsupportedWordBinary';throw 'Unsupported or malformed Word FIB (possibly pre-Word97).'}
        $flags=[BitConverter]::ToUInt16($fib,10)
        if(($flags -band 0x8100) -ne 0){$result.Format='EncryptedOLE';throw 'Encrypted/obfuscated Word document; not decrypted or inspected.'}
        $result.Format='Word97PlusOLE';$projects=@($state.Map.Values | Where-Object {$_.Type -eq 1 -and $_.Name -ieq 'VBA'})
        $sources=0
        foreach($project in $projects){
            $prefix=$project.Path+'/';$parentPrefix='';if($project.Parent){$parentPrefix=$project.Parent+'/'}
            foreach($required in @(($prefix+'dir'),($prefix+'_VBA_PROJECT'),($parentPrefix+'PROJECT'))){if(-not $state.Map.ContainsKey($required) -or $state.Map[$required].Type -ne 2){throw "VBA-like storage is incomplete: $required is missing."}}
            $result.ProjectCount++
            $projectHeader=Read-LwStream $state $state.Map[$prefix+'_VBA_PROJECT'] 7 7
            if($projectHeader.Length -lt 7 -or [BitConverter]::ToUInt16($projectHeader,0) -ne 0x61cc){throw 'Invalid VBA project header.'}
            [void](Read-LwStream $state $state.Map[$parentPrefix+'PROJECT'] 1048576)
            $compressed=Read-LwStream $state $state.Map[$prefix+'dir'] 4194304
            $dir=Expand-LwVba $compressed 0 $state 4194304;$info=Get-LwModules $dir $state
            foreach($module in $info.Modules){
                $key=$prefix+$module.Stream;if(-not $state.Map.ContainsKey($key)){throw 'Declared VBA module stream is missing.'}
                $bytes=Read-LwStream $state $state.Map[$key]
                $sourceBytes=Expand-LwVba $bytes $module.Offset $state
                $source=$info.Encoding.GetString($sourceBytes);if(-not [string]::IsNullOrWhiteSpace($source)){$sources++}
                $source=$null;$sourceBytes=$null
            }
        }
        # Suspicious project streams without a linked VBA storage must not become negatives.
        if(-not $projects.Count -and @($state.Map.Values|Where-Object {$_.Name -ieq '_VBA_PROJECT'}).Count){throw 'Unassociated VBA project stream; incomplete inspection.'}
        $result.SourceModules=$sources;$result.Status='Scanned'
        if($projects.Count){$result.VBA='Detected';$result.Macro='VBAIndicatorsDetected'}else{$result.VBA='NotDetected';$result.Macro='NoVBAIndicatorsDetected'}
    }catch{$result.Error=$_.Exception.Message;$result.Status='Unknown';$result.VBA='Unknown';$result.Macro='Unknown';$result.SourceModules=$null}
    finally{if($state.Stream){$state.Stream.Dispose()}}
    return [pscustomobject]$result
}

if($PSBoundParameters.ContainsKey('Path')){
    $ErrorActionPreference='Stop';Set-StrictMode -Version 2
    $result=Get-LegacyWordInspection $Path ([long]$MaxFileMB*1MB) $TimeoutSeconds
    if($ResultPath){
        # Never overwrite an existing path, including the input document by mistake.
        $output=[IO.FileStream]::new((ConvertTo-LwNativePath $ResultPath),[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try{$writer=[IO.StreamWriter]::new($output,[Text.UTF8Encoding]::new($false));try{$writer.Write(($result|ConvertTo-Json -Depth 4))}finally{$writer.Dispose()}}finally{$output.Dispose()}
    }else{$result}
}
