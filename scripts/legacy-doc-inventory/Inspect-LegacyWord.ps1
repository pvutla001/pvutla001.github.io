#Requires -Version 5.1 
<#
.SYNOPSIS
Identify and inspect local Word containers using PowerShell and built-in .NET.
.DESCRIPTION
Read-only CFB/OVBA, RTF, Word OpenXML, Word XML and MHTML parser. No Python, Word, COM automation, network,
third-party modules, source export or macro execution. This helper replaces
Inspect-LegacyWord.py for inventory v3.10.1. It is not an antivirus or Word validator.
Supported scope: Word 97+ OLE hierarchies, recognized RTF embedded OLE objects,
Word OpenXML/Flat OPC VBA parts, Word 2003 XML and contained MHTML Office data.
HTML sidecar references remain unresolved; simple text is identified separately.
ZIP/ActiveMime support uses built-in .NET compression; ActiveMime compiles a small
input-stream adapter with Add-Type. No separate SDK is required.
WordBasic, packed embedded files, external templates and
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

function Initialize-LwLocalIO {
    if('LegacyWordPathV3101' -as [type]){return}
    # Handle-based streams avoid .NET Framework's legacy device-path validation.
    # This helper uses only Windows/.NET; no SharePoint assembly or extra file.
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
public static class LegacyWordPathV3101 {
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
    private static extern uint GetFullPathNameW(string path, uint size, StringBuilder buffer, IntPtr part);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
    private static extern SafeFileHandle CreateFileW(string path, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
        private static IOException Failure(string operation, string path, int code) {
            return new IOException(operation+" failed (Win32 "+code+") for '"+path+"': "+new Win32Exception(code).Message, new Win32Exception(code));
        }
        public static string FullPath(string path) {
            if(String.IsNullOrWhiteSpace(path)) throw new ArgumentException("Local path is empty.","path");
            path=path.Replace('/','\\');
            if(path.StartsWith(@"\\?\UNC\",StringComparison.OrdinalIgnoreCase)) path=@"\\"+path.Substring(8);
            else if(path.StartsWith(@"\\?\")) path=path.Substring(4);
            if(path.StartsWith(@"\\.\") || path.StartsWith(@"\??\")) throw new ArgumentException("Device paths are not document paths.","path");
            for(int i=0;i<path.Length;i++) {
                char c=path[i];
                if(c<32 || c=='"' || c=='<' || c=='>' || c=='|' || c=='*' || c=='?' || (c==':' && (i!=1 || !Char.IsLetter(path[0])))) throw new ArgumentException("Invalid character or alternate data stream in local path.","path");
            }
            var buffer=new StringBuilder(32768);
            uint count=GetFullPathNameW(path,(uint)buffer.Capacity,buffer,IntPtr.Zero);
            if(count==0) throw Failure("Normalize path",path,Marshal.GetLastWin32Error());
            if(count>=buffer.Capacity || count>32750) throw new PathTooLongException("Local path exceeds the supported Windows path length.");
            string full=buffer.ToString();
            int root=RootLength(full); return full.Length>root ? full.TrimEnd('\\') : full;
        }
        private static int RootLength(string full) {
            if(full.Length>=3 && Char.IsLetter(full[0]) && full[1]==':' && full[2]=='\\') return 3;
            if(full.StartsWith(@"\\")) {
                int server=full.IndexOf('\\',2); if(server<=2 || server==full.Length-1) throw new ArgumentException("UNC path requires server and share.");
                int share=full.IndexOf('\\',server+1); return share<0 ? full.Length : share+1;
            }
            throw new ArgumentException("Expected a drive or UNC path.");
        }
        public static string Extended(string path) {
            path=FullPath(path);
            if(path.StartsWith(@"\\")) return @"\\?\UNC\"+path.Substring(2);
            return @"\\?\"+path;
        }
        private static FileStream Open(string path, bool write) {
            SafeFileHandle handle=CreateFileW(Extended(path),write ? 0x40000000U : 0x80000000U,write ? 0U : 1U,IntPtr.Zero,write ? 1U : 3U,0x08000080U,IntPtr.Zero);
            if(handle.IsInvalid) { int code=Marshal.GetLastWin32Error(); handle.Dispose(); throw Failure(write ? "Create new file" : "Open file for reading",path,code); }
            try { return new FileStream(handle,write ? FileAccess.Write : FileAccess.Read,81920,false); }
            catch { handle.Dispose(); throw; }
        }
        public static FileStream OpenRead(string path) { return Open(path,false); }
        public static FileStream CreateNew(string path) { return Open(path,true); }

    public static byte[] ReadAllBytes(string path, long maximum) {
        using(var stream=OpenRead(path)) {
            if(stream.Length>maximum || stream.Length>Int32.MaxValue) throw new IOException("Input exceeds bounded inspection size limit.");
            byte[] bytes=new byte[(int)stream.Length]; int offset=0;
            while(offset<bytes.Length) {
                int read=stream.Read(bytes,offset,bytes.Length-offset);
                if(read==0) throw new EndOfStreamException("Input changed or was truncated during inspection.");
                offset+=read;
            }
            if(stream.ReadByte()!=-1) throw new IOException("Input grew during inspection.");
            return bytes;
        }
    }
}
'@
}
function ConvertTo-LwNativePath([string]$Value) {
    Initialize-LwLocalIO
    return [LegacyWordPathV3101]::Extended($Value)
}
function ConvertTo-LwFullPath([string]$Value) {
    Initialize-LwLocalIO
    return [LegacyWordPathV3101]::FullPath($Value)
}
function Assert-LwTime($State) {
    if($State.Clock.Elapsed.TotalSeconds -gt $State.Timeout){throw 'Parser time limit exceeded.'}
}
function Read-LwAt($State,[long]$Offset,[int]$Count) {
    $State.FileOffset=$Offset
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
    $State.StreamPath=$Entry.Path
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
    $State.CompressedOffset=$Offset
    if($Offset -lt 0 -or $Offset -ge $Bytes.Length -or $Bytes[$Offset] -ne 1){throw 'Invalid VBA compression signature/offset.'}
    $output=[IO.MemoryStream]::new();$pos=$Offset+1;$chunk=[byte[]]::new(4096)
    try {
        while($pos -lt $Bytes.Length){
            $State.CompressedOffset=$pos
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
                    # MS-OVBA 2.4.1.3.4 permits a flag byte at the chunk end.
                    # Bounds and copy-token checks still reject actual truncation.
                    if($pos -eq $end){break}
                    for($bit=0;$bit -lt 8 -and $pos -lt $end;$bit++){
                        $State.CompressedOffset=$pos
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
    $recordOffset=$Reader.BaseStream.Position
    $id=Read-LwNumber $Reader;$size=Read-LwNumber $Reader 4
    if($id -ne $Expected -or ($Fixed -ge 0 -and $size -ne $Fixed) -or $size -gt $Reader.BaseStream.Length-$Reader.BaseStream.Position){throw ('Unexpected VBA directory record at decompressed offset {0}: expected ID=0x{1:X4}, found ID=0x{2:X4}, declared size={3}, required size={4}, remaining bytes={5}.' -f $recordOffset,$Expected,$id,$size,$Fixed,($Reader.BaseStream.Length-$Reader.BaseStream.Position))}
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

function New-LwDiagnostic($Record,$State,[string]$Code='ParserError') {
    $exception=$Record.Exception
    [pscustomobject][ordered]@{
        Severity='Error';Code=$Code;Stage=$State.Stage;Message=$exception.Message
        ExceptionType=$exception.GetType().FullName;ExceptionDetails=$exception.ToString()
        FullyQualifiedErrorId=$Record.FullyQualifiedErrorId;ScriptStackTrace=$Record.ScriptStackTrace
        ProjectPath=$State.ProjectPath;StreamPath=$State.StreamPath;ModuleName=$State.ModuleName
        FileOffset=$State.FileOffset;CompressedOffset=$State.CompressedOffset
    }
}
function Get-LwNonOleFormat([byte[]]$Header) {
    if(-not $Header.Length){return 'EmptyFile'}
    $text=[Text.Encoding]::UTF8.GetString($Header).TrimStart([char]0xfeff,[char]0x20,[char]0x09,[char]0x0d,[char]0x0a)
    if($Header.Length -ge 2 -and $Header[0] -eq 255 -and $Header[1] -eq 254){$text=[Text.Encoding]::Unicode.GetString($Header).TrimStart([char]0xfeff,[char]0x20,[char]0x0d,[char]0x0a)}
    if($Header.Length -ge 2 -and $Header[0] -eq 254 -and $Header[1] -eq 255){$text=[Text.Encoding]::BigEndianUnicode.GetString($Header).TrimStart([char]0xfeff,[char]0x20,[char]0x0d,[char]0x0a)}
    if($text.StartsWith('ActiveMime')){return 'ActiveMime'}
    if($text -match '(?im)^(MIME-Version\s*:|Content-Type\s*:\s*multipart/)'){return 'MHTML'}
    if($text.StartsWith('{\rtf',[StringComparison]::OrdinalIgnoreCase)){return 'RTF'}
    if($text -match '^(?is)(<!DOCTYPE\s+html|<html\b)'){return 'HTML'}
    if($text -match '^(?is)(<\?xml\b|<!DOCTYPE\b|<[A-Za-z_][A-Za-z0-9_.:-]*(?:\s|/?>))'){return 'XML'}
    if($text.StartsWith('%PDF-')){return 'PDF'}
    if($Header.Length -ge 4 -and $Header[0] -eq 80 -and $Header[1] -eq 75 -and (($Header[2] -eq 3 -and $Header[3] -eq 4) -or ($Header[2] -eq 5 -and $Header[3] -eq 6))){return 'ZIPContainer'}
    if($Header.Length -ge 2 -and (($Header[0] -eq 255 -and $Header[1] -eq 254) -or ($Header[0] -eq 254 -and $Header[1] -eq 255))){return 'TextCandidate'}
    if($text -notmatch '[\x00-\x08\x0b\x0e-\x1f\ufffd]'){return 'TextCandidate'}
    return 'NotOLE'
}
function Get-LwCompoundInspection([string]$FilePath,[long]$MaximumBytes=104857600,[int]$Seconds=120,[byte[]]$Content,[switch]$VbaContainer) {
    $result=[ordered]@{
        SchemaVersion='2.0';Status='Unknown';Format='';VBA='Unknown';Macro='Unknown'
        SourceModules=$null;SourceModulesParsed=0;ProjectCount=$null;ProjectsInspected=0
        Scanner='PowerShell-CFB-OVBA-1.4.1';FilePath=(ConvertTo-LwFullPath $FilePath);FileLength=$null;SignatureHex=''
        DetectionScope='OuterOLEHierarchy';EmbeddedObjects=0;FormatDetails='';FIBVersion=$null;EffectiveFIBVersion=$null;IsTemplate=$null
        ElapsedMilliseconds=0;Error='';ErrorCode='';Diagnostics=@()
        Notes='Static VBA project/source indicators in the outer OLE hierarchy. Not exhaustive WordBasic, packed embedded objects, external-template or compiled-only/stomped-code detection. NotDetected is not a safety verdict.'
    }
    $issues=[Collections.Generic.List[object]]::new()
    $state=@{Stream=$null;Clock=[Diagnostics.Stopwatch]::StartNew();Timeout=$Seconds;ExpandedBytes=0L;Stage='OpenLocalFile';ProjectPath='';StreamPath='';ModuleName='';FileOffset=$null;CompressedOffset=$null}
    try {
        if($PSBoundParameters.ContainsKey('Content')){$state.Stream=[IO.MemoryStream]::new($Content,$false)}
        else{$state.Stream=[LegacyWordPathV3101]::OpenRead($FilePath)}
        $result.FileLength=$state.Stream.Length
        if($state.Stream.Length -gt $MaximumBytes){throw 'File exceeds the configured size limit.'}
        $state.Stage='IdentifyFormat';$header=Read-LwAt $state 0 ([int][math]::Min(512,$state.Stream.Length))
        if($header.Length){$result.SignatureHex=[BitConverter]::ToString($header,0,[math]::Min(16,$header.Length))}
        if($header.Length -lt 8 -or [BitConverter]::ToString($header,0,8) -ne 'D0-CF-11-E0-A1-B1-1A-E1'){
            $result.Format=Get-LwNonOleFormat $header;$result.Status='Unsupported';$result.VBA='Unknown';$result.Macro='Unknown';$result.DetectionScope='IdentificationOnly'
            throw ('Not a legacy OLE Word binary. Detected format/signature: {0}; bytes={1}; header={2}. The .doc extension does not establish the internal format; this scanner does not inspect VBA in this format.' -f $result.Format,$result.FileLength,$result.SignatureHex)
        }
        $result.Format='MalformedOLE';$state.Stage='ReadCompoundDirectory';Initialize-LwCompound $state
        if($state.Map.ContainsKey('EncryptedPackage') -or $state.Map.ContainsKey('EncryptionInfo')){$result.Format='EncryptedOLE';throw 'Encrypted package; not decrypted or inspected.'}
        if($VbaContainer){$result.Format='VBAProjectOLE'}else{
        if(-not $state.Map.ContainsKey('WordDocument')){$result.Format='OtherOLE';$result.Status='Unsupported';throw 'OLE container has no root WordDocument stream.'}
        $state.Stage='ReadWordHeader';$fib=Read-LwStream $state $state.Map['WordDocument'] 32 32
        if($fib.Length -lt 32 -or [BitConverter]::ToUInt16($fib,0) -notin @(0xa5ec,0xa5dc)){$result.Format='MalformedWordBinary';throw 'Malformed Word FIB: missing Word signature or truncated header.'}
        $result.FIBVersion=[int][BitConverter]::ToUInt16($fib,2)
        if($result.FIBVersion -lt 0xc1){
            $result.Format='PreWord97OLE';if($result.FIBVersion -in @(101,104)){$result.Format='Word6Or95OLE'}
            $result.Status='Unsupported';$result.DetectionScope='IdentificationOnly'
            throw ('Pre-Word97 document (nFib={0}); WordBasic macro analysis is not implemented. VBA/Macro=Unknown, not a negative finding.' -f $result.FIBVersion)
        }
        if([BitConverter]::ToUInt16($fib,0) -ne 0xa5ec){throw 'Word 97+ nFib does not match the Word header signature.'}
        if(([BitConverter]::ToUInt16($fib,10) -band 0x8100) -ne 0){$result.Format='EncryptedOLE';throw 'Encrypted/obfuscated Word document; not decrypted or inspected.'}
        $result.Format='Word97PlusOLE'
        $result.FormatDetails='Word 97 or later compound binary; nFib='+$result.FIBVersion
        $result.IsTemplate=(([BitConverter]::ToUInt16($fib,10) -band 1) -ne 0)
        $result.EffectiveFIBVersion=$result.FIBVersion
        # The newer revision can be in FibRgCswNew rather than FibBase. This is
        # format-generation evidence, never a document creation date.
        $extendedFib=Read-LwStream $state $state.Map['WordDocument'] 8192 8192
        if($extendedFib.Length -ge 34){
            $at=34+2*[int][BitConverter]::ToUInt16($extendedFib,32)
            if($at+2 -le $extendedFib.Length){$at+=2+4*[int][BitConverter]::ToUInt16($extendedFib,$at)
                if($at+2 -le $extendedFib.Length){$at+=2+8*[int][BitConverter]::ToUInt16($extendedFib,$at)
                    if($at+4 -le $extendedFib.Length -and [BitConverter]::ToUInt16($extendedFib,$at) -gt 0){$result.EffectiveFIBVersion=[int][BitConverter]::ToUInt16($extendedFib,$at+2)}
                }
            }
        }
        $generations=@{193='Word 97';217='Word 2000';257='Word 2002/XP';268='Word 2003';274='Word 2007 and later legacy-save'}
        $generation='Unmapped revision';if($generations.ContainsKey($result.EffectiveFIBVersion)){$generation=$generations[$result.EffectiveFIBVersion]}
        $result.FormatDetails+='; effective nFib='+$result.EffectiveFIBVersion+' ('+$generation+' binary format); template='+$result.IsTemplate+'. Format revision does not establish the authoring product/year.'
        }
        $projects=@($state.Map.Values | Where-Object {$_.Type -eq 1 -and $_.Name -ieq 'VBA'} | Sort-Object Path)
        if($VbaContainer -and -not $projects.Count){throw 'Declared VBA project part contains no VBA storage; content is incomplete or invalid.'}
        # Count discovered VBA storage paths ONCE, before attempting their sources.
        $result.ProjectCount=$projects.Count
        if($projects.Count){$result.VBA='Detected';$result.Macro='VBAIndicatorsDetected'}
        $sources=0
        foreach($project in $projects){
            $state.ProjectPath=$project.Path;$state.ModuleName='';$state.CompressedOffset=$null
            try {
                $state.Stage='ValidateVbaProject';$prefix=$project.Path+'/';$parentPrefix='';if($project.Parent){$parentPrefix=$project.Parent+'/'}
                foreach($required in @(($prefix+'dir'),($prefix+'_VBA_PROJECT'),($parentPrefix+'PROJECT'))){if(-not $state.Map.ContainsKey($required) -or $state.Map[$required].Type -ne 2){throw "VBA-like storage is incomplete: $required is missing."}}
                $projectHeader=Read-LwStream $state $state.Map[$prefix+'_VBA_PROJECT'] 7 7
                if($projectHeader.Length -lt 7 -or [BitConverter]::ToUInt16($projectHeader,0) -ne 0x61cc){throw 'Invalid VBA project header.'}
                [void](Read-LwStream $state $state.Map[$parentPrefix+'PROJECT'] 1048576)
                $state.Stage='VbaDirectoryDecompression';$compressed=Read-LwStream $state $state.Map[$prefix+'dir'] 4194304
                $dir=Expand-LwVba $compressed 0 $state 4194304
                $state.Stage='VbaDirectoryRecords';$info=Get-LwModules $dir $state
                foreach($module in $info.Modules){
                    $state.ModuleName=$module.Stream;$state.CompressedOffset=$module.Offset;$state.Stage='VbaModuleSource'
                    try {
                        $key=$prefix+$module.Stream;$state.StreamPath=$key
                        if(-not $state.Map.ContainsKey($key)){throw 'Declared VBA module stream is missing.'}
                        $bytes=Read-LwStream $state $state.Map[$key];$sourceBytes=Expand-LwVba $bytes $module.Offset $state
                        $source=$info.Encoding.GetString($sourceBytes);if(-not [string]::IsNullOrWhiteSpace($source)){$sources++}
                    } catch {$issues.Add((New-LwDiagnostic $_ $state 'VbaModuleParseError'))}
                    finally {$source=$null;$sourceBytes=$null}
                }
                $result.ProjectsInspected++
            } catch {$issues.Add((New-LwDiagnostic $_ $state 'VbaProjectParseError'))}
        }
        if(-not $projects.Count -and @($state.Map.Values|Where-Object {$_.Name -ieq '_VBA_PROJECT'}).Count){$state.Stage='ValidateVbaProject';throw 'Unassociated VBA project stream; incomplete inspection.'}
        $result.SourceModulesParsed=$sources
        if($issues.Count){$result.Status='Partial';$result.Notes+=' VBA storage detected; source analysis was incomplete. ProjectCount is the number of discovered VBA storages; SourceModulesParsed is a lower bound.'}
        else{
            $result.SourceModules=$sources;$result.Status='Scanned'
            if(-not $projects.Count){$result.VBA='NotDetected';$result.Macro='NoVBAIndicatorsDetected'}
        }
    } catch {
        $code='ParserError';if($result.Status -eq 'Unsupported'){$code='UnsupportedFormat'}elseif($result.Format -eq 'EncryptedOLE'){$code='EncryptedDocument'}elseif($_.Exception.Message -match 'limit'){$code='ParserResourceLimit'}
        $issues.Add((New-LwDiagnostic $_ $state $code));$result.SourceModules=$null
        if($result.Status -ne 'Unsupported'){
            if($null -ne $result.ProjectCount -and $result.ProjectCount -gt 0){$result.Status='Partial';$result.VBA='Detected';$result.Macro='VBAIndicatorsDetected'}
            else{$result.Status='Unknown';$result.VBA='Unknown';$result.Macro='Unknown'}
        }
    } finally {if($state.Stream){$state.Stream.Dispose()};$result.ElapsedMilliseconds=$state.Clock.ElapsedMilliseconds}
    $result.Diagnostics=$issues.ToArray()
    if($issues.Count){$result.Error=($issues | ForEach-Object {'['+$_.Stage+'] '+$_.Message}) -join ' | ';$result.ErrorCode=(@($issues | ForEach-Object {$_.Code} | Select-Object -Unique) -join '; ')}
    return [pscustomobject]$result
}
function Read-LwZipEntry($Entry,$Budget,$State) {
    if($Entry.Length -gt 33554432 -or $Budget.Bytes+$Entry.Length -gt 67108864){throw 'Embedded content allocation limit exceeded (32 MiB per entry, 64 MiB total).'}
    $inputStream=$Entry.Open();$outputStream=[IO.MemoryStream]::new();$buffer=[byte[]]::new(65536)
    try{while(($n=$inputStream.Read($buffer,0,$buffer.Length)) -gt 0){
        Assert-LwTime $State;$Budget.Bytes+=$n
        if($outputStream.Length+$n -gt 33554432 -or $Budget.Bytes -gt 67108864){throw 'Expanded ZIP content exceeds allocation limit.'}
        $outputStream.Write($buffer,0,$n)
    };if($outputStream.Length -ne $Entry.Length){throw 'ZIP entry size differs from its declared length.'};return ,$outputStream.ToArray()}
    finally{$inputStream.Dispose();$outputStream.Dispose()}
}
function Read-LwSafeXml([byte[]]$Bytes) {
    $settings=[Xml.XmlReaderSettings]::new();$settings.DtdProcessing=[Xml.DtdProcessing]::Prohibit;$settings.XmlResolver=$null;$settings.MaxCharactersInDocument=4194304
    $ms=[IO.MemoryStream]::new($Bytes,$false);$reader=$null
    try{$reader=[Xml.XmlReader]::Create($ms,$settings);$doc=[Xml.XmlDocument]::new();$doc.XmlResolver=$null;$doc.Load($reader);return ,$doc}
    finally{if($reader){$reader.Dispose()};$ms.Dispose()}
}
function Get-LwRtfObjects([byte[]]$Bytes,$State) {
    # Tokenize RTF structure; do not search arbitrary document text for a CFB signature.
    # Unsupported escapes/controls inside objdata fail closed instead of silently dropping bytes.
    $text=[Text.Encoding]::GetEncoding(28591).GetString($Bytes)
    $lexer=[regex]::new('\\(?:[A-Za-z]+-?\d* ?|''[0-9a-fA-F]{2}|[^A-Za-z])|[{}]|[^{}\\]+',[Text.RegularExpressions.RegexOptions]::CultureInvariant,[TimeSpan]::FromSeconds(2))
    $objects=[Collections.Generic.List[byte[]]]::new();$stack=[Collections.Generic.Stack[object]]::new()
    $frame=@{Data=$null;Owns=$false};$pos=0;$tokens=0;$total=0L;$objectDeclarations=0
    while($pos -lt $text.Length){
        if((++$tokens % 256) -eq 0){Assert-LwTime $State};if($tokens -gt 2000000){throw 'RTF token limit exceeded.'}
        $m=$lexer.Match($text,$pos);if(-not $m.Success -or $m.Index -ne $pos){throw ('Invalid RTF token at byte '+$pos)}
        $token=$m.Value;$pos+= $m.Length
        if($token -eq '{'){
            if($stack.Count -ge 256){throw 'RTF nesting limit exceeded.'};$stack.Push($frame);$frame=@{Data=$frame.Data;Owns=$false};continue
        }
        if($token -eq '}'){
            if(-not $stack.Count){throw 'Unbalanced RTF closing brace.'}
            if($frame.Owns){
                $hex=$frame.Data.ToString();if($hex.Length % 2){throw 'Odd hex nibble count in RTF objdata.'}
                $data=[byte[]]::new($hex.Length/2)
                for($j=0;$j -lt $data.Length;$j++){if(($j -band 16383) -eq 0){Assert-LwTime $State};$data[$j]=[Convert]::ToByte($hex.Substring($j*2,2),16)}
                $total+=$data.Length;if($total -gt 67108864 -or $objects.Count -ge 32){throw 'RTF embedded content limit exceeded.'};$objects.Add($data)
            }
            $frame=$stack.Pop();continue
        }
        if($token.StartsWith('\')){
            if($token -match '^\\bin(-?\d+) ?$'){
                $n=[long]$Matches[1];if($n -lt 0 -or $n -gt $text.Length-$pos){throw 'Invalid/truncated RTF binary data.'}
                if($null -ne $frame.Data){if($frame.Data.Length+$n*2 -gt 67108864){throw 'RTF object allocation limit exceeded.'};[void]$frame.Data.Append([BitConverter]::ToString($Bytes,$pos,[int]$n).Replace('-',''))};$pos+=[int]$n;continue
            }
            if($token -match '^\\object ?$'){$objectDeclarations++;continue}
            if($token -match '^\\objdata ?$'){
                if($null -ne $frame.Data -or -not $stack.Count){throw 'Nested/misplaced RTF objdata.'};$frame.Data=[Text.StringBuilder]::new();$frame.Owns=$true;continue
            }
            if($null -ne $frame.Data -and $token -notmatch '^\\[\r\n]$'){throw ('Unsupported RTF objdata control '+$token+'; embedded content was not completely assessed.')}
            continue
        }
        if($null -ne $frame.Data){
            if($token -match '[^0-9a-fA-F\s]'){throw 'Non-hexadecimal content in RTF objdata.'}
            $hex=[regex]::Replace($token,'\s','');if($frame.Data.Length+$hex.Length -gt 67108864){throw 'RTF object allocation limit exceeded.'};[void]$frame.Data.Append($hex)
        }
    }
    if($stack.Count){throw 'Unbalanced/truncated RTF groups.'}
    if($objectDeclarations -ne $objects.Count){throw ('RTF declares {0} objects but exposes {1} inspectable objdata groups; linked/unresolved objects are not fetched.' -f $objectDeclarations,$objects.Count)}
    return ,$objects
}
function Get-LwOle1Native([byte[]]$Bytes) {
    if($Bytes.Length -ge 8 -and [BitConverter]::ToString($Bytes,0,8) -eq 'D0-CF-11-E0-A1-B1-1A-E1'){return ,$Bytes}
    $ms=[IO.MemoryStream]::new($Bytes,$false);$reader=[IO.BinaryReader]::new($ms)
    try{
        [void](Read-LwNumber $reader 4);$kind=Read-LwNumber $reader 4
        if($kind -ne 2){throw ('OLE1 object is not an embedded native object (FormatID='+$kind+'); linked/static objects are not inspected.')}
        for($i=0;$i -lt 3;$i++){$n=Read-LwNumber $reader 4;if($n -gt 1048576){throw 'OLE1 header string exceeds limit.'};Skip-LwBytes $reader $n}
        $size=Read-LwNumber $reader 4;if($size -gt 33554432 -or $size -gt $ms.Length-$ms.Position){throw 'OLE1 native data length is truncated or exceeds limit.'}
        return ,$reader.ReadBytes([int]$size)
    }finally{$reader.Dispose();$ms.Dispose()}
}
function ConvertFrom-LwText([byte[]]$Bytes) {
    $encoding=[Text.UTF8Encoding]::new($false,$true);$start=0
    if($Bytes.Length -ge 2 -and $Bytes[0] -eq 255 -and $Bytes[1] -eq 254){$encoding=[Text.UnicodeEncoding]::new($false,$true,$true);$start=2}
    elseif($Bytes.Length -ge 2 -and $Bytes[0] -eq 254 -and $Bytes[1] -eq 255){$encoding=[Text.UnicodeEncoding]::new($true,$true,$true);$start=2}
    elseif($Bytes.Length -ge 3 -and $Bytes[0] -eq 239 -and $Bytes[1] -eq 187 -and $Bytes[2] -eq 191){$start=3}
    return $encoding.GetString($Bytes,$start,$Bytes.Length-$start)
}
function ConvertFrom-LwBase64([string]$Value,$Budget,$State) {
    Assert-LwTime $State
    if($Value.Length -gt 46000000){throw 'Base64 input exceeds bounded allocation limit.'}
    $bytes=[Convert]::FromBase64String($Value)
    $Budget.Bytes+=$bytes.Length
    if($bytes.Length -gt 33554432 -or $Budget.Bytes -gt 67108864){throw 'Decoded embedded content exceeds allocation limit.'}
    return ,$bytes
}
function Expand-LwActiveMime([byte[]]$Bytes,$Budget,$State) {
    # ActiveMime layout is not an Office open specification. These bounded header
    # offsets are interoperability observations documented by oletools' authors.
    # No brute-force signature searching; zlib header and Adler32 are validated.
    if($Bytes.Length -lt 54 -or [Text.Encoding]::ASCII.GetString($Bytes,0,10) -ne 'ActiveMime'){throw 'Truncated/invalid ActiveMime header.'}
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    # DeflateStream can read ahead into ActiveMime's following metadata records.
    # A one-byte input adapter preserves the exact end of the zlib member, so its
    # checksum can be checked without mistaking the outer trailer for Adler32.
    if(-not ('LwDeflateInputV14' -as [type])){Add-Type -TypeDefinition @'
using System;
using System.IO;
public sealed class LwDeflateInputV14 : Stream {
 private readonly byte[] data; private readonly int start; private int cursor;
 public LwDeflateInputV14(byte[] value,int offset){data=value;start=offset;cursor=offset;}
 public int AbsoluteOffset {get{return cursor;}}
 public override bool CanRead {get{return true;}} public override bool CanSeek {get{return false;}} public override bool CanWrite {get{return false;}}
 public override long Length {get{return data.Length-start;}}
 public override long Position {get{return cursor-start;}set{throw new NotSupportedException();}}
 public override int Read(byte[] buffer,int offset,int count){if(count==0||cursor>=data.Length)return 0;buffer[offset]=data[cursor++];return 1;}
 public override void Flush(){} public override long Seek(long o,SeekOrigin s){throw new NotSupportedException();}
 public override void SetLength(long l){throw new NotSupportedException();} public override void Write(byte[] b,int o,int c){throw new NotSupportedException();}
}
'@ -ErrorAction Stop}
    $offsets=@(([int][BitConverter]::ToUInt16($Bytes,30)+46),50,554)|Select-Object -Unique
    $lastError='No supported zlib offset.'
    foreach($offset in $offsets){
        Assert-LwTime $State
        if($offset+6 -gt $Bytes.Length){continue}
        $cmf=[int]$Bytes[$offset];$flg=[int]$Bytes[$offset+1]
        if(($cmf -band 15) -ne 8 -or ($cmf -shr 4) -gt 7 -or (($cmf*256+$flg)%31) -ne 0 -or ($flg -band 32)){continue}
        $inputStream=[LwDeflateInputV14]::new($Bytes,$offset+2);$outputStream=[IO.MemoryStream]::new();$inflate=$null
        try{
            $inflate=[IO.Compression.DeflateStream]::new($inputStream,[IO.Compression.CompressionMode]::Decompress)
            $buffer=[byte[]]::new(65536)
            while(($n=$inflate.Read($buffer,0,$buffer.Length)) -gt 0){
                Assert-LwTime $State;$Budget.Bytes+=$n
                if($outputStream.Length+$n -gt 33554432 -or $Budget.Bytes -gt 67108864){throw 'ActiveMime decompression allocation limit exceeded.'}
                $outputStream.Write($buffer,0,$n)
            }
            $expanded=$outputStream.ToArray();$a=1L;$b=0L
            for($i=0;$i -lt $expanded.Length;$i++){$a=($a+$expanded[$i])%65521;$b=($b+$a)%65521;if(($i -band 16383) -eq 0){Assert-LwTime $State}}
            $checksum=$b*65536+$a;$end=$inputStream.AbsoluteOffset+4
            if($end -gt $Bytes.Length){throw 'Truncated ActiveMime zlib checksum.'}
            $expected=[long]$Bytes[$end-4]*16777216+[long]$Bytes[$end-3]*65536+[long]$Bytes[$end-2]*256+[long]$Bytes[$end-1]
            if($checksum -ne $expected){throw ('ActiveMime zlib checksum mismatch or trailing/truncated payload. Offset={0}; compressed bytes={1}; expanded bytes={2}; computed Adler32={3:X8}; end Adler32={4:X8}; final bytes={5}.' -f $offset,($Bytes.Length-$offset),$expanded.Length,$checksum,$expected,[BitConverter]::ToString($Bytes,[math]::Max(0,$Bytes.Length-24),[math]::Min(24,$Bytes.Length)))}
            if($expanded.Length -lt 8 -or [BitConverter]::ToString($expanded,0,8) -ne 'D0-CF-11-E0-A1-B1-1A-E1'){throw 'ActiveMime payload is not a supported compound VBA/Word container.'}
            return ,$expanded
        }catch{$lastError=$_.Exception.Message;if($lastError -match 'limit'){throw}}
        finally{if($inflate){$inflate.Dispose()};$inputStream.Dispose();$outputStream.Dispose()}
    }
    throw ('ActiveMime could not be decoded at supported offsets. '+$lastError)
}
function Add-LwEmbeddedInspection($Children,$Issues,[string]$Location,[byte[]]$Bytes,[bool]$VbaContainer,$Context) {
    $s=$Context.State;$s.StreamPath=$Location
    try{
        if((++$Context.Budget.Objects) -gt 32 -or $Context.Depth -ge 3){throw 'Embedded object count/nesting limit exceeded.'}
        if($Bytes.Length -ge 10 -and [Text.Encoding]::ASCII.GetString($Bytes,0,10) -eq 'ActiveMime'){
            $s.Stage='DecompressActiveMime';$Bytes=Expand-LwActiveMime $Bytes $Context.Budget $s;$VbaContainer=$true
        }
        $seconds=[math]::Max(1,($Context.Budget.Timeout-[int]$Context.Budget.Clock.Elapsed.TotalSeconds))
        if($VbaContainer){$child=Get-LwCompoundInspection -FilePath $Context.Path -MaximumBytes $Context.Max -Seconds $seconds -Content $Bytes -VbaContainer}
        else{$child=Get-LegacyWordInspection -FilePath $Context.Path -MaximumBytes $Context.Max -Seconds $seconds -Content $Bytes -Depth ($Context.Depth+1) -Budget $Context.Budget}
        $Children.Add([pscustomobject]@{Location=$Location;Result=$child})
    }catch{$Issues.Add((New-LwDiagnostic $_ $s 'EmbeddedPartError'))}
}
function Get-LwHtmlReferences([string]$Text) {
    $refs=[Collections.Generic.List[string]]::new()
    $regex=[regex]::new('<(?:link|object|script)\b[^>]*>',[Text.RegularExpressions.RegexOptions]::IgnoreCase,[TimeSpan]::FromSeconds(2))
    foreach($m in $regex.Matches($Text)){
        if($m.Value -match '(?i)Edit-Time-Data|\bobject\b|\bscript\b'){
            if($m.Value -match '(?i)\b(?:href|data|src)\s*=\s*(?:"([^"]*)"|''([^'']*)''|([^\s>]+))'){$value=$Matches[1]+$Matches[2]+$Matches[3];$refs.Add([Net.WebUtility]::HtmlDecode($value))}
            else{$refs.Add('(inline script/object requiring separate review)')}
        }
    }
    if($Text -match '(?i)(?:ActiveMime|<\s*(?:\w+:)?binData\b|data:application/)'){$refs.Add('(inline encoded Office content requiring separate review)')}
    return ,$refs
}
function Read-LwMimeParts([string]$Text,$State,[int]$Level=0) {
    if($Level -gt 4){throw 'MIME nesting limit exceeded.'}
    $separator=[regex]::Match($Text,'\r?\n\r?\n')
    if(-not $separator.Success -or $separator.Index -gt 65536){throw 'Missing/oversized MIME headers.'}
    $raw=$Text.Substring(0,$separator.Index);$body=$Text.Substring($separator.Index+$separator.Length);$headers=@{}
    $raw=[regex]::Replace($raw,'\r?\n[ \t]+',' ')
    foreach($line in ($raw -split '\r?\n')){if($line -notmatch '^([^:\s]+):\s*(.*)$'){throw 'Malformed MIME header.'};$name=$Matches[1];if($headers.ContainsKey($name)){throw ('Duplicate MIME header: '+$name)};$headers[$name]=$Matches[2]}
    $type=[string]$headers['Content-Type'];$parts=[Collections.Generic.List[object]]::new()
    if($type -match '(?i)^multipart/'){
        if($type -notmatch '(?i)\bboundary\s*=\s*(?:"([^"]+)"|([^;\s]+))'){throw 'Multipart MIME boundary is missing.'}
        $boundary=$Matches[1]+$Matches[2];if($boundary.Length -gt 200){throw 'MIME boundary exceeds limit.'}
        $regex=[regex]::new('(?m)^--'+[regex]::Escape($boundary)+'(?<end>--)?[ \t]*\r?$', [Text.RegularExpressions.RegexOptions]::CultureInvariant,[TimeSpan]::FromSeconds(2))
        $markers=$regex.Matches($body)
        if($markers.Count -lt 2 -or $markers.Count -gt 129 -or -not $markers[$markers.Count-1].Groups['end'].Success){throw 'Truncated MIME boundary sequence or part count limit exceeded.'}
        for($i=0;$i -lt $markers.Count-1;$i++){
            Assert-LwTime $State;if($markers[$i].Groups['end'].Success){throw 'MIME has content after closing boundary.'}
            $start=$markers[$i].Index+$markers[$i].Length;if($start -lt $body.Length -and $body[$start] -eq "`n"){$start++}
            $part=$body.Substring($start,$markers[$i+1].Index-$start).TrimEnd([char]13,[char]10)
            foreach($p in (Read-LwMimeParts $part $State ($Level+1))){$parts.Add($p);if($parts.Count -gt 128){throw 'MIME part count limit exceeded.'}}
        }
    }else{$parts.Add([pscustomobject]@{Headers=$headers;Body=$body})}
    return ,$parts
}
function ConvertFrom-LwMimeBody($Part,$Budget,$State) {
    $encoding=[string]$Part.Headers['Content-Transfer-Encoding'];$body=$Part.Body
    if($encoding -ieq 'base64'){return ,(ConvertFrom-LwBase64 $body $Budget $State)}
    if($encoding -ieq 'quoted-printable'){
        $ms=[IO.MemoryStream]::new()
        try{
            $regex=[regex]::new('[^=]+|=(?:\r\n|\n|[0-9a-fA-F]{2})',[Text.RegularExpressions.RegexOptions]::CultureInvariant,[TimeSpan]::FromSeconds(2));$pos=0
            while($pos -lt $body.Length){
                Assert-LwTime $State;$m=$regex.Match($body,$pos);if(-not $m.Success -or $m.Index -ne $pos){throw 'Malformed quoted-printable MIME data.'};$pos+=$m.Length
                if($m.Value.StartsWith('=')){if($m.Value -match '^=[0-9a-fA-F]{2}$'){$ms.WriteByte([Convert]::ToByte($m.Value.Substring(1),16))}}
                else{$b=[Text.Encoding]::GetEncoding(28591).GetBytes($m.Value);$ms.Write($b,0,$b.Length)}
                if($ms.Length -gt 33554432){throw 'MIME decoded part limit exceeded.'}
            }
            $bytes=$ms.ToArray()
        }finally{$ms.Dispose()}
    }elseif(-not $encoding -or $encoding -in @('7bit','8bit','binary')){$bytes=[Text.Encoding]::GetEncoding(28591).GetBytes($body)}
    else{throw ('Unsupported MIME transfer encoding: '+$encoding)}
    $Budget.Bytes+=$bytes.Length;if($bytes.Length -gt 33554432 -or $Budget.Bytes -gt 67108864){throw 'MIME aggregate allocation limit exceeded.'};return ,$bytes
}
function Invoke-LwAdditionalContainer($Result,[byte[]]$Content,[string]$FilePath,[long]$MaximumBytes,[int]$Seconds,[int]$Depth,$Budget,$State,$Children,$Issues) {
    $context=@{Path=$FilePath;Max=$MaximumBytes;Depth=$Depth;Budget=$Budget;State=$State}
    switch($Result.Format){
        'ActiveMime' {
            $Result.DetectionScope='ActiveMimeVBAStorage';$State.Stage='DecompressActiveMime'
            Add-LwEmbeddedInspection $Children $Issues 'ActiveMime' $Content $true $context
        }
        'XML' {
            $State.Stage='ParseOfficeXml';$xml=Read-LwSafeXml $Content;$root=$xml.DocumentElement
            if($root.LocalName -eq 'wordDocument' -and $root.NamespaceURI -eq 'http://schemas.microsoft.com/office/word/2003/wordml'){
                $Result.Format='Word2003XML';$Result.DetectionScope='Word2003XmlEmbeddedVBA';$Result.FormatDetails='Word 2003 XML; recognized binary VBA/ActiveMime parts are inspected without opening Word.'
                $nodes=$xml.GetElementsByTagName('binData',$root.NamespaceURI)
                foreach($node in $nodes){
                    $State.Stage='DecodeWord2003BinaryData';$name=$node.GetAttribute('name',$root.NamespaceURI);$State.StreamPath='Word2003XML/'+$name
                    try{
                        $bytes=ConvertFrom-LwBase64 $node.InnerText $Budget $State
                        if($node.ParentNode.LocalName -in @('docSuppData','docOleData') -or $name -match '(?i)\.(mso|bin|doc)$' -or ($bytes.Length -ge 10 -and [Text.Encoding]::ASCII.GetString($bytes,0,10) -eq 'ActiveMime')){
                            $Result.EmbeddedObjects++;Add-LwEmbeddedInspection $Children $Issues $State.StreamPath $bytes ($node.ParentNode.LocalName -eq 'docSuppData') $context
                        }elseif($name -notmatch '(?i)\.(png|jpe?g|gif|bmp|wmf|emf|tiff?)$'){throw ('Unrecognized Word XML binary payload: '+$name)}
                    }catch{$Issues.Add((New-LwDiagnostic $_ $State 'WordXmlPartError'))}
                }
            }elseif($root.LocalName -eq 'package' -and $root.NamespaceURI -eq 'http://schemas.microsoft.com/office/2006/xmlPackage'){
                $Result.Format='WordFlatOPC';$Result.DetectionScope='FlatOpcWordVBAAndEmbeddings';$Result.FormatDetails='Word OpenXML package serialized as XML; VBA and recognized embedded parts are decoded in memory.'
                $ns=$root.NamespaceURI;$parts=@($root.ChildNodes|Where-Object{$_ -is [Xml.XmlElement] -and $_.LocalName -eq 'part' -and $_.NamespaceURI -eq $ns});$names=@{}
                if($parts.Count -gt 10000){throw 'Flat OPC part count limit exceeded.'}
                foreach($part in $parts){$name=$part.GetAttribute('name',$ns);if(-not $name -or $names.ContainsKey($name)){throw 'Duplicate/empty Flat OPC part name.'};$names[$name]=$part}
                if(-not $names.ContainsKey('/word/document.xml')){throw 'Flat OPC does not contain a Word document part.'}
                foreach($part in $parts){
                    $name=$part.GetAttribute('name',$ns);$type=$part.GetAttribute('contentType',$ns);$vba=($type -match '(?i)vbaProject$' -or $name -match '(?i)/vbaProject\.bin$')
                    if($vba -or $name -match '(?i)^/word/embeddings/'){
                        $State.Stage='DecodeFlatOpcPart';$State.StreamPath=$name
                        try{
                            $payload=@($part.ChildNodes|Where-Object{$_ -is [Xml.XmlElement] -and $_.LocalName -eq 'binaryData' -and $_.NamespaceURI -eq $ns})
                            if($payload.Count -ne 1){throw 'Missing/ambiguous Flat OPC binaryData.'}
                            $bytes=ConvertFrom-LwBase64 $payload[0].InnerText $Budget $State
                            if(-not $vba){$Result.EmbeddedObjects++};Add-LwEmbeddedInspection $Children $Issues $name $bytes $vba $context
                        }catch{$Issues.Add((New-LwDiagnostic $_ $State 'FlatOpcPartError'))}
                    }
                }
            }else{$Result.Status='Unsupported';throw ('XML root is not a supported Word container: {'+$root.NamespaceURI+'}'+$root.LocalName)}
        }
        'HTML' {
            $Result.DetectionScope='HtmlNativeAndOfficeReferenceIndicators';$Result.FormatDetails='HTML has no native VBA project. Office binary sidecars and active objects require separate content.'
            $State.Stage='InspectHtmlReferences';$text=[Text.Encoding]::GetEncoding(28591).GetString($Content)
            if($Content.Length -ge 2 -and (($Content[0] -eq 255 -and $Content[1] -eq 254) -or ($Content[0] -eq 254 -and $Content[1] -eq 255))){$text=ConvertFrom-LwText $Content}
            $refs=Get-LwHtmlReferences $text
            if($refs.Count){throw ('HTML references content requiring separate inspection: '+($refs -join '; ')+'. Sidecars/URLs are not downloaded or opened.')}
        }
        'TextCandidate' {
            $State.Stage='ValidatePlainText';$text=ConvertFrom-LwText $Content
            if($text -match '[\x00-\x08\x0b\x0e-\x1f]'){throw 'Binary control bytes prevent classification as plain text.'}
            $Result.Format='PlainText';$Result.DetectionScope='PlainTextNoOfficeProject';$Result.FormatDetails='Text content, not a Word compound/ZIP container; no native Office VBA storage.'
            if($text -match '(?im)^\s*<|^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function)\s+\w+\s*\(|<script\b|<html\b|<\?xml\b'){
                $Result.Status='Unsupported';throw 'Text resembles source code or markup; separate script/content analysis is required.'
            }
        }
        'MHTML' {
            $Result.DetectionScope='MhtmlContainedOfficeVBA';$Result.FormatDetails='Word MIME web archive; recognized Office parts are decoded locally. External resources are not fetched.'
            $State.Stage='ParseMime';$text=[Text.Encoding]::GetEncoding(28591).GetString($Content);$parts=Read-LwMimeParts $text $State
            $locations=@{};$htmlParts=[Collections.Generic.List[object]]::new();$wordPart=$false
            foreach($part in $parts){
                $location=[string]$part.Headers['Content-Location'];if(-not $location){$location=[string]$part.Headers['Content-ID']}
                if($location){if($locations.ContainsKey($location)){throw 'Ambiguous MIME content location.'};$locations[$location]=$true}
            }
            $i=0
            foreach($part in $parts){
                $i++;$location=[string]$part.Headers['Content-Location'];if(-not $location){$location='MIME/part['+$i+']'};$State.StreamPath=$location;$State.Stage='DecodeMimePart'
                try{
                    $bytes=ConvertFrom-LwMimeBody $part $Budget $State;$type=[string]$part.Headers['Content-Type']
                    $prefix=[byte[]]::new([math]::Min(512,$bytes.Length));if($prefix.Length){[Buffer]::BlockCopy($bytes,0,$prefix,0,$prefix.Length)};$sig=Get-LwNonOleFormat $prefix
                    if($sig -eq 'ActiveMime' -or $type -match '(?i)application/(?:x-mso|vnd.ms-office.vbaProject)' -or $location -match '(?i)\.(?:mso|doc|docm|dotm)$'){
                        $Result.EmbeddedObjects++;Add-LwEmbeddedInspection $Children $Issues $location $bytes ($type -match '(?i)vbaProject') $context
                    }elseif($type -match '(?i)^text/html'){
                        $html=[Text.Encoding]::GetEncoding(28591).GetString($bytes);if($html -match '(?i)Word\.Document|schemas-microsoft-com:office:word'){$wordPart=$true}
                        $htmlParts.Add([pscustomobject]@{Location=$location;References=(Get-LwHtmlReferences $html)})
                    }elseif($type -match '(?i)^(image/|text/(?:xml|css)|application/vnd\.ms-officetheme)'){
                        # Images, style/theme and Office file-list metadata have no native VBA project in this scope.
                    }else{throw ('Unrecognized MIME part content type: '+$type)}
                }catch{$Issues.Add((New-LwDiagnostic $_ $State 'MimePartError'))}
            }
            if(-not $wordPart){throw 'MIME archive does not identify a Word HTML document.'}
            foreach($html in $htmlParts){foreach($reference in $html.References){
                $resolved=$reference
                try{$base=[uri]$html.Location;if($base.IsAbsoluteUri){$resolved=([uri]::new($base,$reference)).AbsoluteUri}}catch{}
                if(-not $locations.ContainsKey($resolved) -and -not $locations.ContainsKey($reference)){$State.StreamPath=$html.Location;try{throw ('Unresolved HTML reference in MIME archive: '+$reference+'. External content is not fetched.')}catch{$Issues.Add((New-LwDiagnostic $_ $State 'UnresolvedMimeReference'))}}
            }}
        }
    }
}
function Invoke-LwInspection([string]$FilePath,[long]$MaximumBytes=104857600,[int]$Seconds=120,[byte[]]$Content,[int]$Depth=0,$Budget) {
    if(-not $Budget){$Budget=@{Bytes=0L;Objects=0;Clock=[Diagnostics.Stopwatch]::StartNew();Timeout=$Seconds}}
    $args=@{FilePath=$FilePath;MaximumBytes=$MaximumBytes;Seconds=[math]::Max(1,($Budget.Timeout-[int]$Budget.Clock.Elapsed.TotalSeconds))}
    if($PSBoundParameters.ContainsKey('Content')){$args.Content=$Content}
    $result=Get-LwCompoundInspection @args
    if($result.Format -notin @('RTF','ZIPContainer','XML','HTML','MHTML','ActiveMime','TextCandidate','NotOLE','PDF','EmptyFile')){return $result}
    $result.DetectionScope='IdentificationOnly'
    if($result.Format -notin @('RTF','ZIPContainer','XML','HTML','MHTML','ActiveMime','TextCandidate')){
        $result.FormatDetails='Signature identifies '+$result.Format+'. This format is not inspected for embedded VBA; Unknown does not mean absent.'
        $result.Notes='Format identification only. Embedded/encoded VBA and external content in this format are not inspected or fetched. Unknown is not a negative or safety finding.'
        return $result
    }
    $issues=[Collections.Generic.List[object]]::new();$children=[Collections.Generic.List[object]]::new()
    $state=@{Clock=$Budget.Clock;Timeout=$Budget.Timeout;Stage='InspectContainer';ProjectPath='';StreamPath='';ModuleName='';FileOffset=$null;CompressedOffset=$null}
    $result.Status='Unknown';$result.Error='';$result.ErrorCode='';$result.Diagnostics=@();$result.ProjectCount=$null
    $complete=$false;$container=$result.Format
    try{
        Assert-LwTime $state
        if($Depth -gt 3){throw 'Embedded nesting limit exceeded (3 levels).'}
        if($result.FileLength -gt $MaximumBytes -or $result.FileLength -gt 33554432){throw 'Container exceeds bounded inspection size limit (32 MiB).'}
        if(-not $PSBoundParameters.ContainsKey('Content')){$Content=[LegacyWordPathV3101]::ReadAllBytes($FilePath,[math]::Min($MaximumBytes,33554432))}
        if($container -in @('XML','HTML','MHTML','ActiveMime','TextCandidate')){
            Invoke-LwAdditionalContainer $result $Content $FilePath $MaximumBytes $Seconds $Depth $Budget $state $children $issues
        }elseif($container -eq 'RTF'){
            $result.DetectionScope='RTFNativeAndRecognizedOLEObjects';$result.FormatDetails='RTF has no native VBA project; inspect recognized embedded OLE Word documents.'
            $state.Stage='ParseRtfObjects';$objects=Get-LwRtfObjects $Content $state;$result.EmbeddedObjects=$objects.Count
            $i=0
            foreach($object in $objects){
                $i++;$state.StreamPath='RTF/objdata['+$i+']'
                try{
                    if((++$Budget.Objects) -gt 32){throw 'Embedded object count limit exceeded.'}
                    $native=Get-LwOle1Native $object;$Budget.Bytes+=$native.Length;if($Budget.Bytes -gt 67108864){throw 'Aggregate embedded allocation limit exceeded.'}
                    $child=Get-LegacyWordInspection -FilePath $FilePath -MaximumBytes $MaximumBytes -Seconds $Seconds -Content $native -Depth ($Depth+1) -Budget $Budget
                    $children.Add([pscustomobject]@{Location=$state.StreamPath;Result=$child})
                }catch{$issues.Add((New-LwDiagnostic $_ $state 'EmbeddedObjectError'))}
            }
        }else{
            Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
            $ms=[IO.MemoryStream]::new($Content,$false);$zip=$null
            try{
                $zip=[IO.Compression.ZipArchive]::new($ms,[IO.Compression.ZipArchiveMode]::Read,$false);$names=@{}
                if($zip.Entries.Count -gt 10000){throw 'ZIP entry count limit exceeded.'}
                foreach($entry in $zip.Entries){Assert-LwTime $state;if($names.ContainsKey($entry.FullName)){throw 'ZIP has duplicate/ambiguous entry names.'};$names[$entry.FullName]=$entry}
                if(-not $names.ContainsKey('[Content_Types].xml') -or -not $names.ContainsKey('word/document.xml')){throw 'ZIP is not a recognized Word OpenXML document package.'}
                $result.Format='WordOpenXML';$result.DetectionScope='WordOpenXMLVBAAndRecognizedEmbeddings';$result.FormatDetails='ZIP-based Word document despite its filename extension. Standard VBA parts and word/embeddings are inspected.'
                $types=Read-LwSafeXml (Read-LwZipEntry $names['[Content_Types].xml'] $Budget $state)
                if($types.DocumentElement.LocalName -ne 'Types' -or $types.DocumentElement.NamespaceURI -ne 'http://schemas.openxmlformats.org/package/2006/content-types'){throw 'Invalid OpenXML content-types manifest.'}
                $vbaParts=@{};$defaults=@{};$overrides=@{}
                foreach($node in $types.DocumentElement.ChildNodes){
                    if($node -isnot [Xml.XmlElement]){continue}
                    if($node.LocalName -eq 'Default'){$defaults[$node.GetAttribute('Extension')]=$node.GetAttribute('ContentType')}
                    elseif($node.LocalName -eq 'Override'){
                        $part=[uri]::UnescapeDataString($node.GetAttribute('PartName').TrimStart('/'))
                        $overrides[$part]=$node.GetAttribute('ContentType')
                        if($overrides[$part] -match '(?i)vbaProject$' -and -not $names.ContainsKey($part)){throw ('Declared VBA part is missing: '+$part)}
                    }
                }
                foreach($entry in $zip.Entries){$type=$defaults[[IO.Path]::GetExtension($entry.FullName).TrimStart('.')];if($overrides.ContainsKey($entry.FullName)){$type=$overrides[$entry.FullName]};if($type -match '(?i)vbaProject$'){$vbaParts[$entry.FullName]=$true}}
                $targets=@($zip.Entries|Where-Object {$vbaParts.ContainsKey($_.FullName) -or $_.FullName -match '(?i)(^|/)vbaProject\.bin$|^word/embeddings/[^/]+$'})
                foreach($entry in $targets){
                    $state.Stage='InspectOpenXmlPart';$state.StreamPath=$entry.FullName
                    try{
                        if((++$Budget.Objects) -gt 32){throw 'Embedded object count limit exceeded.'}
                        $bytes=Read-LwZipEntry $entry $Budget $state
                        if($vbaParts.ContainsKey($entry.FullName) -or $entry.FullName -match '(?i)(^|/)vbaProject\.bin$'){
                            $child=Get-LwCompoundInspection -FilePath $FilePath -MaximumBytes $MaximumBytes -Seconds ([math]::Max(1,($Budget.Timeout-[int]$Budget.Clock.Elapsed.TotalSeconds))) -Content $bytes -VbaContainer
                        }else{
                            $result.EmbeddedObjects++
                            $child=Get-LegacyWordInspection -FilePath $FilePath -MaximumBytes $MaximumBytes -Seconds $Seconds -Content $bytes -Depth ($Depth+1) -Budget $Budget
                        }
                        $children.Add([pscustomobject]@{Location=$entry.FullName;Result=$child})
                    }catch{$issues.Add((New-LwDiagnostic $_ $state 'OpenXmlPartError'))}
                }
            }finally{if($zip){$zip.Dispose()};$ms.Dispose()}
        }
        $complete=$true
    }catch{$issues.Add((New-LwDiagnostic $_ $state 'ContainerInspectionError'))}
    $projects=0;$sources=0;$parsed=0;$positive=$false;$countKnown=$complete
    foreach($item in $children){
        $c=$item.Result
        if($c.VBA -eq 'Detected'){$positive=$true}
        if($null -ne $c.ProjectCount){$projects+=$c.ProjectCount}else{$countKnown=$false}
        $parsed+=$c.SourceModulesParsed;if($null -ne $c.SourceModules){$sources+=$c.SourceModules}
        $result.ProjectsInspected+=$c.ProjectsInspected
        foreach($issue in $c.Diagnostics){$issue.StreamPath=$item.Location+'!/'+$issue.StreamPath;$issues.Add($issue)}
        if($c.Status -ne 'Scanned'){$countKnown=$false}
    }
    $result.SourceModulesParsed=$parsed
    if($countKnown -and -not $issues.Count){$result.ProjectCount=$projects;$result.SourceModules=$sources;$result.Status='Scanned';$result.VBA='NotDetected';$result.Macro='NoVBAIndicatorsDetected'}
    elseif($positive){$result.ProjectCount=$projects;$result.Status='Partial'}else{$result.ProjectCount=$null;if($result.Status -ne 'Unsupported'){$result.Status='Unknown'}}
    if($positive){$result.VBA='Detected';$result.Macro='VBAIndicatorsDetected'}
    $result.Notes='Static inspection only within DetectionScope. External links/templates, WordBasic, packed packages, VBA stomping and application-specific parsing tricks are not exhaustively assessed. NotDetected is not a safety verdict. Partial counts are lower bounds.'
    $result.Diagnostics=$issues.ToArray();$result.Error=($issues|ForEach-Object {'['+$_.Stage+'] '+$_.Message}) -join ' | '
    $result.ErrorCode=(@($issues|ForEach-Object {$_.Code}|Select-Object -Unique) -join '; ');$result.ElapsedMilliseconds=$Budget.Clock.ElapsedMilliseconds
    return $result
}
function Get-LegacyWordInspection([string]$FilePath,[long]$MaximumBytes=104857600,[int]$Seconds=120,[byte[]]$Content,[int]$Depth=0,$Budget) {
    # This is the per-document exception boundary, including failures before the
    # container-specific parser's try block. Never abort a library because one
    # document has an unfamiliar or malformed format.
    $clock=[Diagnostics.Stopwatch]::StartNew()
    try{return Invoke-LwInspection @PSBoundParameters}
    catch{
        $state=@{Stage='InspectionBoundary';ProjectPath='';StreamPath='';ModuleName='';FileOffset=$null;CompressedOffset=$null}
        $diagnostic=New-LwDiagnostic $_ $state 'UnhandledDocumentInspectionError'
        $identity=$FilePath;try{$identity=(ConvertTo-LwFullPath $FilePath)}catch{}
        return [pscustomobject][ordered]@{SchemaVersion='2.0';Status='Unknown';Format='Unidentified';VBA='Unknown';Macro='Unknown';SourceModules=$null;SourceModulesParsed=0;ProjectCount=$null;ProjectsInspected=0;Scanner='PowerShell-CFB-OVBA-1.4.1';FilePath=$identity;FileLength=$null;SignatureHex='';DetectionScope='Incomplete';EmbeddedObjects=0;FormatDetails='Document inspection could not complete; see diagnostics.';FIBVersion=$null;EffectiveFIBVersion=$null;IsTemplate=$null;ElapsedMilliseconds=$clock.ElapsedMilliseconds;Error=$diagnostic.Message;ErrorCode=$diagnostic.Code;Diagnostics=@($diagnostic);Notes='No negative VBA or safety finding can be made from an incomplete inspection.'}
    }
}
if($PSBoundParameters.ContainsKey('Path')){
    $ErrorActionPreference='Stop';Set-StrictMode -Version 2
    $result=Get-LegacyWordInspection $Path ([long]$MaxFileMB*1MB) $TimeoutSeconds
    if($ResultPath){
        # Never overwrite an existing path, including the input document by mistake.
        $output=[LegacyWordPathV3101]::CreateNew($ResultPath)
        try{$writer=[IO.StreamWriter]::new($output,[Text.UTF8Encoding]::new($false));try{$writer.Write(($result|ConvertTo-Json -Depth 12))}finally{$writer.Dispose()}}finally{$output.Dispose()}
    }else{$result}
}
