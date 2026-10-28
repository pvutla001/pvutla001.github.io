using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Net;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using Microsoft.SharePoint.Client;

namespace SPOInventoryV310 {
    public static class BuildInfo { public const string Version = "3.10.1"; }
    // One shared gate per run; every worker respects service-wide Retry-After.
    public sealed class RetryGate {
        private readonly object sync = new object();
        private DateTime until = DateTime.MinValue;
        public long Retries;
        public long Requests;
        public void Wait() {
            while (true) {
                double ms; lock(sync) { ms = (until-DateTime.UtcNow).TotalMilliseconds; }
                if (ms <= 0) return;
                Thread.Sleep((int)Math.Min(1000, Math.Ceiling(ms)));
            }
        }
        public double RemainingSeconds { get { lock(sync) { return Math.Max(0,(until-DateTime.UtcNow).TotalSeconds); } } }
        public bool Backoff(Exception error, int attempt) {
            WebException web = null;
            for(Exception e=error; e!=null; e=e.InnerException) { if(e is WebException) { web=(WebException)e; break; } }
            if(web == null) return false;
            HttpWebResponse response = web.Response as HttpWebResponse;
            int code = response == null ? 0 : (int)response.StatusCode;
            if(code != 429 && code != 503 && web.Status != WebExceptionStatus.Timeout && web.Status != WebExceptionStatus.ConnectionClosed && web.Status != WebExceptionStatus.ReceiveFailure && web.Status != WebExceptionStatus.ConnectFailure) return false;
            double seconds = Math.Min(120,Math.Pow(2,Math.Min(attempt+1,7)));
            if(response != null) {
                string header=response.Headers["Retry-After"]; double number; DateTimeOffset date;
                if(Double.TryParse(header, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out number)) seconds=Math.Max(seconds,number);
                else if(DateTimeOffset.TryParse(header,out date)) seconds=Math.Max(seconds,(date.UtcDateTime-DateTime.UtcNow).TotalSeconds);
            }
            seconds += new Random(Guid.NewGuid().GetHashCode()).NextDouble();
            lock(sync) { DateTime next=DateTime.UtcNow.AddSeconds(seconds); if(next>until) until=next; }
            Interlocked.Increment(ref Retries); return true;
        }
        public void ExecuteRead(ClientContext context, int maxRetries) {
            ClientRequest retry = null;
            for(int attempt=0;;attempt++) {
                Wait(); Interlocked.Increment(ref Requests);
                try { if(retry==null) context.ExecuteQuery(); else context.RetryQuery(retry); return; }
                catch(Exception error) {
                    ClientRequestWrapper wrapper=null;
                    for(Exception e=error; e!=null; e=e.InnerException) {
                        if(e.Data.Contains("ClientRequest")) wrapper=e.Data["ClientRequest"] as ClientRequestWrapper;
                    }
                    if(attempt>=maxRetries || wrapper==null || wrapper.Value==null || !Backoff(error,attempt)) throw;
                    retry=wrapper.Value; // replay the original CSOM request, not an empty ExecuteQuery
                }
            }
        }
    }

    public static class LocalIO {
        // Use Unicode Win32 paths and handle-based streams. Passing a device path
        // to a .NET Framework string-path API can fail in legacy compatibility mode.
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
        private static extern uint GetFullPathNameW(string path, uint size, StringBuilder buffer, IntPtr part);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
        private static extern uint GetFileAttributesW(string path);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
        private static extern bool CreateDirectoryW(string path, IntPtr security);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
        private static extern bool DeleteFileW(string path);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
        private static extern bool MoveFileW(string source, string destination);
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
        public static string Parent(string path) {
            string full=FullPath(path); int root=RootLength(full);
            if(full.Length<=root) return null;
            int slash=full.LastIndexOf('\\'); return slash<root ? full.Substring(0,root) : full.Substring(0,slash);
        }
        private static uint Attributes(string path, bool allowMissing) {
            uint value=GetFileAttributesW(Extended(path));
            if(value!=UInt32.MaxValue) return value;
            int code=Marshal.GetLastWin32Error();
            if(allowMissing && (code==2 || code==3)) return UInt32.MaxValue;
            throw Failure("Read attributes",path,code);
        }
        public static bool FileExists(string path) {
            uint value=Attributes(path,true);
            return value!=UInt32.MaxValue && (value & (uint)FileAttributes.Directory)==0;
        }
        private static void CheckDirectory(string path, uint attributes) {
            if((attributes & (uint)FileAttributes.Directory)==0) throw new IOException("Directory path is occupied by a file: "+path);
            if((attributes & (uint)FileAttributes.ReparsePoint)!=0) throw new IOException("Directory link/reparse point refused: "+path);
        }
        public static void CreateDirectory(string path) {
            string full=FullPath(path); int root=RootLength(full);
            string cursor=full.Substring(0,root); CheckDirectory(cursor,Attributes(cursor,false));
            foreach(string part in full.Substring(root).Split(new char[]{'\\'},StringSplitOptions.RemoveEmptyEntries)) {
                cursor=cursor.TrimEnd('\\')+@"\"+part;
                uint attributes=Attributes(cursor,true);
                if(attributes==UInt32.MaxValue) {
                    if(!CreateDirectoryW(Extended(cursor),IntPtr.Zero)) {
                        int code=Marshal.GetLastWin32Error(); if(code!=183) throw Failure("Create directory",cursor,code);
                    }
                    attributes=Attributes(cursor,false);
                }
                CheckDirectory(cursor,attributes);
            }
        }
        private static FileStream Open(string path, bool write) {
            SafeFileHandle handle=CreateFileW(Extended(path),write ? 0x40000000U : 0x80000000U,write ? 0U : 1U,IntPtr.Zero,write ? 1U : 3U,0x08000080U,IntPtr.Zero);
            if(handle.IsInvalid) { int code=Marshal.GetLastWin32Error(); handle.Dispose(); throw Failure(write ? "Create new file" : "Open file for reading",path,code); }
            try { return new FileStream(handle,write ? FileAccess.Write : FileAccess.Read,81920,false); }
            catch { handle.Dispose(); throw; }
        }
        public static FileStream OpenRead(string path) { return Open(path,false); }
        public static FileStream CreateNew(string path) { return Open(path,true); }
        public static long Length(string path) { using(var stream=OpenRead(path)) return stream.Length; }
        public static void Delete(string path) {
            if(!DeleteFileW(Extended(path))) { int code=Marshal.GetLastWin32Error(); if(code!=2 && code!=3) throw Failure("Delete file",path,code); }
        }
        public static void Move(string source,string destination) {
            if(!MoveFileW(Extended(source),Extended(destination))) throw Failure("Move file to "+destination,source,Marshal.GetLastWin32Error());
        }
        private static bool IsReservedDeviceName(string component) {
            string stem=component;
            int dot=stem.IndexOf('.');
            if(dot>=0) stem=stem.Substring(0,dot);
            return System.Text.RegularExpressions.Regex.IsMatch(stem,@"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$",System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }
        private static string Hex4(char value) { return ((int)value).ToString("X4",System.Globalization.CultureInfo.InvariantCulture); }
        private static string Hash12(string value) {
            using(var sha=System.Security.Cryptography.SHA256.Create()) {
                byte[] digest=sha.ComputeHash(Encoding.UTF8.GetBytes(value));
                var text=new StringBuilder(12);
                for(int n=0;n<6;n++) text.Append(digest[n].ToString("x2",System.Globalization.CultureInfo.InvariantCulture));
                return text.ToString();
            }
        }
        private static string SafeComponent(string original, out bool changed) {
            if(String.IsNullOrEmpty(original) || original=="." || original=="..") throw new IOException("Empty or relative path component cannot be mirrored.");
            string normalized=original.Normalize(NormalizationForm.FormC);
            changed=!String.Equals(original,normalized,StringComparison.Ordinal);
            char[] invalid=System.IO.Path.GetInvalidFileNameChars();
            var result=new StringBuilder(normalized.Length+16);
            for(int n=0;n<normalized.Length;n++) {
                char c=normalized[n];
                bool trailing=(n==normalized.Length-1 && (c=='.' || c==' '));
                bool unsafeChar=c=='~' || c<32 || Array.IndexOf(invalid,c)>=0 || trailing;
                if(unsafeChar) { result.Append('~').Append(Hex4(c)); changed=true; }
                else result.Append(c);
            }
            string safe=result.ToString();
            if(IsReservedDeviceName(safe)) { safe="~005F"+safe; changed=true; }
            // Windows file-system components are normally limited to 255 UTF-16 code units.
            // Keep margin for the deterministic suffix and avoid splitting a surrogate pair.
            if(safe.Length>240) {
                string suffix="~h"+Hash12(original);
                int keep=240-suffix.Length;
                if(keep>0 && keep<safe.Length && Char.IsHighSurrogate(safe[keep-1])) keep--;
                safe=safe.Substring(0,keep)+suffix;
                changed=true;
            }
            if(String.IsNullOrEmpty(safe)) throw new IOException("Path component became empty after Windows-safe mapping.");
            return safe;
        }
        public static bool RequiresMapping(string relative) {
            if(relative==null) throw new ArgumentNullException("relative");
            bool any=false;
            foreach(string part in relative.Split('/')) { bool changed; SafeComponent(part,out changed); any|=changed; }
            return any;
        }
        public static string Mirror(string root, string relative) {
            if(String.IsNullOrWhiteSpace(root)) throw new ArgumentException("Download root is empty.","root");
            if(relative==null) throw new ArgumentNullException("relative");
            root=FullPath(root).TrimEnd('\\');
            string target=root;
            foreach(string part in relative.Split('/')) {
                bool changed;
                string safe=SafeComponent(part,out changed);
                target=target+@"\"+safe;
            }
            // Check every existing component, including the root's parents and target.
            string cursor=target;
            while(!String.IsNullOrEmpty(cursor)) {
                uint attributes=Attributes(cursor,true);
                if(attributes!=UInt32.MaxValue && (attributes & (uint)FileAttributes.ReparsePoint)!=0) throw new IOException("Directory link/reparse point refused: "+cursor);
                cursor=Parent(cursor);
            }
            return target;
        }
        public static void CopyLimited(Stream input, string destination, long maximum) {
            byte[] buffer=new byte[81920]; long total=0;
            using(var output=CreateNew(destination)) {
                int read; while((read=input.Read(buffer,0,buffer.Length))>0) { total+=read; if(total>maximum) throw new IOException("Downloaded bytes exceed MaxScanMB."); output.Write(buffer,0,read); }
            }
        }
        public static string Hash(string path) {
            using(var hash=System.Security.Cryptography.SHA256.Create()) using(var stream=OpenRead(path)) return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-","");
        }
    }

    // Spool worksheet XML to disk, then ZIP it. No EPPlus workbook/shared-string
    // table or complete inventory is held in RAM. Inline strings never formulas.
    public sealed class StreamingWorkbook : IDisposable {
        private const string ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private sealed class Sheet { public string Name,Path; public XmlWriter Xml; public long Rows; public int Columns; }
        private readonly List<Sheet> sheets=new List<Sheet>();
        private readonly string destination;
        private readonly string spool;
        private bool completed;
        public StreamingWorkbook(string path) { destination=path; spool=path+".spool_"+Guid.NewGuid().ToString("N"); Directory.CreateDirectory(spool); }
        private static string Column(int n) { string s=""; while(n>0) { n--; s=(char)('A'+n%26)+s; n/=26; } return s; }
        private static string Safe(string text) {
            if(text==null) return "";
            if(text.Length>32767) text=text.Substring(0,32720)+" [TRUNCATED: see JSONL]";
            var s=new StringBuilder(text.Length);
            for(int i=0;i<text.Length;i++) { char c=text[i]; if(XmlConvert.IsXmlChar(c)) s.Append(c); else if(Char.IsHighSurrogate(c) && i+1<text.Length && Char.IsLowSurrogate(text[i+1])) { s.Append(c); s.Append(text[++i]); } }
            return s.ToString();
        }
        public int AddSheet(string name,string[] headers) {
            var sheet=new Sheet {Name=name,Path=System.IO.Path.Combine(spool,"sheet"+(sheets.Count+1)+".xml"),Columns=headers.Length};
            sheet.Xml=XmlWriter.Create(sheet.Path,new XmlWriterSettings {Encoding=new UTF8Encoding(false),CloseOutput=true});
            var x=sheet.Xml; x.WriteStartDocument(); x.WriteStartElement("worksheet",ns);
            x.WriteStartElement("sheetViews"); x.WriteStartElement("sheetView"); x.WriteAttributeString("workbookViewId","0");
            x.WriteStartElement("pane"); x.WriteAttributeString("ySplit","1"); x.WriteAttributeString("topLeftCell","A2"); x.WriteAttributeString("state","frozen"); x.WriteEndElement(); x.WriteEndElement(); x.WriteEndElement();
            x.WriteStartElement("cols");
            for(int i=0;i<headers.Length;i++) {
                double width=24;
                if(headers[i].IndexOf("Path",StringComparison.OrdinalIgnoreCase)>=0 || headers[i]=="Url") width=58;
                else if(headers[i]=="Message" || headers[i]=="Error" || headers[i]=="Notes" || headers[i]=="Value") width=72;
                else if(headers[i]=="FileName" || headers[i]=="Property") width=38;
                x.WriteStartElement("col"); x.WriteAttributeString("min",(i+1).ToString()); x.WriteAttributeString("max",(i+1).ToString()); x.WriteAttributeString("width",width.ToString(System.Globalization.CultureInfo.InvariantCulture)); x.WriteAttributeString("customWidth","1"); x.WriteEndElement();
            }
            x.WriteEndElement();
            x.WriteStartElement("sheetData"); sheets.Add(sheet); int index=sheets.Count-1; AddRow(index,headers); return index;
        }
        public void AddRow(int index,string[] values) {
            var s=sheets[index]; if(s.Rows>=1048576) throw new InvalidOperationException("Excel row limit reached; rotate the workbook.");
            s.Rows++; var x=s.Xml; x.WriteStartElement("row"); x.WriteAttributeString("r",s.Rows.ToString());
            if(s.Rows==1) { x.WriteAttributeString("ht","30"); x.WriteAttributeString("customHeight","1"); }
            for(int i=0;i<values.Length;i++) { x.WriteStartElement("c"); x.WriteAttributeString("r",Column(i+1)+s.Rows); x.WriteAttributeString("s",s.Rows==1?"1":(s.Rows%2==0?"2":"0")); x.WriteAttributeString("t","inlineStr"); x.WriteStartElement("is"); x.WriteStartElement("t"); x.WriteAttributeString("xml","space",null,"preserve"); x.WriteString(Safe(values[i])); x.WriteEndElement(); x.WriteEndElement(); x.WriteEndElement(); }
            x.WriteEndElement();
        }
        private static void Put(ZipArchive zip,string name,string text) { using(var writer=new StreamWriter(zip.CreateEntry(name,CompressionLevel.Fastest).Open(),new UTF8Encoding(false))) writer.Write(text); }
        public void Complete() {
            if(completed) return;
            foreach(var s in sheets) { var x=s.Xml; x.WriteEndElement(); if(s.Columns>0) { x.WriteStartElement("autoFilter"); x.WriteAttributeString("ref","A1:"+Column(s.Columns)+s.Rows); x.WriteEndElement(); } x.WriteEndElement(); x.WriteEndDocument(); x.Dispose(); s.Xml=null; }
            string temp=destination+".partial";
            using(var output=new FileStream(temp,FileMode.CreateNew)) using(var zip=new ZipArchive(output,ZipArchiveMode.Create)) {
                var types=new StringBuilder("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
                var workbook=new StringBuilder("<workbook xmlns=\""+ns+"\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
                var rels=new StringBuilder("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for(int i=0;i<sheets.Count;i++) { int n=i+1; var s=sheets[i]; types.Append("<Override PartName=\"/xl/worksheets/sheet"+n+".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"); workbook.Append("<sheet name=\""+System.Security.SecurityElement.Escape(s.Name)+"\" sheetId=\""+n+"\" r:id=\"rId"+n+"\"/>"); rels.Append("<Relationship Id=\"rId"+n+"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet"+n+".xml\"/>"); using(var input=System.IO.File.OpenRead(s.Path)) using(var target=zip.CreateEntry("xl/worksheets/sheet"+n+".xml",CompressionLevel.Fastest).Open()) input.CopyTo(target); }
                types.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
                rels.Append("<Relationship Id=\"rStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
                Put(zip,"xl/styles.xml","<styleSheet xmlns=\""+ns+"\"><fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><color rgb=\"FFFFFFFF\"/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"4\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF17365D\"/><bgColor indexed=\"64\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF0F4F8\"/><bgColor indexed=\"64\"/></patternFill></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"top\" wrapText=\"1\"/></xf><xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" wrapText=\"1\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyFill=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\" wrapText=\"1\"/></xf></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>");
                Put(zip,"[Content_Types].xml",types+"</Types>"); Put(zip,"xl/workbook.xml",workbook+"</sheets></workbook>"); Put(zip,"xl/_rels/workbook.xml.rels",rels+"</Relationships>");
                Put(zip,"_rels/.rels","<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            }
            System.IO.File.Move(temp,destination); completed=true;
            foreach(var s in sheets) System.IO.File.Delete(s.Path); Directory.Delete(spool,false);
        }
        public void Dispose() { foreach(var s in sheets) if(s.Xml!=null) { s.Xml.Dispose(); s.Xml=null; } }
    }
}
