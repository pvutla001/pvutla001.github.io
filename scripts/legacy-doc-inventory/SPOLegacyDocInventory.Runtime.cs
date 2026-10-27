using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Net;
using System.Threading;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace SPOInventoryV3 {
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
        public static string Extended(string path) {
            if(path.StartsWith(@"\\?\")) return path;
            if(path.StartsWith(@"\\")) return @"\\?\UNC\"+path.Substring(2);
            return @"\\?\"+path;
        }
        public static string Mirror(string root, string relative) {
            root=System.IO.Path.GetFullPath(root).TrimEnd('\\');
            string target=root;
            foreach(string part in relative.Split('/')) {
                if(String.IsNullOrEmpty(part) || part=="." || part==".." || part.IndexOfAny(System.IO.Path.GetInvalidFileNameChars())>=0 || part.EndsWith(".") || part.EndsWith(" ") || System.Text.RegularExpressions.Regex.IsMatch(part,@"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)",System.Text.RegularExpressions.RegexOptions.IgnoreCase)) throw new IOException("Cannot mirror Windows filename: "+part);
                target+="\\"+part;
            }
            // Check every existing component, including the root's parents and target.
            string cursor=target;
            string volume=System.IO.Path.GetPathRoot(root).TrimEnd('\\');
            while(!String.IsNullOrEmpty(cursor)) {
                try { if((System.IO.File.GetAttributes(Extended(cursor)) & FileAttributes.ReparsePoint)!=0) throw new IOException("Directory link/reparse point refused: "+cursor); }
                catch(FileNotFoundException) {} catch(DirectoryNotFoundException) {}
                if(String.Equals(cursor.TrimEnd('\\'),volume,StringComparison.OrdinalIgnoreCase)) break;
                int slash=cursor.LastIndexOf('\\'); if(slash<3) break; cursor=cursor.Substring(0,slash);
            }
            return target;
        }
        public static void CopyLimited(Stream input, string destination, long maximum) {
            byte[] buffer=new byte[81920]; long total=0;
            using(var output=new FileStream(Extended(destination),FileMode.CreateNew,FileAccess.Write,FileShare.None)) {
                int read; while((read=input.Read(buffer,0,buffer.Length))>0) { total+=read; if(total>maximum) throw new IOException("Downloaded bytes exceed MaxScanMB."); output.Write(buffer,0,read); }
            }
        }
        public static string Hash(string path) {
            using(var hash=System.Security.Cryptography.SHA256.Create()) using(var stream=System.IO.File.OpenRead(Extended(path))) return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-","");
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
            x.WriteStartElement("sheetData"); sheets.Add(sheet); int index=sheets.Count-1; AddRow(index,headers); return index;
        }
        public void AddRow(int index,string[] values) {
            var s=sheets[index]; if(s.Rows>=1048576) throw new InvalidOperationException("Excel row limit reached; rotate the workbook.");
            s.Rows++; var x=s.Xml; x.WriteStartElement("row"); x.WriteAttributeString("r",s.Rows.ToString());
            for(int i=0;i<values.Length;i++) { x.WriteStartElement("c"); x.WriteAttributeString("r",Column(i+1)+s.Rows); x.WriteAttributeString("t","inlineStr"); x.WriteStartElement("is"); x.WriteStartElement("t"); x.WriteAttributeString("xml","space",null,"preserve"); x.WriteString(Safe(values[i])); x.WriteEndElement(); x.WriteEndElement(); x.WriteEndElement(); }
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
                Put(zip,"[Content_Types].xml",types+"</Types>"); Put(zip,"xl/workbook.xml",workbook+"</sheets></workbook>"); Put(zip,"xl/_rels/workbook.xml.rels",rels+"</Relationships>");
                Put(zip,"_rels/.rels","<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            }
            System.IO.File.Move(temp,destination); completed=true;
            foreach(var s in sheets) System.IO.File.Delete(s.Path); Directory.Delete(spool,false);
        }
        public void Dispose() { foreach(var s in sheets) if(s.Xml!=null) { s.Xml.Dispose(); s.Xml=null; } }
    }
}
