using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using System.Linq;
using System.Collections.Generic;
public class ImportCell {public object Value;public string Address;public bool Formula;}
public class ImportSheet {public string Name;public Dictionary<int,Dictionary<int,ImportCell>> Rows=new Dictionary<int,Dictionary<int,ImportCell>>();}
public class ImportBook {public List<ImportSheet> Sheets=new List<ImportSheet>();public List<string> Findings=new List<string>();}
public static class NaturalXlsx {
 static XNamespace n="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
 static XDocument Doc(ZipArchive z,string path){var e=z.GetEntry(path);if(e==null)throw new Exception("Missing XLSX part: "+path);using(var s=e.Open())return XDocument.Load(s);}
 public static ImportBook Read(string path){
  var book=new ImportBook();using(var z=ZipFile.OpenRead(path)){
   if(z.Entries.Count>512||z.Entries.Sum(e=>e.Length)>32*1024*1024)throw new Exception("XLSX size/part limit");
   foreach(var e in z.Entries){if(e.FullName.Contains("..")||e.FullName.Contains("vbaProject")||e.FullName.Contains("externalLink"))throw new Exception("Unsupported active/external XLSX content");
    if(e.FullName.EndsWith(".xml")||e.FullName.EndsWith(".rels"))using(var sr=new StreamReader(e.Open())){string t=sr.ReadToEnd();
     if(e.FullName.EndsWith(".rels"))foreach(var rel in XDocument.Parse(t).Root.Elements())if((string)rel.Attribute("TargetMode")=="External"){
      if((string)rel.Attribute("Type")!="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")throw new Exception("External relationship refused");
      book.Findings.Add("하이퍼링크 미실행: "+e.FullName);
     }
     if(System.Text.RegularExpressions.Regex.IsMatch(t,@"(?:01[016789][- ]?\d{3,4}[- ]?\d{4}|\b\d{6}[- ]?[1-4]\d{6}\b|[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})"))book.Findings.Add("개인정보 패턴: "+e.FullName);
    }
   }
   var strings=new List<string>();if(z.GetEntry("xl/sharedStrings.xml")!=null)foreach(var s in Doc(z,"xl/sharedStrings.xml").Descendants(n+"si"))strings.Add(String.Concat(s.Descendants(n+"t").Select(t=>t.Value)));
   var rels=Doc(z,"xl/_rels/workbook.xml.rels").Root.Elements().ToDictionary(e=>(string)e.Attribute("Id"),e=>(string)e.Attribute("Target"));
   foreach(var s in Doc(z,"xl/workbook.xml").Descendants(n+"sheet")){
    var sheet=new ImportSheet{Name=(string)s.Attribute("name")};string rid=(string)s.Attribute(XName.Get("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships"));string target=rels[rid].Replace('\\','/');string part=target.StartsWith("/")?target.TrimStart('/'):("xl/"+target);int count=0;
    foreach(var c in Doc(z,part).Descendants(n+"c")){
     string a=(string)c.Attribute("r");var match=System.Text.RegularExpressions.Regex.Match(a??"",@"^([A-Z]+)(\d+)$");if(!match.Success)throw new Exception("Bad cell address");int col=0;foreach(char ch in match.Groups[1].Value)col=col*26+ch-'A'+1;int row=int.Parse(match.Groups[2].Value);if(row>20000||col>16384||++count>2000000)throw new Exception("Sheet bounds exceeded");
     string type=(string)c.Attribute("t"),v=(string)c.Element(n+"v");object value=null;
     if(type=="s"&&v!=null)value=strings[int.Parse(v)];else if(type=="inlineStr")value=String.Concat(c.Descendants(n+"t").Select(t=>t.Value));else if(type=="b")value=v=="1";else if(type=="str"||type=="e")value=v;else if(v!=null){double d;if(double.TryParse(v,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture,out d))value=d;else value=v;}
     if(!sheet.Rows.ContainsKey(row))sheet.Rows[row]=new Dictionary<int,ImportCell>();sheet.Rows[row][col]=new ImportCell{Address=a,Value=value,Formula=c.Element(n+"f")!=null};
    }book.Sheets.Add(sheet);
   }
  }return book;
 }
}
