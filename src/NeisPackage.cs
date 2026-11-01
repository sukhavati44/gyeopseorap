using System;using System.IO;using System.IO.Compression;using System.Collections;using System.Text;using System.Text.RegularExpressions;using System.Xml.Linq;using System.Linq;
public static class NeisPackage {
 public static void Write(string source,string output,Hashtable changes){
  Write(source,output,changes,null);
 }
 public static void Write(string source,string output,Hashtable changes,string filterRange){
  using(var z=ZipFile.OpenRead(source))using(var fs=new FileStream(output,FileMode.CreateNew))using(var dest=new ZipArchive(fs,ZipArchiveMode.Create)){
   foreach(var e in z.Entries){var n=dest.CreateEntry(e.FullName,CompressionLevel.Optimal);n.LastWriteTime=e.LastWriteTime;n.ExternalAttributes=e.ExternalAttributes;using(var input=e.Open())using(var o=n.Open()){
    if(e.FullName!="xl/worksheets/sheet1.xml"){input.CopyTo(o);continue;}
    string xml;using(var reader=new StreamReader(input,Encoding.UTF8,true,4096,true))xml=reader.ReadToEnd();
    foreach(DictionaryEntry kv in changes){string cell=(string)kv.Key,value=(string)kv.Value;var re=new Regex("<c\\b(?=[^>]*\\br=\""+Regex.Escape(cell)+"\")([^>]*?)(?:/>|>(.*?)</c>)",RegexOptions.Singleline);var matches=re.Matches(xml);if(matches.Count!=1)throw new Exception("Expected one score cell: "+cell);var m=matches[0];if(m.Value.Contains("<f"))throw new Exception("Formula score cell protected: "+cell);
     string attrs=m.Groups[1].Value;string content=m.Groups[2].Value;
     if(value=="" && content=="")continue;
     attrs=Regex.Replace(attrs,"\\s+t=\"[^\"]*\"","");string replacement="<c"+attrs+(value==""?"/>":"><v>"+value+"</v></c>");xml=xml.Substring(0,m.Index)+replacement+xml.Substring(m.Index+m.Length);
    }
    if(!String.IsNullOrEmpty(filterRange)){
     var doc=XDocument.Parse(xml,LoadOptions.PreserveWhitespace);XNamespace ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
     var filter=doc.Root.Element(ns+"autoFilter");if(filter==null){filter=new XElement(ns+"autoFilter");doc.Root.Element(ns+"sheetData").AddAfterSelf(filter);}filter.SetAttributeValue("ref",filterRange);xml=doc.ToString(SaveOptions.DisableFormatting);
    }
    var bytes=new UTF8Encoding(false).GetBytes(xml);o.Write(bytes,0,bytes.Length);
   }}
  }
 }
}
