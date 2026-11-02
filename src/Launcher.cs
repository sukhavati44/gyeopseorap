using System;
using System.IO;
using System.Windows.Forms;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using System.Reflection;
[assembly: AssemblyTitle("겹서랍")]
[assembly: AssemblyProduct("겹서랍")]
[assembly: AssemblyDescription("수행평가 다층 작업기억 관리 프로그램")]
[assembly: AssemblyVersion("0.0.2.0")]
[assembly: AssemblyFileVersion("0.0.2.0")]
class Launcher {
 [STAThread] static void Main() {
  try {
   string root=AppDomain.CurrentDomain.BaseDirectory;
   using(Runspace rs=RunspaceFactory.CreateRunspace()) {
    rs.ApartmentState=ApartmentState.STA;rs.ThreadOptions=PSThreadOptions.UseCurrentThread;rs.Open();
    using(PowerShell ps=PowerShell.Create()) {
     ps.Runspace=rs;ps.AddCommand(Path.Combine(root,"gui.ps1"));ps.Invoke();
     if(ps.HadErrors) { string errors=String.Join(Environment.NewLine,ps.Streams.Error); File.WriteAllText(Path.Combine(root,"evidence","launcher-error.txt"),errors);MessageBox.Show(errors,"겹서랍 0.0.2 · 실행 오류"); }
    }
   }
  } catch(Exception ex) { MessageBox.Show(ex.ToString(),"겹서랍 0.0.2 · 실행 오류"); }
 }
}
