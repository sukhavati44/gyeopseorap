Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$base=$PSScriptRoot;[void](New-Item -ItemType Directory -Path (Join-Path $base 'evidence') -Force);$script:ws=$null;$script:meta=$null;$script:current=$null;$script:busy=$false
. (Join-Path $base 'product-info.ps1')
. (Join-Path $base 'student-codes.ps1')
if(!('AssessmentGrid' -as [type])){Add-Type -Path (Join-Path $base 'GridNavigation.cs') -ReferencedAssemblies System.Windows.Forms,System.Drawing,System.Core}
$f=New-Object Windows.Forms.Form;$f.Text=$productTitle;$f.Size=New-Object Drawing.Size(1300,920);$f.StartPosition='CenterScreen';$f.Font=New-Object Drawing.Font('맑은 고딕',10)
function Label($hostControl,$text,$x,$y,$w,$h=25){$c=New-Object Windows.Forms.Label;$c.Text=$text;$c.SetBounds($x,$y,$w,$h);$hostControl.Controls.Add($c);return $c}
function Button($hostControl,$text,$x,$y,$w,$handler){$c=New-Object Windows.Forms.Button;$c.Text=$text;$c.SetBounds($x,$y,$w,36);$c|Add-Member -NotePropertyName ActionScript -NotePropertyValue $handler
 $c.Add_Click({
  $rootAction=($this.FindForm() -eq $f)
  if($rootAction -and ($busy -or $uiActionDepth -gt 0 -or $uiLoadingDepth -gt 0)){return}
  if($rootAction){$script:uiActionDepth++;Buttons-Enabled $f $false}
  try{& $this.ActionScript}finally{if($rootAction){$script:uiActionDepth=[math]::Max(0,$uiActionDepth-1);if(!$busy -and !$uiActionDepth -and !$uiLoadingDepth){$f.UseWaitCursor=$false;$f.Cursor=[Windows.Forms.Cursors]::Default;Buttons-Enabled $f $true};if(Get-Command NRefreshActions -ErrorAction SilentlyContinue){NRefreshActions}}}
 });$hostControl.Controls.Add($c);return $c}
function Combo($hostControl,$x,$y,$w){$c=New-Object Windows.Forms.ComboBox;$c.DropDownStyle='DropDownList';$c.DisplayMember='Text';$c.SetBounds($x,$y,$w,30);$hostControl.Controls.Add($c);return $c}
function Choice($c,$id,$text){[void]$c.Items.Add([pscustomobject]@{Id=$id;Text=$text})}
function Choose($c,$id){for($i=0;$i -lt $c.Items.Count;$i++){if($c.Items[$i].Id -eq $id){$c.SelectedIndex=$i;return}}}
function Grid($hostControl,$x,$y,$w,$h,$columns){
 $c=New-Object AssessmentGrid;$c.SetBounds($x,$y,$w,$h);$c.ReadOnly=$true;$c.AllowUserToAddRows=$false;$c.AllowUserToDeleteRows=$false;$c.RowHeadersVisible=$false;$c.AutoSizeColumnsMode='Fill';$c.SelectionMode='FullRowSelect';$c.BackgroundColor=[Drawing.Color]::White;$c.MultiSelect=$false
 foreach($col in $columns){[void]$c.Columns.Add($col,$col)};$hostControl.Controls.Add($c);return $c
}
function Map($o){if($null -eq $o){return $null};if($o -is [Management.Automation.PSCustomObject]){$h=@{};foreach($p in $o.PSObject.Properties){$h[$p.Name]=Map $p.Value};return $h};if($o -is [array]){return ,@($o|ForEach-Object {Map $_})};return $o}
function Read-Meta{$script:meta=Map (Get-Content -Raw -Encoding utf8 (Join-Path $ws 'relations.json')|ConvertFrom-Json)}
function LName($id){if(!$id){return '상위 없음'};$name=$meta.layers[$id].name;if(!$name){$name='이름 없는 표'};return $name}
function FName($id,$field){$name=$meta.labels["$id|$field"];if(!$name){$name=$field};return $name}
function ValueText($v){if($null -eq $v){return '빈칸 (null)'};if($v -is [ValueType] -and $v -eq 0){return '0'};return [string]$v}
function Buttons-Enabled($control,$yes){foreach($c in $control.Controls){if($c -is [Windows.Forms.Button]){$c.Enabled=$yes};if($c.Controls.Count){Buttons-Enabled $c $yes}}}
function Run($action,$p=@{},$workspaceOverride=''){
 if($dashProfileId -and $action -in @('importpreview','importcommit','stagingpreview','stagingcommit')){$p.Name=$dashProfileId;$p.Area=$dashArea}
 if(!$ws -and !$workspaceOverride){throw '먼저 프로젝트를 여세요.'};if($busy){throw '처리 중입니다.'};$script:busy=$true;Buttons-Enabled $f $false;$f.UseWaitCursor=$true;$status.Text='처리 중 — Excel을 저장하고 닫아 두세요.';if($dashMessage){$dashMessage.Text='자료를 처리하고 있습니다. 잠시 기다려 주세요.'};[Windows.Forms.Application]::DoEvents()
 try{
  # Excel runs in a separate STA process. Never block the window's STA on COM.
  function Quote-WorkerArg($value){$v=[string]$value;return ('"'+[regex]::Replace([regex]::Replace($v,'(\\*)"','$1$1\"'),'(\\+)$','$1$1')+'"')}
  $argsList=@('-NoProfile','-NonInteractive','-STA','-ExecutionPolicy','Bypass','-File',(Join-Path $base 'engine.ps1'),'-Action',$action,'-Workspace',$(if($workspaceOverride){$workspaceOverride}else{$ws}))
  foreach($key in $p.Keys){if([string]$p[$key] -ne ''){$argsList+=@(('-'+$key),([string]$p[$key]))}}
  $start=New-Object Diagnostics.ProcessStartInfo;$start.FileName=Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe';$start.Arguments=($argsList|ForEach-Object {Quote-WorkerArg $_}) -join ' ';$start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true;$start.StandardOutputEncoding=[Text.Encoding]::UTF8;$start.StandardErrorEncoding=[Text.Encoding]::UTF8
  $worker=New-Object Diagnostics.Process;$worker.StartInfo=$start
  $timer=[Diagnostics.Stopwatch]::StartNew();$lastSecond=-1
  $wasEnabled=@{};foreach($control in $f.Controls){$wasEnabled[$control]=$control.Enabled;$control.Enabled=$false}
  try{
   [void]$worker.Start();$outTask=$worker.StandardOutput.ReadToEndAsync();$errTask=$worker.StandardError.ReadToEndAsync()
   while(!$worker.HasExited){
    $second=[int]$timer.Elapsed.TotalSeconds
    if($second -ne $lastSecond){$lastSecond=$second;$status.Text='처리 중 · '+$second+'초 · 기존 자료 보호 중';if($dashMessage){$dashMessage.Text='Excel 자료 처리 중 · '+$second+'초. 완료되면 화면을 갱신합니다.'}}
    [Windows.Forms.Application]::DoEvents();[Threading.Thread]::Sleep(40)
    if($timer.Elapsed.TotalSeconds -gt 180){$worker.Kill();$worker.WaitForExit();throw '처리 시간이 180초를 초과하여 별도 작업을 중단했습니다. 부분 완료로 간주하지 않습니다. 다시 생성하지 말고 중단 체크포인트를 확인하세요.'}
   }
   $worker.WaitForExit();$output=$outTask.Result;$errors=$errTask.Result
   if($worker.ExitCode -ne 0){throw $errors};return $output
  }finally{foreach($control in $wasEnabled.Keys){$control.Enabled=$wasEnabled[$control]};$timer.Stop();$worker.Dispose()}
 }
 finally{$script:busy=$false;$f.UseWaitCursor=$false;Buttons-Enabled $f $true;$status.Text='준비';if($dashMessage -and ($dashMessage.Text -like 'Excel 자료 처리 중*' -or $dashMessage.Text -eq '자료를 처리하고 있습니다. 잠시 기다려 주세요.')){$dashMessage.Text='처리 완료 · 다음 작업을 선택하세요.'};if(Get-Command NRefreshActions -ErrorAction SilentlyContinue){NRefreshActions}}
}
function Log($s){$log.AppendText((Get-Date -Format 'HH:mm:ss')+' '+[string]$s+[Environment]::NewLine);$log.SelectionStart=$log.TextLength;$log.ScrollToCaret()}
function Input-Name($title,$prompt,$default){
 $d=New-Object Windows.Forms.Form;$d.Text=$title;$d.Size=New-Object Drawing.Size(570,200);$d.StartPosition='CenterParent';[void](Label $d $prompt 15 15 530 40)
 $t=New-Object Windows.Forms.TextBox;$t.Text=$default;$t.SetBounds(15,65,520,30);$d.Controls.Add($t)
 $ok=Button $d '확인' 330 110 95 {$d.Tag=$t.Text;$d.DialogResult='OK';$d.Close()};$cancel=Button $d '취소' 440 110 95 {$d.DialogResult='Cancel';$d.Close()};$d.AcceptButton=$ok;$d.CancelButton=$cancel
 if($(Ui12Dialog $d;$d.ShowDialog($f)) -eq 'OK'){return [string]$d.Tag};return ''
}

# Only the active table, local import and recovery surfaces are constructed.
. (Join-Path $base 'gui-management.ps1')
$status=New-Object Windows.Forms.Label
$tabs=New-Object Windows.Forms.TabControl;$tabs.SetBounds(210,110,1045,700);$f.Controls.Add($tabs)
$helpTab=New-Object Windows.Forms.TabPage;$helpTab.Text='작업 기록·도움말';[void]$tabs.TabPages.Add($helpTab)
$log=TextBox $helpTab 15 15 990 565;$log.ReadOnly=$true
function Refresh-Navigation($select){Read-Meta;$script:current=$select}
function Open-Workspace($path){
 if(Get-Command MemoryClear -ErrorAction SilentlyContinue){MemoryClear}
 if(!(Test-Path -LiteralPath (Join-Path $path 'relations.json'))){throw '프로젝트 폴더를 선택하세요.'}
 $script:ws=[IO.Path]::GetFullPath($path);Read-Meta;$script:current='L1'
}
. (Join-Path $base 'gui-staging.ps1')
. (Join-Path $base 'gui-dashboard.ps1')
. (Join-Path $base 'gui-workspaces.ps1')
. (Join-Path $base 'gui-projects.ps1')
. (Join-Path $base 'gui-trash-fold.ps1')
. (Join-Path $base 'gui-selection.ps1')
. (Join-Path $base 'gui-memory.ps1')
. (Join-Path $base 'gui-light.ps1')
. (Join-Path $base 'gui-pull-cells.ps1')
. (Join-Path $base 'gui-neis.ps1')
. (Join-Path $base 'gui-paste.ps1')
. (Join-Path $base 'gui-cleanup.ps1')
. (Join-Path $base 'gui-operations.ps1')
. (Join-Path $base 'gui-layout12.ps1')
. (Join-Path $base 'gui-theme12.ps1')
. (Join-Path $base 'gui-memos.ps1')
. (Join-Path $base 'gui-column-tools.ps1')
. (Join-Path $base 'gui-code-management.ps1')
. (Join-Path $base 'gui-import-mapping.ps1')
. (Join-Path $base 'gui-branding.ps1')
. (Join-Path $base 'gui-score-notes.ps1')
. (Join-Path $base 'gui-student-export.ps1')
[IO.File]::WriteAllText((Join-Path $base 'evidence/startup-diagnostic.json'),(@{tabs=@($tabs.TabPages|ForEach-Object {$_.Text});errors=@($Error|ForEach-Object {$_.ToString()});root=$base}|ConvertTo-Json -Depth 5));[void]$f.ShowDialog()
