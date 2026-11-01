param(
 [ValidateSet('cleanupscan','cleanupcommit','projectarchive','neispreview','neisimport','neisexport','projectcreate','projectinspect','projectbackup','projectrestore','init','create','register','scan','link','sync','open','query','status','rename','student','row','relocate','recover','preview','commit','additem','importpreview','importcommit','withdraw','membership','importrecover','stagingpreview','stagingcommit','courseconfigure','tablesnapshot','gridcommit','tableexport','tablepreview','tablecommit','coursework','spacecolumntype','spacecolumn','spacecreate','spacerename','parentpreview','parentcommit','studentview','trashpreview','trashcommit','trashlist','trashread')][string]$Action='status',
 [string]$Layer='L1',[string]$Parent='', [string]$File='', [string]$Item='A.score',
 [string]$Source='', [string]$SourceItem='', [string]$Target='', [string]$TargetItem='',
 [string]$Class='', [string]$Area='', [string]$Student='', [string]$AsOf='2026-10-31',
 [string]$Workspace=(Join-Path $PSScriptRoot 'sample'),[string]$Name='',[string]$PreviewFingerprint='',[string]$PayloadFile=''
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$Root=[IO.Path]::GetFullPath($Workspace).TrimEnd('\')
$MetaPath=Join-Path $Root 'relations.json'
. (Join-Path $PSScriptRoot 'student-codes.ps1')
function Save-Meta {
 $tmp=Join-Path $Root ('.relations-'+[guid]::NewGuid()+'.tmp')
 [IO.File]::WriteAllText($tmp,($script:M|ConvertTo-Json -Depth 30),[Text.UTF8Encoding]::new($true))
 if(Test-Path -LiteralPath $MetaPath){[IO.File]::Replace($tmp,$MetaPath,(Join-Path $Root ('backups/meta-'+[guid]::NewGuid()+'.json')))}else{[IO.File]::Move($tmp,$MetaPath)}
}
function To-Map($o){
 if($null -eq $o){return $null}
 if($o -is [System.Management.Automation.PSCustomObject]){$h=@{};foreach($p in $o.PSObject.Properties){$h[$p.Name]=To-Map $p.Value};return $h}
 if($o -is [array]){return ,@($o|ForEach-Object {To-Map $_})};return $o
}
function Load-Meta { $script:M=To-Map (Get-Content -Raw -Encoding utf8 -LiteralPath $MetaPath | ConvertFrom-Json) }
function Same-Value($a,$b){
 if($null -eq $a -or $null -eq $b){return ($null -eq $a -and $null -eq $b)}
 if($a -is [ValueType] -and $b -is [ValueType]){return ([decimal]$a -eq [decimal]$b)}
 return ([string]$a -ceq [string]$b)
}
function Resolve-Local($rel) {
 $p=[IO.Path]::GetFullPath((Join-Path $Root $rel))
 if(!$p.StartsWith($Root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Outside project root'}
 return $p
}
function Excel { $script:XL=New-Object -ComObject Excel.Application; $XL.DisplayAlerts=$false }
function Close-Excel { if($script:XL){$XL.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($XL);$script:XL=$null;[GC]::Collect()} }
function Layer-Info($id) { if(!$M.layers.Contains($id)){throw "Unknown layer: $id"};return $M.layers[$id] }
function Open-Book($id,[bool]$ReadOnly=$false) {
 $p=Resolve-Local (Layer-Info $id).path
 if(Test-Path -LiteralPath (Join-Path (Split-Path $p) ('~$'+(Split-Path $p -Leaf)))){throw "Close Excel workbook first: $p"}
 $b=$XL.Workbooks.Open($p,0,$ReadOnly)
 if($b.ReadOnly -and !$ReadOnly){$b.Close($false);throw 'Read-only or concurrent editing'}
 try{if([string]$b.Worksheets.Item('_identity').Cells.Item(1,2).Value2 -ne $id){throw 'Layer identity mismatch'}}catch{$b.Close($false);throw "Layer identity missing/mismatch: $id"}
 return $b
}
function Read-Table($b) {
 try{$s=$b.Worksheets.Item('Data')}catch{throw 'Data 시트가 없거나 이름이 바뀌었습니다. 원래 이름 Data로 복구하세요.'};$cols=@{};$warnings=[Collections.Generic.List[string]]::new()
 $lastColumn=$s.Cells.Item(1,$s.Columns.Count).End(-4159).Column
 $values=$s.Range($s.Cells.Item(1,1),$s.Cells.Item($s.UsedRange.Rows.Count,$lastColumn)).Value2
 $formulas=$s.Range($s.Cells.Item(1,1),$s.Cells.Item($s.UsedRange.Rows.Count,$lastColumn)).Formula
 for($c=1;$c -le $lastColumn;$c++){
  $h=[string]$values[1,$c]
  if(!$h){$warnings.Add("빈 헤더 열 $c : 자동 등록하지 않음");continue};if($cols.ContainsKey($h)){throw "Duplicate item ID/header: $h"};$cols[$h]=$c
 }
 if(!$cols.ContainsKey('student_id')){throw 'Missing student_id header'}
 $rows=@{};$bad=@{}
 for($r=2;$r -le $s.UsedRange.Rows.Count;$r++){
  $id=[string]$values[$r,$cols.student_id]
  if(!$id -or !$M.students.Contains($id)){$warnings.Add("Unknown/missing student ID at row $r : $id");continue}
  if($rows.ContainsKey($id)){$bad[$id]=$true;$warnings.Add("Duplicate student ID: $id");continue};$rows[$id]=$r
 }
 foreach($id in $bad.Keys){$rows.Remove($id)}
 return @{sheet=$s;cols=$cols;rows=$rows;warnings=$warnings;values=$values;formulas=$formulas}
}
function Refresh-Colors($t,$layerId){
 foreach($id in $t.rows.Keys){foreach($field in $t.cols.Keys){
  $cell=$t.sheet.Cells.Item($t.rows[$id],$t.cols[$field])
  if($cell.Interior.Color -ne 14811135){continue}
  $p=$M.provenance["$layerId|$id|$field"];$v=$cell.Value2
  $same=$p -and (Same-Value $v $p.value)
  if(!$same){$cell.Interior.ColorIndex=-4142}
 }}
}
function New-Layer($id,$parent,$items) {
 if($M.layers.Contains($id)){throw 'Layer already exists'}
 if($parent){[void](Layer-Info $parent)}
 $rel="books/$id.xlsx";$path=Resolve-Local $rel
 if(Test-Path -LiteralPath $path){throw 'Refuse overwrite'}
 $b=$XL.Workbooks.Add();$s=$b.Worksheets.Item(1);$s.Name='Data'
 while($b.Worksheets.Count -gt 1){$b.Worksheets.Item($b.Worksheets.Count).Delete()}
 $identity=$b.Worksheets.Add();$identity.Name='_identity';$identity.Cells.Item(1,1).Value2='layer_id';$identity.Cells.Item(1,2).Value2=$id;$identity.Visible=2
 $headers=@('student_id','virtual_name','code')+@($items)
 for($c=0;$c -lt $headers.Count;$c++){$s.Cells.Item(1,$c+1).Value2=$headers[$c]}
 $r=2;foreach($id2 in ($M.students.Keys | Sort-Object)){
  $s.Cells.Item($r,1).Value2=$id2;$s.Cells.Item($r,2).Value2=$M.students[$id2].name;$s.Cells.Item($r,3).Value2=$M.students[$id2].code;$r++
 }
 $s.Range($s.Cells.Item(1,1),$s.Cells.Item($M.students.Count+1,$headers.Count)).AutoFilter() | Out-Null
 $s.Rows.Item(1).Font.Bold=$true;$s.Rows.Item(1).Interior.Color=15773696
 $s.UsedRange.Columns.ColumnWidth=19;$s.UsedRange.Font.Size=11
 $s.Activate();$XL.ActiveWindow.SplitRow=1;$XL.ActiveWindow.FreezePanes=$true
 $b.SaveAs($path,51);$b.Close($false)
 $M.layers[$id]=@{path=$rel;parent=$parent;items=@($items);name=$id};Save-Meta
}
. (Join-Path $PSScriptRoot 'plan.ps1')
. (Join-Path $PSScriptRoot 'project.ps1')
. (Join-Path $PSScriptRoot 'cleanup.ps1')
if($Action -in @('cleanupscan','cleanupcommit')){$req=To-Map (Get-Content -LiteralPath $PayloadFile -Raw -Encoding UTF8|ConvertFrom-Json);if($Action -eq 'cleanupscan'){CleanupPlan $req|ConvertTo-Json -Depth 10}else{CleanupCommit $req|ConvertTo-Json -Depth 10};exit}
if($Action -eq 'projectcreate'){PCreate (To-Map (Get-Content $PayloadFile -Raw -Encoding UTF8|ConvertFrom-Json))|ConvertTo-Json -Depth 10;exit}
if($Action -eq 'projectarchive'){PArchive (To-Map (Get-Content $PayloadFile -Raw -Encoding UTF8|ConvertFrom-Json))|ConvertTo-Json -Depth 10;exit}
if($Action -eq 'projectinspect'){PInspect $Root|ConvertTo-Json -Depth 10;exit}
if($Action -eq 'projectbackup'){PBackup|ConvertTo-Json -Depth 10;exit}
if($Action -eq 'projectrestore'){PRestore|ConvertTo-Json -Depth 10;exit}
if($Action -eq 'init'){
 if(Test-Path $MetaPath){throw 'Already initialized'}
 New-Item -ItemType Directory -Path (Join-Path $Root 'books'),(Join-Path $Root 'backups'),(Join-Path $Root 'logs') -Force | Out-Null
 $script:M=@{version=1;students=@{};layers=@{};links=@();provenance=@{};memberships=@()}
 for($n=1;$n -le 10;$n++){
  $id='S{0:000}' -f $n;$M.students[$id]=@{name="가상학생$n";code=$id}
  $M.memberships+=@{student=$id;class=$(if($n -le 5){'1반'}else{'2반'});from='2026-03-01';until=$(if($n -eq 9){'2026-10-15'}else{$null})}
 }
 $M.memberships[9].from='2026-10-20';$M.memberships[9].decision='과거평가 교사판단 대기'
 Save-Meta;Excel
 try {New-Layer 'L1' '' @('A.score','B.score','C.score')}finally{Close-Excel}
 'Initialized virtual roster: 10, classes: 2, areas: 3';exit
}
$guard=$null
try{
if($Action -notin @('open','status','preview','importpreview','stagingpreview','tablesnapshot','tablepreview','parentpreview','studentview','trashpreview','trashlist','trashread')){
 try{$guard=[IO.File]::Open((Join-Path $Root '.manager.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)}catch{throw 'Workspace busy: another manager is writing'}
}
Load-Meta
. (Join-Path $PSScriptRoot 'trash.ps1')
if(!$M.workspaceScopes){$M.workspaceScopes=@{}}
TRecover
if($Action -in @('trashpreview','trashcommit','trashlist','trashread')){
 $r=@{};if($PayloadFile){$r=To-Map (Get-Content $PayloadFile -Raw -Encoding UTF8|ConvertFrom-Json)}
 $result=switch($Action){'trashpreview'{TPlan $r};'trashcommit'{TCommit $r};'trashlist'{TList};'trashread'{TRead $r}};$result|ConvertTo-Json -Depth 45;exit
}
if($Action -in @('rename','link','sync','commit','additem')){foreach($pair in @(@($Layer,$Item),@($Source,$SourceItem),@($Target,$TargetItem))){if($M.retiredItems -and $M.retiredItems.ContainsKey(([string]$pair[0]+'|'+[string]$pair[1]))){throw '휴지통으로 이동한 항목은 수정하거나 다시 연결하지 않습니다. 휴지통 읽기 전용 기록을 확인하세요.'}}}
if((Test-Path (Join-Path $Root 'pending-import.json')) -and $Action -notin @('status','query','open','recover','preview','importpreview','stagingpreview','tablesnapshot','tablepreview','importrecover')){throw '중단된 반입이 있습니다. pending-import.json의 복구 사본을 사용하세요.'}
if((Test-Path -LiteralPath (Join-Path $Root 'pending-sync.json')) -and $Action -notin @('status','query','open','recover','preview')){throw '이전 갱신 중단 감지: pending-sync.json의 체크포인트로 복구 사본을 만드세요. 현재 파일은 추가 변경하지 않습니다.'}
if($Action -in @('create','register','scan','link','sync','rename','student','row','relocate','commit','additem','importcommit','withdraw','membership','stagingcommit','courseconfigure','gridcommit','tablecommit','coursework')){
 # ZIP archives may omit empty workspace folders. Initialize them before any write.
 New-Item -ItemType Directory -Path (Join-Path $Root 'backups'),(Join-Path $Root 'logs') -Force|Out-Null
}
if($Action -eq 'recover'){
 $checkpoint=Resolve-Local $File
 if(!(Test-Path -LiteralPath (Join-Path $checkpoint 'relations.json'))){throw 'Checkpoint relations.json missing'}
 $dest=Join-Path $Root ('recovered-'+[guid]::NewGuid().ToString('N'))
 New-Item -ItemType Directory $dest|Out-Null;Copy-Item -Path (Join-Path $checkpoint '*') -Destination $dest -Recurse
 New-Item -ItemType Directory (Join-Path $dest 'backups'),(Join-Path $dest 'logs') -Force|Out-Null
 $restoredMeta=To-Map (Get-Content (Join-Path $dest 'relations.json') -Raw -Encoding utf8|ConvertFrom-Json)
 foreach($batchId in $restoredMeta.importBatches.Keys){if($batchId -notmatch '^B[0-9a-f]{20}$'){throw '복구 배치 ID 오류'};$sourceDir=Join-Path $Root ('imports/'+$batchId);if(!(Test-Path $sourceDir)){throw '복구할 원본/교환 기록 누락'};New-Item -ItemType Directory (Join-Path $dest 'imports') -Force|Out-Null;Copy-Item -LiteralPath $sourceDir -Destination (Join-Path $dest ('imports/'+$batchId)) -Recurse}
 $dest;exit
}
if(!$M.ContainsKey('labels')){$M.labels=@{}}
if(!$M.ContainsKey('reviewItems')){$M.reviewItems=@{}}
if(!$M.ContainsKey('participation')){$M.participation=@{}}
if(!$M.ContainsKey('manualOverrides')){$M.manualOverrides=@{}}
if(!$M.ContainsKey('importRecords')){$M.importRecords=@{}}
. (Join-Path $PSScriptRoot 'natural-import.ps1')
. (Join-Path $PSScriptRoot 'management.ps1')
. (Join-Path $PSScriptRoot 'staging.ps1')
. (Join-Path $PSScriptRoot 'course.ps1')
. (Join-Path $PSScriptRoot 'leaf-import.ps1')
. (Join-Path $PSScriptRoot 'workspace.ps1')
. (Join-Path $PSScriptRoot 'neis.ps1')
if($Action -eq 'neispreview'){$n=NeisRead;@{areas=$n.areas;info=$n.info}|ConvertTo-Json -Depth 10;exit}
if($Action -eq 'neisimport'){NeisImport|ConvertTo-Json -Depth 15;exit}
if($Action -eq 'neisexport'){NeisExport|ConvertTo-Json -Depth 15;exit}
if($Action -in @('spacecolumntype','spacecolumn','spacecreate','spacerename','parentpreview','parentcommit','studentview')){
 New-Item -ItemType Directory -Path (Join-Path $Root 'backups'),(Join-Path $Root 'logs') -Force|Out-Null
 $r=@{};if($PayloadFile){$r=To-Map (Get-Content $PayloadFile -Raw -Encoding utf8|ConvertFrom-Json)}
 $result=switch($Action){
 'spacecolumntype'{WColumnType $r};'spacecolumn'{WColumn $r};'spacecreate'{WCreate $r};'spacerename'{WRename $r};'parentpreview'{WParentPlan};'parentcommit'{WParentCommit $r};'studentview'{WStudentView $r}
 };$result|ConvertTo-Json -Depth 25;exit
}
if($Action -eq 'coursework'){$r=@{};if($PayloadFile){$r=To-Map (Get-Content -LiteralPath $PayloadFile -Raw -Encoding utf8|ConvertFrom-Json)};CWork $r|ConvertTo-Json -Depth 12;exit}
if($Action -eq 'tablesnapshot'){CSnapshot|ConvertTo-Json -Depth 20;exit}
if($Action -eq 'tableexport'){CExport|ConvertTo-Json -Depth 10;exit}
if($Action -eq 'tablepreview'){CTablePreview|ConvertTo-Json -Depth 20;exit}
if($Action -eq 'tablecommit'){$p=CTablePreview;if(!$PreviewFingerprint -or $PreviewFingerprint -ne $p.fingerprint){throw '업로드 미리보기가 바뀌었습니다.'};CCommit @{expected_fingerprint=$p.expected_fingerprint;changes=@($p.changes);new_fields=@($p.new_fields);import_hash=$p.source_hash}|ConvertTo-Json -Depth 20;exit}
if($Action -in @('courseconfigure','gridcommit')){$r=To-Map (Get-Content -LiteralPath $PayloadFile -Raw -Encoding utf8|ConvertFrom-Json);if($Action -eq 'courseconfigure'){CConfigure $r|ConvertTo-Json -Depth 20}else{CCommit $r|ConvertTo-Json -Depth 20};exit}
if($Action -eq 'stagingpreview'){SPreview|ConvertTo-Json -Depth 30;exit}
if($Action -eq 'stagingcommit'){if(!$PayloadFile){throw '변경 자료 필요'};$r=To-Map (Get-Content -LiteralPath $PayloadFile -Raw -Encoding utf8|ConvertFrom-Json);$saved=SCommit $r;WTagImported $saved;$saved|ConvertTo-Json -Depth 30;exit}
if($Action -in @('importpreview','importcommit')){$p=NPreview;if($Action -eq 'importpreview'){$p|ConvertTo-Json -Depth 25}else{$saved=NCommit $p;WTagImported $saved;$saved|ConvertTo-Json -Depth 25};exit}
if($Action -eq 'importrecover'){NRecover|ConvertTo-Json -Depth 8;exit}
if($Action -eq 'withdraw'){NWithdraw $Name|ConvertTo-Json -Depth 8;exit}
if($Action -eq 'membership'){if(!$PayloadFile){throw '변경 자료 필요'};Manage-Membership (To-Map (Get-Content $PayloadFile -Raw -Encoding utf8|ConvertFrom-Json))|ConvertTo-Json -Depth 12;exit}
if($Action -eq 'rename'){
 if(!$Name){throw 'Display name required'}
 if($Item -eq '@layer'){$M.layers[$Layer].name=$Name}else{$M.labels["$Layer|$Item"]=$Name};Save-Meta;'Display name saved';exit
}
if($Action -eq 'student'){
 if(!$Student -or !$Name -or $M.students.Contains($Student)){throw 'New unique student ID and name required'}
 $M.students[$Student]=@{name=$Name;code=$Student};$M.memberships+=@{student=$Student;class=$Class;from=$AsOf;until=$null};Save-Meta;'Roster registered';exit
}

if($Action -eq 'status'){$M | ConvertTo-Json -Depth 20;exit}
if($Action -eq 'open'){
 $l=Layer-Info $Layer;Start-Process (Resolve-Local $l.path)
 "Layer $Layer; parent=$($l.parent); children=$((@($M.layers.Keys | Where-Object {$M.layers[$_].parent -eq $Layer})) -join ',')";exit
}
if($Action -eq 'register'){
 if($M.layers.Contains($Layer)){throw 'Layer exists'}
 $path=Resolve-Local $File;if(!(Test-Path $path)){throw 'File missing'}
 if($Parent){[void](Layer-Info $Parent)}
 if(@($M.layers.Values|Where-Object {(Resolve-Local $_.path) -eq $path}).Count){throw 'File already registered'}
 Excel
 try{
  $b=$XL.Workbooks.Open($path,0,$false)
  try{
   if($b.ReadOnly){throw 'File busy/read-only'};$t=Read-Table $b
   $b.SaveCopyAs((Join-Path $Root ('backups/'+[guid]::NewGuid()+'.xlsx')))
   try{$identity=$b.Worksheets.Item('_identity');throw 'Existing layer identity: register a distinct new file'}catch{if($_.Exception.Message -match 'Existing layer identity'){throw}}
   $identity=$b.Worksheets.Add();$identity.Name='_identity';$identity.Cells.Item(1,1).Value2='layer_id';$identity.Cells.Item(1,2).Value2=$Layer;$identity.Visible=2;$b.Save()
  }finally{$b.Close($false)}
 }finally{Close-Excel}
 $M.layers[$Layer]=@{path=$File;parent=$Parent;items=@();name=$Layer};Save-Meta;'Registered; run scan';exit
}
Excel
try{
 if($Action -in @('preview','commit','sync')){
  $plan=Build-Plan
  if($Action -eq 'preview'){$plan|ConvertTo-Json -Depth 16}else{(Apply-Plan $plan)|ConvertTo-Json -Depth 16};exit
 }
 switch($Action){
 'create' {New-Layer $Layer $Parent @($Item);"Created $Layer"}
 'relocate' {
  $p=Resolve-Local $File;if(!(Test-Path -LiteralPath $p)){throw 'Selected file missing'}
  if(Test-Path -LiteralPath (Join-Path (Split-Path $p) ('~$'+(Split-Path $p -Leaf)))){throw 'Close selected Excel file first'}
  $b=$XL.Workbooks.Open($p,0,$false)
  try{if($b.ReadOnly -or [string]$b.Worksheets.Item('_identity').Cells.Item(1,2).Value2 -ne $Layer){throw 'Selected file is not the same layer'};[void](Read-Table $b)}finally{$b.Close($false)}
  Copy-Item -LiteralPath $MetaPath -Destination (Join-Path $Root ('backups/'+[guid]::NewGuid()+'.json'))
  $M.layers[$Layer].path=$File;Save-Meta;'File location repaired after identity check'
 }
 'row' {
  if(!$M.students.Contains($Student)){throw 'Register roster first'};$b=Open-Book $Layer
  try{$t=Read-Table $b;if($t.warnings.Count){throw 'Fix ID warnings before adding a row'};if($t.rows.ContainsKey($Student)){throw 'Student row already exists'}
   $b.SaveCopyAs((Join-Path $Root ('backups/'+[guid]::NewGuid()+'.xlsx')))
   $r=$t.sheet.UsedRange.Rows.Count+1;$t.sheet.Cells.Item($r,$t.cols.student_id).Value2=$Student
   $t.sheet.Cells.Item($r,$t.cols.virtual_name).Value2=[string]$M.students[$Student].name
   $t.sheet.Cells.Item($r,$t.cols.code).Value2=$Student;$b.Save();'Layer row added'
  }finally{$b.Close($false)}
 }
 'scan' {
  if($M.retiredItems -and @($M.retiredItems.Keys|Where-Object {$_.StartsWith($Layer+'|')}).Count){throw '보관용 원래 열이 남아 있어 자동 재등록을 차단합니다. 활성 작업표를 사용하세요.'}
  $b=Open-Book $Layer
  try{$t=Read-Table $b;$items=@($t.cols.Keys | Where-Object {$_ -notin @('student_id','virtual_name','code')} | Sort-Object);$M.layers[$Layer].items=$items;Save-Meta;@{items=$items;warnings=@($t.warnings)}|ConvertTo-Json -Depth 8}finally{$b.Close($false)}
 }
 'link' {
  [void](Layer-Info $Source);[void](Layer-Info $Target)
  if($M.layers[$Source].parent -ne $Target){throw 'Source must be a direct child of target'}
  if(!$SourceItem -or !$TargetItem){throw 'Item IDs required'}
  if(@($M.links | Where-Object {$_.target -eq $Target -and $_.targetItem -eq $TargetItem}).Count){throw 'Target item already linked'}
  foreach($pair in @(@($Source,$SourceItem),@($Target,$TargetItem))){$b=Open-Book $pair[0];try{$t=Read-Table $b;if(!$t.cols.ContainsKey($pair[1])){throw "Item missing: $($pair[1])"}}finally{$b.Close($false)}}
  $M.links+=@{source=$Source;sourceItem=$SourceItem;target=$Target;targetItem=$TargetItem};Save-Meta;'Explicit link registered; sync required'
 }
 'additem' {
  if(!$Name -or !$Area){throw '표시명과 영역을 선택하세요.'}
  if(@($M.layers[$Layer].items|Where-Object {(Label-For $Layer $_) -eq $Name}).Count){throw '같은 표에 동일한 표시명이 있습니다. 구별되는 이름을 쓰세요.'}
  $b=Open-Book $Layer
  try{$t=Read-Table $b;$id=$Area+'.i'+[guid]::NewGuid().ToString('N').Substring(0,8);$c=($t.cols.Values|Measure-Object -Maximum).Maximum+1
   $b.SaveCopyAs((Join-Path $Root ('backups/'+[guid]::NewGuid()+'.xlsx')))
   $t.sheet.Cells.Item(1,$c).Value2=$id;$t.sheet.Columns.Item($c).ColumnWidth=22;$b.Save()
   $M.layers[$Layer].items+=,$id;$M.labels["$Layer|$id"]=$Name;Save-Meta
   @{item=$id;label=$Name}|ConvertTo-Json
  }finally{$b.Close($false)}
 }

 'query' {
  $b=Open-Book $Layer
  try{
   $t=Read-Table $b;$out=@()
   foreach($id in ($t.rows.Keys|Sort-Object)){
    $person=$M.students[$id];$membership=@($M.memberships|Where-Object {$_.student -eq $id -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)})
    if($Student -and $Student -notin @($id,$person.name,$person.code)){continue}
    if($Class -and $Class -notin @($membership.class)){continue}
    foreach($field in $t.cols.Keys|Sort-Object){
     if($field -in @('student_id','virtual_name','code')){continue};if($Area -and !$field.StartsWith($Area+'.')){continue}
     $val=$t.values[$t.rows[$id],$t.cols[$field]];$p=$M.provenance["$Layer|$id|$field"]
     $origin=if($p -and (Same-Value $val $p.value)){'참조'}elseif($p){'충돌/직접수정'}else{'직접입력'}
     $state=if($null -eq $val){'미입력'}else{'입력됨'}
     $out+=@{canonical="$Layer|$id|$field";student=$id;name=$person.name;class=@($membership.class);item=$field;label=$M.labels["$Layer|$field"];value=$val;state=$state;origin=$origin;source=$p.source;sourceItem=$p.item;refreshed=$p.refreshed;review=$M.reviewItems["$Layer|$field"];reviewNote=$M.manualOverrides["$Layer|$id|$field"].note;path=(Layer-Info $Layer).path}
    }
   };$fingerprint=(NHash ((Shared-Hash $MetaPath)+'|'+(Shared-Hash (Resolve-Local (Layer-Info $Layer).path))));@{data=$out;warnings=@($t.warnings);fingerprint=$fingerprint}|ConvertTo-Json -Depth 12
  }finally{$b.Close($false)}
 }
 }
}finally{Close-Excel}
}finally{if($guard){$guard.Dispose()}}
