$script:nView='project';$script:nLastArea='';$script:nLastSpace='';$script:nTabLoading=$false
$f.Text=$productTitle;$f.MinimumSize=New-Object Drawing.Size(1300,940)
$nBar=New-Object Windows.Forms.FlowLayoutPanel;$nBar.Dock='Bottom';$nBar.Height=52;$nBar.BackColor=$pale;$nBar.WrapContents=$false;$f.Controls.Add($nBar)
$nTabs=@{};foreach($spec in @(@('project','프로젝트'),@('summary','종합표'),@('detail','영역 세부'),@('workspace','작업공간'),@('student','학생 개별조회'),@('import','반입 검토'),@('help','오류·복구 안내'))){$b=DButton $nBar $spec[1] 0 0 175 {DSafe {NSelect ([string]$this.Tag)}};$b.Tag=$spec[0];$b.Height=44;$b.Margin=New-Object Windows.Forms.Padding(3,4,3,4);$nTabs[$spec[0]]=$b}
$nHome=New-Object Windows.Forms.Panel;$nHome.Dock='Fill';$nHome.BackColor=[Drawing.Color]::White;$f.Controls.Add($nHome)
$nHomeTitle=Label $nHome $productTitle 35 30 900 50;$nHomeTitle.Font=New-Object Drawing.Font('맑은 고딕',22,[Drawing.FontStyle]::Bold);$nHomeTitle.ForeColor=$navy
$nHomeDescription=Label $nHome $productDescription 300 22 700 26;$nHomeDescription.ForeColor=$navy
$nFlow=Label $nHome '① 프로젝트 생성/불러오기 → ② 명부 반입 → ③ 평가영역·하위 작업표 → ④ 점수 입력·저장' 40 92 1150 40
DButton $nHome '새 프로젝트 만들기' 40 150 360 {DSafe {NNewProject}}|Out-Null
DButton $nHome '프로젝트 불러오기' 420 150 360 {DSafe {NOpenPicker}}|Out-Null
$nCurrent=TextBox $nHome 40 220 1110 175;$nCurrent.ReadOnly=$true
[void](Label $nHome '최근 프로젝트 — 선택 후 불러오기' 40 420 700 30)
$nRecent=Grid $nHome 40 460 1110 160 @('프로젝트','실제 PC 폴더');$nRecent.Columns[1].FillWeight=300
$nOpenSelected=DButton $nHome '선택 프로젝트 불러오기' 40 640 360 {DSafe {if(!$nRecent.CurrentRow){throw '최근 프로젝트를 선택하세요.'};NOpen ([string]$nRecent.CurrentRow.Tag)}}
$nBackup=DButton $nHome '현재 프로젝트 백업' 420 640 360 {DSafe {DClean;if(!$ws){throw '백업할 프로젝트를 불러오세요.'};$r=(Run 'projectbackup')|ConvertFrom-Json;$nCurrent.Text+=''+"`r`n완료된 백업: "+$r.backup+"`r`n백업 "+$r.files+'파일 해시 일치. 외부 전송 없음.'}}
DButton $nHome '백업을 새 폴더로 복원' 800 640 350 {DSafe {NRestorePicker}}|Out-Null
[void](Label $nHome 'Drive·OneDrive·네트워크 폴더는 로컬 전용 보관이 아닙니다. 자동 동기화·다른 PC 사용은 아직 검증하지 않았습니다.' 40 710 1120 60)
# Replace the old handler by a dedicated header button, retaining legacy tools behind it.
$nProjectHeader=DButton $dashboard '프로젝트' 1090 12 155 {DSafe {NSelect 'project'}}

$nProjectLabel=Label $dashboard '' 210 48 1020 22;$nProjectLabel.ForeColor=$navy
function NRun($action,$args2=@{},$target=''){$script:nErrorContext=if($target){$target}else{$ws};try{return Run $action $args2 $target}finally{}}
function NConfigFile($value){$dir=Join-Path $base 'evidence/project-requests';New-Item -ItemType Directory $dir -Force|Out-Null;$path=Join-Path $dir ('config-'+[guid]::NewGuid().ToString('N')+'.json');$value|ConvertTo-Json -Depth 12|Set-Content $path -Encoding UTF8;return $path}
function NRecentEntries($value){foreach($x in $value){if($x.path -is [string] -and $x.label -is [string]){$x}elseif($x.value){NRecentEntries $x.value}}}
function NRecents{
 $script:nRecentLoading=$true;try{$nRecent.Rows.Clear();$path=Join-Path $base 'recent-projects.json';if(Test-Path $path){$raw=Get-Content $path -Raw -Encoding UTF8|ConvertFrom-Json;foreach($x in @(NRecentEntries $raw)){$display=[string]$x.label;if(!(Test-Path -LiteralPath (Join-Path ([string]$x.path) 'relations.json'))){$display+=' (자료 누락 또는 폴더 없음)'};$r=$nRecent.Rows[$nRecent.Rows.Add([object[]]@($display,[string]$x.path))];$r.Tag=[string]$x.path}}}finally{$script:nRecentLoading=$false};if(Get-Command NRefreshActions -ErrorAction SilentlyContinue){NRefreshActions}
}
function NRemember($info){
 $path=Join-Path $base 'recent-projects.json';$list=@();if(Test-Path $path){$raw=Get-Content $path -Raw -Encoding UTF8|ConvertFrom-Json;$list=@(NRecentEntries $raw|Where-Object {$_.path -ne $info.root})};$list=@(@{label=$info.project.label;path=$info.root})+$list;$list=@($list|Select-Object -First 12);[IO.File]::WriteAllText($path,(ConvertTo-Json -InputObject $list -Depth 5),[Text.UTF8Encoding]::new($true));NRecents
}
function NOpen($path){
 DClean;$info=(NRun 'projectinspect' @{} $path)|ConvertFrom-Json
 $previous=@{root=$ws;profile=$dashProfileId;mode=$dashMode;area=$dashArea;target=$wTarget;student=$wStudent};try{
 $script:wStudent='';$script:wCodes=@{};$script:wIdentityMode=0;$script:wTarget='';$script:nLastArea='';$script:nLastSpace=''
 Open-Workspace $info.root;DInitialize '';NRemember $info
 $nCurrent.Text='현재 프로젝트: '+$info.project.label+"`r`n실제 폴더: "+$info.root+"`r`n학생 "+$info.students+'명 / 작업표 '+$info.layers+"개`r`n데이터 books / 관계 relations.json / 반입 원본 imports / 명부 원본 originals / 복구 backups / 로그 logs"+$(if($info.project.legacy){"`r`n기존 형식: 원본 폴더 구조와 불변 연결을 그대로 사용합니다."}else{"`r`n프로젝트 ID: "+$info.project.id})
 $nProjectLabel.Text=$info.project.label+'  |  '+$info.root;NActivate 'summary'
 if(!$info.students){$dashMessage.Text='② 명부 반입: 프로젝트 화면의 나이스 학급 반입을 누르세요. 학생 0명인 빈 표도 정상 생성된 상태입니다.'}
 }catch{$failure=$_;NCloseProject;if($previous.root -and $previous.root -ne $info.root -and (Test-Path -LiteralPath (Join-Path $previous.root 'relations.json'))){try{Open-Workspace $previous.root;$script:wTarget=$previous.target;$script:wStudent=$previous.student;DInitialize $previous.profile;DLoad $previous.mode $previous.area;NActivate $previous.mode}catch{NCloseProject;NActivate 'project'}}else{NActivate 'project'};throw $failure}
}
function NOpenPicker{
 DClean;$d=New-Object Windows.Forms.FolderBrowserDialog;$d.Description='프로젝트 이름의 폴더를 선택하세요. 내부 파일을 고를 필요가 없습니다.';$d.ShowNewFolderButton=$false
 if($ws -and (Test-Path -LiteralPath $ws -PathType Container)){$d.SelectedPath=$ws}elseif(Test-Path -LiteralPath (Join-Path $base 'projects')){$d.SelectedPath=Join-Path $base 'projects'}
 try{if($d.ShowDialog($f) -eq 'OK'){$path=$d.SelectedPath;if(!(Test-Path -LiteralPath (Join-Path $path 'relations.json') -PathType Leaf)){[void][Windows.Forms.MessageBox]::Show($f,'프로젝트 이름의 폴더를 선택해 주세요. 상위 보관 폴더를 선택했다면 그 안의 프로젝트 폴더를 선택하세요.','프로젝트 폴더 선택','OK','Information');return};NOpen $path}}finally{$d.Dispose()}
}
function NNewProject{
 DClean;$d=New-Object Windows.Forms.Form;$d.Text='새 프로젝트 만들기';$d.Size=New-Object Drawing.Size(900,760);$d.StartPosition='CenterParent';$d.Font=$f.Font
 [void](Label $d '① 평가 설정 → ② 저장 위치 → ③ 명부 선택 → 프로젝트 생성' 20 15 840 32)
 [void](Label $d '① 프로젝트명' 20 68 125);$nameBox=TextBox $d 150 62 690;$nameBox.Text='2026_2학기_3학년_역사'
 [void](Label $d '학년도 / 학기 / 학년 / 과목' 20 115 290);$yearBox=TextBox $d 305 110 90;$yearBox.Text='2026';$termBox=Combo $d 410 110 90;Choice $termBox '1' '1학기';Choice $termBox '2' '2학기';$termBox.SelectedIndex=1;$gradeBox=Combo $d 515 110 90;foreach($g in 1..12){Choice $gradeBox ([string]$g) ($g.ToString()+'학년')};$gradeBox.SelectedIndex=2;$subjectBox=TextBox $d 620 110 220;$subjectBox.Text='역사'
 [void](Label $d '평가영역 (쉼표·; 구분)' 20 167 220);$areaBox=TextBox $d 250 162 590;$areaBox.Text='주제 탐구 활동;역사적 사고력 활동;과정형 활동'
 [void](Label $d '② 저장 위치' 20 225 125);$pathBox=TextBox $d 150 220 510;$pathBox.ReadOnly=$true
 $folderButton=DButton $d '저장 폴더 선택' 675 217 165 {$pick=New-Object Windows.Forms.FolderBrowserDialog;$pick.Description='새 프로젝트 폴더를 만들 상위 위치를 선택하세요.';$pick.ShowNewFolderButton=$true;if($pathBox.Text){$pick.SelectedPath=$pathBox.Text}else{$pick.SelectedPath=Join-Path $base 'projects'};try{if($pick.ShowDialog($d) -eq 'OK'){$pathBox.Text=$pick.SelectedPath}}finally{$pick.Dispose()};Update-Creation}
 [void](Label $d '선택한 위치 안에 프로젝트명으로 새 폴더가 생깁니다. 순서를 바꿔 선택해도 됩니다.' 150 255 690 40)
 [void](Label $d '③ 시작 명부' 20 310 125);$rosterBox=Combo $d 150 305 690;Choice $rosterBox 'xlsx' 'Excel 명부 (.xlsx · 이름/성명, 반, 번호)';Choice $rosterBox 'empty' '명부 없이 시작 → 생성 후 나이스 학급 반입';Choice $rosterBox 'csv' 'CSV 명부 (name,class,number)';$rosterBox.SelectedIndex=1
 $rosterFile=TextBox $d 150 355 690;$rosterFile.ReadOnly=$true
 $rosterButton=DButton $d '명부 파일 선택' 20 352 125 {$pick=New-Object Windows.Forms.OpenFileDialog;$pick.Title='Excel 또는 CSV 명부 선택';$pick.Filter='Excel 명부|*.xlsx|CSV 명부|*.csv';try{if($pick.ShowDialog($d) -eq 'OK'){$chosen=$pick.FileName;Choose $rosterBox ([IO.Path]::GetExtension($chosen).TrimStart('.').ToLowerInvariant());$rosterFile.Text=$chosen}}finally{$pick.Dispose()};Update-Creation}
 $emptyRosterButton=DButton $d '명부 없이 시작 선택' 150 395 250 {Choose $rosterBox 'empty';Update-Creation}
 $nextLabel=Label $d '' 20 445 820 65;$nextLabel.BackColor=$pale;$nextLabel.ForeColor=$navy
 $errorLabel=Label $d '' 20 515 820 85;$errorLabel.ForeColor=[Drawing.Color]::Firebrick
 $createButton=DButton $d '프로젝트 생성' 475 620 180 {
  $problem=Creation-Problem;if($problem){$nextLabel.Text=$problem;return};$errorLabel.Text=''
  try{$req=@{label=$nameBox.Text.Trim();year=$yearBox.Text.Trim();term=$termBox.SelectedItem.Id;grade=$gradeBox.SelectedItem.Id;subject=$subjectBox.Text.Trim();areas=@($areaBox.Text -split '[,;，；\r\n]'|ForEach-Object {$_.Trim()}|Where-Object {$_});roster=$rosterBox.SelectedItem.Id;roster_file=$(if($rosterBox.SelectedItem.Id -eq 'empty'){''}else{$rosterFile.Text})};$target=Join-Path $pathBox.Text $nameBox.Text.Trim();$r=(NRun 'projectcreate' @{PayloadFile=(NConfigFile $req)} $target)|ConvertFrom-Json;$d.Tag=$r.root;$d.DialogResult='OK';$d.Close()}
  catch{$errorLabel.Text='생성하지 못했습니다: '+($_.Exception.Message -split "`r?`n")[0]+"`r`n입력 내용을 유지했습니다. 수정한 뒤 다시 생성하거나 돌아가기를 누르세요.";Log $_.Exception.Message;Update-Creation}
 }
 $cancelButton=DButton $d '돌아가기' 675 620 165 {$d.Close()};$cancelButton.DialogResult='Cancel';$d.CancelButton=$cancelButton
 function Creation-Problem{
  $n=$nameBox.Text.Trim();if(!$n -or $n.Length -gt 80 -or $n.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $n -in @('.','..') -or $n.EndsWith('.') -or $n -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)'){return '① 프로젝트명을 확인하세요. 폴더 이름으로 쓸 수 있는 이름을 입력합니다.'}
  if($yearBox.Text.Trim() -notmatch '^20\d\d$' -or !$subjectBox.Text.Trim()){return '① 학년도(4자리)와 과목을 입력하세요.'}
  $a=@($areaBox.Text -split '[,;，；\r\n]'|ForEach-Object {$_.Trim()}|Where-Object {$_});if(!$a.Count -or $a.Count -gt 12 -or @($a|Sort-Object -Unique).Count -ne $a.Count){return '① 서로 다른 평가영역을 1~12개 입력하세요. 영역 사이는 쉼표 또는 ;로 구분합니다.'}
  if(!$pathBox.Text -or !(Test-Path -LiteralPath $pathBox.Text -PathType Container)){return '② 저장 폴더 선택을 눌러 저장 위치를 고르세요.'}
  if(Test-Path -LiteralPath (Join-Path $pathBox.Text $n)){return '② 같은 이름의 폴더가 있습니다. 프로젝트명을 바꾸거나 다른 저장 위치를 선택하세요. 기존 폴더는 덮어쓰지 않습니다.'}
  if($rosterBox.SelectedItem.Id -ne 'empty'){
   if(!$rosterFile.Text -or !(Test-Path -LiteralPath $rosterFile.Text -PathType Leaf)){return '③ 명부 파일을 선택하세요. 나이스 학급 파일을 나중에 반입하려면 명부 없이 시작을 고르세요.'}
   if([IO.Path]::GetExtension($rosterFile.Text).TrimStart('.').ToLowerInvariant() -ne $rosterBox.SelectedItem.Id){return '③ 선택한 명부 종류에 맞는 파일을 다시 선택하세요.'}
  };return ''
 }
 function Update-Creation{
  $empty=($rosterBox.SelectedItem.Id -eq 'empty');$rosterButton.Enabled=$true;$rosterFile.Enabled=$true
  $problem=Creation-Problem;$createButton.Enabled=(!$problem -and !$busy)
  $nextLabel.Text=if($problem){'다음 할 일: '+$problem}else{'평가영역 '+@($areaBox.Text -split '[,;，；\r\n]'|Where-Object {$_.Trim()}).Count+'개 / 준비 완료: '+(Join-Path $pathBox.Text $nameBox.Text.Trim())+"`r`n"+$(if($empty){'명부 없이 생성합니다. 다음 단계는 나이스 학급 반입입니다.'}else{'선택한 명부로 생성한 뒤 종합표를 엽니다.'})}
 }
 foreach($c in @($nameBox,$yearBox,$subjectBox,$areaBox,$pathBox,$rosterFile)){$c.Add_TextChanged({$errorLabel.Text='';Update-Creation})}
 $rosterBox.Add_SelectedIndexChanged({$errorLabel.Text='';if($rosterBox.SelectedItem.Id -eq 'empty' -or ($rosterFile.Text -and [IO.Path]::GetExtension($rosterFile.Text).TrimStart('.').ToLowerInvariant() -ne $rosterBox.SelectedItem.Id)){$rosterFile.Text=''};Update-Creation})
 Update-Creation
 try{if($(Ui12Dialog $d;$d.ShowDialog($f)) -eq 'OK'){NOpen ([string]$d.Tag)}}finally{$d.Dispose()}
}
function NRestorePicker{
 DClean;$d=New-Object Windows.Forms.OpenFileDialog;$d.Title='완료된 project-backup.json 선택';$d.Filter='프로젝트 백업|project-backup.json';if($ws){$d.InitialDirectory=Join-Path $ws 'backups'};if($(Ui12Dialog $d;$d.ShowDialog($f)) -ne 'OK'){return}
 $default=Join-Path (Join-Path $base 'projects') ('복원-'+[guid]::NewGuid().ToString('N').Substring(0,8));$target=Input-Name '백업 복원' '아직 없는 새 프로젝트 최상위 폴더를 지정하세요. 원본은 보존합니다.' $default;if(!$target){return};$r=(NRun 'projectrestore' @{File=(Split-Path $d.FileName)} $target)|ConvertFrom-Json;NOpen $r.root
}
function NActivate($view){
 $script:nView=$view;$nHome.Visible=($view -eq 'project');$dashboard.Visible=($view -ne 'project');if($view -eq 'project'){$nHome.BringToFront()}else{$dashboard.BringToFront()};$nBar.BringToFront();if(Get-Command NRefreshActions -ErrorAction SilentlyContinue){NRefreshActions};foreach($k in $nTabs.Keys){$nTabs[$k].BackColor=if($k -eq $view){$blue}else{$pale};$nTabs[$k].ForeColor=if($k -eq $view){[Drawing.Color]::White}else{$navy}}
}
function NSelect($view){
 DClean;if($view -eq 'project'){NActivate $view;return};if(!$ws){$nCurrent.Text='① 먼저 새 프로젝트를 만들거나 기존 프로젝트를 불러오세요.';NActivate 'project';return}
 switch($view){
  'summary' {DLoad 'summary' '';NActivate 'summary'}
  'detail' {$a=if($dashArea){$dashArea}elseif($nLastArea){$nLastArea}else{$meta.courses[$dashProfileId].areas[0].root_item};DLoad 'detail' $a;NActivate 'detail'}
  'workspace' {if($nLastSpace -and $meta.layers[$nLastSpace]){$script:wTarget=$nLastSpace;DLoad 'workspace' $nLastArea;NActivate 'workspace'}else{throw '왼쪽 트리에서 하위 작업공간을 선택하거나 영역 세부에서 새 작업공간을 만드세요.'}}
  'student' {WStudentDialog}
  'import' {WImportDialog}
  'help' {Ui12Help}
 }
}
function NOnLoad{
 if($neisDownload){$show=($dashMode -eq 'summary' -and $meta.neisSources);$neisDownload.Visible=$show;$dashDown.Visible=!$show}
 if($dashArea){$script:nLastArea=$dashArea};if($dashMode -eq 'workspace'){$script:nLastSpace=$dashLayer};if($nBar){NActivate $dashMode}
}
NRecents;$nCurrent.Text='열린 프로젝트가 없습니다. 새 프로젝트 만들기 또는 프로젝트 불러오기를 선택하세요.';NActivate 'project'

$nDelete=DButton $nHome '프로젝트 삭제' 800 790 350 {DSafe {
 DClean;if(!$nRecent.CurrentRow){throw '삭제할 프로젝트를 목록에서 선택하세요.'};$selectedPath=[string]$nRecent.CurrentRow.Tag
 if(!$selectedPath){throw '목록에서 삭제할 프로젝트를 선택하세요.'}
 $label=[string]$nRecent.CurrentRow.Cells[0].Value
 $warning="선택 프로젝트: $label`r`n위치: $selectedPath`r`n`r`n정상 여부와 관계없이 최근 목록에서 제거합니다. 폴더가 남아 있으면 같은 상위 위치의 .project-trash로 통째로 보관합니다. 영구 삭제하지 않습니다. 계속하시겠습니까?"
 if([Windows.Forms.MessageBox]::Show($f,$warning,'프로젝트 삭제','YesNo','Warning','Button2') -ne 'Yes'){return}
 $result=(NRun 'projectarchive' @{PayloadFile=(NConfigFile @{confirmed=$true;selected_path=$selectedPath})} $selectedPath)|ConvertFrom-Json
 if($ws -and [string]::Equals($ws,$selectedPath,[StringComparison]::OrdinalIgnoreCase)){NCloseProject}
 $recentFile=Join-Path $base 'recent-projects.json';$remaining=@(NRecentEntries (Get-Content $recentFile -Raw -Encoding UTF8|ConvertFrom-Json)|Where-Object {![string]::Equals($_.path,$selectedPath,[StringComparison]::OrdinalIgnoreCase)})
 [IO.File]::WriteAllText($recentFile,(ConvertTo-Json -InputObject $remaining -Depth 5),[Text.UTF8Encoding]::new($true));NRecents;NActivate 'project';$nCurrent.Text=if($result.retained_path){'프로젝트를 목록에서 제거했습니다. 보관 위치: '+$result.retained_path}else{'폴더가 없어 최근 목록에서만 제거했습니다.'}
}}

function NCloseProject{
 $script:dashLoading=$true;$script:uLoading=$true;$script:wTreeLoading=$true
 try{
  $dashGrid.Rows.Clear();$dashGrid.Columns.Clear();$wTree.Nodes.Clear();MemoryClear
  $script:ws=$null;$script:meta=$null;$script:dashSnapshot=$null;$script:dashPending=@{};$script:dashRows=@{};$script:dashProfileId='';$script:dashLayer='';$script:dashArea='';$script:wTarget='';$script:nLastSpace='';$script:nLastArea='';$script:dashSelectedKey='';$script:uReadyLayer='';$nProjectLabel.Text='';$dashSelected.Text=''
 }finally{$script:wTreeLoading=$false;$script:uLoading=$false;$script:dashLoading=$false}
}

# Advisory states never disable a button for a missing workflow prerequisite.
# Busy operations still prevent concurrent writes; existing data validation remains.
function NActionHint($button,$ready,$reason,$caution=$false){
 if(!$button){return}
 $locked=($busy -or $uiActionDepth -gt 0 -or $uiLoadingDepth -gt 0);$button.Enabled=(!$locked)
 $kind=if($caution){'주의'}elseif($ready){'가능'}else{'먼저 확인'}
 $hint=$kind+' · '+$reason
 if($button.AccessibleDescription -ne $hint){
  $button.AccessibleDescription=$hint
  if($nActionTips){$nActionTips.SetToolTip($button,$hint)}
 }
 $button.BackColor=if($locked){[Drawing.Color]::FromArgb(225,229,236)}elseif($caution -or !$ready){[Drawing.Color]::FromArgb(255,231,170)}elseif($button.Parent -eq $nBar -and $button.Tag -ne $nView){$pale}else{$blue}
 $button.ForeColor=if(!$ready -or $caution){$navy}elseif($button.Parent -eq $nBar -and $button.Tag -ne $nView){$navy}else{[Drawing.Color]::White}
 if(Test-Path Function:\Ui12Button){Ui12Button $button $ready $caution $locked}
}
function NRefreshActions{
 if($nRecentLoading){return}
 $open=[bool]($ws -and $meta -and $dashProfileId);$hasStudents=($open -and $meta.students.Count -gt 0);$hasSpaces=($open -and $meta.layers.Count -gt 1);$queried=($open -and $null -ne $dashSnapshot)
 $selectedPath=if($nRecent.CurrentRow){[string]$nRecent.CurrentRow.Tag}else{''}
 $selected=!!$selectedPath
 NActionHint $nOpenSelected $selected $(if($selected){'선택 프로젝트 열기'}else{'목록에서 프로젝트를 선택하세요.'})
 NActionHint $nDelete $selected $(if($selected){'선택 프로젝트를 보관 폴더로 이동합니다. 확인창이 열립니다.'}else{'먼저 삭제할 프로젝트를 선택하세요.'}) $true
 NActionHint $nBackup $open $(if($open){'현재 프로젝트 백업'}else{'프로젝트를 먼저 여세요.'})
 foreach($k in $nTabs.Keys){$ok=($k -eq 'project' -or ($open -and ($k -ne 'workspace' -or $hasSpaces)));NActionHint $nTabs[$k] $ok $(if($ok){'화면 이동'}elseif(!$open){'프로젝트를 먼저 여세요.'}else{'하위 작업공간을 먼저 만들거나 선택하세요.'})}
 NActionHint $neisImportButton ($open -and $meta.layers.Count -eq 1) $(if(!$open){'빈 프로젝트를 먼저 만드세요.'}elseif($meta.layers.Count -ne 1){'현재 반입 방식은 하위 작업공간을 만들기 전 명부 반입을 지원합니다.'}else{'명부 파일을 선택해 학급을 추가하세요.'})
 $tableReason=if(!$open){'프로젝트를 먼저 여세요.'}elseif(!$queried){'조회 버튼으로 현재 조건의 표를 불러오세요.'}elseif(!$hasStudents){'명부를 먼저 반입하세요.'}else{'현재 조회한 학생과 표를 대상으로 실행합니다.'}
 foreach($b in @($dashDown,$dashUp,$dashSave,$wOne,$wAll,$neisDownload,$wRecords)){NActionHint $b ($hasStudents -and $queried) $tableReason}
 foreach($b in @($dashNew,$wNew)){NActionHint $b $queried $(if($queried){'현재 작업표에 항목을 추가합니다.'}else{'프로젝트를 열고 표를 조회하세요.'})}
 $cell=($queried -and $dashGrid.CurrentCell -and $dashGrid.CurrentCell.OwningColumn.Tag.item)
 NActionHint $attachExisting $cell $(if($cell){'선택한 항목에 하위 공간을 연결합니다.'}else{'연결할 점수 항목 셀을 먼저 선택하세요.'})
 $parent=($queried -and $hasStudents -and $meta.layers[$dashLayer].parent)
 NActionHint $wReflect $parent $(if($parent){'현재 하위 합계를 상위 표에 반영합니다. 반영 대상을 확인하세요.'}else{'학생이 있는 하위 작업표를 열고 조회하세요.'}) $true
 if($nRosterPrompt){$nRosterPrompt.Visible=($open -and !$hasStudents -and $dashMode -eq 'summary');NActionHint $nRosterPrompt $open '명부 반입 화면으로 이동합니다.';if($nRosterPrompt.Visible){$nRosterPrompt.BringToFront();$dashDown.Visible=$false}}
 if($nFlow){$nFlow.Text='파랑: 실행 가능 · 노랑: 선행 작업 확인/주의 (클릭 가능) · 마우스를 올리면 이유 표시'+[Environment]::NewLine+$(if(!$open){'① 새 프로젝트를 만들거나 기존 프로젝트를 불러오세요.'}elseif(!$hasStudents){'② 명부를 반입하세요. 빈 프로젝트 생성은 완료됐습니다.'}else{'③ 영역 선택 → 하위 작업표 또는 점수 입력 → 저장'});$nFlow.Height=40}
}
$nRecent.Add_SelectionChanged({if(Get-Command NRefreshActions -ErrorAction SilentlyContinue){NRefreshActions}})
NRefreshActions

$nRosterPrompt=DButton $dashboard '② 명부 반입' 770 175 150 {NActivate 'project'};$nRosterPrompt.Visible=$false
NRefreshActions

$nActionTips=New-Object Windows.Forms.ToolTip
$nActionTips.AutoPopDelay=15000;$nActionTips.InitialDelay=250;$nActionTips.ReshowDelay=100
$nActionTimer=New-Object Windows.Forms.Timer;$nActionTimer.Interval=400
$nActionTimer.Add_Tick({if(!$f.IsDisposed){NRefreshActions}})
$f.Add_Shown({$nActionTimer.Start();NRefreshActions})
$f.Add_FormClosed({$nActionTimer.Stop();$nActionTimer.Dispose();$nActionTips.Dispose()})
NRefreshActions
