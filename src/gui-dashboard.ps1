function DDownloadName {
 $profile=$meta.courses[$dashProfileId];$workspace=if($dashLayer -and $meta.layers[$dashLayer].name){[string]$meta.layers[$dashLayer].name}else{'작업표'}
 $className=if($dashSnapshot.class){[string]$dashSnapshot.class}else{'전체학급'}
 $parts=@('수행평가',[string]$profile.subject,$workspace,$className)
 $clean=@($parts|ForEach-Object {$part=([regex]::Replace($_,'[\x00-\x1f<>:"/\\|?*]','-')).Trim().TrimEnd('.');if($part.Length -gt 45){$part=$part.Substring(0,45).TrimEnd()};if($part){$part}})
 return (($clean -join '_')+'_'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.xlsx')
}
# Teacher-facing table. The existing workspace tools remain behind the detail view.
$script:dashLoading=$true;$script:dashPending=@{};$script:dashRows=@{};$script:dashSnapshot=$null;$script:dashMode='summary';$script:dashArea='';$script:dashProfileId='';$script:dashSelectedKey=''
$blue=[Drawing.Color]::FromArgb(59,103,244);$navy=[Drawing.Color]::FromArgb(22,49,88);$pale=[Drawing.Color]::FromArgb(242,245,250)
$legacyControls=@($f.Controls);foreach($control in $legacyControls){$control.Visible=$false}
$dashboard=New-Object Windows.Forms.Panel;$dashboard.Dock='Fill';$dashboard.BackColor=[Drawing.Color]::White;$f.Controls.Add($dashboard);$dashboard.BringToFront()
$f.Text=$productTitle;$f.MinimumSize=New-Object Drawing.Size(1300,920)
function DButton($hostControl,$text,$x,$y,$width,$handler){$c=Button $hostControl $text $x $y $width $handler;$c.FlatStyle='Flat';$c.FlatAppearance.BorderSize=0;$c.BackColor=$blue;$c.ForeColor=[Drawing.Color]::White;return $c}
function DClean{[void]$dashGrid.EndEdit();if($dashPending.Count){throw '저장하지 않은 변경이 있습니다. 먼저 저장하거나 변경 취소를 누르세요.'}}
$dashTitle=Label $dashboard $productTitle 20 15 580 35;$dashTitle.Font=New-Object Drawing.Font('맑은 고딕',16,[Drawing.FontStyle]::Bold);$dashTitle.ForeColor=$navy
$productCaption=Label $dashboard $productDescription 310 10 605 24;$productCaption.ForeColor=[Drawing.Color]::FromArgb(90,104,122)
$side=New-Object Windows.Forms.Panel;$side.SetBounds(0,65,190,800);$side.BackColor=$pale;$dashboard.Controls.Add($side)
[void](Label $side '평가 업무' 15 15 160 30)

$dashFilters=New-Object Windows.Forms.Panel;$dashFilters.SetBounds(210,65,1045,100);$dashFilters.BackColor=$pale;$dashboard.Controls.Add($dashFilters)
foreach($spec in @(@('학년도',10,10,55),@('학기',190,10,40),@('학년',355,10,40),@('반',515,10,30),@('과목',665,10,40))){[void](Label $dashFilters $spec[0] $spec[1] $spec[2] $spec[3])}
$dy=Combo $dashFilters 70 6 105;$dt=Combo $dashFilters 235 6 105;$dg=Combo $dashFilters 400 6 100;$dc=Combo $dashFilters 550 6 100;$ds=Combo $dashFilters 710 6 135
[void](Label $dashFilters '기준일' 10 58 55);$ddate=DateBox $dashFilters 70 50 150
$dashContext=Label $dashFilters '작업공간을 열어 주세요.' 245 55 650 35
$dashQuery=DButton $dashFilters '조회' 930 6 100 {DSafe {DClean;$nextProfile=DFindProfile;if($nextProfile -ne $dashProfileId){DInitialize $nextProfile}else{DLoad $dashMode $dashArea}}}
$dashHeading=Label $dashboard '수행평가 종합표' 210 178 550 35;$dashHeading.Font=New-Object Drawing.Font('맑은 고딕',13,[Drawing.FontStyle]::Bold);$dashHeading.ForeColor=$navy
$dashDown=DButton $dashboard 'Excel 다운로드' 770 175 150 {DSafe {DClean;if(!$dashSnapshot){throw '먼저 조회하세요.'};$d=New-Object Windows.Forms.SaveFileDialog;$d.Filter='Excel 성적표|*.xlsx';$d.FileName=DDownloadName;if($d.ShowDialog($f) -eq 'OK'){$r=Run 'tableexport' (DArgs @{File=$d.FileName});$dashMessage.Text='조회한 학급·항목의 Excel을 새 파일로 저장했습니다.'}}}
$dashUp=DButton $dashboard 'Excel 값 반영' 930 175 150 {DSafe {DClean;if(!$dashSnapshot){throw '먼저 조회하세요.'};$d=New-Object Windows.Forms.OpenFileDialog;$d.Filter='Excel 자료|*.xlsx';$d.Title='학생코드 또는 명부가 담긴 Excel 반입';if($d.ShowDialog($f) -eq 'OK'){$p=(Run 'tablepreview' (DArgs @{File=$d.FileName}))|ConvertFrom-Json;foreach($issue in @($p.skipped|Where-Object {$_})){Log ('반입 제외: 행 '+$issue.row+' / 열 '+$issue.column+' / '+$issue.reason)};if(!$p.changes.Count){$dashMessage.Text='반영할 값이 없습니다. 제외된 행·셀 '+@($p.skipped|Where-Object {$_}).Count+'건.';return};$saved=(Run 'tablecommit' (DArgs @{File=$d.FileName;PreviewFingerprint=$p.fingerprint}))|ConvertFrom-Json;DLoad $dashMode $dashArea;$dashMessage.Text='Excel 값 '+$saved.saved+'건 / 새 열 '+$saved.new_columns+'개 반입 / 제외 '+@($p.skipped|Where-Object {$_}).Count+'건. 상위 값은 변경하지 않았습니다.'}}}
$dashSave=DButton $dashboard '변경 저장' 1090 175 145 {DSafe {DSave}}
$dashNew=DButton $dashboard '열 추가' 505 218 180 {DSafe {WNewColumn}}
$dashCancel=DButton $dashboard '변경 취소' 700 218 150 {DSafe {if($dashPending.Count -and [Windows.Forms.MessageBox]::Show($f,'화면에서 수정한 미저장 변경을 취소할까요? 저장된 자료는 유지됩니다.','미저장 변경 취소','OKCancel') -ne 'OK'){return};[void]$dashGrid.CancelEdit();$script:dashPending=@{};DLoad $dashMode $dashArea}};$dashCancel.CausesValidation=$false
$dashCount=Label $dashboard '' 870 227 370 30
$dashGrid=Grid $dashboard 210 270 1045 390 @();$dashGrid.ShowCellToolTips=$false;$dashGrid.AllowUserToOrderColumns=$true;$dashGrid.ReadOnly=$false;$dashGrid.MultiSelect=$false;$dashGrid.SelectionMode='CellSelect';$dashGrid.EditMode='EditOnEnter';$dashGrid.AllowUserToResizeRows=$false;$dashGrid.RowTemplate.Height=32;$dashGrid.ColumnHeadersHeight=76;$dashGrid.EnableHeadersVisualStyles=$false;$dashGrid.ColumnHeadersDefaultCellStyle.BackColor=[Drawing.Color]::FromArgb(235,239,246);$dashGrid.ColumnHeadersDefaultCellStyle.ForeColor=$navy;$dashGrid.DefaultCellStyle.SelectionBackColor=[Drawing.Color]::FromArgb(219,232,255);$dashGrid.DefaultCellStyle.SelectionForeColor=$navy
$dashSelected=Label $dashboard '학생의 점수 또는 기록을 선택하세요.' 210 675 1045 45

$dashMessage=Label $dashboard '가상 예제를 열거나 작업공간을 선택하세요.' 210 805 1045 45;$dashMessage.ForeColor=$navy
$dashLegacyBack=DButton $f '수행평가 화면으로' 1010 15 245 {DSafe {DBack}};$dashLegacyBack.Visible=$false
function DLayout{
 $w=$dashboard.ClientSize.Width;$h=$dashboard.ClientSize.Height-52;$rw=[math]::Max(1010,$w-230);$dashFilters.Width=$rw;$side.Height=$h-65;$dashGrid.SetBounds(210,270,$rw,[math]::Max(300,$h-490));$dashSelected.SetBounds(210,$h-205,$rw,45);$dashMessage.SetBounds(210,$h-65,$rw,45)
 $dashDown.Left=$w-515;$dashUp.Left=$w-355;$dashSave.Left=$w-195;
}
$dashboard.Add_Resize({DLayout});DLayout
function DFindProfile{
 if(!$meta -or !$meta.courses){throw '프로젝트의 학년도·과목 설정을 확인하세요.'}
 $matches=@($meta.courses.Values|Where-Object {$_.year -eq $dy.SelectedItem.Id -and $_.term -eq $dt.SelectedItem.Id -and $_.grade -eq $dg.SelectedItem.Id -and $_.subject -eq $ds.SelectedItem.Id});if($matches.Count -ne 1){throw '선택 조건의 교육과정이 없거나 중복입니다.'};return $matches[0].id
}
function DInitialize($selected){
 $script:dashLoading=$true;Read-Meta;foreach($c in @($dy,$dt,$dg,$dc,$ds)){$c.Items.Clear()};$script:dashPending=@{};$script:dashSnapshot=$null;$dashGrid.Rows.Clear();$dashGrid.Columns.Clear()
 if(!$meta.courses -or !$meta.courses.Count){$script:dashLoading=$false;$dashContext.Text='평가영역이 있는 프로젝트를 불러오세요.';return}
 $profiles=@($meta.courses.Values);if(!$selected){$selected=($profiles|Sort-Object year,term,grade,subject|Select-Object -First 1).id};$p=$meta.courses[$selected]
 foreach($v in @($profiles.year|Sort-Object -Unique)){Choice $dy $v $v};foreach($v in @($profiles.term|Sort-Object -Unique)){Choice $dt $v ($v+'학기')};foreach($v in @($profiles.grade|Sort-Object -Unique)){Choice $dg $v ($v+'학년')};foreach($v in @($profiles.subject|Sort-Object -Unique)){Choice $ds $v $v};Choose $dy $p.year;Choose $dt $p.term;Choose $dg $p.grade;Choose $ds $p.subject
 Choice $dc '' '전체 학급';$classes=if($p.classes){@($p.classes)}else{@($meta.memberships|ForEach-Object {$_.class}|Sort-Object -Unique)};foreach($v in $classes){Choice $dc $v $v};$dc.SelectedIndex=if($dc.Items.Count -gt 1){1}else{0}
 $script:dashProfileId=$selected;$script:dashLoading=$false;DLoad 'summary' ''
}
function DArgs($extra=@{}){$a=@{Name=$dashProfileId;Layer=$dashLayer;Area=$dashArea;Class=$dashSnapshot.class;AsOf=$dashSnapshot.as_of;Student=$wStudent};foreach($k in $extra.Keys){$a[$k]=$extra[$k]};return $a}
function DLoad($mode,$area){
 $script:uReadyLayer=''
 if($busy){return}
 if(!$ws){throw '작업공간을 먼저 여세요.'};MemoryEnsure;$profileId=DFindProfile;$p=$meta.courses[$profileId];$script:dashLoading=$true;$script:dashProfileId=$profileId;$script:dashMode=$mode;$script:dashArea=$area;$script:dashSelectedKey='';$script:dashPending=@{};$script:dashRows=@{};$script:dashSnapshot=$null;$dashGrid.Rows.Clear();$dashGrid.Columns.Clear()
 try{
  $script:dashLayer=$p.root;$heading='수행평가 종합표';$dashNew.Visible=$true
  if($mode -in @('detail','workspace')){$a=WDisplayArea $p $area;if(!$a){throw '평가영역을 선택하세요.'};$heading='종합표  ›  '+$a.label;$script:dashLayer=if($mode -eq 'workspace'){$wTarget}else{$a.work_layer};if($mode -eq 'workspace'){$heading+='  ›  '+(LName $wTarget)};$dashNew.Text='열 추가'}
  $dashHeading.Text=$(if($mode -eq 'workspace'){WPath $dashLayer}else{$heading});$dashContext.Text=$p.year+'학년도 '+$p.term+'학기 '+$p.grade+'학년 '+$p.subject+' / '+$dc.SelectedItem.Text;$dashCount.Text=''
  if(!$dashLayer){$dashMessage.Text='이 영역에는 작업표가 없습니다. 종합표에서 해당 영역을 선택하고 [선택 항목에 하위 공간 만들기]를 누르세요.';return}
  $q=MemorySnapshot @{Name=$profileId;Layer=$dashLayer;Area=$area;Class=$dc.SelectedItem.Id;AsOf=$ddate.Value.ToString('yyyy-MM-dd');Student=$wStudent};$script:dashSnapshot=$q
  foreach($spec in @(@('class','반',45),@('number','번호',60),@('name','이름',130))){$c=New-Object Windows.Forms.DataGridViewTextBoxColumn;$c.Name=$spec[0];$c.HeaderText=$spec[1];$c.AutoSizeMode='None';$c.Width=$spec[2];$c.ReadOnly=$true;$c.Frozen=$true;[void]$dashGrid.Columns.Add($c)}
  $sumCol=New-Object Windows.Forms.DataGridViewTextBoxColumn;$sumCol.Name='total';$sumCol.HeaderText=if($mode -eq 'summary'){'합계'}else{"합계`n(상위 반영)"};$sumCol.ReadOnly=$true;$sumCol.Frozen=$true;$sumCol.AutoSizeMode='None';$sumCol.Width=110;[void]$dashGrid.Columns.Add($sumCol)
  foreach($item in $q.fields){$title=if($mode -eq 'summary'){@($p.areas|Where-Object root_item -EQ $item)[0].label}else{FName $dashLayer $item}
   $c=New-Object Windows.Forms.DataGridViewTextBoxColumn;$c.Name='v|'+$item;$c.HeaderText=$title+"`n"+$(if($mode -eq 'summary'){'점수'}elseif($item -in @($p.numeric_items)){'점수'}else{'기록'});$c.FillWeight=110;$c.Tag=@{item=$item;kind='value'};[void]$dashGrid.Columns.Add($c)

  }

  $ids=@($q.data|Group-Object { [string]$_.student }|ForEach-Object {$_.Group[0]}|Sort-Object class,@{Expression={if([string]$_.number -match '^\d+$'){[long]$_.number}else{[long]::MaxValue}}},student|ForEach-Object {$_.student});foreach($sid in $ids){$records=@($q.data|Where-Object student -EQ $sid);$r=$dashGrid.Rows[$dashGrid.Rows.Add()];$r.Tag=$sid;$r.Cells['class'].Value=$records[0].class;$r.Cells['number'].Value=$records[0].number;$r.Cells['name'].Value=$records[0].name
   foreach($x in $records){$key=$sid+'|'+$x.item;$script:dashRows[$key]=$x;$cell=$r.Cells['v|'+$x.item];$cell.Value=if($null -eq $x.value){''}else{[string]$x.value};$cell.ReadOnly=$false;$cell.ToolTipText=$x.origin}

   DTotal $r
  }
  $dashCount.Text='조회 '+$ids.Count+'명 · 미저장 0건';$dashMessage.Text=if($q.warnings.Count){'식별자 경고 '+$q.warnings.Count+'건. 자료·연결 상세에서 확인하세요.'}elseif($mode -eq 'summary'){'연결된 색상 머리글을 두 번 누르면 하위 작업표로 이동합니다. 열 제목 ▼는 정렬·검색입니다. 불일치 방향은 표 위 색상 안내를 확인하세요.'}else{'드래그·Shift·Ctrl로 여러 셀 선택 / Ctrl+V로 표 붙여넣기 / 변경 저장. 상위·하위 불일치는 색으로 표시합니다.'}
 }finally{$script:dashLoading=$false};if(Get-Command WRefreshTree -ErrorAction SilentlyContinue){WRefreshTree;WIdentityApply};DSelection;if(Get-Command NOnLoad -ErrorAction SilentlyContinue){NOnLoad}
}

function DChange($row,$item){
 $key=$row.Tag+'|'+$item;$old=$dashRows[$key];if(!$old){return};$text=[string]$row.Cells['v|'+$item].Value;$v=$null
 if($text -ne ''){if($old.numeric){$num=0.0;if(![double]::TryParse($text,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$num)){throw '점수는 숫자로 입력하세요.'};$v=$num}else{$v=$text}}
 $change=@{student=$row.Tag;item=$item;value=$v;status=$(if($null -eq $v){'미입력'}else{'입력됨'});note='';before=$old.value;beforeStatus='';name=$old.name;label=$old.label}
 if(DValueEqual $v $old.value){$dashPending.Remove($key)}else{$dashPending[$key]=$change}
 if($dashPasting){return}
 if(Get-Command LightCompare -ErrorAction SilentlyContinue){LightCompare}
 $dashCount.Text='조회 '+$dashGrid.Rows.Count+'명 · 미저장 '+$dashPending.Count+'건';DTotal $row;$dashGrid.Invalidate()
}
function DValueEqual($a,$b){if($null -eq $a -or $null -eq $b){return ($null -eq $a -and $null -eq $b)};if($a -is [ValueType] -and $b -is [ValueType]){return ([double]$a -eq [double]$b)};return ([string]$a -ceq [string]$b)}

function DSave{
 if(!$dashGrid.EndEdit()){throw '입력 중인 셀의 형식을 확인하세요.'}
 if(!$dashSnapshot -or !$dashPending.Count){
  $dashMessage.Text=if($script:memoRoot -eq $ws -and $script:memoData.Count){'저장할 변경이 없습니다. 메모는 이미 저장되어 있습니다.'}else{'저장할 변경이 없습니다.'}
  return
 };$changes=@($dashPending.Values)
 $req=Write-Request @{expected_fingerprint=$dashSnapshot.fingerprint;changes=$changes};$saved=(Run 'gridcommit' (DArgs @{PayloadFile=$req}))|ConvertFrom-Json;$script:dashPending=@{};DLoad $dashMode $dashArea;$dashMessage.Text='변경 '+$saved.saved+'건 저장 완료 · 기존 기록·수정 이력 보존'
}
function DAdvanced($kind){Refresh-Navigation $dashLayer;$tabs.SelectedTab=$stagingTab;$dashboard.Visible=$false;foreach($control in $legacyControls){$control.Visible=$true};$dashLegacyBack.Visible=$true;$dashLegacyBack.BringToFront()}
function DBack{$dashLegacyBack.Visible=$false;foreach($control in $legacyControls){$control.Visible=$false};$dashboard.Visible=$true;$dashboard.BringToFront();Read-Meta;if($dashProfileId){DLoad $dashMode $dashArea}}
function DFilterChanged{
 if($dashLoading -or !$dashSnapshot){return};if($dashPending.Count){$script:dashLoading=$true;$p=$meta.courses[$dashProfileId];Choose $dy $p.year;Choose $dt $p.term;Choose $dg $p.grade;Choose $ds $p.subject;Choose $dc $dashSnapshot.class;$ddate.Value=[datetime]$dashSnapshot.as_of;$script:dashLoading=$false;$dashMessage.Text='미저장 변경이 있어 조회 조건을 유지했습니다. 먼저 저장하거나 취소하세요.';return}
 $script:dashSnapshot=$null;$script:dashLoading=$true;try{$dashGrid.Rows.Clear();$dashGrid.Columns.Clear()}finally{$script:dashLoading=$false};$dashSelected.Text='조회 조건이 바뀌었습니다. 조회를 누르세요.';$dashMessage.Text='현재 표를 새 조건으로 다시 조회해 주세요.';if(Get-Command NRefreshActions -ErrorAction SilentlyContinue){NRefreshActions}
}
foreach($combo in @($dy,$dt,$dg,$dc,$ds)){$combo.Add_SelectedIndexChanged({DFilterChanged})};$ddate.Add_ValueChanged({DFilterChanged})

$dashGrid.Add_CellValidating({param($sender,$e)if(!$dashLoading -and $e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0){$tag=$dashGrid.Columns[$e.ColumnIndex].Tag;if($tag -and $tag.kind -eq 'value'){$key=$dashGrid.Rows[$e.RowIndex].Tag+'|'+$tag.item;$old=$dashRows[$key];$text=[string]$e.FormattedValue;$num=0.0;if($old.numeric -and $text -ne '' -and (![double]::TryParse($text,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$num) -or [double]::IsNaN($num) -or [double]::IsInfinity($num))){$e.Cancel=$true;$dashMessage.Text='점수는 유한한 숫자로 입력하세요.'}}}})
$dashGrid.Add_CellValueChanged({param($sender,$e) if(!$dashLoading -and $e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0){$tag=$dashGrid.Columns[$e.ColumnIndex].Tag;if($tag){DSafe {DChange $dashGrid.Rows[$e.RowIndex] $tag.item}}}})
$dashGrid.Add_DataError({param($sender,$e)$e.ThrowException=$false;$dashMessage.Text='입력 형식 또는 값을 확인하세요.'})
$dashGrid.Add_SelectionChanged({DSelection})
$script:dashLoading=$false
$dashNew.Visible=$false
$f.Add_FormClosing({param($sender,$e)if($dashPending.Count -and [Windows.Forms.MessageBox]::Show($f,'저장하지 않은 변경을 남겨 두고 종료할까요?','미저장 변경','YesNo') -ne 'Yes'){$e.Cancel=$true}})
