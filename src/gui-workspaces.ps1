function WColumnNames{
 $d=New-Object Windows.Forms.Form;$d.Text='열 여러 개 추가';$d.Size=New-Object Drawing.Size(680,440);$d.StartPosition='CenterParent';$d.Font=$f.Font
 [void](Label $d '열 이름을 한 줄에 하나씩 입력하세요. 쉼표·세미콜론으로 구분해도 됩니다.' 20 15 630 45)
 $names=TextBox $d 20 65 620 200;$names.Multiline=$true;$names.AcceptsReturn=$true;$names.ScrollBars='Vertical';$names.Text='새 평가항목'
 $info=Label $d '' 20 275 620 40
 $names.Add_TextChanged({$n=@($names.Text -split '[,;，；\r\n]'|ForEach-Object {$_.Trim()}|Where-Object {$_});$info.Text='추가할 열 '+$n.Count+'개 (다음 단계에서 숫자점수 / 문자기록 선택)'})
 DButton $d '다음' 360 330 130 {$n=@($names.Text -split '[,;，；\r\n]'|ForEach-Object {$_.Trim()}|Where-Object {$_});if(!$n.Count -or $n.Count -gt 200 -or @($n|Sort-Object -Unique).Count -ne $n.Count -or @($n|Where-Object {$_.Length -gt 80}).Count){$info.Text='중복 없이 1~200개, 각 이름은 80자 이하로 입력하세요.';return};$d.Tag=$n;$d.DialogResult='OK';$d.Close()}|Out-Null
 $cancel=DButton $d '취소' 510 330 130 {$d.DialogResult='Cancel';$d.Close()};$d.CancelButton=$cancel
 try{if($(Ui12Dialog $d;$d.ShowDialog($f)) -eq 'OK'){return ,@($d.Tag)};return $null}finally{$d.Dispose()}
}
$script:wTarget='';$script:wStudent='';$script:wIdentityMode=0;$script:wCodes=@{};$script:wTreeLoading=$false;$script:wLastTicket=''

$wTree=New-Object Windows.Forms.TreeView;$wTree.SetBounds(8,48,174,360);$wTree.HideSelection=$false;$wTree.ShowNodeToolTips=$true;$side.Controls.Add($wTree)
$wTree.Indent=12;$wTree.Font=New-Object Drawing.Font('맑은 고딕',9)
foreach($button in @($side.Controls|Where-Object {$_ -is [Windows.Forms.Button]})){
 if($button.Text -eq '자료 반입·작업공간'){$button.Visible=$false}
 if($button.Text -eq '하위자료 반영 검토'){$button.Visible=$false}
}
foreach($label in @($side.Controls|Where-Object {$_ -is [Windows.Forms.Label] -and $_.Top -gt 100})){$label.Visible=$false}
$wNew=DButton $side '하위 공간 만들기' 8 420 174 {DSafe {WNewSpace}}
$wRename=DButton $side '이름 바꾸기' 8 462 174 {DSafe {DClean;if($dashMode -ne 'workspace'){throw '이름을 바꿀 하위 작업공간을 선택하세요.'};$label=Input-Name '작업공간 이름' '표시명만 변경합니다. 학생 ID·자료 연결은 유지됩니다.' (LName $dashLayer);if(!$label){return};Run 'spacerename' (DArgs @{PayloadFile=(Write-Request @{label=$label})})|Out-Null;Read-Meta;DLoad 'workspace' $dashArea}}
$wImport=DButton $side '자료 가져오기' 8 514 174 {DSafe {WImportDialog}}
$wReflect=DButton $side '상위에 반영' 8 556 174 {DSafe {WParentDialog}}
$wRecords=DButton $side '기록 내보내기' 8 608 174 {DSafe {StudentExportDialog}}

$wModeButton=DButton $dashboard '이름  ›' 450 12 260 {DSafe {$script:wIdentityMode=($wIdentityMode+1)%3;WIdentityApply;DSelection}}
$dashTitle.Width=400;$wModeButton.BringToFront()
$wOne=DButton $dashboard '선택 학생 보기' 870 218 155 {DSafe {DClean;if(!$dashGrid.CurrentRow){throw '표에서 학생 한 명을 선택하세요.'};$script:wStudent=[string]$dashGrid.CurrentRow.Tag;DLoad $dashMode $dashArea}}
$wAll=DButton $dashboard '학급 전체' 1035 218 110 {DSafe {DClean;$script:wStudent='';DLoad $dashMode $dashArea}}
$dashCount.SetBounds(210,837,1000,22);$dashCount.Visible=$false
function WStudentLabel($sid,$fallback=''){
 if(!$sid){return $fallback};$script:wCodes[$sid]=StudentCode $ws $sid
 if($wIdentityMode -eq 2){return $wCodes[$sid]};$s=$meta.students[$sid];if(!$s){return $fallback};$label=[string]$s.number+'번 '+$s.name;if($wIdentityMode -eq 1){$label+=' / '+$wCodes[$sid]};return $label
}
function WIdentityApply{
 $wModeButton.Text=@('이름  ›','이름+코드  ›','코드  ›')[$wIdentityMode]
 if(!$dashGrid.Columns.Contains('view_code')){$c=New-Object Windows.Forms.DataGridViewTextBoxColumn;$c.Name='view_code';$c.HeaderText='코드';$c.ReadOnly=$true;$c.Frozen=$true;$c.AutoSizeMode='None';$c.Width=280;$dashGrid.Columns.Insert([math]::Min(3,$dashGrid.Columns.Count),$c)}
 $dashGrid.Columns['view_code'].Visible=($wIdentityMode -ne 0)
 foreach($name in @('class','number','name')){if($dashGrid.Columns.Contains($name)){$dashGrid.Columns[$name].Visible=($wIdentityMode -ne 2)}}
 foreach($row in $dashGrid.Rows){[void](WStudentLabel $row.Tag);$row.Cells['view_code'].Value=$wCodes[$row.Tag]}
 $dashDown.Text='Excel 다운로드'
 $dashNew.Visible=$true;$wOne.Text=if($wStudent){'학생 1명 조회 중'}else{'선택 학생 보기'}
 if($dashSnapshot){$dashContext.Text=$dashSnapshot.profile.year+'학년도 '+$dashSnapshot.profile.term+'학기 '+$dashSnapshot.profile.grade+'학년 '+$dashSnapshot.profile.subject+' / '+$dc.SelectedItem.Text+$(if($wStudent){' / '+(WStudentLabel $wStudent)}else{' / 학급 전체'})+' / '+$dashGrid.Rows.Count+'명'}
}
function WEligible($id,$area){
 $scope=if($meta.workspaceScopes){$meta.workspaceScopes[$id]}else{$null};if($scope){return ($scope.course -eq $dashProfileId -and $scope.area -eq $area)}
 $prefix=$area.Split('.')[0];return (@($meta.layers[$id].items|Where-Object {$_.Split('.')[0] -eq $prefix}).Count -gt 0)
}
function WRefreshTree{
 if(!$meta -or !$dashProfileId){return};$script:wTreeLoading=$true;$expanded=@{}
 function Save-Expansion($nodes){foreach($n in $nodes){if($n.IsExpanded){$expanded[$n.Name]=$true};Save-Expansion $n.Nodes}}
 Save-Expansion $wTree.Nodes;$wTree.BeginUpdate();$wTree.Nodes.Clear();$p=$meta.courses[$dashProfileId]
 $homeNode=New-Object Windows.Forms.TreeNode '수행평가 종합표';$homeNode.Name='summary';$homeNode.Tag=@{mode='summary';area='';layer=$p.root};[void]$wTree.Nodes.Add($homeNode)
 function Add-WorkspaceChildren($parentNode,$parentId,$area,$seen){
  if($seen.ContainsKey($parentId)){return};$seen[$parentId]=$true
  foreach($id in @($meta.layers.Keys|Where-Object {$meta.layers[$_].parent -eq $parentId}|Sort-Object {$meta.layers[$_].name})){
   if(!(WEligible $id $area)){continue};$node=New-Object Windows.Forms.TreeNode (LName $id);$node.Name=$area+'|'+$id;$node.Tag=@{mode='workspace';layer=$id;area=$area};$node.ToolTipText=(LName $id)+' / 모든 반 공통 · 학급은 조회 조건입니다.';[void]$parentNode.Nodes.Add($node);Add-WorkspaceChildren $node $id $area $seen;if($expanded[$node.Name]){$node.Expand()}
  }
 }
 foreach($a in $p.areas){$n=New-Object Windows.Forms.TreeNode $a.label;$n.Name=$a.root_item;$n.Tag=@{mode='detail';area=$a.root_item;layer=$a.work_layer};[void]$homeNode.Nodes.Add($n);if($a.work_layer){Add-WorkspaceChildren $n $a.work_layer $a.root_item @{}};if($expanded[$n.Name]){$n.Expand()}}
 if($meta.workspaceScopes){$loose=@($meta.workspaceScopes.Keys|Where-Object {$meta.workspaceScopes[$_].unlinked -and $meta.workspaceScopes[$_].course -eq $dashProfileId -and $meta.layers[$_]});if($loose.Count){$group=New-Object Windows.Forms.TreeNode '미연결 작업공간';$group.Name='unlinked';[void]$wTree.Nodes.Add($group);foreach($id in $loose){$ar=$meta.workspaceScopes[$id].area;$node=New-Object Windows.Forms.TreeNode (LName $id);$node.Name=$ar+'|'+$id;$node.Tag=@{mode='workspace';layer=$id;area=$ar};[void]$group.Nodes.Add($node);Add-WorkspaceChildren $node $id $ar @{}};$group.Expand()}}
 $homeNode.Expand();function Select-CurrentNode($nodes){foreach($n in $nodes){if($n.Tag.mode -eq $dashMode -and $n.Tag.layer -eq $dashLayer -and $n.Tag.area -eq $dashArea){$up=$n.Parent;while($up){$up.Expand();$up=$up.Parent};$wTree.SelectedNode=$n;$n.EnsureVisible()};Select-CurrentNode $n.Nodes}};Select-CurrentNode $wTree.Nodes;$wTree.EndUpdate();$script:wTreeLoading=$false
}
$wTree.Add_AfterSelect({param($sender,$e)if(!$wTreeLoading -and !$busy -and $e.Node.Tag){$navCallback=[Action[object]]{param($nav)if(!$busy -and $ws -and $meta -and $meta.layers.ContainsKey($nav.layer)){DSafe {DClean;$script:wTarget=$nav.layer;DLoad $nav.mode $nav.area}}};[void]$f.BeginInvoke($navCallback,[object[]]@($e.Node.Tag))}})
function WPath($layer){$parts=@();$seen=@{};while($layer){if($seen[$layer]){break};$seen[$layer]=$true;$parts=@((LName $layer))+$parts;$layer=$meta.layers[$layer].parent};return (($parts | Select-Object -Last 3) -join ' > ')}
function WNewColumn{
 DClean;if(!$dashSnapshot){throw '먼저 작업공간을 조회하세요.'}
 $labels=WColumnNames;if(!$labels){return}
 $choice=WColumnMode;if(!$choice){return}
 $numeric=$true;if($choice -eq 'direct'){$kind=WColumnKind;if(!$kind){return};$numeric=($kind -eq 'numeric')}
 $r=(Run 'spacecolumn' (DArgs @{PayloadFile=(Write-Request @{labels=@($labels);numeric=$numeric;child=($choice -eq 'child')})}))|ConvertFrom-Json;Read-Meta
 if($r.child){$script:wTarget=$r.child.layer;if($dashMode -eq 'summary'){$script:dashArea=$r.item};DLoad 'workspace' $dashArea}else{DLoad $dashMode $dashArea}
 $dashMessage.Text='열 '+$r.added+'개 생성 완료 · 기존 기록 유지 · 상위 점수 변경 없음'
}
function WNewSpace{WNewColumn}
function WColumnKind{
 $d=New-Object Windows.Forms.Form;$d.Text='열 종류';$d.ClientSize=New-Object Drawing.Size(620,160);$d.StartPosition='CenterParent';$d.Font=$f.Font
 [void](Label $d '입력할 자료의 종류를 선택하세요.' 20 20 580 40)
 DButton $d '숫자점수' 20 90 180 {$d.Tag='numeric';$d.Close()}|Out-Null
 DButton $d '문자관찰기록' 215 90 230 {$d.Tag='text';$d.Close()}|Out-Null
 $cancel=DButton $d '취소' 460 90 140 {$d.Close()};$d.CancelButton=$cancel
 [void]$(Ui12Dialog $d;$d.ShowDialog($f));return [string]$d.Tag
}
function WFriendly($errorText){
 $technical=[string]$errorText
 $raw=($technical -split '(?m)^\s*(?:위치\s|At\s|\+\s*CategoryInfo|\+\s*FullyQualifiedErrorId)')[0].Trim()
 $problem=$raw;$action='입력한 값과 현재 작업공간을 확인한 뒤 다시 실행하세요.'
 if($raw -match 'Invoke|호출하는 동안|Exception|위치:'){$problem='선택한 자료의 확인 또는 저장을 마치지 못했습니다. 아래 위치와 조치를 확인하세요.'}
 if($raw -match '평가항목이 변경'){$problem='표에서 필요한 평가항목을 찾을 수 없습니다.';$action='원본 Excel의 첫 행 항목 이름을 임의로 바꾸었는지 확인하고 복구한 뒤 [조회]하세요.'}
 if($raw -match '조회|원본|바뀌|fingerprint|지문'){$action='Excel을 저장하고 닫은 뒤 현재 작업공간에서 [조회]를 눌러 최신 값으로 다시 확인하세요.'}

 if($raw -match '학생|ID|코드|중복|누락'){$action='문제 행의 식별정보를 원문과 대조하세요. 반입 파일은 [자료 가져오기]에서 다시 미리보기하거나 [예외 행 해결]에서 학생을 명시적으로 연결하세요.'}

 if($raw -match '문자·자유서술|수식·문자'){$action='숫자 평가항목만 있는 작업표를 선택하세요. 자유서술·메모는 제외하고 식별정보 위험이 불명확한 자료의 반출은 보류합니다.'}
 if($raw -match '프로젝트|명부 CSV|같은 반 번호|백업 파일|데이터 파일'){$action='프로젝트 탭에서 선택한 폴더와 명부 CSV의 name,class,number 열·행을 확인하세요. 기존 폴더는 덮어쓰지 말고 새 경로로 다시 생성하거나 올바른 프로젝트를 불러오세요.'}
 if($raw -match '현재 표에서 학생|학생 한 명을 선택'){$action='현재 점수표에서 학생 행을 한 번 클릭한 뒤 하단 학생 기록을 누르세요. 선택 학생 보기로 한 명만 제한할 수도 있습니다.'}
 if($raw -match '8001010[Dd]|RPC_E_CANTCALLOUT'){$problem='화면 전환 중 Excel 연결을 시작하지 못했습니다.';$action='안내 창을 닫고 현재 작업공간에서 조회를 다시 누르세요. 반복되면 프로그램을 다시 실행하세요. 새 연결·반영을 중복 실행하지 마세요.'}
 if($raw -match '활성 연결|공유 연결|내부 하위 작업공간|종합표는 휴지통'){$problem=$raw;$action='공간 전체 이동은 연결 보호 때문에 제한됩니다. 상위 작업표의 해당 항목 열과 연결 위치를 확인하세요. 하위 공간은 보존한 채 미연결로 만든 후 이동할 수 있습니다.'}
 $dataState='처리를 완료하지 못했습니다. 원본·변경 전 복구 사본을 보존합니다. 정상·격리 건수는 반입 미리보기 결과에서 확인하세요.'
 if($ws -and (Test-Path (Join-Path $ws 'pending-sync.json'))){$dataState='저장 중단 표시가 있습니다. 일부 파일만 변경되었을 수 있습니다.';$action='추가 저장을 중단하세요. 복구 화면에서 표시된 체크포인트를 새 폴더로 복원하고 원본과 대조한 뒤 사용하세요.'}
 $d=New-Object Windows.Forms.Form;$d.Text='자료 확인이 필요합니다';$d.Size=New-Object Drawing.Size(760,420);$d.StartPosition='CenterParent'
 $box=TextBox $d 15 15 710 255;$box.ReadOnly=$true;$box.Text="문제: $problem`r`n`r`n위치: $(if($nErrorContext){$nErrorContext}elseif($dashLayer){WPath $dashLayer}else{'작업공간 선택 전'})`r`n`r`n자료 상태: $dataState`r`n`r`n지금 할 일: $action"
 DButton $d '기술 상세 보기' 15 285 170 {$box.Text=$technical}|Out-Null;$close=DButton $d '닫고 입력 확인' 525 285 200 {$d.Close()};$d.ActiveControl=$close;[void]$(Ui12Dialog $d;$d.ShowDialog($f))
}
function DSafe($op){try{& $op}catch{$dashMessage.Text='자료 확인이 필요합니다. 안내한 위치와 조치를 확인하세요.';Log $_.Exception.Message;WFriendly ($_.Exception.GetBaseException().Message)}}
function Safe($op){try{& $op}catch{Log $_.Exception.Message;WFriendly ($_.Exception.GetBaseException().Message)}}
function WParentDialog{
 DClean;if(!$dashSnapshot){throw '먼저 작업표를 여세요.'};$parent=$meta.layers[$dashLayer].parent
 if(!$parent -or $meta.workspaceScopes[$dashLayer].unlinked){throw '연결된 상위 작업표가 없습니다.'}
 $p=$meta.courses[$dashProfileId];$a=WDisplayArea $p $dashArea
 $targetFields=if($parent -eq $p.root){@($a.root_item)}elseif($parent -eq $a.work_layer){@($a.work_items)}else{@($meta.layers[$parent].items|Where-Object {$_ -notmatch '\.(status|note)$'})}
 $links=@($meta.links|Where-Object {$_.source -eq $dashLayer -and $_.target -eq $parent -and $_.targetItem -in $targetFields -and ($_.sourceItem -eq '__sum__' -or $_.sourceItem -in $dashSnapshot.fields)})
 if($links.Count -eq 1){$sourceItem=$links[0].sourceItem;$targetItem=$links[0].targetItem}else{
  $d=New-Object Windows.Forms.Form;$d.Text='상위 항목 연결';$d.ClientSize=New-Object Drawing.Size(700,210);$d.StartPosition='CenterParent';$d.Font=$f.Font
  [void](Label $d ((LName $dashLayer)+' → '+(LName $parent)) 20 15 660 35)
  $src=Combo $d 20 65 315;$dst=Combo $d 350 65 330;Choice $src '__sum__' '현재 작업표의 합계'
  foreach($item in $dashSnapshot.fields){Choice $src $item (FName $dashLayer $item)};foreach($item in $targetFields){Choice $dst $item (FName $parent $item)}
  $src.SelectedIndex=0;if($dst.Items.Count){$dst.SelectedIndex=0}
  DButton $d '연결하고 반영' 380 145 170 {if($src.SelectedItem -and $dst.SelectedItem){$d.Tag=@{source=$src.SelectedItem.Id;target=$dst.SelectedItem.Id};$d.Close()}}|Out-Null
  DButton $d '취소' 565 145 115 {$d.Close()}|Out-Null;[void]$(Ui12Dialog $d;$d.ShowDialog($f));if(!$d.Tag){return};$sourceItem=$d.Tag.source;$targetItem=$d.Tag.target
 }
 $args2=DArgs @{Item=$sourceItem;TargetItem=$targetItem;PayloadFile=(Write-Request @{direct=$true})}
 $result=(Run 'parentcommit' $args2)|ConvertFrom-Json;DLoad $dashMode $dashArea
 $dashMessage.Text='상위에 '+@($result.rows|Where-Object disposition -EQ '반영').Count+'명 반영했습니다.'
}
function WStudentDialog{
 DClean;$sid=if($wStudent){$wStudent}elseif($dashGrid.CurrentRow){[string]$dashGrid.CurrentRow.Tag}else{''};if(!$sid){throw '현재 표에서 학생을 선택하세요.'}
 $viewItem=if($dashGrid.CurrentCell){[string]$dashGrid.CurrentCell.OwningColumn.Tag.item}else{''};$d=New-Object Windows.Forms.Form;$d.Text='선택 학생 · 필요한 깊이만 조회';$d.Size=New-Object Drawing.Size(1050,650);$d.StartPosition='CenterParent';[void](Label $d ((WStudentLabel $sid)+' / '+(WPath $dashLayer)) 15 15 990 40)
 $depth=Combo $d 15 65 400;Choice $depth 0 '현재 작업표만';Choice $depth 1 '현재와 바로 아래 작업표';Choice $depth 2 '현재에서 두 단계 아래까지';Choice $depth 3 '현재에서 세 단계 아래까지';$depth.SelectedIndex=0;$g=Grid $d 15 120 990 420 @('작업공간','항목','값')
 DButton $d '선택 깊이 조회' 440 63 240 {DSafe {$r=(Run 'studentview' (DArgs @{Student=$sid;PayloadFile=(Write-Request @{depth=[int]$depth.SelectedItem.Id;item=$viewItem})}))|ConvertFrom-Json;$g.Rows.Clear();foreach($x in $r.records){[void]$g.Rows.Add([object[]]@($x.workspace,$x.label,(ValueText $x.value)))}}}|Out-Null;[void]$(Ui12Dialog $d;$d.ShowDialog($f))
}

function WImportDialog{
 DClean;if(!$dashSnapshot -or $dashMode -eq 'summary'){throw '자료를 넣을 평가영역 또는 하위 작업공간을 선택하세요.'};if($wIdentityMode -eq 2){throw '원문에 학생 정보가 있을 수 있습니다. 자료 대조를 위해 이름 모드로 전환한 뒤 반입하세요.'}
 $target=$dashLayer;$d=New-Object Windows.Forms.Form;$d.Text='선택 작업공간에 비정형 자료 반입';$d.Size=New-Object Drawing.Size(1100,710);$d.StartPosition='CenterParent';[void](Label $d ('반입 대상: '+(WPath $target)+' / 파일 전체 학생을 제출별 새 근거표로 보존 · 학급 조회 필터와 별개 / 상위 자동 반영 없음') 15 10 1050 45)
 $fileBox=TextBox $d 15 60 810;$fileBox.ReadOnly=$true;$g=Grid $d 15 160 1050 330 @('처리','학생 또는 원문','항목','값','위치','지금 할 일');$g.DefaultCellStyle.WrapMode='True';$g.AutoSizeRowsMode='AllCells';$summary=Label $d '최하위 작업표에서 가져오기를 권장합니다. 현재 위치에서는 새 근거표를 만듭니다.' 15 112 1050 45
 $state=@{plan=$null;file=''};$args2=DArgs @{Parent=$target;AsOf=$dashSnapshot.as_of}
 $save=DButton $d '정상 기록 반입' 770 565 290 {DSafe {if(!$state.plan){throw '자료를 선택해 미리보기를 확인하세요.'};$a=@{};foreach($k in $args2.Keys){$a[$k]=$args2[$k]};$a.File=$state.file;$a.PreviewFingerprint=$state.plan.fingerprint;$r=(Run 'importcommit' $a)|ConvertFrom-Json;$summary.Text='정상 '+$r.records.Count+'기록 저장 / 격리 '+$state.plan.quarantine.Count+'건 원문 보존. 상위 반영 0건.';$save.Enabled=$false;Read-Meta;WRefreshTree}};$save.Enabled=$false
 DButton $d 'Excel 파일 선택·미리보기' 840 57 220 {DSafe {$pick=New-Object Windows.Forms.OpenFileDialog;$pick.Filter='Excel 자료|*.xlsx';$pick.InitialDirectory=Join-Path $base 'outputs/natural-import-v1';if($pick.ShowDialog($d) -ne 'OK'){return};$state.file=$pick.FileName;$fileBox.Text=$state.file;$a=@{};foreach($k in $args2.Keys){$a[$k]=$args2[$k]};$a.File=$state.file;$state.plan=(Run 'importpreview' $a)|ConvertFrom-Json;$g.Rows.Clear();foreach($r in $state.plan.records){[void]$g.Rows.Add([object[]]@('정상',(WStudentLabel $r.student_id),$r.item_label,(ValueText $r.value),$r.source_cell,'확인 후 정상 기록 반입'))};foreach($r in $state.plan.quarantine){[void]$g.Rows.Add([object[]]@('격리',$r.subject,'','',($r.source_sheet+'!'+$r.source_row),($r.reason+' → 예외 행 해결에서 학생·항목 지정')))};$summary.Text='정상 '+$state.plan.records.Count+'기록 / 격리 '+$state.plan.quarantine.Count+'건. 원본 보존 · 미리보기는 점수를 변경하지 않습니다..';$save.Enabled=(!$state.plan.duplicate_batch -and $state.plan.records.Count -gt 0)}}|Out-Null
 DButton $d '예외 행 해결' 15 565 260 {DSafe {if(!$state.file){throw '자료를 먼저 선택하세요.'};$stagingFile.Text=$state.file;Refresh-Navigation $target;$script:current=$target;$ddate.Value=[datetime]$dashSnapshot.as_of;$script:stagingPlan=$null;$d.Tag='exceptions';$d.Close();DAdvanced '자료';$tabs.SelectedTab=$stagingTab;$stagingSummary.Text='현재 반입 대상: '+(WPath $target)+' / 읽기 전용 검토 후 문제 행을 선택하고 학생·항목을 지정하세요.'}}|Out-Null
 [void]$(Ui12Dialog $d;$d.ShowDialog($f));if($d.Tag -ne 'exceptions'){Read-Meta;DLoad $dashMode $dashArea}
}

$columnMenu=New-Object Windows.Forms.ContextMenuStrip
$renameColumn=$columnMenu.Items.Add('선택 열 이름 바꾸기')
$renameColumn.Add_Click({DSafe {DClean;$tag=$dashGrid.CurrentCell.OwningColumn.Tag;if(!$tag.item){throw '점수 또는 기록 열을 먼저 선택하세요.'};$label=Input-Name '열 이름 바꾸기' '표시명만 변경하며 연결과 점수는 유지합니다.' $(if($dashMode -eq 'summary'){@($meta.courses[$dashProfileId].areas|Where-Object root_item -EQ $tag.item)[0].label}else{FName $dashLayer $tag.item});if(!$label){return};Run 'spacerename' (DArgs @{PayloadFile=(Write-Request @{item=$tag.item;label=$label})})|Out-Null;Read-Meta;DLoad $dashMode $dashArea}})
$dashGrid.ContextMenuStrip=$columnMenu
function WColumnMode{
 $d=New-Object Windows.Forms.Form;$d.Text='새 열의 사용 방식';$d.ClientSize=New-Object Drawing.Size(660,235);$d.StartPosition='CenterParent';$d.Font=$f.Font
 [void](Label $d '새 열을 어떻게 사용할까요?' 20 15 620 30)
 [void](Label $d '현재 공간에 직접 입력: 하위 공간 없이 기록합니다.' 20 55 620 30)
 [void](Label $d '하위 공간 만들기: 세부 작업의 합산점수를 이 항목으로 반영합니다.' 20 90 620 45)
 DButton $d '현재 공간에 직접 입력' 20 160 240 {$d.Tag='direct';$d.Close()}|Out-Null
 DButton $d '하위 공간 만들기' 275 160 240 {$d.Tag='child';$d.Close()}|Out-Null
 DButton $d '취소' 530 160 105 {$d.Close()}|Out-Null
 [void]$(Ui12Dialog $d;$d.ShowDialog($f));return [string]$d.Tag
}

$dashGrid.Add_ColumnHeaderMouseDoubleClick({param($sender,$e)
 if($dashLoading -or $busy -or $e.ColumnIndex -lt 0 -or $e.X -ge ($sender.Columns[$e.ColumnIndex].Width-24)){return}
 DSafe {DClean;$tag=$dashGrid.Columns[$e.ColumnIndex].Tag;if(!$tag.item){return};if((UIsFolded $tag.item)){UFold $tag.item $false;return};$children=@(WColumnChildren $tag.item)
  if($children.Count -eq 1){if($dashMode -eq 'summary'){DLoad 'detail' $tag.item}else{$script:wTarget=$children[0];DLoad 'workspace' $dashArea}}else{$dashMessage.Text='이 열에는 단일 하위 작업공간이 연결되어 있지 않습니다. 왼쪽 트리에서 작업공간을 선택하세요.'}
 }
})
function WAttachExistingColumn{
 DClean;if(!$dashSnapshot -or !$dashGrid.CurrentCell -or !$dashGrid.CurrentCell.OwningColumn.Tag.item){throw '연결할 기존 점수 또는 기록 열을 먼저 선택하세요.'}
 $targetItem=[string]$dashGrid.CurrentCell.OwningColumn.Tag.item;$targetArea=if($dashMode -eq 'summary'){$targetItem}else{$dashArea};$columnLabel=if($dashMode -eq 'summary'){@($meta.courses[$dashProfileId].areas|Where-Object root_item -EQ $targetItem)[0].label}else{FName $dashLayer $targetItem};$label=WAttachName $columnLabel;if(!$label){return}
 $r=(Run 'spacecreate' (DArgs @{Area=$targetArea;PayloadFile=(Write-Request @{label=$label;target_item=$targetItem;numeric=$true})}))|ConvertFrom-Json;Read-Meta;$script:wTarget=$r.layer;if($meta.layers[$r.layer].parent -eq $meta.courses[$dashProfileId].root){DLoad 'detail' $targetArea}else{DLoad 'workspace' $targetArea}
 $dashMessage.Text='하위 공간 생성 완료 · 학생 명렬표 '+$r.students+'명 전수 일치 · 기존 기록·소속 유지 · 상위 점수 변경 없음.'
}
$attachExisting=DButton $dashboard '선택 항목에 하위 공간 만들기' 210 178 345 {DSafe {WAttachExistingColumn}}
$dashHeading.Width=550;$attachExisting.SetBounds(210,218,280,36)
$attachColumnMenu=$columnMenu.Items.Add('선택 항목에 하위 공간 만들기')
$attachColumnMenu.Add_Click({DSafe {WAttachExistingColumn}})

function WAttachName($columnLabel){
 $d=New-Object Windows.Forms.Form;$d.Text='기존 열에 하위 공간 만들기';$d.ClientSize=New-Object Drawing.Size(740,260);$d.StartPosition='CenterParent';$d.Font=$f.Font
 [void](Label $d ('선택 항목: '+$columnLabel) 20 15 690 35)
 $notice=Label $d '전체 학생에게 하위 공간을 만듭니다. 기존 기록은 유지됩니다.' 20 55 690 35;$notice.Font=New-Object Drawing.Font($f.Font,[Drawing.FontStyle]::Bold)
 [void](Label $d '상위 점수는 자동 변경되지 않습니다. [상위에 반영]을 눌러 값을 보냅니다.' 20 95 690 35)
 $nameBox=TextBox $d 20 140 690 35;$nameBox.Text=$columnLabel+' 세부 작업'
 $ok=DButton $d '기존 기록 유지하고 만들기' 305 205 260 {$d.Tag=$nameBox.Text;$d.Close()};$cancel=DButton $d '취소' 580 205 130 {$d.Close()};$d.AcceptButton=$ok;$d.CancelButton=$cancel
 [void]$(Ui12Dialog $d;$d.ShowDialog($f));return [string]$d.Tag
}
function WColumnChildren($item){
 $ids=@($meta.workspaceScopes.Keys|Where-Object {$meta.layers[$_].parent -eq $dashLayer -and $meta.workspaceScopes[$_].parent_item -eq $item})
 $ids+=@($meta.links|Where-Object {$_.target -eq $dashLayer -and $_.targetItem -eq $item}|ForEach-Object source)
 foreach($p in $meta.courses.Values){if($p.root -eq $dashLayer){$ids+=@($p.areas|Where-Object root_item -EQ $item|ForEach-Object work_layer)}}
 return @($ids|Where-Object {$_ -and $meta.layers.ContainsKey($_) -and !$meta.workspaceScopes[$_].unlinked}|Sort-Object -Unique)
}
function WSelectedHeader{
 if($dashLoading){return};$selected=if($dashGrid.CurrentCell){$dashGrid.CurrentCell.OwningColumn.Tag.item}else{''};$attachExisting.Enabled=[bool]$selected
 foreach($col in $dashGrid.Columns){if($col.Tag.item){
  $active=$selected -and $selected -eq $col.Tag.item;$linked=@(WColumnChildren $col.Tag.item).Count -gt 0
  $col.HeaderCell.Style.BackColor=if($linked){if($active){[Drawing.Color]::FromArgb(184,224,214)}else{[Drawing.Color]::FromArgb(220,242,238)}}elseif($active){[Drawing.Color]::FromArgb(218,228,239)}else{[Drawing.Color]::FromArgb(235,239,246)}
  $col.HeaderCell.Style.ForeColor=if($linked){[Drawing.Color]::FromArgb(15,85,75)}else{$navy};$col.HeaderCell.Style.SelectionBackColor=$col.HeaderCell.Style.BackColor;$col.HeaderCell.Style.SelectionForeColor=$col.HeaderCell.Style.ForeColor
  $col.HeaderCell.ToolTipText=if($linked){'하위 작업공간 연결 · 머리글 더블클릭으로 이동 · ▼는 정렬·검색'}else{'현재 표에서 직접 입력 · ▼는 정렬·검색'}
 }};$dashGrid.Invalidate()
}
$dashGrid.Add_SelectionChanged({WSelectedHeader})

$columnTypeMenu=$columnMenu.Items.Add('열 종류 변경 · 숫자점수 / 문자기록')
$columnTypeMenu.Add_Click({DSafe {
 DClean;$cell=$dashGrid.CurrentCell;if(!$cell -or !$cell.OwningColumn.Tag.item){throw '종류를 바꿀 점수·기록 셀을 선택하세요.'}
 $item=[string]$cell.OwningColumn.Tag.item;$kind=WColumnKind;if(!$kind){return}
 $result=(Run 'spacecolumntype' (DArgs @{PayloadFile=(Write-Request @{item=$item;kind=$kind})}))|ConvertFrom-Json
 MemoryClear;Read-Meta;DLoad $dashMode $dashArea
 $dashMessage.Text='열 종류: '+$(if($kind -eq 'numeric'){'숫자점수 · 합계 포함'}else{'문자기록 · 합계 제외'})+' / 기존 값·연결 유지 / 상위 값은 변경하지 않았습니다.'
}})
