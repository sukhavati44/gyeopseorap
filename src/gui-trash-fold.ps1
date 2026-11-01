function WDisplayArea($p,$area){$a=@($p.areas|Where-Object root_item -EQ $area);if($a.Count){return $a[0]};return $meta.detachedAreas[$dashProfileId+'|'+$area]}
function UColumnRisk($plan){
 if(@($plan.children|Where-Object {$_}).Count -or @($plan.removed_links|Where-Object {$_}).Count){return 'linked'}
 if(@($plan.records|Where-Object {($null -ne $_.value -and [string]$_.value -ne '') -or $_.formula}).Count){return 'data'}
 if($script:memoRoot -eq $ws -and @($script:memoData.Values|Where-Object {$_.layer -eq $plan.layer -and $_.item -eq $plan.item}).Count){return 'data'}
 return 'empty'
}
function ULinkedColumnChoice($plan){
 $d=New-Object Windows.Forms.Form;$d.Text='연결된 열 삭제';$d.ClientSize=New-Object Drawing.Size(600,225);$d.StartPosition='CenterParent';$d.Font=$f.Font
 $heading=Label $d ('주의: '+$plan.label+' 열에 연결이 있습니다.') 18 16 564 32;$heading.ForeColor=[Drawing.Color]::FromArgb(155,44,44);$heading.Font=New-Object Drawing.Font($f.Font,[Drawing.FontStyle]::Bold)
 [void](Label $d ("삭제하면 이 열의 데이터가 작업표에서 제외되고 연결이 끊어집니다.`r`n하위 작업공간은 미연결 작업공간으로 남습니다.`r`n자료와 연결을 유지하려면 [숨김(접기)]을 권장합니다.") 18 58 564 95)
 $hide=Button $d '숨김(접기) · 권장' 18 171 190 {$d.Tag='hide';$d.Close()}
 $delete=Button $d '연결 끊고 열 삭제' 222 171 190 {$d.Tag='delete';$d.Close()}
 $cancel=Button $d '취소' 426 171 156 {$d.Tag='cancel';$d.Close()};$d.CancelButton=$cancel
 try{Ui12Dialog $d;[void]$d.ShowDialog($f);return [string]$d.Tag}finally{$d.Dispose()}
}
function UTrash($kind,$deleteColumn=$false){
 if($kind -eq 'space' -and (!$wTree.SelectedNode.Tag -or $wTree.SelectedNode.Tag.layer -ne $dashLayer)){throw '이동할 작업공간을 먼저 열고 조회가 끝난 뒤 다시 선택하세요.'}
 DClean;if(!$dashSnapshot){throw '작업표를 먼저 조회하세요.'};$item='';if($kind -eq 'column'){if(!$dashGrid.CurrentCell.OwningColumn.Tag.item){throw '이동할 점수·상태 항목을 선택하세요.'};$item=$dashGrid.CurrentCell.OwningColumn.Tag.item}
 if($kind -eq 'space'){$reason=UTrashSpaceBlock $meta $dashLayer;if($reason){[Windows.Forms.MessageBox]::Show($f,$reason,'휴지통 이동 제한','OK','Information')|Out-Null;$dashMessage.Text='휴지통 이동하지 않음 · 연결 보호 · 기존 자료 유지';return}}
 $plan=(Run 'trashpreview' (DArgs @{PayloadFile=(Write-Request @{kind=$kind;item=$item})}))|ConvertFrom-Json
 $text='대상: '+$plan.label+"`r`n원래 위치: "+$plan.location+"`r`n학생기록: "+$plan.student_records+"명`r`n하위 작업공간: "+$plan.children.Count+"개 / 해제할 연결: "+$plan.removed_links.Count+"개`r`n`r`n"+$plan.warning+"`r`n`r`n별도 보관·체크포인트 검증 후 활성 항목에서 제외합니다. 휴지통에서 열람하거나 경고 확인 후 보관 사본을 영구 삭제할 수 있습니다."
 if($deleteColumn){
  switch(UColumnRisk $plan){
   'linked'{$choice=ULinkedColumnChoice $plan;if($choice -eq 'hide'){UFold $item $true;$dashMessage.Text='열을 숨김(접기) 처리했습니다. 데이터와 하위 연결은 유지됩니다.';return};if($choice -ne 'delete'){return}}
   'data'{if([Windows.Forms.MessageBox]::Show($f,('['+$plan.label+"] 열의 데이터·메모가 작업표에서 삭제됩니다.`r`n삭제한 열은 휴지통에 보관합니다."),'데이터가 있는 열 삭제','OKCancel','Warning') -ne 'OK'){return}}
  }
 }elseif([Windows.Forms.MessageBox]::Show($f,$text,'휴지통으로 이동 · 경고','OKCancel','Warning') -ne 'OK'){return}
 $req=@{id=$plan.id;kind=$kind;item=$item;fingerprint=$plan.fingerprint;confirmed=$true};$r=(Run 'trashcommit' (DArgs @{PayloadFile=(Write-Request $req)}))|ConvertFrom-Json
 if(!$r.completed){throw '휴지통 이동 검증이 완료되지 않았습니다. 추가 저장을 멈추세요.'};$previousMode=$dashMode;$previousArea=$dashArea;Read-Meta;if($deleteColumn -and $previousMode -ne 'summary' -and $meta.layers.ContainsKey($dashLayer)){DLoad $previousMode $previousArea}else{DInitialize $dashProfileId};$dashMessage.Text=if($deleteColumn){'열 삭제 완료 · 휴지통 보관 · 연결된 하위 공간은 미연결 작업공간으로 보존'}else{'휴지통 이동 완료 · 기록·하위 공간 보존 · 다른 점수 변경 없음'}
}
function UTrashSpaceBlock($data,$id){
 $li=$data.layers[$id]
 if(!$li){return '선택한 작업공간을 찾을 수 없습니다. 다시 조회하세요.'}
 if(!$li.parent -and !$data.workspaceScopes[$id].unlinked){return '수행평가 종합표 전체는 휴지통으로 이동할 수 없습니다. 이동할 평가항목의 열을 선택하세요.'}
 return ''
}
$trashColumn=$columnMenu.Items.Add('열 삭제');$trashColumn.Add_Click({DSafe {UTrash 'column' $true}})
$columnMenu.Add_Opening({$trashColumn.Enabled=(!$busy -and !$dashLoading -and $null -ne $dashSnapshot -and $null -ne $dashGrid.CurrentCell -and [bool]$dashGrid.CurrentCell.OwningColumn.Tag.item)})
$treeMenu=New-Object Windows.Forms.ContextMenuStrip;$trashSpace=$treeMenu.Items.Add('휴지통으로 이동');$trashSpace.Add_Click({DSafe {UTrash 'space'}});$wTree.ContextMenuStrip=$treeMenu
$trashButton=DButton $side '휴지통' 8 705 174 {DSafe {UTrashList}}
function UTrashList{
 DClean;$q=(Run 'trashlist')|ConvertFrom-Json;$d=New-Object Windows.Forms.Form;$d.Text='휴지통 · 열람 및 비우기';$d.Size=New-Object Drawing.Size(1120,720);$d.StartPosition='CenterParent';$d.Font=$f.Font
 [void](Label $d '열람하거나 휴지통 비우기로 보관 사본을 영구 삭제할 수 있습니다. 미연결 작업공간은 보존합니다.' 15 10 1060 40)
 $g=Grid $d 15 55 1070 210 @('종류','이름','원래 위치','처리 시각','학생기록','검증','이전 연결');foreach($e in $q.entries){$row=$g.Rows[$g.Rows.Add([object[]]@($(if($e.kind -eq 'column'){'항목'}else{'작업공간'}),$e.label,$e.location,$e.at,$e.student_records,'보관 검증 완료',('연결 '+$e.previous_links.Count+'개 해제 · 하위 '+$e.children.Count+'개 보존')))];$row.Tag=$e.id}
 $body=TextBox $d 15 325 1070 290;$body.ReadOnly=$true
 DButton $d '선택 기록 열람' 15 275 230 {DSafe {if(!$g.CurrentRow){return};$r=(Run 'trashread' @{PayloadFile=(Write-Request @{id=$g.CurrentRow.Tag})})|ConvertFrom-Json;$lines=@('읽기 전용 · '+$r.entry.label+' / '+$r.entry.location);foreach($record in $r.records){$lines+=((WStudentLabel $record.student)+' / 값: '+(ValueText $record.value)+' / 수정 이력: '+@($record.manual.history).Count+'건')};$body.Text=$lines -join "`r`n`r`n"}}|Out-Null
 if($q.pending -and $q.pending.stage -ne 'completed'){$body.Text='미완료 요청: '+$q.pending.id+' / '+$q.pending.stage+"`r`n활성 자료는 변경 전 상태입니다. 동일 대상·요청을 확인한 뒤에만 재시도하세요.";$retry=DButton $d '미완료 요청 확인·재시도' 260 275 330 {DSafe {$p=$q.pending.plan;if([Windows.Forms.MessageBox]::Show($d,('대상: '+$p.label+"`r`n"+$p.warning),'미완료 이동 재확인','OKCancel','Warning') -ne 'OK'){return};Run 'trashcommit' @{Layer=$p.layer;PayloadFile=(Write-Request @{id=$q.pending.id;kind=$p.kind;item=$p.item;fingerprint=$p.fingerprint;confirmed=$true})}|Out-Null;$d.Close();Read-Meta;DInitialize $dashProfileId}}}
 DButton $d '휴지통 비우기' 15 625 230 {DSafe {$d.Close();UCleanupDialog 'workspace'}}|Out-Null
 DButton $d '닫기' 920 625 160 {$d.Close()}|Out-Null;try{[void]$(Ui12Dialog $d;$d.ShowDialog($f))}finally{$d.Dispose()}
}

# Presentation preferences are independent of records and relations.json.
$script:uLoading=$false;$script:uProject='';$script:uLayout=@{tables=@{};trees=@{}}
$dashGrid.ScrollBars='Both';$dashGrid.AllowUserToResizeColumns=$true
function ULoadPrefs{
 if($uProject -eq $ws){return};$script:uProject=$ws;$script:uLayout=@{tables=@{};trees=@{}};$path=Join-Path $ws 'ui-layout.json'
 if(Test-Path $path){try{$script:uLayout=Map (Get-Content $path -Raw -Encoding UTF8|ConvertFrom-Json)}catch{$dashMessage.Text='화면 설정을 읽지 못했습니다. 기록은 유지하고 기본 너비로 표시합니다.'}}
}
function USavePrefs{
 if(!$ws -or $uLoading -or $dashLoading){return};$dir=Join-Path $ws '.ui-history';New-Item -ItemType Directory $dir -Force|Out-Null;$tmp=Join-Path $ws ('.ui-'+[guid]::NewGuid().ToString('N')+'.tmp');$dest=Join-Path $ws 'ui-layout.json'
 [IO.File]::WriteAllText($tmp,($uLayout|ConvertTo-Json -Depth 15),[Text.UTF8Encoding]::new($true));if(Test-Path $dest){[IO.File]::Replace($tmp,$dest,(Join-Path $dir ([guid]::NewGuid().ToString('N')+'.json')))}else{[IO.File]::Move($tmp,$dest)}
}
function UTablePrefs{
 ULoadPrefs;if(!$uLayout.tables[$dashLayer]){$uLayout.tables[$dashLayer]=@{widths=@{};folded=@{};row=0;scroll=0}};return $uLayout.tables[$dashLayer]
}
function UIsFolded($item){if(!$item -or !$ws -or !$dashLayer){return $false};$p=UTablePrefs;return [bool]$p.folded[$item]}
function UColumnLook($col,$fold){
 if(!$col.Tag.item){return};if(!$col.Tag.ContainsKey('normal_header')){$col.Tag.normal_header=$col.HeaderText}
 if($fold){$col.MinimumWidth=8;$col.Width=12;$col.HeaderText='▸';$col.DefaultCellStyle.BackColor=$navy;$col.DefaultCellStyle.ForeColor=$navy;$col.HeaderCell.Style.BackColor=$navy;$col.HeaderCell.Style.ForeColor=[Drawing.Color]::White;foreach($cell in $dashGrid.Rows|ForEach-Object {$_.Cells[$col.Index]}){if(!$cell.Tag){$cell.Tag=@{back=$cell.Style.BackColor;fore=$cell.Style.ForeColor}};$cell.Style.BackColor=$navy;$cell.Style.ForeColor=$navy};$col.HeaderCell.ToolTipText='접힌 항목 · 머리글 두 번 클릭 또는 메뉴에서 펼치기'}else{$col.MinimumWidth=35;$col.HeaderText=$col.Tag.normal_header;$col.DefaultCellStyle.BackColor=[Drawing.Color]::White;$col.DefaultCellStyle.ForeColor=$navy;foreach($cell in $dashGrid.Rows|ForEach-Object {$_.Cells[$col.Index]}){if($cell.Tag){$cell.Style.BackColor=$cell.Tag.back;$cell.Style.ForeColor=$cell.Tag.fore;$cell.Tag=$null}}}
}
function UFold($item,$fold){
 if(!$item){throw '접을 점수·기록 항목을 선택하세요.'};$p=UTablePrefs;$script:uLoading=$true
 try{foreach($col in $dashGrid.Columns){if($col.Tag.item -eq $item){if($fold -and !$p.folded[$item]){$p.widths[$col.Name]=$col.Width};if($col.Tag.kind -eq 'value'){UColumnLook $col $fold;if(!$fold){$col.Width=if($p.widths[$col.Name]){[int]$p.widths[$col.Name]}else{170}}}else{$col.Visible=!$fold;if(!$fold -and $p.widths[$col.Name]){$col.Width=[int]$p.widths[$col.Name]}}}};$p.folded[$item]=$fold}finally{$script:uLoading=$false};USavePrefs
}
function UApplyPrefs{
 if(!$ws -or !$dashSnapshot){return};$p=UTablePrefs;$script:uLoading=$true;$oldTree=$wTreeLoading;$script:wTreeLoading=$true
 try{
  foreach($col in $dashGrid.Columns){$col.AutoSizeMode='None';if($col.Tag.item){$col.MinimumWidth=35;$col.Width=if($p.widths[$col.Name]){[int]$p.widths[$col.Name]}elseif($col.Tag.kind -eq 'value'){170}else{140};if($p.folded[$col.Tag.item]){if($col.Tag.kind -eq 'value'){UColumnLook $col $true}else{$col.Visible=$false}}}elseif($col.Name -eq 'total'){$col.Width=if($dashMode -eq 'summary'){90}else{110}}elseif($p.widths[$col.Name]){$col.Width=[int]$p.widths[$col.Name]}}
  function Apply-Utree($nodes){foreach($n in $nodes){if($uLayout.trees.ContainsKey($n.Name)){if($uLayout.trees[$n.Name]){$n.Expand()}else{$n.Collapse()}};Apply-Utree $n.Nodes}};Apply-Utree $wTree.Nodes
  if($dashGrid.Rows.Count -gt 0){$dashGrid.FirstDisplayedScrollingRowIndex=[math]::Min([int]$p.row,$dashGrid.Rows.Count-1)};$dashGrid.HorizontalScrollingOffset=[math]::Max(0,[int]$p.scroll)
 }finally{$script:uLoading=$false;$script:wTreeLoading=$oldTree;$script:uReadyLayer=$dashLayer}
}
function UCapturePrefs{
 if($uLoading -or $dashLoading -or !$dashSnapshot -or $uReadyLayer -ne $dashLayer){return};$p=UTablePrefs
 foreach($col in $dashGrid.Columns){if(!$col.Tag.item -or !$p.folded[$col.Tag.item]){$p.widths[$col.Name]=$col.Width}}
 $p.row=[math]::Max(0,$dashGrid.FirstDisplayedScrollingRowIndex);$p.scroll=$dashGrid.HorizontalScrollingOffset;USavePrefs
}
$script:uOriginalOnLoad=${function:NOnLoad};function NOnLoad{& $uOriginalOnLoad;UApplyPrefs}
$dashGrid.Add_ColumnWidthChanged({UCapturePrefs});$dashGrid.Add_Scroll({UCapturePrefs})
$dashGrid.Add_ColumnDividerDoubleClick({param($sender,$e)if($e.ColumnIndex -ge 0 -and !($dashGrid.Columns[$e.ColumnIndex].Tag.item -and (UIsFolded $dashGrid.Columns[$e.ColumnIndex].Tag.item))){$dashGrid.AutoResizeColumn($e.ColumnIndex,[Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells);$e.Handled=$true;UCapturePrefs}})
$foldMenu=$columnMenu.Items.Add('접기 / 펼치기');$foldMenu.Add_Click({DSafe {$item=$dashGrid.CurrentCell.OwningColumn.Tag.item;UFold $item (!(UIsFolded $item))}})
$allColumns=$columnMenu.Items.Add('열 모두 펼치기');$allColumns.Add_Click({DSafe {$p=UTablePrefs;foreach($item in @($p.folded.Keys)){if($p.folded[$item]){UFold $item $false}}}})
$treeExpand=$treeMenu.Items.Add('작업공간 펼치기');$treeExpand.Add_Click({if($wTree.SelectedNode){$wTree.SelectedNode.Expand()}})
$treeCollapse=$treeMenu.Items.Add('작업공간 접기');$treeCollapse.Add_Click({if($wTree.SelectedNode){$wTree.SelectedNode.Collapse()}})
$treeAll=$treeMenu.Items.Add('작업공간 모두 펼치기');$treeAll.Add_Click({$wTree.ExpandAll()})
function UTreeChanged($node,$expanded){if($uLoading -or $wTreeLoading -or !$ws){return};ULoadPrefs;$uLayout.trees[$node.Name]=$expanded;USavePrefs}
$wTree.Add_AfterCollapse({param($sender,$e)UTreeChanged $e.Node $false});$wTree.Add_AfterExpand({param($sender,$e)UTreeChanged $e.Node $true})
$dashGrid.Add_CellMouseDown({param($sender,$e)if($e.Button -eq [Windows.Forms.MouseButtons]::Right -and $e.ColumnIndex -ge 0 -and $dashGrid.Rows.Count){$row=if($e.RowIndex -ge 0){$e.RowIndex}elseif($dashGrid.CurrentCell){$dashGrid.CurrentCell.RowIndex}else{0};if(!$dashGrid.Rows[$row].Cells[$e.ColumnIndex].Selected){$dashGrid.ClearSelection();$dashGrid.CurrentCell=$dashGrid.Rows[$row].Cells[$e.ColumnIndex];$dashGrid.CurrentCell.Selected=$true}}})
$dashGrid.Add_EditingControlShowing({param($sender,$e)$e.Control.ContextMenuStrip=$columnMenu})

$dashGrid.Add_CurrentCellChanged({if(!$dashLoading -and !$uLoading){WSelectedHeader;DSelection}})
