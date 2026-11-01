. (Join-Path $base 'import-layout.ps1')
$script:legacyImportDialog=${function:WImportDialog}
function ImportAtLeaf {return $dashSnapshot -and $dashMode -ne 'summary' -and !@($meta.layers.Values|Where-Object {$_.parent -eq $dashLayer}).Count -and !@($meta.links|Where-Object {$_.target -eq $dashLayer}).Count}
function WImportDialog {
 DClean;if(!$dashSnapshot -or $dashMode -eq 'summary'){$dashMessage.Text='자료 가져오기는 하위 작업공간이 없는 최하위 작업표에서 권장합니다. 왼쪽 트리에서 자료를 넣을 표를 선택하세요.';return}
 if(!(ImportAtLeaf)){
  $notice=New-Object Windows.Forms.Form;$notice.Text='자료 가져오기 · 위치 선택';$notice.ClientSize=New-Object Drawing.Size(620,220);$notice.StartPosition='CenterParent';$notice.Font=$f.Font
  [void](Label $notice '최하위 작업표에서 자료 가져오기를 권장합니다.' 18 16 580 30)
  [void](Label $notice "현재 표에는 연결된 하위 작업공간이 있습니다.`r`n왼쪽 트리에서 더 아래의 작업표를 선택하면, 가져온 열을 그 표에 직접 넣을 수 있습니다.`r`n현재 표에서 계속하면 기존 방식으로 새 근거표를 만듭니다." 18 52 580 96)
  DButton $notice '최하위 표 선택으로 돌아가기' 18 166 280 {$notice.Close()}|Out-Null
  DButton $notice '현재 위치에서 새 근거표 만들기' 315 166 285 {$notice.Tag='legacy';$notice.Close()}|Out-Null
  Ui12Style $notice;[void]$notice.ShowDialog($f);if($notice.Tag -eq 'legacy'){& $legacyImportDialog};return
 }
 MemoryReader;$d=New-Object Windows.Forms.Form;$d.Text='자료 가져오기 · 학생과 열 연결';$d.ClientSize=New-Object Drawing.Size(1120,760);$d.MinimumSize=New-Object Drawing.Size(1050,740);$d.StartPosition='CenterParent';$d.Font=$f.Font;$d.BackColor=[Drawing.Color]::White
 $state=@{file='';book=$null;plan=$null;request='';loading=$false;done=$false}
 $targetLabel=Label $d ('넣을 위치: '+(WPath $dashLayer)) 16 10 1088 28;$targetLabel.Font=New-Object Drawing.Font($f.Font,[Drawing.FontStyle]::Bold)
 [void](Label $d '최하위 작업표에서 가져오기를 권장합니다. 지금은 최하위 표입니다.  ① 파일 선택 → ② 학생·열 연결 확인 → ③ 현재 표에 저장' 16 38 1088 24)
 $pathBox=TextBox $d 16 70 850;$pathBox.ReadOnly=$true
 $pickButton=DButton $d '① 파일 선택' 880 68 224 {ImportTry {ImportPick}}
 [void](Label $d '시트' 16 113 40);$sheetBox=Combo $d 58 109 180
 [void](Label $d '제목 행' 252 113 64);$headBox=New-Object Windows.Forms.NumericUpDown;$headBox.Minimum=1;$headBox.Maximum=20000;$headBox.SetBounds(320,109,64,28);$d.Controls.Add($headBox)
 [void](Label $d '학생코드 열' 405 113 90);$idBox=Combo $d 498 109 250
 $applyLayout=DButton $d '이 연결로 열 다시 읽기' 850 107 254 {ImportTry {ImportBuildLayout}}
 [void](Label $d '코드가 없으면' 16 152 115)
 [void](Label $d '반' 135 152 24);$classBox=Combo $d 160 148 165
 [void](Label $d '번호' 340 152 40);$numberBox=Combo $d 383 148 165
 [void](Label $d '이름' 564 152 40);$nameBox=Combo $d 607 148 165
 $blankBox=New-Object Windows.Forms.CheckBox;$blankBox.Text='빈칸으로 기존 값도 지우기';$blankBox.SetBounds(805,148,299,28);$d.Controls.Add($blankBox)
 $info=Label $d '파일을 선택하면 학생코드를 자동으로 찾아 명부에 연결합니다. 행 순서와 이름만으로 맞추지 않습니다.' 16 186 1088 38;$info.Font=New-Object Drawing.Font('맑은 고딕',10,[Drawing.FontStyle]::Bold)
 $tab=New-Object Windows.Forms.TabControl;$tab.SetBounds(16,232,1088,382);$tab.Anchor='Top,Bottom,Left,Right';$d.Controls.Add($tab)
 $mapPage=New-Object Windows.Forms.TabPage;$mapPage.Text='② 열 연결 · 어디에 넣나요?';[void]$tab.TabPages.Add($mapPage)
 $studentPage=New-Object Windows.Forms.TabPage;$studentPage.Text='학생 연결 확인';[void]$tab.TabPages.Add($studentPage)
 $changePage=New-Object Windows.Forms.TabPage;$changePage.Text='변경할 값';[void]$tab.TabPages.Add($changePage)
 $issuePage=New-Object Windows.Forms.TabPage;$issuePage.Text='제외 항목 · 이유';[void]$tab.TabPages.Add($issuePage)
 $mapGrid=Grid $mapPage 0 0 1075 345 @('엑셀 열','값 예시');$mapGrid.Dock='Fill';$mapGrid.ReadOnly=$false;$mapGrid.EditMode='EditOnEnter';$mapGrid.RowTemplate.Height=32;$mapGrid.AutoSizeColumnsMode='Fill';$mapGrid.Columns[0].ReadOnly=$true;$mapGrid.Columns[1].ReadOnly=$true;$mapGrid.Columns[0].FillWeight=90;$mapGrid.Columns[1].FillWeight=95
 $destLookup=@{};$destTitles=@{}
 $destCol=New-Object Windows.Forms.DataGridViewComboBoxColumn;$destCol.Name='destination';$destCol.HeaderText='현재 작업표의 어느 열에 넣나요?';[void]$destCol.Items.Add('자동 연결 · 없으면 새 열');foreach($item in $dashSnapshot.fields){$title='기존 열: '+(FName $dashLayer $item);if($destLookup.ContainsKey($title)){$title+=' ('+($destLookup.Count+1)+')'};$destLookup[$title]=$item;$destTitles[$item]=$title;[void]$destCol.Items.Add($title)};$destCol.FillWeight=170;[void]$mapGrid.Columns.Add($destCol)
 $kindCol=New-Object Windows.Forms.DataGridViewComboBoxColumn;$kindCol.Name='kind';$kindCol.HeaderText='가져올 내용 / 합계 포함 여부';[void]$kindCol.Items.AddRange([object[]]@('숫자 · 합계에 포함','문자 · 합계에서 제외','가져오지 않음'));$kindCol.FillWeight=120;[void]$mapGrid.Columns.Add($kindCol)
 $studentsGrid=Grid $studentPage 0 0 1075 345 @('원본 행','파일의 학생코드','파일 이름','연결된 명부 학생','연결 결과');$studentsGrid.Dock='Fill'
 $changesGrid=Grid $changePage 0 0 1075 345 @('학생','작업표 열','현재 값','가져올 값','원본 위치');$changesGrid.Dock='Fill'
 $issuesGrid=Grid $issuePage 0 0 1075 345 @('원본 행','열','제외 이유');$issuesGrid.Dock='Fill'
 $help=Label $d "학생코드가 같으면 같은 학생입니다. 이름은 로컬 명부에서 표시합니다. 새 제목은 새 열, 같은 제목은 기존 열에 연결합니다.`r`n합계·최종 결과·판정 등은 중복 합산을 피하도록 문자로 제안합니다. 숫자 열만 합계에 들어갑니다. 빈칸은 기본적으로 기존 값을 유지합니다.`r`n[현재 표에 저장]은 기존 값을 바꾸고 즉시 저장합니다. 원본·변경 전 사본은 보존하며, 상위 표는 자동 반영하지 않습니다." 16 622 1088 68;$help.Anchor='Bottom,Left,Right'
 $refresh=DButton $d '② 미리보기 갱신' 16 698 208 {ImportTry {ImportPreview}};$refresh.Anchor='Bottom,Left';$refresh.Enabled=$false
 $save=DButton $d '③ 현재 표에 저장' 790 698 200 {ImportTry {ImportSave}};$save.Anchor='Bottom,Right';$save.Enabled=$false
 $close=DButton $d '닫기' 1004 698 100 {$d.Close()};$close.Anchor='Bottom,Right';$d.CancelButton=$close
 function ImportTry($op){if($state.loading){return};try{& $op}catch{$state.plan=$null;$save.Enabled=$false;$info.ForeColor=[Drawing.Color]::Firebrick;$info.Text=$_.Exception.GetBaseException().Message;Log ('자료 가져오기: '+$info.Text)}}
 function ImportDirty {if(!$state.loading){$state.plan=$null;$save.Enabled=$false;$info.ForeColor=$navy;$info.Text='연결 설정을 바꿨습니다. [② 미리보기 갱신]으로 반영할 내용을 확인하세요.'}}
 function ImportConfig {
  [void]$mapGrid.EndEdit();$cols=@();foreach($row in $mapGrid.Rows){$kind=switch([string]$row.Cells['kind'].Value){'숫자 · 합계에 포함'{'number'};'문자 · 합계에서 제외'{'text'};default{'skip'}};$target='';$destination=[string]$row.Cells['destination'].Value;if($destLookup.ContainsKey($destination)){$target=$destLookup[$destination]};$cols+=@{column=$row.Tag.column;label=$row.Tag.label;kind=$kind;target=$target}}
  return @{sheet=[string]$sheetBox.SelectedItem.Id;header=[int]$headBox.Value;roles=@{id=[int]$idBox.SelectedItem.Id;class=[int]$classBox.SelectedItem.Id;number=[int]$numberBox.SelectedItem.Id;name=[int]$nameBox.SelectedItem.Id};columns=$cols;replace_blanks=$blankBox.Checked}
 }
 function ImportBuildLayout {
  $settings=ImportConfig;$layout=ImportLayout $state.book $settings;$state.loading=$true
  try{
   $mapGrid.Rows.Clear()
   foreach($col in $layout.columns){if($col.column -in @($layout.roles.Values)){continue};$destination='자동 연결 · 없으면 새 열';$item='';if($col.key.StartsWith('value|') -and $col.key.Substring(6) -in $dashSnapshot.fields){$item=$col.key.Substring(6)}else{$same=@($dashSnapshot.fields|Where-Object {(FName $dashLayer $_) -ceq $col.label});if($same.Count -eq 1){$item=$same[0]}}
    $kind=$col.kind;if($item){$destination=$destTitles[$item];$kind=if($item -in $meta.courses[$dashProfileId].numeric_items){'number'}else{'text'}}
    $idx=$mapGrid.Rows.Add([object[]]@(($col.letter+' · '+$col.label),$col.sample,$destination,$(if($kind -eq 'number'){'숫자 · 합계에 포함'}else{'문자 · 합계에서 제외'})));$mapGrid.Rows[$idx].Tag=$col
   }
  }finally{$state.loading=$false};ImportPreview
 }
 function ImportPick {
  $pick=New-Object Windows.Forms.OpenFileDialog;$pick.Filter='Excel 자료 (*.xlsx)|*.xlsx';$pick.Title='현재 작업표에 가져올 자료 선택';if($state.file){$pick.InitialDirectory=Split-Path $state.file}
  if($pick.ShowDialog($d) -ne 'OK'){return};$state.plan=$null;$save.Enabled=$false;$state.file=$pick.FileName;$pathBox.Text=$state.file;$state.book=[NaturalXlsx]::Read($state.file);try{$layout=ImportLayout $state.book}catch{$candidate=@($state.book.Sheets|Where-Object {$_.Name -notlike '_*'})[0];if(!$candidate){throw '자료 시트가 없습니다.'};$rowNumber=@($candidate.Rows.Keys|Sort-Object|Select-Object -First 50|Sort-Object {$candidate.Rows[$_].Count} -Descending)[0];$layout=ImportLayout $state.book @{sheet=$candidate.Name;header=$rowNumber;allow_unmatched=$true}}
  $state.loading=$true;try{
   $sheetBox.Items.Clear();foreach($s in $layout.sheets){Choice $sheetBox $s $s};Choose $sheetBox $layout.sheet.Name;$headBox.Value=$layout.header
   foreach($combo in @($idBox,$classBox,$numberBox,$nameBox)){$combo.Items.Clear();Choice $combo 0 '선택 안 함';foreach($col in $layout.columns){Choice $combo $col.column ($col.letter+' · '+$col.label)}}
   Choose $idBox ([int]$layout.roles.id);Choose $classBox ([int]$layout.roles.class);Choose $numberBox ([int]$layout.roles.number);Choose $nameBox ([int]$layout.roles.name)
  }finally{$state.loading=$false};$refresh.Enabled=$true;ImportBuildLayout
 }
 function ImportStructureChanged {
  if($state.loading -or !$state.book -or !$sheetBox.SelectedItem){return};ImportDirty;$mapGrid.Rows.Clear();$studentsGrid.Rows.Clear();$changesGrid.Rows.Clear();$issuesGrid.Rows.Clear()
  try{$layout=ImportLayout $state.book @{sheet=[string]$sheetBox.SelectedItem.Id;header=[int]$headBox.Value;allow_unmatched=$true};$state.loading=$true
   foreach($pair in @(@($idBox,'id'),@($classBox,'class'),@($numberBox,'number'),@($nameBox,'name'))){$combo=$pair[0];$combo.Items.Clear();Choice $combo 0 '선택 안 함';foreach($col in $layout.columns){Choice $combo $col.column ($col.letter+' · '+$col.label)};Choose $combo ([int]$layout.roles[$pair[1]])}
   $info.Text='제목 행을 바꿨습니다. 학생코드 열을 확인하고 [이 연결로 열 다시 읽기]를 누르세요.'
  }catch{$info.Text='선택한 제목 행이 없습니다. 행 번호를 확인하세요.'}finally{$state.loading=$false}
 }
 function ImportPreview {
  if(!$state.file){return};$save.Enabled=$false;$info.Text='학생과 열을 연결하는 중입니다…';$info.Refresh();$state.request=Write-Request (ImportConfig)
  $args2=DArgs @{File=$state.file;PayloadFile=$state.request;Class='';Student=''};$state.loading=$true
  try{$state.plan=(Run 'tablepreview' $args2)|ConvertFrom-Json}finally{$state.loading=$false}
  $state.loading=$true;try{foreach($row in $mapGrid.Rows){$mapped=@($state.plan.column_mapping|Where-Object {$_.column -eq $row.Tag.column});if($mapped.Count -and [string]$row.Cells['kind'].Value -ne '가져오지 않음'){$row.Cells['kind'].Value=if($mapped[0].numeric){'숫자 · 합계에 포함'}else{'문자 · 합계에서 제외'}}}}finally{$state.loading=$false}
  $studentsGrid.Rows.Clear();foreach($r in $state.plan.student_rows){[void]$studentsGrid.Rows.Add([object[]]@($r.row,$r.code,$r.source_name,($r.class+' '+$r.number.ToString()+'번 '+$r.name),$r.result))}
  $changesGrid.Rows.Clear();foreach($r in @($state.plan.changes|Select-Object -First 1000)){[void]$changesGrid.Rows.Add([object[]]@((WStudentLabel $r.student),$r.label,(ValueText $r.before),(ValueText $r.value),($r.source_sheet+'!'+(ImportColumnName $r.source_column)+$r.source_row)))}
  $issuesGrid.Rows.Clear();foreach($r in $state.plan.skipped){[void]$issuesGrid.Rows.Add([object[]]@($r.row,$r.column,$r.reason))}
  $info.ForeColor=$navy;$changePage.Text='변경할 값 ('+$state.plan.changes.Count+'칸 · 최대 1000칸 표시)';$issuePage.Text='제외 항목 · 이유 ('+$state.plan.skipped.Count+'건)';$info.Text='학생 '+$state.plan.matched_students+'명 연결 · 새 열 '+$state.plan.new_fields.Count+'개 · 변경 '+$state.plan.changes.Count+'칸 · 제외 '+$state.plan.skipped.Count+'건 · 상위 표는 그대로 유지'
  if(!$state.plan.changes.Count){$info.Text+=' / 반영할 변경이 없습니다.'};$save.Enabled=$state.plan.changes.Count -gt 0
 }
 function ImportSave {
  if(!$state.plan){return};$save.Enabled=$false;$args2=DArgs @{File=$state.file;PayloadFile=$state.request;Class='';Student='';PreviewFingerprint=$state.plan.fingerprint};$state.loading=$true
  try{$saved=(Run 'tablecommit' $args2)|ConvertFrom-Json}finally{$state.loading=$false}
  $state.done=$true;$state.plan=$null;$info.ForeColor=[Drawing.Color]::DarkGreen;$info.Text='저장 완료 · '+$saved.saved+'칸 / 새 열 '+$saved.new_columns+'개. 현재 작업표에 저장했습니다. 상위 표는 [상위에 반영]을 눌러야 바뀝니다.';$save.Enabled=$false;$refresh.Enabled=$false;$pickButton.Enabled=$false;$applyLayout.Enabled=$false;$mapGrid.ReadOnly=$true
 }
 foreach($combo in @($idBox,$classBox,$numberBox,$nameBox)){$combo.Add_SelectedIndexChanged({ImportDirty})};$sheetBox.Add_SelectedIndexChanged({ImportStructureChanged});$headBox.Add_ValueChanged({ImportStructureChanged});$blankBox.Add_CheckedChanged({ImportDirty})
 $mapGrid.Add_CurrentCellDirtyStateChanged({if($mapGrid.IsCurrentCellDirty){[void]$mapGrid.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit)}});$mapGrid.Add_CellValueChanged({ImportDirty})
 $mapGrid.Add_DataError({param($sender,$e) $e.ThrowException=$false;$info.Text='열 연결 선택을 다시 확인하세요.'})
 Ui12Style $d;foreach($grid in @($mapGrid,$studentsGrid,$changesGrid,$issuesGrid)){$grid.DefaultCellStyle.SelectionBackColor=[Drawing.Color]::FromArgb(228,237,247);$grid.DefaultCellStyle.SelectionForeColor=$navy};foreach($button in @($pickButton,$applyLayout,$refresh,$save,$close)){$button.Add_EnabledChanged({Ui12Button $this})};[void]$d.ShowDialog($f);if($state.done){Read-Meta;DLoad $dashMode $dashArea;$dashMessage.Text='자료 가져오기 완료 · 현재 작업표에 저장했습니다. 상위 표는 자동으로 바뀌지 않습니다.'}
}
$script:roundTripAction=$dashUp.ActionScript
$dashUp.ActionScript={DSafe {if(ImportAtLeaf){WImportDialog}else{& $roundTripAction}}}
