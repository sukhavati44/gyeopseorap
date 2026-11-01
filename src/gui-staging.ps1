$stagingTab=New-Object Windows.Forms.TabPage;$stagingTab.Text='자료 반입·검토';[void]$tabs.TabPages.Add($stagingTab);$tabs.SelectedTab=$stagingTab
[void](Label $stagingTab '원문을 대조하고 학생·항목·값을 지정한 기록만 새 하위층에 저장합니다.' 15 8 880)
$stagingFile=TextBox $stagingTab 15 37 640;$stagingFile.ReadOnly=$true
Button $stagingTab '로컬 자료 선택' 675 33 220 {Safe {$d=New-Object Windows.Forms.OpenFileDialog;$d.Filter='Excel 또는 교환자료|*.xlsx;*.json';if($d.ShowDialog($f) -eq 'OK'){$stagingFile.Text=$d.FileName;$script:stagingPlan=$null;$script:stagingRowMappings=@{};$stagingGrid.Rows.Clear();$stagingDetail.Clear();$stagingSave.Enabled=$false}}}|Out-Null
Button $stagingTab '읽기 전용 검토' 15 75 250 {Safe {
 $script:stagingParent=$current;$script:stagingDate=$ddate.Value.ToString('yyyy-MM-dd');$script:stagingPlan=(Run 'stagingpreview' @{File=$stagingFile.Text;Parent=$stagingParent;AsOf=$stagingDate})|ConvertFrom-Json
 $script:stagingRowMappings=@{};Read-Meta;Fill-StagingStudents;Fill-StagingPreview
}}|Out-Null
$stagingSummary=Label $stagingTab '현재 작업층을 고른 뒤 승인된 로컬 자료를 선택하세요.' 280 80 610 55
$stagingSummary.SetBounds(280,75,335,60)
Button $stagingTab '미등록 식별번호로 이동' 625 75 270 {Safe {
 if(!$stagingPlan){throw '읽기 전용 검토를 먼저 누르세요.'};$found=-1
 for($ri=0;$ri -lt $stagingPlan.records.Count;$ri++){$rec=$stagingPlan.records[$ri];if(!$rec.resolved -and !$meta.students.ContainsKey([string]$rec.source_subject_key)){$found=$ri;break}}
 if($found -lt 0){$stagingSummary.Text='미등록 식별번호 기록이 없습니다. 원문 위치를 대조해 필요한 셀을 선택하세요.';return}
 $stagingGrid.CurrentCell=$stagingGrid.Rows[$found].Cells[0];$stagingGrid.FirstDisplayedScrollingRowIndex=$found;Load-StagingSelection
 $stagingSummary.Text='이 원문 식별번호는 명부와 자동 연결할 수 없습니다. 해당 행의 점수 셀을 선택하고 아래 학생·항목를 지정하세요.'
}}|Out-Null
$stagingGrid=Grid $stagingTab 15 140 880 95 @('원문 식별번호','원문 항목','원문값','출처','검토 상태')
$stagingDetail=TextBox $stagingTab 15 242 880 65;$stagingDetail.ReadOnly=$true
$stagingClass=Combo $stagingTab 15 317 150;$stagingStudent=Combo $stagingTab 180 317 715
[void](Label $stagingTab '평가항목' 15 360 80);$stagingItem=TextBox $stagingTab 95 355 800
[void](Label $stagingTab '확정할 값' 15 400 80);$stagingValue=TextBox $stagingTab 95 395 410
$stagingKind=Combo $stagingTab 520 395 140;Choice $stagingKind 'number' '숫자';Choice $stagingKind 'text' '문자';$stagingKind.SelectedIndex=1
$stagingClear=New-Object Windows.Forms.CheckBox;$stagingClear.Text='빈값(null)';$stagingClear.SetBounds(680,395,210,30);$stagingTab.Controls.Add($stagingClear)

$stagingSave=Button $stagingTab '선택 기록 확인·저장 → 다음' 15 490 410 {Safe {
 if(!$stagingPlan -or $current -ne $stagingParent -or $ddate.Value.ToString('yyyy-MM-dd') -ne $stagingDate){throw '현재 작업층·기준일에서 다시 검토하세요.'}
 $row=$stagingGrid.CurrentRow;if(!$row -or !$row.Tag){throw '원문 기록을 선택하세요.'};$rec=@($stagingPlan.records|Where-Object source_record_id -CEQ $row.Tag)[0]
 if(!$stagingStudent.SelectedItem -or !$stagingStudent.SelectedItem.Id){throw '등록 학생을 명시적으로 선택하세요.'}
 if(!$stagingItem.Text.Trim()){throw '평가항목을 입력하세요.'}
 $v=$null;if(!$stagingClear.Checked){if($stagingKind.SelectedItem.Id -eq 'number'){$d=0.0;if(![double]::TryParse($stagingValue.Text,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){throw '숫자 형식을 확인하세요.'};$v=$d}else{$v=$stagingValue.Text}}

 $payload=Write-Request @{expected_fingerprint=$stagingPlan.fingerprint;source_record_id=$rec.source_record_id;student_id=$stagingStudent.SelectedItem.Id;item_label=$stagingItem.Text;value=$v;status=$(if($null -eq $v){'미입력'}else{'입력됨'});review_note=''}
 $saved=(Run 'stagingcommit' @{File=$stagingFile.Text;Parent=$stagingParent;AsOf=$stagingDate;PayloadFile=$payload})|ConvertFrom-Json
 # Remember only the teacher-confirmed identity for this exact source row.
 $script:stagingRowMappings[$rec.context_id]=[string]$stagingStudent.SelectedItem.Id
 $savedSourceHash=$stagingPlan.source_hash
 $script:stagingPlan=$null;$stagingSave.Enabled=$false
 Refresh-Navigation $stagingParent
 $stagingSummary.Text='선택 기록 저장 완료. 다음 기록을 읽고 있습니다.'
 $script:stagingPlan=(Run 'stagingpreview' @{File=$stagingFile.Text;Parent=$stagingParent;AsOf=$stagingDate})|ConvertFrom-Json
 if($stagingPlan.source_hash -ne $savedSourceHash){$script:stagingRowMappings=@{}}
 Fill-StagingPreview
 $stagingSummary.Text='1건 저장 완료 · '+@($stagingPlan.records|Where-Object {!$_.resolved}).Count+'건 남음. 같은 원문 행은 확인한 학생 연결을 이어 씁니다. 저장 전 다시 확인하세요.'

}};$stagingSave.Enabled=$false
[void](Label $stagingTab '학생·항목·값을 지정하고 저장하세요.' 445 494 450 30)
function Fill-StagingPreview{
 $stagingGrid.Rows.Clear();$first=-1;$resolved=0
 foreach($r in $stagingPlan.records){
  $i=$stagingGrid.Rows.Add([object[]]@($r.source_subject_key,$r.item_label,(ValueText $r.raw_value),$r.source_cell,$(if($r.resolved){'저장 완료'}else{$r.reason})))
  $stagingGrid.Rows[$i].Tag=$r.source_record_id
  if($r.resolved){$resolved++;$stagingGrid.Rows[$i].DefaultCellStyle.BackColor=[Drawing.Color]::Honeydew}elseif($first -lt 0){$first=$i}
 }
 if($first -ge 0){$stagingGrid.CurrentCell=$stagingGrid.Rows[$first].Cells[0];$stagingGrid.FirstDisplayedScrollingRowIndex=$first}
 Load-StagingSelection
 $stagingSummary.Text='전체 '+$stagingPlan.records.Count+'건 · 저장 완료 '+$resolved+'건 · 원문 확인 필요 '+$stagingPlan.quarantine.Count+'건. 외부 전송 없음.'
}
function Fill-StagingStudents{
 $stagingClass.Items.Clear();Choice $stagingClass '' '모든 반';foreach($c in @($meta.memberships|ForEach-Object {$_.class}|Sort-Object -Unique)){Choice $stagingClass $c $c};$stagingClass.SelectedIndex=0;Filter-StagingStudents
}
function Filter-StagingStudents{
 $stagingStudent.Items.Clear();Choice $stagingStudent '' '학생을 선택하세요 (자동 연결 안 함)';$cls=if($stagingClass.SelectedItem){$stagingClass.SelectedItem.Id}else{''}
 if($meta){foreach($id in ($meta.students.Keys|Sort-Object)){$s=$meta.students[$id];$member=@($meta.memberships|Where-Object {$_.student -eq $id -and $_.from -le $stagingDate -and (!$_.until -or $_.until -ge $stagingDate)});if($cls -and $cls -notin @($member.class)){continue};Choice $stagingStudent $id ($s.name+' · '+($member.class -join ',')+' · '+$s.number+'번')}};$stagingStudent.SelectedIndex=0
}
function Load-StagingSelection{
 $stagingValue.Clear();$stagingDetail.Clear();$stagingItem.Clear();$stagingSave.Enabled=$false
 if(!$stagingPlan -or !$stagingGrid.CurrentRow -or !$stagingGrid.CurrentRow.Tag){return};$r=@($stagingPlan.records|Where-Object source_record_id -CEQ $stagingGrid.CurrentRow.Tag)[0];if(!$r){return}
 $stagingDetail.Text='원문 위치: '+$r.source_cell+"`r`n"+(($stagingPlan.contexts.($r.context_id)|ForEach-Object {if($_.label){$_.label+': '+[string]$_.value}else{'원문값: '+(ValueText $_.raw_value)+"`r`n변환값: "+(ValueText $_.value)+"`r`n항목: "+$_.item_label+"`r`n변환 결과의 근거: "+$_.review_note}})-join "`r`n")
 $stagingItem.Text=$r.item_label;$stagingValue.Text=if($null -eq $r.value){''}else{[string]$r.value};$stagingClear.Checked=($null -eq $r.value);Choose $stagingKind $(if($r.value -is [ValueType] -and $r.value -isnot [bool]){'number'}else{'text'});if($stagingStudent.Items.Count){$stagingStudent.SelectedIndex=0};if($script:stagingRowMappings -and $script:stagingRowMappings.ContainsKey($r.context_id)){Choose $stagingStudent $script:stagingRowMappings[$r.context_id]};$stagingSave.Enabled=(!$r.resolved)
}
$stagingClass.Add_SelectedIndexChanged({Filter-StagingStudents});$stagingGrid.Add_SelectionChanged({Load-StagingSelection})

