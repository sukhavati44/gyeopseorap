function DPasteRows([string]$text){
 $rows=[Collections.Generic.List[object]]::new();$row=[Collections.Generic.List[string]]::new();$value=[Text.StringBuilder]::new();$quoted=$false;$start=$true
 for($i=0;$i -lt $text.Length;$i++){$ch=$text[$i]
  if($ch -eq '"'){if($quoted -and $i+1 -lt $text.Length -and $text[$i+1] -eq '"'){[void]$value.Append('"');$i++;continue};if($quoted){$quoted=$false;continue};if($start){$quoted=$true;$start=$false;continue}}
  if(!$quoted -and ($ch -eq "`t" -or $ch -eq "`r" -or $ch -eq "`n")){$row.Add($value.ToString());[void]$value.Clear();$start=$true;if($ch -ne "`t"){$rows.Add($row.ToArray());$row=[Collections.Generic.List[string]]::new();if($ch -eq "`r" -and $i+1 -lt $text.Length -and $text[$i+1] -eq "`n"){$i++}};continue}
  [void]$value.Append($ch);$start=$false
 }
 if($quoted){throw '클립보드의 따옴표가 닫히지 않았습니다.'}
 if($value.Length -or $row.Count -or !$rows.Count -or ($text.Length -and $text[$text.Length-1] -eq "`t")){$row.Add($value.ToString());$rows.Add($row.ToArray())}
 return ,$rows.ToArray()
}
function DPasteText([string]$text){
 if($busy -or $uiLoadingDepth -gt 0 -or !$dashSnapshot -or !$dashGrid.CurrentCell){return}
 if(!$dashGrid.EndEdit()){throw '입력 중인 값을 확인하세요.'}
 $matrix=DPasteRows $text;$columns=@($dashGrid.Columns|Where-Object {$_.Visible}|Sort-Object DisplayIndex);$startCol=[array]::IndexOf($columns,$dashGrid.CurrentCell.OwningColumn);$startRow=$dashGrid.CurrentCell.RowIndex
 $visibleRows=@($dashGrid.Rows|Where-Object Visible);$startRow=[array]::IndexOf($visibleRows,$dashGrid.CurrentRow);$plan=[Collections.Generic.List[object]]::new();$count=0
 for($r=0;$r -lt $matrix.Count;$r++){for($c=0;$c -lt $matrix[$r].Count;$c++){
  if(++$count -gt 100000){throw '한 번에 100,000셀 이하로 붙여넣으세요.'}
  if($startRow+$r -ge $visibleRows.Count -or $startCol+$c -ge $columns.Count){throw '붙여넣을 행·열이 부족합니다. 열을 추가하거나 범위를 줄이세요. 붙여넣기는 적용하지 않았습니다.'}
  $col=$columns[$startCol+$c];$row=$visibleRows[$startRow+$r];$cell=$row.Cells[$col.Index];$old=$dashRows[([string]$row.Tag+'|'+$col.Tag.item)]
  if($cell.ReadOnly -or !$col.Tag.item -or !$old){throw '점수·기록 셀에서 시작하세요. 이름·코드·합계 등 읽기 전용 열은 붙여넣지 않습니다.'}
  $v=[string]$matrix[$r][$c];if($v.Length -gt 32767){throw '한 셀의 문자는 32767자 이하로 입력하세요.'}
  if($old.numeric -and $v -ne ''){$n=0.0;if(![double]::TryParse($v,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n) -or [double]::IsNaN($n) -or [double]::IsInfinity($n)){throw ('숫자 열의 값을 확인하세요: 붙여넣기 '+($r+1)+'행 '+($c+1)+'열. 붙여넣기는 적용하지 않았습니다.')}}
  $plan.Add(@{cell=$cell;row=$row;item=$col.Tag.item;value=$v})
 }}
 $wasLoading=$dashLoading;$script:dashLoading=$true;$script:dashPasting=$true
 try{foreach($entry in $plan){$entry.cell.Value=$entry.value;DChange $entry.row $entry.item}}finally{$script:dashLoading=$wasLoading;$script:dashPasting=$false}
 LightCompare;foreach($row in @($plan|ForEach-Object {$_.row}|Sort-Object Index -Unique)){DTotal $row}
 $dashGrid.ClearSelection();foreach($entry in $plan){$entry.cell.Selected=$true};$dashGrid.Invalidate();DSelection
 $dashCount.Text='조회 '+$dashGrid.Rows.Count+'명 · 미저장 '+$dashPending.Count+'건';$dashMessage.Text='클립보드 '+$matrix.Count+'행 / '+$plan.Count+'셀 붙여넣기 완료. [변경 저장] 또는 [변경 취소]를 선택하세요. 상위 값은 변경하지 않았습니다.'
}
function DPasteClipboard {if([Windows.Forms.Clipboard]::ContainsText()){DPasteText ([Windows.Forms.Clipboard]::GetText())}}
$dashGrid.MultiSelect=$true;$dashGrid.SelectionMode='CellSelect'
$dashGrid.Add_KeyDown({param($sender,$e)if($e.Control -and $e.KeyCode -eq [Windows.Forms.Keys]::V){$e.Handled=$true;$e.SuppressKeyPress=$true;DSafe {DPasteClipboard}}})
$script:pasteEditorHandler=[Windows.Forms.KeyEventHandler]{param($sender,$e)if($e.Control -and $e.KeyCode -eq [Windows.Forms.Keys]::V -and [Windows.Forms.Clipboard]::ContainsText()){$text=[Windows.Forms.Clipboard]::GetText();if($text.Contains("`t") -or $text.Contains("`n") -or $text.Contains("`r")){$e.Handled=$true;$e.SuppressKeyPress=$true;DSafe {DPasteText $text}}}}
$dashGrid.Add_EditingControlShowing({param($sender,$e)$e.Control.Remove_KeyDown($pasteEditorHandler);$e.Control.Add_KeyDown($pasteEditorHandler)})
$pasteMenu=$columnMenu.Items.Add('클립보드 표 붙여넣기 (Ctrl+V)');$pasteMenu.Add_Click({DSafe {DPasteClipboard}})
