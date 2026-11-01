# Pull from the upper table: selected score cells only. No implicit cascade.
function LightPullPlan($selections){
 $changes=@();$hashes=@{};$refs=@{};$seen=@{}
 foreach($selected in $selections){
  $sid=[string]$selected.student;$item=[string]$selected.item;$key=$sid+'|'+$item
  if($seen[$key]){continue};$seen[$key]=$true
  if($item -notin $dashSnapshot.fields){continue}
  $links=@($meta.links|Where-Object {$_.target -eq $dashLayer -and $_.targetItem -eq $item -and $_.sourceItem -notmatch '\.(status|note)$'})
  if(!$links.Count){continue};if($links.Count -ne 1){throw '한 셀에 연결된 하위 값이 여러 개입니다.'};$ln=$links[0];$t=$memoryTables[$ln.source]
  if(!$t -or !$t.rows.ContainsKey($sid)){throw '하위 작업표에 선택 학생의 행이 없습니다.'}
  if($ln.sourceItem -eq '__sum__'){
   $numeric=LightNumericItems $ln.source;$items=$numeric
   foreach($profile in $meta.courses.Values){$a=@($profile.areas|Where-Object {$_.work_layer -eq $ln.source -and $_.root_item -eq $ln.targetItem});if($a.Count -eq 1){$items=@($a[0].work_items|Where-Object {$_ -in $numeric})}}
   $value=LightSum $ln.source $sid $items
  }else{$value=LightValue $ln.source $sid $ln.sourceItem}
  $before=LightValue $dashLayer $sid $item
  if(MemorySame $before $value){continue}
  $changes+=@{student=$sid;item=$item;value=$value;before=$before;note=''}
  $hashes[$ln.source]=$t.hash;$refs[$key]=@{source=$ln.source;item=$ln.sourceItem}
 }
 return @{expected_fingerprint=(MemoryTextHash ($memoryMetaHash+'|'+$memoryTables[$dashLayer].hash));changes=@($changes);source_fingerprints=$hashes;source_refs=$refs}
}
function LightPullSelected{
 DClean;if(!$dashSnapshot){throw '먼저 작업표를 여세요.'}
 $selected=@($dashGrid.SelectedCells|Where-Object {$_.OwningColumn.Tag.item}|ForEach-Object {@{student=[string]$_.OwningRow.Tag;item=[string]$_.OwningColumn.Tag.item}})
 if(!$selected.Count){throw '하위와 다른 점수 셀을 선택하세요. Ctrl로 여러 셀을 선택할 수 있습니다.'}
 MemoryEnsure
 foreach($id in @($memoryTables.Keys)){$t=$memoryTables[$id];$hash=MemoryHash $t.path;if($hash -cne $t.hash){$script:memoryTables[$id]=MemoryReadBook $id $t.path $hash;$script:memoryViews=@{}}}
 $request=LightPullPlan $selected
 if(!$request.changes.Count){LightCompare;$dashGrid.Invalidate();$dashMessage.Text='선택 셀은 하위 값과 같거나 하위 연결이 없습니다.';return}
 $notice=([string]$request.changes.Count)+'개 셀의 기존 값을 하위 값으로 바꾸고 즉시 저장합니다.'+"`r`n"+'[변경 취소]로 되돌릴 수 없습니다.'
 if([Windows.Forms.MessageBox]::Show($f,$notice,'하위값 반영 · 즉시 저장',[Windows.Forms.MessageBoxButtons]::OKCancel,[Windows.Forms.MessageBoxIcon]::Information) -ne [Windows.Forms.DialogResult]::OK){$dashMessage.Text='하위값 반영을 취소했습니다. 기존 값은 유지됩니다.';return}
 $result=(Run 'gridcommit' (DArgs @{PayloadFile=(Write-Request $request)}))|ConvertFrom-Json
 DLoad $dashMode $dashArea;$dashMessage.Text='선택한 '+$result.saved+'개 셀을 하위 값으로 바꾸고 저장했습니다. [변경 취소] 대상이 아닙니다.'
}
$pullCellsMenu=$columnMenu.Items.Insert(0,(New-Object Windows.Forms.ToolStripMenuItem '선택 셀에 하위값 반영'))
# Insert returns void; keep the actual menu object for enabling and tooltips.
$pullCellsMenu=$columnMenu.Items[0]
$pullCellsMenu.ToolTipText='선택 셀의 기존 값을 하위 값으로 바꾸고 즉시 저장합니다. [변경 취소]로 되돌릴 수 없습니다.'
$pullCellsMenu.Add_Click({DSafe {LightPullSelected}})
$columnMenu.Add_Opening({
 $pullCellsMenu.Enabled=(@($dashGrid.SelectedCells|Where-Object {$item=$_.OwningColumn.Tag.item;$item -and ([int]$script:lightMismatchKinds[$dashLayer+'|'+$_.OwningRow.Tag+'|'+$item] -band 2)}).Count -gt 0)
})
