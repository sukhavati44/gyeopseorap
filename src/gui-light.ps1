# Lightweight presentation. Historical status metadata is ignored, never used to gate values.
$script:lightMismatch=@{};$script:lightMismatchKinds=@{}

function LightValue($layer,$sid,$item){
 if($layer -eq $dashLayer -and $dashPending[$sid+'|'+$item]){return $dashPending[$sid+'|'+$item].value}
 $t=$memoryTables[$layer];if(!$t -or !$t.rows.ContainsKey($sid) -or !$t.cols.ContainsKey($item)){return $null}
 return $t.sheet.Rows[$t.rows[$sid]][[int]$t.cols[$item]].Value
}
function LightNumericItems($layer){
 $numeric=@($meta.courses.Values|ForEach-Object {$_.numeric_items}|Sort-Object -Unique)
 return ,@($meta.layers[$layer].items|Where-Object {$_ -in $numeric -and $_ -notmatch '\.(status|note)$'})
}
function LightSum($layer,$sid,$items=$null){
 $sum=[decimal]0;$count=0
 if($null -eq $items){$items=LightNumericItems $layer}
 foreach($item in $items){$v=LightValue $layer $sid $item;if($v -is [ValueType] -and $v -isnot [bool]){$sum += [decimal]$v;$count++}}
 if($count){return $sum};return $null
}
function LightMark($key,$kind,$message){
 $script:lightMismatchKinds[$key]=[int]$script:lightMismatchKinds[$key] -bor $kind
 if($script:lightMismatch[$key]){$script:lightMismatch[$key]+=' · '+$message}else{$script:lightMismatch[$key]=$message}
}
function LightMismatchColor($key){
 switch([int]$script:lightMismatchKinds[$key]){
  1 {return [Drawing.Color]::FromArgb(215,235,255)}
  2 {return [Drawing.Color]::FromArgb(255,232,179)}
  3 {return [Drawing.Color]::FromArgb(232,216,250)}
  default {return [Drawing.Color]::White}
 }
}
function LightCompare{
 $script:lightMismatch=@{};$script:lightMismatchKinds=@{}
 foreach($ln in $meta.links){if($ln.sourceItem -match '\.(status|note)$' -or $ln.targetItem -match '\.(status|note)$'){continue}
   $numericSource=LightNumericItems $ln.source;$sumItems=$numericSource
   foreach($profile in $meta.courses.Values){$ap=@($profile.areas|Where-Object {$_.work_layer -eq $ln.source -and $_.root_item -eq $ln.targetItem});if($ap.Count -eq 1){$sumItems=@($ap[0].work_items|Where-Object {$_ -in $numericSource})}}
  foreach($sid in $meta.students.Keys){
   $st=$memoryTables[$ln.source];$tt=$memoryTables[$ln.target]
   if(!$st -or !$tt -or !$st.rows.ContainsKey($sid) -or !$tt.rows.ContainsKey($sid)){continue}
   $lower=if($ln.sourceItem -eq '__sum__'){LightSum $ln.source $sid $sumItems}else{LightValue $ln.source $sid $ln.sourceItem}
   $upper=LightValue $ln.target $sid $ln.targetItem
   if(!(MemorySame $lower $upper)){
    LightMark ($ln.source+'|'+$sid+'|'+$ln.sourceItem) 1 ('상위와 다름: 현재 '+(ValueText $lower)+' / 상위 '+(ValueText $upper))
    LightMark ($ln.target+'|'+$sid+'|'+$ln.targetItem) 2 ('하위와 다름: 하위 '+(ValueText $lower)+' / 현재 '+(ValueText $upper))
   }
  }
 }
}
function DTotal($row){$numericHere=LightNumericItems $dashLayer;$items=@($dashSnapshot.fields|Where-Object {$_ -in $numericHere});$v=LightSum $dashLayer ([string]$row.Tag) $items;$row.Cells['total'].Value=$v}
function DSelection{
 if($dashLoading){return};$cell=$dashGrid.CurrentCell
 if($dashMessage.Text -like '하위값 반영: *'){$dashMessage.Text='값을 입력하거나 Excel 값을 반영하세요. 상위·하위 불일치는 색으로 표시합니다.'}
 if(!$cell){$dashSelected.Text='학생의 점수 또는 기록을 선택하세요.';return}
 $item=if($cell.OwningColumn.Name -eq 'total'){'__sum__'}else{$cell.OwningColumn.Tag.item};$sid=[string]$cell.OwningRow.Tag
 $script:dashSelectedKey=$sid+'|'+$item
 $dashSelected.Text=(WStudentLabel $sid)+' / '+$cell.OwningColumn.HeaderText.Replace("`n",' ')
 $different=$script:lightMismatch[$dashLayer+'|'+$sid+'|'+$item]
 if($different){$dashSelected.Text+=' / '+$different}
 if($cell.OwningColumn.Tag.item -and ([int]$script:lightMismatchKinds[$dashLayer+'|'+$sid+'|'+$item] -band 2)){
  $dashMessage.Text='하위값 반영: 셀 우클릭 → [선택 셀에 하위값 반영] · 여러 셀은 드래그 선택 · 반영 즉시 저장'
 }
}
$script:lightOriginalLoad=${function:DLoad}
function DLoad($mode,$area){
 & $lightOriginalLoad $mode $area
 # Read changed linked files into RAM before comparing; no Excel/worker for navigation.
 foreach($id in @($memoryTables.Keys)){$t=$memoryTables[$id];$hash=MemoryHash $t.path;if($hash -cne $t.hash){$script:memoryTables[$id]=MemoryReadBook $id $t.path $hash;$script:memoryViews=@{}}}
 LightCompare;foreach($row in $dashGrid.Rows){DTotal $row};$dashGrid.Invalidate();DSelection
}
$script:lightOriginalLayout=${function:DLayout}
function DLayout{
 & $lightOriginalLayout
 $h=$dashboard.ClientSize.Height-52
 $dashGrid.Top=285;$dashGrid.Height=[math]::Max(300,$h-420);$dashSelected.SetBounds(210,$h-115,$dashGrid.Width,42)
 
}

foreach($spec in @(@('상위 값과 다름',215,235,255),@('하위 값과 다름',255,232,179),@('상·하위 모두 다름',232,216,250))){$x=210+(@('상위 값과 다름','하위 값과 다름','상·하위 모두 다름').IndexOf($spec[0]))*220;$legend=Label $dashboard $spec[0] $x 258 205 20;$legend.BackColor=[Drawing.Color]::FromArgb($spec[1],$spec[2],$spec[3]);$legend.TextAlign='MiddleCenter'}
$dashTitle.Text=$productTitle
$f.Text=$productTitle
DLayout

# Keep only source import and recovery behind the advanced view.
foreach($page in @($tabs.TabPages)){if($page -ne $stagingTab -and $page -ne $helpTab){$tabs.TabPages.Remove($page)}}
