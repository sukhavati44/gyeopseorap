function Graph-Depth($id){$seen=@{};$n=0;while($id){if($seen.ContainsKey($id)){throw '층 관계 순환을 발견했습니다.'};$seen[$id]=$true;$id=(Layer-Info $id).parent;$n++};return $n}
function Shared-Hash($path){
 $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite);$h=[Security.Cryptography.SHA256]::Create()
 try{return ([BitConverter]::ToString($h.ComputeHash($stream))).Replace('-','')}finally{$h.Dispose();$stream.Dispose()}
}
function Fingerprint {
 $parts=[Collections.Generic.List[string]]::new();$parts.Add((Shared-Hash $MetaPath))
 foreach($id in ($M.layers.Keys|Sort-Object)){$parts.Add($id+'|'+(Shared-Hash (Resolve-Local $M.layers[$id].path)))}
 $parts.Add("$Source|$SourceItem|$Target|$TargetItem")
 $sha=[Security.Cryptography.SHA256]::Create();try{return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`n"))))).Replace('-','')}finally{$sha.Dispose()}
}
function Snapshot($id){
 if($script:parentFastSnapshot){
  $book=[NaturalXlsx]::Read((Resolve-Local $M.layers[$id].path));$s=@($book.Sheets|Where-Object Name -EQ 'Data')[0];$identity=@($book.Sheets|Where-Object Name -EQ '_identity')[0]
  if(!$s -or !$identity -or [string]$identity.Rows[1][2].Value -cne $id){throw '작업표 ID를 확인하세요.'}
  $cols=@{};$rows=@{};$warnings=@();$lastRow=1;$lastCol=1
  foreach($cn in $s.Rows[1].Keys){$key=[string]$s.Rows[1][$cn].Value;if($cols.ContainsKey($key)){throw '중복 항목 ID'};$cols[$key]=[int]$cn;$lastCol=[math]::Max($lastCol,[int]$cn)}
  if(!$cols.student_id){throw '학생 ID 열 누락'}
  foreach($rn in $s.Rows.Keys){$lastRow=[math]::Max($lastRow,[int]$rn);if($rn -le 1){continue};$sid=[string]$s.Rows[$rn][$cols.student_id].Value;if(!$M.students.ContainsKey($sid) -or $rows.ContainsKey($sid)){throw '학생 ID 누락 또는 중복'};$rows[$sid]=[int]$rn}
  $values=New-Object 'object[,]' ($lastRow+1),($lastCol+1);$formulas=New-Object 'object[,]' ($lastRow+1),($lastCol+1)
  foreach($rn in $s.Rows.Keys){foreach($cn in $s.Rows[$rn].Keys){if($cn -le $lastCol){$values[[int]$rn,[int]$cn]=$s.Rows[$rn][$cn].Value;$formulas[[int]$rn,[int]$cn]=$s.Rows[$rn][$cn].Formula}}}
  return @{cols=$cols;rows=$rows;values=$values;formulas=$formulas;warnings=$warnings}
 }
 $b=Open-Book $id $true
 try{$t=Read-Table $b;return @{cols=$t.cols;rows=$t.rows;values=$t.values;formulas=$t.formulas;warnings=@($t.warnings)}}finally{$b.Close($false)}
}
function Label-For($layerId,$field){$label=$M.labels["$layerId|$field"];if(!$label){$label=$field};return $label}
function Source-State($t,$id,$field,$value){if($null -eq $value){return '미입력'};return '입력됨'}
function Build-Plan {
 $before=Fingerprint;$links=@($M.links|Where-Object {$_.sourceItem -notmatch '\.(status|note)$' -and $_.targetItem -notmatch '\.(status|note)$'});$candidate=$null
 if($Source -or $Target -or $SourceItem -or $TargetItem){
  if(!$Source -or !$Target -or !$SourceItem -or !$TargetItem){throw '하위 열과 상위 열을 모두 선택하세요.'}
  if((Layer-Info $Source).parent -ne $Target){throw '직접 부모 표에만 연결할 수 있습니다.'}
  $existing=@($links|Where-Object {$_.target -eq $Target -and $_.targetItem -eq $TargetItem})
  if($existing.Count -and ($existing[0].source -ne $Source -or $existing[0].sourceItem -ne $SourceItem)){throw '이미 다른 열이 연결된 상위 항목입니다. 자동 교체하지 않습니다.'}
  if(!$existing.Count){$candidate=@{source=$Source;sourceItem=$SourceItem;target=$Target;targetItem=$TargetItem};$links+=,$candidate}
 }
 foreach($ln in $links){if(($M.retiredItems -and ($M.retiredItems.ContainsKey($ln.source+'|'+$ln.sourceItem) -or $M.retiredItems.ContainsKey($ln.target+'|'+$ln.targetItem))) -or !$M.layers.ContainsKey($ln.source) -or !$M.layers.ContainsKey($ln.target)){throw '휴지통 항목의 연결은 사용할 수 없습니다.'}}
 $ordered=@($links|Sort-Object @{Expression={Graph-Depth $_.source};Descending=$true},source,sourceItem)
 $tables=@{};$warnings=[Collections.Generic.List[object]]::new()
 foreach($key in $M.reviewItems.Keys){$warnings.Add(@{layer=$key.Split('|')[0];message=('검토 권장: '+$M.reviewItems[$key])})}
 foreach($layerId in ($M.layers.Keys|Sort-Object)){
  $tables[$layerId]=Snapshot $layerId
  foreach($w in $tables[$layerId].warnings){$warnings.Add(@{layer=$layerId;message=$w})}
 }
 $rows=[Collections.Generic.List[object]]::new();$order=[Collections.Generic.List[string]]::new()
 foreach($ln in $ordered){
  if((Layer-Info $ln.source).parent -ne $ln.target){throw '등록된 연결이 부모 관계와 일치하지 않습니다.'}
  $st=$tables[$ln.source];$tt=$tables[$ln.target]
  if(($ln.sourceItem -ne '__sum__' -and !$st.cols.ContainsKey($ln.sourceItem)) -or !$tt.cols.ContainsKey($ln.targetItem)){throw "연결된 항목이 삭제되거나 헤더ID가 바뀌었습니다: $($ln.sourceItem) -> $($ln.targetItem)"}
  $path="$($ln.source)|$($ln.sourceItem)|$($ln.target)|$($ln.targetItem)";$order.Add($path)
  foreach($id in ($M.students.Keys|Sort-Object)){
   $targetKey="$($ln.target)|$id|$($ln.targetItem)";$sourceKey="$($ln.source)|$id|$($ln.sourceItem)"
   $old=$null;$value=$null;$formula='';$reason='반영';$state='미입력'
   if($tt.rows.ContainsKey($id)){$old=$tt.values[$tt.rows[$id],$tt.cols[$ln.targetItem]];$formula=[string]$tt.formulas[$tt.rows[$id],$tt.cols[$ln.targetItem]]}
   if($ln.sourceItem -ne '__sum__' -and $st.rows.ContainsKey($id)){$value=$st.values[$st.rows[$id],$st.cols[$ln.sourceItem]];$state=Source-State $st $id $ln.sourceItem $value;if($M.importRecords -and $M.importRecords.ContainsKey($sourceKey)){$state=$M.importRecords[$sourceKey].status};if($M.manualOverrides.ContainsKey($sourceKey)){$state=$M.manualOverrides[$sourceKey].status}}
   if($ln.sourceItem -eq '__sum__'){
    $numeric=@($M.courses.Values|ForEach-Object {$_.numeric_items}|Sort-Object -Unique);$sumItems=@((Layer-Info $ln.source).items|Where-Object {$_ -in $numeric -and $_ -notmatch '\.(status|note)$'})
    if($Name -and $M.courses[$Name]){$ap=@($M.courses[$Name].areas|Where-Object {$_.work_layer -eq $ln.source -and $_.root_item -eq $ln.targetItem});if($ap.Count -eq 1){$sumItems=@($ap[0].work_items|Where-Object {$_ -in $numeric})}}
    $value=$null;$state='미입력';$complete=$true;$sumWhy=@();$total=0.0;$count=0
    foreach($si in $sumItems){if($st.rows.ContainsKey($id) -and $st.cols.ContainsKey($si)){$sv=$st.values[$st.rows[$id],$st.cols[$si]];if($sv -is [ValueType] -and $sv -isnot [bool]){$total += [double]$sv;$count++}}}
    if($count){$value=$total;$state='입력됨'}

   }
   if(!$tt.rows.ContainsKey($id)){$reason='대상ID누락/격리'}elseif(!$st.rows.ContainsKey($id)){$reason='원본ID누락/격리'}
   $after=$old;if($reason -eq '반영'){$after=$value;$tt.values[$tt.rows[$id],$tt.cols[$ln.targetItem]]=$value}
   $review=''
   $membership=@($M.memberships|Where-Object {$_.student -eq $id -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)})
   $studentLabel=$M.students[$id].name+' · '+($membership.class -join ',')+' · '+$M.students[$id].number+'번'
   $rows.Add(@{canonical=$targetKey;sourceKey=$sourceKey;student=$id;name=$M.students[$id].name;studentLabel=$studentLabel;item=$ln.targetItem;itemLabel=(Label-For $ln.target $ln.targetItem);source=$ln.source;sourceItem=$ln.sourceItem;target=$ln.target;path=$path;before=$old;proposed=$value;after=$after;disposition=$reason;state=$state;review=$review})
  }
 }
 $afterFingerprint=Fingerprint;if($before -ne $afterFingerprint){throw '미리보기 중 파일이 바뀌었습니다. 저장하고 닫은 뒤 다시 확인하세요.'}
 $counts=@{};foreach($r in $rows){if(!$counts.ContainsKey($r.disposition)){$counts[$r.disposition]=0};$counts[$r.disposition]++}
 return @{version='0.3';fingerprint=$before;candidate=$candidate;order=$order.ToArray();rows=$rows.ToArray();warnings=$warnings.ToArray();counts=$counts;readOnly=$true}
}
function Apply-Plan($plan){
 if($PreviewFingerprint -and $PreviewFingerprint -ne $plan.fingerprint){throw '미리보기 이후 파일/관계가 바뀌어 반영을 중단했습니다. 다시 미리보기를 누르세요.'}
 if($Action -eq 'commit' -and !$PreviewFingerprint){throw '확인한 미리보기가 없습니다.'}
 $cpRel='backups/sync-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8);$cp=Resolve-Local $cpRel
 New-Item -ItemType Directory $cp|Out-Null;Copy-Item -LiteralPath $MetaPath -Destination (Join-Path $cp 'relations.json')
 $books=@{};$tables=@{}
 try{
  $targets=@($plan.rows|Where-Object disposition -EQ '반영'|ForEach-Object target|Sort-Object -Unique)
  foreach($id in ($M.layers.Keys|Sort-Object)){$dest=Join-Path $cp $M.layers[$id].path;New-Item -ItemType Directory (Split-Path $dest) -Force|Out-Null;if($id -in $targets){$books[$id]=Open-Book $id;$tables[$id]=Read-Table $books[$id];$books[$id].SaveCopyAs($dest)}else{Copy-Item -LiteralPath (Resolve-Local $M.layers[$id].path) -Destination $dest}}
  if((Fingerprint) -ne $plan.fingerprint){throw '반영 직전 파일이 바뀌었습니다. 새 미리보기가 필요합니다.'}
  @{checkpoint=$cpRel;started=(Get-Date -Format o)}|ConvertTo-Json|Set-Content (Join-Path $Root 'pending-sync.json') -Encoding utf8
  if($plan.candidate){$M.links+=,$plan.candidate}
  $changed=@{};$now=Get-Date -Format o;$groups=@{}
  foreach($r in $plan.rows){
   if($r.disposition -ne '반영'){continue}
   $t=$tables[$r.target]

   $v=$r.after
   $g=$r.target+'|'+$r.item;if(!$groups.ContainsKey($g)){$groups[$g]=[Collections.Generic.List[object]]::new()};$groups[$g].Add($r);$changed[$r.target]=$true
   $M.manualOverrides.Remove($r.canonical)
   $M.provenance[$r.canonical]=@{source=$r.source;item=$r.sourceItem;value=$v;refreshed=$now}
  }
  function Flush-Segment($segment,$table,$column){
   if(!$segment.Count){return};$a=New-Object 'object[,]' $segment.Count,1
   for($i=0;$i -lt $segment.Count;$i++){$v=$segment[$i].after;if($null -eq $v){$a[$i,0]=$null}elseif($v -is [bool]){$a[$i,0]=[bool]$v}elseif($v -is [ValueType]){$a[$i,0]=[double]$v}else{$a[$i,0]=[string]$v}}
   $start=$table.rows[$segment[0].student];$end=$table.rows[$segment[$segment.Count-1].student]
   $range=$table.sheet.Range($table.sheet.Cells.Item($start,$column),$table.sheet.Cells.Item($end,$column));$range.ClearContents()|Out-Null;for($j=0;$j -lt $segment.Count;$j++){$v=$segment[$j].after;$cell=$table.sheet.Cells.Item(($start+$j),$column);if($null -eq $v){[void]$cell.ClearContents()}elseif($v -is [ValueType]){$cell.Value2=[double]$v}else{$cell.Value2="'"+[string]$v}};$range.Interior.Pattern=-4142
  }
  foreach($g in $groups.Keys){
   $rows=@($groups[$g]|Sort-Object @{Expression={$tables[$_.target].rows[$_.student]}});$t=$tables[$rows[0].target];$col=$t.cols[$rows[0].item];$segment=[Collections.Generic.List[object]]::new();$previous=-2
   foreach($r in $rows){$pos=$t.rows[$r.student];if($pos -ne $previous+1 -and $segment.Count){Flush-Segment $segment $t $col;$segment.Clear()};$segment.Add($r);$previous=$pos};Flush-Segment $segment $t $col
  }
  foreach($id in $changed.Keys){$books[$id].Save()};Save-Meta
  $plan.readOnly=$false;$plan.checkpoint=$cpRel;$plan.appliedAt=$now
  $plan|ConvertTo-Json -Depth 16|Set-Content (Join-Path $Root ('logs/sync-'+(Get-Date -Format 'yyyyMMdd-HHmmss-fff')+'.json')) -Encoding utf8
  Move-Item -LiteralPath (Join-Path $Root 'pending-sync.json') -Destination (Join-Path $Root ('logs/completed-'+[guid]::NewGuid().ToString('N')+'.json'))
  return $plan
 }finally{foreach($book in $books.Values){$book.Close($false)}}
}
