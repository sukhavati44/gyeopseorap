# Shared workspace structure, scoped reflection and local synthetic code exchange.
function WArea($p){$a=@(TArea $p $Area|Where-Object {$_});if($a.Count -ne 1){throw '평가영역을 선택하세요.'};return $a[0]}
function WWithin($child,$ancestor){$seen=@{};while($child){if($child -eq $ancestor){return $true};if($seen[$child]){throw '작업공간 관계 순환'};$seen[$child]=$true;$child=(Layer-Info $child).parent};return $false}
function WFields($p){
 if($Layer -eq $p.root){return ,@($p.areas|ForEach-Object {$_.root_item})};$a=WArea $p
 if($Layer -eq $a.work_layer){return ,@($a.work_items)}
 if(!(TDetached $Layer) -and (!$a.work_layer -or !(WWithin $Layer $a.work_layer))){throw '선택한 교육과정·영역 밖 작업공간입니다.'}
 if($M.workspaceScopes -and $M.workspaceScopes[$Layer] -and ($M.workspaceScopes[$Layer].course -ne $Name -or $M.workspaceScopes[$Layer].area -ne $Area)){throw '다른 영역에 속한 작업공간입니다.'}
 $items=@((Layer-Info $Layer).items|Where-Object {$_ -notmatch '\.(status|note)$'})
 if(!$M.workspaceScopes -or !$M.workspaceScopes[$Layer]){$prefix=$a.root_item.Split('.')[0];$items=@($items|Where-Object {$_.Split('.')[0] -eq $prefix})}
 return ,$items
}
function WCreate($req){
 $p=CProfile;$fields=WFields $p;$a=WArea $p
 if($req.target_item -and $req.target_item -notin $fields){throw '연결할 상위 열을 확인하세요.'};if($req.target_item -and @($M.links|Where-Object {$_.target -eq $Layer -and $_.targetItem -eq $req.target_item}).Count){throw '이미 연결된 열입니다.'}
 if($req.target_item){
  if($M.workspaceScopes -and @($M.workspaceScopes.Keys|Where-Object {$M.layers[$_].parent -eq $Layer -and $M.workspaceScopes[$_].parent_item -eq $req.target_item}).Count){throw ((Layer-Info $Layer).name+' / '+(Label-For $Layer $req.target_item)+' 열에는 이미 하위 공간이 있습니다. 기존 공간을 여세요. 점수와 연결은 유지했습니다.')}
  Excel;try{$check=Snapshot $Layer;if($check.warnings.Count -or !$check.cols.ContainsKey($req.target_item)){throw '선택 열의 학생·항목 식별자를 확인하세요. 기존 자료는 유지했습니다.'};foreach($sid in $check.rows.Keys){if(([string]$check.formulas[$check.rows[$sid],$check.cols[$req.target_item]]).StartsWith('=')){throw ((Layer-Info $Layer).name+' / '+(Label-For $Layer $req.target_item)+' 열에 수식이 있습니다. 수식을 보호하기 위해 새 연결을 만들지 않았습니다. 다른 점수 열을 선택하세요.')}}}finally{Close-Excel}
 }
 $label=([string]$req.label).Trim();if(!$label -or $label.Length -gt 80 -or $label -match '[\r\n]'){throw '작업공간 이름은 1~80자로 입력하세요.'}
 $metaBefore=Shared-Hash $MetaPath;$participation=($M.participation|ConvertTo-Json -Depth 10 -Compress);$membership=($M.memberships|ConvertTo-Json -Depth 20 -Compress);$oldBooks=@{};foreach($lid in $M.layers.Keys){$oldBooks[$lid]=Shared-Hash (Resolve-Local $M.layers[$lid].path)}
 $id='W'+[guid]::NewGuid().ToString('N');$prefix=$a.root_item.Split('.')[0];$item=$prefix+'.w'+[guid]::NewGuid().ToString('N').Substring(0,12)
 $cpRel='backups/space-'+[guid]::NewGuid().ToString('N');$cp=Resolve-Local $cpRel;New-Item -ItemType Directory $cp|Out-Null;Copy-Item $MetaPath (Join-Path $cp 'relations.json')
 foreach($existingId in $M.layers.Keys){$relative=$M.layers[$existingId].path;$destination=Join-Path $cp $relative;New-Item -ItemType Directory (Split-Path $destination) -Force|Out-Null;Copy-Item -LiteralPath (Resolve-Local $relative) -Destination $destination}
 @{checkpoint=$cpRel;action='space_create';created_layer=$id}|ConvertTo-Json|Set-Content (Join-Path $Root 'pending-sync.json') -Encoding utf8
 Excel;try{New-Layer $id $Layer @($item,($prefix+'.status'));$structure=WVerifyRoster $id @($item,($prefix+'.status'))}finally{Close-Excel}
 $M.layers[$id].name=$label;$M.labels["$id|$item"]='점수 1';if(!$M.workspaceScopes){$M.workspaceScopes=@{}};$M.workspaceScopes[$id]=@{course=$Name;area=$Area;created_at=(Get-Date -Format o);target_students=@($M.students.Keys|Sort-Object);structure_only=$true}
 if($req.numeric){$p.numeric_items+=,$item};if($req.target_item){$M.workspaceScopes[$id].parent_item=[string]$req.target_item;$M.workspaceScopes[$id].aggregation='sum';if($Layer -eq $p.root){$a.work_layer=$id;$a.work_items=@($item)}}
 if(($M.participation|ConvertTo-Json -Depth 10 -Compress) -cne $participation){throw '참여 상태가 변경되어 생성 확인을 중단했습니다.'}
 if(($M.memberships|ConvertTo-Json -Depth 20 -Compress) -cne $membership){throw '소속 이력 보존 확인이 필요합니다. 복구 사본을 대조하세요.'};foreach($lid in $oldBooks.Keys){if((Shared-Hash (Resolve-Local $M.layers[$lid].path)) -ne $oldBooks[$lid]){throw '기존 작업표가 생성 도중 변경되었습니다. 추가 저장을 멈추고 복구 사본을 대조하세요.'}}
 Save-Meta;Move-Item (Join-Path $Root 'pending-sync.json') (Join-Path $cp 'completed.json')
 return @{layer=$id;item=$item;structure_verified=$structure;students=$M.students.Count;participation_changed=$false;values_assigned=0;checkpoint=$cpRel;automatic_parent_update=$false}
}
function WRename($req){
 $p=CProfile;$fields=WFields $p;$label=([string]$req.label).Trim();if(!$label -or $label.Length -gt 80 -or $label -match '[\r\n]'){throw '작업공간 이름은 1~80자로 입력하세요.'}
  if($req.ContainsKey('item')){if($req.item -notin $fields){throw '이름을 바꿀 열을 선택하세요.'};$old=$M.labels["$Layer|$($req.item)"];$M.labels["$Layer|$($req.item)"]=$label;if($Layer -eq $p.root){@($p.areas|Where-Object root_item -EQ $req.item)[0].label=$label};Save-Meta;return @{layer=$Layer;item=$req.item;before=$old;after=$label;identity_unchanged=$true;books_unchanged=$true}}
 $old=(Layer-Info $Layer).name;$M.layers[$Layer].name=$label
 if(!$M.structureHistory){$M.structureHistory=@()};$M.structureHistory+=@{action='rename';layer=$Layer;before=$old;after=$label;at=(Get-Date -Format o)};Save-Meta
 return @{layer=$Layer;before=$old;after=$label;identity_unchanged=$true;books_unchanged=$true}
}
function WParentPlan{
 $p=CProfile;$fields=WFields $p;$li=Layer-Info $Layer;$parent=$li.parent
 if(!$parent -or ($Item -ne '__sum__' -and $Item -notin $fields)){throw '반영할 하위 항목을 선택하세요.'}
 $binding=if($M.workspaceScopes){$M.workspaceScopes[$Layer]}else{$null};if($binding.parent_item -and ($TargetItem -ne $binding.parent_item -or $Item -ne '__sum__')){throw '이 작업공간은 연결한 상위 열로 학생별 합산점수만 반영합니다.'};$saveLayer=$Layer;$script:Layer=$parent;try{$targetFields=WFields $p}finally{$script:Layer=$saveLayer}
 if($TargetItem -notin $targetFields){throw '현재 영역의 바로 상위 항목을 선택하세요.'}
 $script:Source=$Layer;$script:SourceItem=$Item;$script:Target=$parent
 $allLinks=$M.links;$M.links=@($allLinks|Where-Object {$_.source -eq $Layer -and $_.sourceItem -eq $Item -and $_.target -eq $parent -and $_.targetItem -eq $TargetItem})
 # Existing target connections must never be silently replaced.
 if(@($allLinks|Where-Object {$_.target -eq $parent -and $_.targetItem -eq $TargetItem -and ($_.source -ne $Layer -or $_.sourceItem -ne $Item)}).Count){throw '상위 항목에 다른 자료가 연결돼 있습니다. 연결 관리에서 먼저 확인하세요.'}
 $script:parentFastSnapshot=$true
 try{$plan=Build-Plan}finally{$script:parentFastSnapshot=$false;$M.links=$allLinks}
 $allowed=@{};foreach($sid in $M.students.Keys){$member=@($M.memberships|Where-Object {$_.student -eq $sid -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)});if((!$Student -or $Student -eq $sid) -and (!$Class -or $Class -in @($member.class)) -and (!$p.classes -or @($member|Where-Object {$_.class -in @($p.classes)}).Count)){$allowed[$sid]=$true}}
 $plan.rows=@($plan.rows|Where-Object {$allowed[$_.student]});if(!$plan.rows.Count){throw '반영 범위에 학생이 없습니다.'};$plan.counts=@{};foreach($r in $plan.rows){if($r.disposition -eq '반영'){try{CValidateValue $r.proposed $r.state ($TargetItem -in @($p.numeric_items))}catch{$r.disposition='값·상태 검토 필요';$r.after=$r.before;$r.review=$_.Exception.Message}};$k=$r.disposition;$plan.counts[$k]=1+$plan.counts[$k]}
 $plan.original_fingerprint=$plan.fingerprint;$plan.scope=@{layer=$Layer;parent=$parent;source_item=$Item;target_item=$TargetItem;class=$Class;student=$Student;date=$AsOf};$plan.fingerprint=NHash ($plan.fingerprint+"|$Layer|$parent|$Item|$TargetItem|$Class|$Student|$AsOf")
 return $plan
}
function WParentCommit($request=@{}){
 $plan=WParentPlan;if(!$request.direct -and (!$PreviewFingerprint -or $PreviewFingerprint -ne $plan.fingerprint)){throw '반영 범위나 원본이 바뀌었습니다. 다시 비교하세요.'}
 foreach($row in $plan.rows){if($row.disposition -eq '반영' -and (Same-Value $row.before $row.after)){$row.disposition='동일값 유지'}}
 if(!@($plan.rows|Where-Object disposition -EQ '반영').Count -and !$plan.candidate){$plan.readOnly=$false;$plan.unchanged=$true;return $plan}
 # Apply-Plan independently rechecks its original file fingerprint.
 $plan.fingerprint=$plan.original_fingerprint;$save=$PreviewFingerprint;$script:PreviewFingerprint=$plan.fingerprint
 try{Excel;try{return Apply-Plan $plan}finally{Close-Excel}}finally{$script:PreviewFingerprint=$save}
}
function WStudentView($req){
 if(!$Student -or !$M.students.ContainsKey($Student) -or $req.depth -notin @(0,1,2,3)){throw '학생과 조회 깊이(0~3)를 선택하세요.'}
 $selectedItem=[string]$req['item'];$p=CProfile;[void](WFields $p);$start=$Layer;$savedArea=$Area;$ids=@($start);$front=@($start)
 for($d=0;$d -lt [int]$req.depth;$d++){
  $next=@($M.layers.Keys|Where-Object {$M.layers[$_].parent -in $front})
  if($d -eq 0 -and $selectedItem){$next=@($next|Where-Object {$M.workspaceScopes[$_].parent_item -eq $selectedItem -or $_ -in @($p.areas|Where-Object root_item -EQ $selectedItem|ForEach-Object {$_.work_layer})})}
  $ids+=$next;$front=$next
 }
 $records=[Collections.Generic.List[object]]::new()
 try{foreach($id in $ids){
  $script:Layer=$id;$script:Area=$savedArea
  if($id -ne $p.root){
   if($M.workspaceScopes[$id].area){$script:Area=[string]$M.workspaceScopes[$id].area}
   else{$cursor=$id;$seen=@{};while($cursor -and !$seen[$cursor]){$seen[$cursor]=$true;$match=@($p.areas|Where-Object work_layer -EQ $cursor);if($match.Count){$script:Area=[string]$match[0].root_item;break};$cursor=$M.layers[$cursor].parent}}
  }
  [void](WFields $p);$q=CSnapshot
  foreach($r in $q.data){if($id -eq $start -and $selectedItem -and $r.item -ne $selectedItem){continue};$r.workspace=(Layer-Info $id).name;$r.layer=$id;$records.Add($r)}
 }}finally{$script:Layer=$start;$script:Area=$savedArea}
 return @{records=$records.ToArray();depth=[int]$req.depth;workspace_count=$ids.Count;student=$Student;global_search=$false}
}




function WTagImported($result){
 if(!$Name -or !$Area -or !$M.courses -or !$M.courses[$Name]){return};$p=CProfile;$save=$Layer;$script:Layer=$Parent;try{[void](WFields $p)}finally{$script:Layer=$save}
 if(!$M.workspaceScopes){$M.workspaceScopes=@{}};foreach($id in $result.layers){$M.workspaceScopes[$id]=@{course=$Name;area=$Area;imported_under=$Parent;created_at=(Get-Date -Format o)}};Save-Meta
}


function WColumn($req){
 $p=CProfile;$fields=WFields $p;$labels=@($(if($req.labels){$req.labels}else{$req.label})|ForEach-Object {([string]$_).Trim()}|Where-Object {$_})
 if(!$labels.Count -or $labels.Count -gt 200 -or @($labels|Where-Object {$_.Length -gt 80 -or $_ -match '[\r\n]'}).Count -or @($labels|Sort-Object -Unique).Count -ne $labels.Count){throw '서로 다른 열 이름을 1~200개 입력하세요. 각 이름은 80자 이하입니다.'}
 $columns=@();foreach($label in $labels){$prefix=if($Layer -eq $p.root){'E'+[guid]::NewGuid().ToString('N').Substring(0,10)}else{$Area.Split('.')[0]};$columns+=@{label=$label;item=($prefix+'.i'+[guid]::NewGuid().ToString('N').Substring(0,12))}}
 $cpRel='backups/column-'+[guid]::NewGuid().ToString('N');$cp=Resolve-Local $cpRel;New-Item -ItemType Directory $cp|Out-Null;Copy-Item $MetaPath (Join-Path $cp 'relations.json')
 foreach($lid in $M.layers.Keys){$rel=$M.layers[$lid].path;$dest=Join-Path $cp $rel;New-Item -ItemType Directory (Split-Path $dest) -Force|Out-Null;Copy-Item (Resolve-Local $rel) $dest}
 @{checkpoint=$cpRel;action='space_column'}|ConvertTo-Json|Set-Content (Join-Path $Root 'pending-sync.json') -Encoding utf8
 Excel;try{$b=Open-Book $Layer;try{$t=Read-Table $b;$c=($t.cols.Values|Measure-Object -Maximum).Maximum+1;foreach($column in $columns){$t.sheet.Cells.Item(1,$c).Value2=[string]$column.item;$c++};$b.Save()}finally{$b.Close($false)}}finally{Close-Excel}
 foreach($column in $columns){$item=$column.item;$M.layers[$Layer].items+=,$item;$M.labels["$Layer|$item"]=$column.label;if($req.numeric){$p.numeric_items+=,$item}
  if($Layer -eq $p.root){$p.areas+=,@{label=$column.label;root_item=$item;work_layer='';work_items=@()}}else{$a=WArea $p;if($Layer -eq $a.work_layer){$a.work_items+=,$item}}
 }
 Save-Meta;Move-Item (Join-Path $Root 'pending-sync.json') (Join-Path $cp 'completed.json')
 $children=@();if($req.child){foreach($column in $columns){$oldArea=$Area;if($Layer -eq $p.root){$script:Area=$column.item};try{$children+=,(WCreate @{label=($column.label+' 세부 작업');numeric=$true;target_item=$column.item})}finally{$script:Area=$oldArea}}}
 return @{layer=$Layer;item=$columns[0].item;items=@($columns|ForEach-Object {$_.item});added=$columns.Count;child=$(if($children.Count -eq 1){$children[0]}else{$null});children=$children;checkpoint=$cpRel;values_assigned=0;automatic_parent_update=$false}
}
function WVerifyRoster($id,$items){
 $t=Snapshot $id;$expected=@($M.students.Keys|Sort-Object);$actual=@($t.rows.Keys|Sort-Object)
 if($t.warnings.Count -or $actual.Count -ne $expected.Count -or ($actual -join '|') -cne ($expected -join '|')){throw '전체 학생 ID 전수 대조에 실패했습니다. 새 공간을 사용하지 말고 복구 사본을 확인하세요.'}
 foreach($field in @('student_id','virtual_name','code')+@($items)){if(!$t.cols.ContainsKey($field)){throw '공통 필드가 누락되었습니다. 새 공간을 사용하지 마세요.'}}
 foreach($sid in $actual){foreach($field in $items){if($null -ne $t.values[$t.rows[$sid],$t.cols[$field]]){throw '구조 생성 중 평가값이 들어갔습니다. 새 공간을 사용하지 마세요.'}}}
 return @{student_ids=$actual;exact_id_set=$true;no_duplicates=$true;common_fields=@($items);blank_values=$true}
}

function WColumnType($req){
 $p=CProfile;$fields=WFields $p;$item=[string]$req.item;if($item -notin $fields){throw '종류를 바꿀 점수·기록 열을 선택하세요.'};if($req.kind -notin @('numeric','text')){throw '숫자점수 또는 문자기록을 선택하세요.'}
 $numeric=($req.kind -eq 'numeric');if(($item -in @($p.numeric_items)) -eq $numeric){return @{changed=$false;item=$item;kind=$req.kind}}
 Excel;try{$b=Open-Book $Layer;try{$t=Read-Table $b;$col=$t.cols[$item];$changes=@();$bad=@()
  foreach($sid in $t.rows.Keys){$r=$t.rows[$sid];$v=$t.values[$r,$col];if($null -eq $v -or [string]$v -eq ''){continue}
   if($numeric){$n=0.0;if(![double]::TryParse([string]$v,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n) -or [double]::IsNaN($n) -or [double]::IsInfinity($n)){$bad+=($M.students[$sid].name+' ('+$sid+') · '+[string]$v);continue};$value=$n}else{$value=[string]$v}
   if(!([string]$t.formulas[$r,$col]).StartsWith('=')){$changes+=@{row=$r;value=$value}}
  }
  if($bad.Count){throw ('숫자로 바꿀 수 없는 값 '+$bad.Count+'건이 있습니다. 기존 값과 열 종류는 유지했습니다. '+(($bad|Select-Object -First 5) -join ' / '))}
  $cpRel='backups/type-'+[guid]::NewGuid().ToString('N');$cp=Resolve-Local $cpRel;[void][IO.Directory]::CreateDirectory($cp);Copy-Item -LiteralPath $MetaPath -Destination (Join-Path $cp 'relations.json');$rel=$M.layers[$Layer].path;$dest=Join-Path $cp $rel;[void][IO.Directory]::CreateDirectory((Split-Path $dest));$b.SaveCopyAs($dest)
  @{checkpoint=$cpRel;action='column_type';layer=$Layer;item=$item}|ConvertTo-Json|Set-Content (Join-Path $Root 'pending-sync.json') -Encoding UTF8
  foreach($change in $changes){if($numeric){$t.sheet.Cells.Item($change.row,$col).Value2=[double]$change.value}else{$t.sheet.Cells.Item($change.row,$col).Value2="'"+[string]$change.value}}
  $p.numeric_items=@($p.numeric_items|Where-Object {$_ -ne $item});if($numeric){$p.numeric_items+=,$item}
  $b.Save();Save-Meta;Move-Item -LiteralPath (Join-Path $Root 'pending-sync.json') -Destination (Join-Path $cp 'completed.json')
  return @{changed=$true;item=$item;kind=$req.kind;converted=$changes.Count;automatic_parent_update=$false;checkpoint=$cpRel}
 }finally{$b.Close($false)}}finally{Close-Excel}
}
