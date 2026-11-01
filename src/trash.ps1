# Reversible archival: workbook bytes never change during trash operations.
function TJson($path,$value){$tmp=$path+'.'+[guid]::NewGuid().ToString('N')+'.tmp';[IO.File]::WriteAllText($tmp,($value|ConvertTo-Json -Depth 45),[Text.UTF8Encoding]::new($true));if(Test-Path $path){[IO.File]::Replace($tmp,$path,($tmp+'.previous'))}else{[IO.File]::Move($tmp,$path)}}
function TBooks{ $h=@{};foreach($id in $M.layers.Keys){$h[$M.layers[$id].path]=Shared-Hash (Resolve-Local $M.layers[$id].path)};return $h }
function TCheckBooks($hashes,$prefix=''){foreach($rel in $hashes.Keys){$path=if($prefix){Join-Path $prefix $rel}else{Resolve-Local $rel};if(!(Test-Path $path) -or (Shared-Hash $path) -ne $hashes[$rel]){throw ('보관 대조 실패: '+$rel+'. 추가 저장을 중단하고 체크포인트를 확인하세요.')}}}
function TArchiveCheck($j){TCheckBooks $j.books (Resolve-Local $j.archive);TCheckBooks $j.books (Resolve-Local $j.checkpoint);foreach($dir in @($j.archive,$j.checkpoint)){if((Shared-Hash (Join-Path (Resolve-Local $dir) 'relations.json')) -ne $j.before_hash){throw '관계 보관 사본 검증 실패'}};if((Shared-Hash (Join-Path (Resolve-Local $j.archive) 'records.json')) -ne $j.record_hash){throw '휴지통 기록 검증 실패'}}
function TState{ $h=TBooks;$sha=[Security.Cryptography.SHA256]::Create();try{return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(((Shared-Hash $MetaPath)+'|'+(($h.Keys|Sort-Object|ForEach-Object {$_+'='+$h[$_]}) -join '|')))))).Replace('-','')}finally{$sha.Dispose()} }
function TDetached($id){$seen=@{};while($id){if($seen[$id]){throw '작업공간 순환'};$seen[$id]=$true;if($M.workspaceScopes[$id].unlinked){return $true};$id=$M.layers[$id].parent};return $false}
function TArea($p,$area){$a=@($p.areas|Where-Object root_item -EQ $area);if($a.Count){return $a[0]};if($M.detachedAreas){return $M.detachedAreas[$Name+'|'+$area]};return $null}
function TPath($id){$names=@();$seen=@{};while($id){if($seen[$id]){throw '작업공간 순환'};$seen[$id]=$true;$names=@($M.layers[$id].name)+$names;if($M.workspaceScopes[$id].unlinked){$names=@('미연결 작업공간')+$names};$id=$M.layers[$id].parent};return ($names -join ' > ')}
function TSpaceContext($id){
 $scope=$M.workspaceScopes[$id];$course=if($scope.course){$scope.course}else{$Name};$area=if($scope.area){$scope.area}else{$Area}
 if(!$scope.area -and $course -and $M.courses[$course]){
  $prefixes=@($M.layers[$id].items|Where-Object {$_ -notmatch '\.(status|note)$'}|ForEach-Object {$_.Split('.')[0]}|Sort-Object -Unique)
  $matches=@($M.courses[$course].areas|Where-Object {$_.root_item.Split('.')[0] -in $prefixes})
  if($matches.Count -eq 1){$area=$matches[0].root_item}
 }
 return @{course=$course;area=$area}
}
function TRecover{
 $pending=Join-Path $Root 'pending-trash.json';if(!(Test-Path $pending)){return};$j=To-Map (Get-Content $pending -Raw -Encoding UTF8|ConvertFrom-Json)
 if($j.stage -eq 'completed'){return}
 $actual=Shared-Hash $MetaPath
 if($j.after_hash -and $actual -eq $j.after_hash){TCheckBooks $j.books;TArchiveCheck $j;$j.stage='completed';$j.verified_at=Get-Date -Format o;TJson (Resolve-Local $j.journal) $j;TJson $pending $j;return}
 if($actual -ne $j.before_hash){throw '휴지통 처리 중 자료 상태가 달라졌습니다. 자동 복구하지 않습니다. 추가 저장을 중단하고 별도 체크포인트를 확인하세요.'}
 if($Action -notin @('trashpreview','trashcommit','trashlist','trashread','tablesnapshot','parentpreview','status','query','open')){throw ('미완료 휴지통 처리 '+$j.id+'가 있습니다. 휴지통에서 해당 요청을 다시 확인하세요. 활성 자료는 변경 전 상태입니다.')}
}
function TPlan($req){
 $metaVersion=Shared-Hash $MetaPath;Load-Meta;if((Shared-Hash $MetaPath) -ne $metaVersion){throw '대상 확인 중 관계 파일이 변경됐습니다. 다시 확인하세요.'}
 if((Test-Path (Join-Path $Root 'pending-sync.json')) -or (Test-Path (Join-Path $Root 'pending-import.json'))){throw '다른 저장 중단 표시가 있습니다. 휴지통 이동 전에 기존 체크포인트를 확인하세요.'}
 $li=Layer-Info $Layer;$kind=[string]$req.kind;if($kind -notin @('column','space')){throw '휴지통 대상을 선택하세요.'};$fields=@($li.items|Where-Object {$_ -notmatch '\.(status|note)$'});$item=[string]$req.item;$children=@();$removed=@();$areaInfo=$null
 if($kind -eq 'column'){
  if($item -notin $fields){throw '이미 제외되었거나 활성 항목이 아닙니다. 다시 조회하세요.'}
  foreach($p in $M.courses.Values){$a=@($p.areas|Where-Object root_item -EQ $item);if($Layer -eq $p.root -and $a.Count){$areaInfo=$a[0];if($a[0].work_layer){$children+=,$a[0].work_layer}}}
  $children+=@($M.workspaceScopes.Keys|Where-Object {$M.layers[$_].parent -eq $Layer -and $M.workspaceScopes[$_].parent_item -eq $item})
  $children+=@($M.links|Where-Object {$_.target -eq $Layer -and $_.targetItem -eq $item}|ForEach-Object {$_.source});$children=@($children|Sort-Object -Unique)
  foreach($child in $children){if($M.layers[$child].parent -ne $Layer -or @($M.links|Where-Object {$_.source -eq $child -and $_.target -eq $Layer -and $_.targetItem -ne $item}).Count){throw ('공유 연결 영향이 불명확하여 이동을 차단합니다: '+$li.name+' / '+(Label-For $Layer $item)+'. 다른 연결 위치를 먼저 확인하세요.')}}
  $removed=@($M.links|Where-Object {($_.target -eq $Layer -and $_.targetItem -eq $item) -or ($_.source -eq $Layer -and $_.sourceItem -eq $item)})
 }else{
  if(!$li.parent -and !($M.workspaceScopes[$Layer].unlinked)){throw '수행평가 종합표는 휴지통으로 이동할 수 없습니다.'}
  $refs=@($M.links|Where-Object {$_.source -eq $Layer -or $_.target -eq $Layer});$desc=@($M.layers.Keys|Where-Object {$M.layers[$_].parent -eq $Layer})
  $areaRefs=@($M.courses.Values|ForEach-Object {$_.areas}|Where-Object work_layer -EQ $Layer)
  $children=$desc;$removed=$refs
 }
 $childContexts=@{};foreach($child in $children){$childContexts[$child]=TSpaceContext $child}
 $before=TState;if((Shared-Hash $MetaPath) -ne $metaVersion){throw '대상 확인 중 관계 파일이 변경됐습니다. 다시 확인하세요.'};Excel;try{$t=Snapshot $Layer}finally{Close-Excel};if($t.warnings.Count){throw '학생 식별자 경고가 있어 이동을 중단했습니다. 원본 학생 명렬표를 대조하세요.'}
 $records=@();foreach($sid in @($t.rows.Keys|Sort-Object)){foreach($f in $(if($kind -eq 'column'){@($item)}else{$fields})){$key="$Layer|$sid|$f";$records+=@{student=$sid;item=$f;value=$t.values[$t.rows[$sid],$t.cols[$f]];formula=$t.formulas[$t.rows[$sid],$t.cols[$f]];manual=$M.manualOverrides[$key];import=$M.importRecords[$key];provenance=$M.provenance[$key];state=(Source-State $t $sid $f $t.values[$t.rows[$sid],$t.cols[$f]])}}}
 if((TState) -ne $before){throw '확인 중 자료가 변경되었습니다. 다시 확인하세요.'}
 return @{id=[guid]::NewGuid().ToString('N');kind=$kind;layer=$Layer;item=$item;label=$(if($kind -eq 'column'){Label-For $Layer $item}else{$li.name});location=(TPath $Layer);student_records=$t.rows.Count;records=$records;children=$children;removed_links=$removed;scope_course=$Name;scope_area=$Area;child_contexts=$childContexts;area_info=$areaInfo;fingerprint=$before;warning=$(if($kind -eq 'column'){'기존 기록은 보관 사본에 유지됩니다. 하위 작업공간은 남지만 연결이 해제되며, 자동 복구되지 않습니다. 다른 점수는 자동 재합산하지 않습니다.'}else{'선택 공간의 상위·하위 연결을 해제한 뒤 이 공간만 휴지통으로 이동합니다. 하위 공간과 전체 학생 기록·내부 연결은 미연결 작업공간으로 보존하며 독립 열람·Excel 다운로드가 가능합니다. 상위의 기존 점수·항목은 그대로 유지합니다. 자동 복원·재연결·영구 삭제는 하지 않습니다.'})}
}
function TFail($req,$stage){if((Test-Path (Join-Path $Root '.synthetic-test')) -and $req.test_failure -eq $stage){throw ('가상 실패 시험: '+$stage)};if((Test-Path (Join-Path $Root '.synthetic-test')) -and $req.test_exit -eq $stage){[Environment]::Exit(77)}}
function TCommit($req){
 if(!$req.confirmed -or $req.id -notmatch '^[a-f0-9]{32}$'){throw '대상·연결 영향 경고를 확인한 요청이 필요합니다.'}
 if($M.trash -and $M.trash[$req.id]){$e=$M.trash[$req.id];TArchiveCheck $e;return @{completed=$true;duplicate=$true;entry=$e}}
 $pending=Join-Path $Root 'pending-trash.json';$j=$null;if(Test-Path $pending){$j=To-Map (Get-Content $pending -Raw -Encoding UTF8|ConvertFrom-Json);if($j.stage -ne 'completed' -and $j.id -ne $req.id){throw '다른 미완료 휴지통 요청을 먼저 확인하세요.'}}
 if(!$j -or $j.id -ne $req.id -or $j.stage -eq 'completed'){
  $plan=TPlan $req;if($plan.fingerprint -ne $req.fingerprint){throw '미리보기 이후 자료가 변경되었습니다. 다시 확인하세요.'}
  $j=@{id=$req.id;stage='started';before_hash=(Shared-Hash $MetaPath);books=(TBooks);archive=('trash/'+$req.id);checkpoint=('backups/trash-'+$req.id);journal=('transactions/trash-'+$req.id+'.json');plan=$plan;at=(Get-Date -Format o)}
  foreach($rel in @($j.archive,$j.checkpoint,'transactions')){New-Item -ItemType Directory (Resolve-Local $rel) -Force|Out-Null};TJson (Resolve-Local $j.journal) $j;TJson $pending $j
 }
 TCheckBooks $j.books;if((Shared-Hash $MetaPath) -ne $j.before_hash){throw '활성 자료가 변경되어 휴지통 재시도를 중단했습니다.'}
 if($j.plan.layer -ne $Layer -or $j.plan.kind -ne $req.kind -or $j.plan.item -ne [string]$req.item -or $j.plan.fingerprint -ne $req.fingerprint){throw '원래 확인한 동일 대상·요청으로만 다시 시도하세요.'}
 $archive=Resolve-Local $j.archive;$cp=Resolve-Local $j.checkpoint
 foreach($dir in @($archive,$cp)){if((Test-Path -LiteralPath $dir) -and !(Test-Path -LiteralPath $dir -PathType Container)){throw ('보관 폴더 위치에 파일이 있습니다: '+$dir+'. 활성 자료는 유지했습니다. 보관 경로를 확인한 뒤 동일 요청으로 다시 시도하세요.')};New-Item -ItemType Directory -Path $dir -Force|Out-Null}
 TFail $req 'archive'
 foreach($dir in @($archive,$cp)){
  Copy-Item -LiteralPath $MetaPath -Destination (Join-Path $dir 'relations.json');foreach($rel in $j.books.Keys){$dest=Join-Path $dir $rel;New-Item -ItemType Directory (Split-Path $dest) -Force|Out-Null;Copy-Item -LiteralPath (Resolve-Local $rel) -Destination $dest};TCheckBooks $j.books $dir;if((Shared-Hash (Join-Path $dir 'relations.json')) -ne $j.before_hash){throw '관계 보관 사본 대조 실패'}
 }
 TJson (Join-Path $archive 'records.json') $j.plan.records
 $j.record_hash=Shared-Hash (Join-Path $archive 'records.json');TArchiveCheck $j
 $plan=$j.plan;if(!$M.trash){$M.trash=@{}};if(!$M.retiredItems){$M.retiredItems=@{}};if(!$M.retiredSpaces){$M.retiredSpaces=@{}};if(!$M.detachedAreas){$M.detachedAreas=@{}}
 $entry=@{id=$req.id;kind=$plan.kind;label=$plan.label;location=$plan.location;layer=$plan.layer;item=$plan.item;student_records=$plan.student_records;at=$j.at;archive=$j.archive;checkpoint=$j.checkpoint;previous_links=$plan.removed_links;children=$plan.children;books=$j.books;before_hash=$j.before_hash;record_hash=$j.record_hash;verified=$true}
 if($plan.kind -eq 'column'){
  $M.retiredItems[$Layer+'|'+$plan.item]=$req.id;$M.layers[$Layer].items=@($M.layers[$Layer].items|Where-Object {$_ -ne $plan.item})
  foreach($p in $M.courses.Values){$a=@($p.areas|Where-Object root_item -EQ $plan.item);if($Layer -eq $p.root -and $a.Count){$M.detachedAreas[$p.id+'|'+$plan.item]=$a[0];$p.areas=@($p.areas|Where-Object root_item -NE $plan.item)};foreach($area in $p.areas){if($area.work_layer -eq $Layer){$area.work_items=@($area.work_items|Where-Object {$_ -ne $plan.item})}}}
  $M.links=@($M.links|Where-Object {!(($_.target -eq $Layer -and $_.targetItem -eq $plan.item) -or ($_.source -eq $Layer -and $_.sourceItem -eq $plan.item))})
  foreach($child in $plan.children){$M.layers[$child].parent='';if(!$M.workspaceScopes[$child]){$M.workspaceScopes[$child]=@{course=$Name;area=$plan.item}};$M.workspaceScopes[$child].unlinked=$true;$M.workspaceScopes[$child].former_parent=$Layer;$M.workspaceScopes[$child].former_parent_item=$plan.item;$M.workspaceScopes[$child].Remove('parent_item')}
 }else{
  if(!$M.workspaceScopes){$M.workspaceScopes=@{}}
  foreach($child in $plan.children){
   $previousParentItem=$M.workspaceScopes[$child].parent_item
   $M.layers[$child].parent=''
   if(!$M.workspaceScopes[$child]){$context=$plan.child_contexts[$child];$M.workspaceScopes[$child]=@{course=$context.course;area=$context.area}}
   $M.workspaceScopes[$child].unlinked=$true;$M.workspaceScopes[$child].former_parent=$Layer
   if($previousParentItem){$M.workspaceScopes[$child].former_parent_item=$previousParentItem}
   $M.workspaceScopes[$child].Remove('parent_item')
  }
  $M.links=@($M.links|Where-Object {$_.source -ne $Layer -and $_.target -ne $Layer})
  foreach($profile in $M.courses.Values){foreach($area in $profile.areas){if($area.work_layer -eq $Layer){$area.work_layer='';$area.work_items=@()}}}
  foreach($area in $M.detachedAreas.Values){if($area.work_layer -eq $Layer){$area.work_layer='';$area.work_items=@()}}
  $M.retiredSpaces[$Layer]=$req.id;$M.layers.Remove($Layer);$M.workspaceScopes.Remove($Layer)
 }
 $M.trash[$req.id]=$entry;$after=Join-Path $archive 'after-relations.json';TJson $after $M;$j.after_hash=Shared-Hash $after;$j.stage='prepared';TJson (Resolve-Local $j.journal) $j;TJson $pending $j;TFail $req 'prepared'
 TCheckBooks $j.books;if((Shared-Hash $MetaPath) -ne $j.before_hash){throw '변경 직전 관계 파일이 바뀌었습니다.'}
 $tmp=Join-Path $Root ('.trash-'+$req.id+'.tmp');Copy-Item -LiteralPath $after -Destination $tmp;[IO.File]::Replace($tmp,$MetaPath,(Join-Path $cp 'atomic-before.json'));TFail $req 'swapped'
 if((Shared-Hash $MetaPath) -ne $j.after_hash){throw '관계 변경 사후 대조 실패'};TCheckBooks $j.books;TCheckBooks $j.books $archive;TCheckBooks $j.books $cp
 TArchiveCheck $j;$j.stage='completed';$j.verified_at=Get-Date -Format o;TJson (Resolve-Local $j.journal) $j;TJson $pending $j
 return @{completed=$true;duplicate=$false;entry=$entry;books_unchanged=$true;automatic_score_changes=0}
}
function TList{ $pending=$null;$p=Join-Path $Root 'pending-trash.json';if(Test-Path $p){$pending=To-Map (Get-Content $p -Raw -Encoding UTF8|ConvertFrom-Json)};return @{entries=@($M.trash.Values);pending=$pending;read_only=$true;permanent_delete=$false;automatic_restore=$false}}
function TRead($req){if(!$M.trash[$req.id]){throw '휴지통 항목이 없습니다.'};$e=$M.trash[$req.id];TArchiveCheck $e;return @{entry=$e;records=(To-Map (Get-Content (Join-Path (Resolve-Local $e.archive) 'records.json') -Raw -Encoding UTF8|ConvertFrom-Json));read_only=$true}}
