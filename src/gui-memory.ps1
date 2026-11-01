# Session-only project cache. No snapshots or row values are written to disk.
$script:memoryRoot='';$script:memoryMetaHash='';$script:memoryTables=@{};$script:memoryViews=@{};$script:memoryMemberships=@{}
$script:memoryStats=@{loads=0;book_reads=0;view_builds=0;view_hits=0;load_ms=0;last_view_ms=0}
function MemoryClear{$script:memoryRoot='';$script:memoryMetaHash='';$script:memoryTables=@{};$script:memoryViews=@{};$script:memoryMemberships=@{}}
function MemoryHash($path){$s=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite);$h=[Security.Cryptography.SHA256]::Create();try{return ([BitConverter]::ToString($h.ComputeHash($s))).Replace('-','')}finally{$s.Dispose();$h.Dispose()}}
function MemoryTextHash($text){$h=[Security.Cryptography.SHA256]::Create();try{$bytes=$h.ComputeHash([Text.Encoding]::UTF8.GetBytes($text));return ([BitConverter]::ToString($bytes)).Replace('-','').ToLowerInvariant()}finally{$h.Dispose()}}
function MemorySame($a,$b){if($null -eq $a -or $null -eq $b){return ($null -eq $a -and $null -eq $b)};if($a -is [ValueType] -and $b -is [ValueType]){return ([decimal]$a -eq [decimal]$b)};return ([string]$a -ceq [string]$b)}
function MemoryPath($rel){$p=[IO.Path]::GetFullPath((Join-Path $ws $rel));if(!$p.StartsWith($ws.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw '프로젝트 경계 밖 자료는 메모리에 읽지 않습니다.'};$walk=$p;while($walk -and $walk.Length -gt $ws.Length){if((Get-Item -LiteralPath $walk -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw '연결 폴더의 자료는 메모리에 읽지 않습니다.'};$walk=Split-Path $walk};return $p}
function MemoryReader{
 if(!('NaturalXlsx' -as [type])){Add-Type -AssemblyName System.IO.Compression.FileSystem;Add-Type -Path (Join-Path $base 'XlsxReader.cs') -ReferencedAssemblies System.Xml.Linq,System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem}
}
function MemoryReadBook($id,$path,$hash){
 MemoryReader;$book=[NaturalXlsx]::Read($path);$identity=@($book.Sheets|Where-Object Name -EQ '_identity');$sheets=@($book.Sheets|Where-Object Name -EQ 'Data')
 if($identity.Count -ne 1 -or [string]$identity[0].Rows[1][2].Value -cne $id){throw ('작업표 ID 불일치: '+$id)}
 if($sheets.Count -ne 1){throw ('Data 시트를 확인하세요: '+$id)};$sheet=$sheets[0];$cols=@{};$rows=@{};$bad=@{};$warnings=[Collections.Generic.List[string]]::new()
 if(!$sheet.Rows.ContainsKey(1)){throw ('항목 제목행 누락: '+$id)}
 foreach($c in $sheet.Rows[1].Keys){$name=[string]$sheet.Rows[1][$c].Value;if(!$name){continue};if($cols.ContainsKey($name)){throw ('항목 ID 중복: '+$id+' / '+$name)};$cols[$name]=[int]$c}
 if(!$cols.ContainsKey('student_id')){throw ('학생 ID 열 누락: '+$id)}
 foreach($r in $sheet.Rows.Keys){if($r -eq 1){continue};$sid=[string]$sheet.Rows[$r][$cols.student_id].Value;if(!$sid -or !$meta.students.ContainsKey($sid)){$warnings.Add('Unknown/missing student ID at row '+$r+' : '+$sid);continue};if($rows.ContainsKey($sid)){$bad[$sid]=$true;$warnings.Add('Duplicate student ID: '+$sid);continue};$rows[$sid]=[int]$r}
 foreach($sid in $bad.Keys){$rows.Remove($sid)}
 foreach($field in $meta.layers[$id].items){if(!$cols.ContainsKey($field)){throw ('활성 항목이 작업표에 없습니다: '+$id+' / '+$field)}}
 if((MemoryHash $path) -cne $hash){throw ('읽는 동안 작업표가 바뀌었습니다. 다시 조회하세요: '+$id)}
 $script:memoryStats.book_reads++
 return @{path=$path;hash=$hash;sheet=$sheet;cols=$cols;rows=$rows;warnings=$warnings.ToArray()}
}
function MemoryEnsure{
 if(!$ws){throw '프로젝트를 먼저 여세요.'}
 $trashMarker=Join-Path $ws 'pending-trash.json';if(Test-Path -LiteralPath $trashMarker){$pending=Get-Content -LiteralPath $trashMarker -Raw -Encoding UTF8|ConvertFrom-Json;if($pending.stage -ne 'completed'){if(Get-Command Run -ErrorAction SilentlyContinue){[void](Run 'trashlist')}else{throw '미완료 휴지통 처리를 먼저 검증해야 합니다.'}}}
 foreach($marker in @('pending-sync.json','pending-import.json')){if(Test-Path -LiteralPath (Join-Path $ws $marker)){throw ('미완료 저장이 있어 메모리 읽기를 중단했습니다. 복구 사본을 확인하세요: '+$marker)}}
 $metaPath=Join-Path $ws 'relations.json';$version=MemoryHash $metaPath
 if($memoryRoot -eq $ws -and $memoryMetaHash -ceq $version){return}
 $clock=[Diagnostics.Stopwatch]::StartNew();$oldTables=if($memoryRoot -eq $ws){$memoryTables}else{@{}}
 Read-Meta;$tables=@{};$ids=@($meta.layers.Keys|Sort-Object);$n=0
 try{
  foreach($id in $ids){
   $seen=@{};$ancestor=$id;while($ancestor){if($seen[$ancestor] -or !$meta.layers.ContainsKey($ancestor)){throw ('작업공간 관계가 올바르지 않습니다: '+$id)};$seen[$ancestor]=$true;$ancestor=$meta.layers[$ancestor].parent}
   $path=MemoryPath $meta.layers[$id].path;$hash=MemoryHash $path
   if($oldTables[$id] -and $oldTables[$id].path -ceq $path -and $oldTables[$id].hash -ceq $hash){$tables[$id]=$oldTables[$id]}else{$tables[$id]=MemoryReadBook $id $path $hash}
   $n++;$progress='프로젝트 검증·메모리 적재 '+$n+' / '+$ids.Count;if($dashMessage){$dashMessage.Text=$progress;$dashMessage.Refresh()};if($nCurrent){$nCurrent.Text=$progress;$nCurrent.Refresh()}
  }
  foreach($link in $meta.links){if(!$tables[$link.source] -or !$tables[$link.target] -or ($link.sourceItem -ne '__sum__' -and !$tables[$link.source].cols.ContainsKey($link.sourceItem)) -or !$tables[$link.target].cols.ContainsKey($link.targetItem)){throw ('작업표 연결 항목을 확인하세요: '+$link.source+' / '+$link.sourceItem+' → '+$link.target+' / '+$link.targetItem)}}
  if((MemoryHash $metaPath) -cne $version){throw '적재 중 프로젝트 관계가 바뀌었습니다. 다시 여세요.'}
  $memberships=@{};foreach($member in $meta.memberships){$memberships[$member.student]=@($memberships[$member.student]|Where-Object {$_})+@($member)}
  $script:memoryRoot=$ws;$script:memoryMetaHash=$version;$script:memoryTables=$tables;$script:memoryViews=@{};$script:memoryMemberships=$memberships;$script:memoryStats.loads++;$script:memoryStats.load_ms=$clock.Elapsed.TotalMilliseconds
 }catch{MemoryClear;throw}finally{$clock.Stop()}
}
function MemoryFields($p,$layer,$area){
 if($layer -eq $p.root){return ,@($p.areas|ForEach-Object {$_.root_item})}
 $a=@($p.areas|Where-Object root_item -EQ $area);if(!$a.Count -and $meta.detachedAreas[$p.id+'|'+$area]){$a=@($meta.detachedAreas[$p.id+'|'+$area])};if($a.Count -ne 1){throw '평가영역을 선택하세요.'};$a=$a[0]
 if($layer -eq $a.work_layer){return ,@($a.work_items)}
 $cursor=$layer;$detached=$false;$within=$false;while($cursor){if($meta.workspaceScopes[$cursor].unlinked){$detached=$true};if($cursor -eq $a.work_layer){$within=$true};$cursor=$meta.layers[$cursor].parent}
 if(!$detached -and (!$a.work_layer -or !$within)){throw '선택한 교육과정·영역 밖 작업공간입니다.'}
 $scope=$meta.workspaceScopes[$layer];if($scope -and ($scope.course -ne $p.id -or $scope.area -ne $area)){throw '다른 영역에 속한 작업공간입니다.'}
 $fields=@($meta.layers[$layer].items|Where-Object {$_ -notmatch '\.(status|note)$'});if(!$scope){$prefix=$a.root_item.Split('.')[0];$fields=@($fields|Where-Object {$_.Split('.')[0] -eq $prefix})};return ,$fields
}
function MemorySnapshot($args2){
 $clock=[Diagnostics.Stopwatch]::StartNew();MemoryEnsure
 $layer=[string]$args2.Layer;$p=$meta.courses[$args2.Name];if(!$p -or !$memoryTables[$layer]){throw '프로젝트 또는 작업표를 다시 선택하세요.'}
 $t=$memoryTables[$layer];$hash=MemoryHash $t.path
 if($hash -cne $t.hash){$script:memoryTables[$layer]=MemoryReadBook $layer $t.path $hash;$t=$memoryTables[$layer];$script:memoryViews=@{}}
 $key=(@($args2.Name,$layer,$args2.Area,$args2.Class,$args2.AsOf,$args2.Student,$memoryMetaHash,$hash) -join '|')
 if($memoryViews.ContainsKey($key)){$script:memoryStats.view_hits++;$script:memoryStats.last_view_ms=$clock.Elapsed.TotalMilliseconds;return $memoryViews[$key]}
 $fields=MemoryFields $p $layer $args2.Area;$data=[Collections.Generic.List[object]]::new()
 foreach($sid in ($t.rows.Keys|Sort-Object)){
  if($args2.Student -and $sid -ne $args2.Student){continue};$s=$meta.students[$sid];$members=@($memoryMemberships[$sid]|Where-Object {$_.from -le $args2.AsOf -and (!$_.until -or $_.until -ge $args2.AsOf)})
  if($args2.Class -and $args2.Class -notin @($members.class)){continue};if($p.classes -and !@($members|Where-Object {$_.class -in @($p.classes)}).Count){continue}
  foreach($item in $fields){
   $row=$t.sheet.Rows[$t.rows[$sid]];$cell=$row[[int]$t.cols[$item]];$v=$cell.Value;$canonical=$layer+'|'+$sid+'|'+$item;$manual=$meta.manualOverrides[$canonical];$import=$meta.importRecords[$canonical];$origin=$meta.provenance[$canonical]
   $state=if($null -eq $v){'미입력'}else{'입력됨'}
   $label=if($layer -eq $p.root){@($p.areas|Where-Object root_item -EQ $item)[0].label}else{$meta.labels[$layer+'|'+$item]};if(!$label){$label=$item}
   $same=MemorySame $origin.value $v
   $data.Add([pscustomobject]@{student=$sid;name=$s.name;number=$s.number;class=($members.class -join ',');item=$item;label=$label;value=$v;state=$state;note=$(if($manual){$manual.note}else{''});formula=[bool]$cell.Formula;numeric=($item -in @($p.numeric_items));origin=$(if($manual){'교사 직접입력'}elseif($import){'자료 반입'}elseif($origin -and $same){'하위 작업표 참조'}elseif($origin){'직접수정·검토 필요'}else{'직접입력'});source=$origin.source;sourceItem=$origin.item;sourceFile=$meta.layers[$layer].path;recordId=$(if($manual){$manual.record_id}elseif($import){$import.record_id}else{''});sourceCell=$import.source_cell})
  }
 }
 if((MemoryHash (Join-Path $ws 'relations.json')) -cne $memoryMetaHash -or (MemoryHash $t.path) -cne $hash){throw '조회 중 원본이 변경됐습니다. 다시 조회하세요.'}
 $q=[pscustomobject]@{data=$data.ToArray();fields=@($fields);fingerprint=(MemoryTextHash ($memoryMetaHash+'|'+$hash));warnings=@($t.warnings);profile=$p;layer=$layer;area=$args2.Area;class=$args2.Class;as_of=$args2.AsOf}
 $script:memoryViews[$key]=$q;$script:memoryStats.view_builds++;$script:memoryStats.last_view_ms=$clock.Elapsed.TotalMilliseconds;return $q
}
if($f){$f.Text=$productTitle;$f.Add_FormClosed({MemoryClear})}
