# Track B: deterministic local interpretation; no model or network calls.
Add-Type -AssemblyName System.IO.Compression.FileSystem
if(!('NaturalXlsx' -as [type])){if($PSVersionTable.PSVersion.Major -ge 6){Add-Type -Path (Join-Path $PSScriptRoot 'XlsxReader.cs')}else{Add-Type -Path (Join-Path $PSScriptRoot 'XlsxReader.cs') -ReferencedAssemblies System.Xml.Linq,System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem}}
function NHash([string]$text){$h=[Security.Cryptography.SHA256]::Create();try{return ([BitConverter]::ToString($h.ComputeHash([Text.Encoding]::UTF8.GetBytes($text)))).Replace('-','').ToLower()}finally{$h.Dispose()}}
function NNorm($v){return ([string]$v).Trim().Replace(' ','').Replace("`n",'').ToLower()}
function NRole($v){switch -Regex (NNorm $v){'^(학생id|student_id|학생코드|학생식별코드|학생식별자|식별코드|고유번호)$'{'id';break};'^(이름|성명|학생명|virtual_name)$'{'name';break};'^(학급|반)$'{'class';break};'^(번호|출석번호)$'{'number';break};'^(상태|제출상태|처리상태)$'{'status';break};'^(제출차수|회차|재제출)$'{'attempt';break};'^(평가항목|항목명)$'{'item';break};'^(관찰값|입력값|값)$'{'value';break};default{'field'}}}
function NCell($row,$column){if($column -and $row.ContainsKey([int]$column)){return $row[[int]$column].Value};return $null}
function NStudent($row,$roles,$sheetName){
 $id=[string](NCell $row $roles.id);$name=[string](NCell $row $roles.name);$cls=[string](NCell $row $roles.class);$num=[string](NCell $row $roles.number)
 if(!$cls -and !$id -and $sheetName -match '^(\d+)반'){$cls=$Matches[1]+'반'};if($cls -match '^\d+$'){$cls+='반'}
 if($id){$mapped=$naturalCodeMap[$id];if(!$mapped){return @{error=('이 프로젝트에 없는 학생코드: '+$id)}};$ids=@($mapped)}else{
  if(!$name -or !$cls -or !$num){return @{error='학생ID 또는 이름·반·번호의 정확한 조합 필요'}}
  $ids=@($M.students.Keys|Where-Object {$sid=$_;$s=$M.students[$sid];$s.name -ceq $name -and [string]$s.number -eq $num -and @($M.memberships|Where-Object {$_.student -eq $sid -and $_.class -eq $cls -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)}).Count -eq 1})
 }
 if($ids.Count -ne 1){return @{error='신원 미매칭/모호성'}};$sid=$ids[0];$s=$M.students[$sid]
 if(($name -and $s.name -cne $name) -or ($num -and [string]$s.number -ne $num) -or ($cls -and !@($M.memberships|Where-Object {$_.student -eq $sid -and $_.class -eq $cls -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)}).Count)){return @{error='학생ID와 표시 신원 불일치'}}
 return @{id=$sid}
}
function NValue($raw,$state){
 if($null -eq $raw -or [string]$raw -eq ''){return @{value=$null;status=$(if($state){$state}else{'미입력'})}}
 if([string]$raw -in @('미제출','면제','대기','미입력','미참여')){return @{value=$null;status=[string]$raw}}
 $v=$raw;if($raw -is [string]){$d=0.0;if([double]::TryParse($raw,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){$v=$d}}
 return @{value=$v;status=$(if($state){$state}elseif($v -is [ValueType] -and $v -eq 0){'0점'}else{'입력됨'})}
}
function NPreview {
 if(!(Test-Path -LiteralPath $File) -or [IO.Path]::GetExtension($File) -ne '.xlsx'){throw '.xlsx 자료 파일을 선택하세요.'}
 if((Get-Item -LiteralPath $File).Length -gt 16MB){throw '16MB 이하의 Excel 자료를 선택하세요.'}
 $fileHash=Shared-Hash $File;$batch='B'+$fileHash.Substring(0,20).ToLower();$book=[NaturalXlsx]::Read([IO.Path]::GetFullPath($File));if([IO.Path]::GetFileName($File) -match '(01[016789][- ]?\d{3,4}[- ]?\d{4}|\d{6}[- ]?[1-4]\d{6}|@)'){$book.Findings.Add('파일명 개인정보 패턴')};$records=[Collections.Generic.List[object]]::new();$isolated=[Collections.Generic.List[object]]::new();$structure=[Collections.Generic.List[object]]::new();$attempts=@{}
 $naturalCodeMap=@{};if(Test-Path -LiteralPath (Join-Path $Root 'project.json')){$naturalCodeMap=StudentCodeMap $Root $M.students}else{foreach($sid in $M.students.Keys){$naturalCodeMap[$sid]=$sid}}
 foreach($sheet in @($book.Sheets|Where-Object {$_.Name -notlike '_*'})){
  $header=0;$roles=@{};$fields=@{};$dup=@{}
  foreach($rn in ($sheet.Rows.Keys|Sort-Object|Select-Object -First 50)){
   $rr=@{};$ff=@{};$seen=@{};$dd=@{};foreach($cn in ($sheet.Rows[$rn].Keys|Sort-Object)){$v=$sheet.Rows[$rn][$cn].Value;if($null -eq $v -or [string]$v -eq ''){continue};$role=NRole $v;$key=NNorm $v;if($seen.ContainsKey($key)){$dd[$key]=$true};$seen[$key]=$true;if($role -ne 'field'){if($rr.ContainsKey($role)){$dd['@'+$role]=$true};$rr[$role]=$cn}else{$ff[$cn]=[string]$v}}
   if($rr.id -or ($rr.name -and $rr.number)){$header=$rn;$roles=$rr;$fields=$ff;$dup=$dd;break}
  }
  if(!$header){$isolated.Add(@{source_sheet=$sheet.Name;source_row=0;reason='학생 식별 머리글을 찾지 못함';raw=$null});continue}
  $structure.Add(@{sheet=$sheet.Name;header_row=$header;method='로컬 별칭/구조 규칙';roles=$roles;fields=@($fields.Keys|Sort-Object|ForEach-Object {@{column=[int]$_;label=$fields[$_]}})})
  foreach($rn in ($sheet.Rows.Keys|Sort-Object|Where-Object {$_ -gt $header})){
   $row=$sheet.Rows[$rn];if(!@($row.Values|Where-Object {$null -ne $_.Value -and [string]$_.Value -ne ''}).Count){continue}
   if((NRole (NCell $row $roles.id)) -eq 'id' -or (NRole (NCell $row $roles.name)) -eq 'name'){continue}
   foreach($cn0 in $row.Keys){if($cn0 -notin @($roles.Values) -and !$fields.ContainsKey($cn0) -and $null -ne $row[$cn0].Value -and [string]$row[$cn0].Value -ne ''){$isolated.Add(@{source_sheet=$sheet.Name;source_row=$rn;source_column=$cn0;reason='빈 머리글의 값: 자동 항목 지정 안 함';raw=$row[$cn0].Value})}}
   if(@($dup.Keys|Where-Object {$_.StartsWith('@')}).Count){$isolated.Add(@{source_sheet=$sheet.Name;source_row=$rn;reason='중복 식별/상태 머리글';raw=@($row.Values|ForEach-Object {$_.Value})});continue}
   if($roles.name){$nameValue=[string](NCell $row $roles.name);if($nameValue -and $nameValue -notmatch '^가상'){$book.Findings.Add('실명 가능성: '+$sheet.Name+'!'+$rn)}}
   $person=NStudent $row $roles $sheet.Name;if($person.error){$isolated.Add(@{source_sheet=$sheet.Name;source_row=$rn;reason=$person.error;raw=@($row.Values|ForEach-Object {$_.Value})});continue};$sid=$person.id
   if(!$attempts.ContainsKey($sid)){$attempts[$sid]=0};$attempts[$sid]++;$slot=$attempts[$sid];$state=[string](NCell $row $roles.status)
   $active=@{};foreach($cn in $fields.Keys){$active[$cn]=$fields[$cn]};if($roles.item -and $roles.value){$label=[string](NCell $row $roles.item);if(!$label){$isolated.Add(@{source_sheet=$sheet.Name;source_row=$rn;reason='빈 항목명';raw=$null});continue};$active=@{};$active[$roles.value]=$label}
   foreach($cn in ($active.Keys|Sort-Object)){
    $label=$active[$cn];$raw=NCell $row $cn;$loc=([string]$sheet.Name)+'!'+$rn+':'+$cn
    if($dup.ContainsKey((NNorm $label))){$isolated.Add(@{source_sheet=$sheet.Name;source_row=$rn;source_column=$cn;reason='중복 항목 머리글';raw=$raw});continue}
    if($row.ContainsKey([int]$cn) -and $row[[int]$cn].Formula){$isolated.Add(@{source_sheet=$sheet.Name;source_row=$rn;source_column=$cn;reason='수식은 자동 확정하지 않음';raw=$raw});continue}
    if($raw -is [bool]){$isolated.Add(@{source_sheet=$sheet.Name;source_row=$rn;source_column=$cn;reason='논리값의 평가 의미 수동 검토';raw=$raw});continue}
    $item='A.raw'+(NHash (NNorm $label)).Substring(0,12);$parsed=NValue $raw $state;$rid='R'+(NHash ($fileHash+'|'+$loc+'|'+$item)).Substring(0,24)
    $records.Add(@{student_id=$sid;item_id=$item;record_id=$rid;import_batch_id=$batch;value=$parsed.value;raw_value=$raw;source_file=[IO.Path]::GetFileName($File);source_file_hash=$fileHash;source_sheet=$sheet.Name;source_row=$rn;source_column=$cn;source_cell=$loc;status=$parsed.status;review_note='기계적 구조 변환. 교사 확정 점수 아님';item_label=$label;submission_slot=$slot;submission_label=[string](NCell $row $roles.attempt);created_at=$null;modified_at=$null;history=@()})
   }
  }
 }
 if((Shared-Hash $File) -ne $fileHash){throw '해석 도중 원본이 바뀌었습니다.'}
 $metaHash=Shared-Hash $MetaPath;$token=NHash ($fileHash+'|'+$metaHash+'|'+$Parent+'|'+$AsOf)
 $duplicate=$false;if($M.importBatches -and $M.importBatches.ContainsKey($batch)){$duplicate=$true}
 return @{schema_version='1.0';batch_id=$batch;fingerprint=$token;source_hash=$fileHash;metadata_hash=$metaHash;parent=$Parent;as_of=$AsOf;records=$records.ToArray();quarantine=$isolated.ToArray();structure=$structure.ToArray();privacy_findings=@($book.Findings);external_transmission_allowed=$false;privacy_scan_is_anonymity_guarantee=$false;duplicate_batch=$duplicate;read_only=$true}
}
function NCommit($p){
 if(!$PreviewFingerprint -or $p.fingerprint -ne $PreviewFingerprint){throw '원본/관계/기준일이 바뀌었거나 미리보기가 없습니다.'};if($p.duplicate_batch){throw '동일 원본의 반입 이력이 있습니다. 중복 생성하지 않습니다.'};[void](Layer-Info $Parent)
 if(!$p.records.Count){throw '정상 반입할 기록이 없습니다. 격리 사유를 확인하세요.'}
 $dir=Join-Path $Root ('imports/'+$p.batch_id);if(Test-Path $dir){throw '배치 디렉터리가 이미 존재합니다.'};New-Item -ItemType Directory $dir|Out-Null
 $sourceCopy='source'+[IO.Path]::GetExtension($File).ToLowerInvariant();Copy-Item -LiteralPath $MetaPath -Destination (Join-Path $dir 'before-relations.json');Copy-Item -LiteralPath $File -Destination (Join-Path $dir $sourceCopy)
 if((Shared-Hash (Join-Path $dir $sourceCopy)) -ne $p.source_hash){throw '원본 사본의 해시가 다릅니다. 다시 확인하세요.'}
 @{batch=$p.batch_id;recovery=$dir}|ConvertTo-Json|Set-Content (Join-Path $Root 'pending-import.json') -Encoding utf8
 if(!$M.ContainsKey('importBatches')){$M.importBatches=@{}};if(!$M.ContainsKey('importRecords')){$M.importRecords=@{}};$newLayers=[Collections.Generic.List[string]]::new();$now=Get-Date -Format o
 try{
  Excel
  foreach($group in ($p.records|Group-Object { [string]$_.submission_slot })){
   $id='I'+$p.batch_id.Substring(1,10)+'_'+$group.Name;$rel='books/'+$id+'.xlsx';$path=Resolve-Local $rel;if(Test-Path $path){throw '반입 대상 파일 존재'}
   $items=@($group.Group|ForEach-Object {$_.item_id}|Sort-Object -Unique);$sids=@($group.Group|ForEach-Object {$_.student_id}|Sort-Object -Unique);$headers=@('student_id','virtual_name','code')+$items
   $a=New-Object 'object[,]' ($sids.Count+1),$headers.Count;for($c=0;$c -lt $headers.Count;$c++){$a[0,$c]=[string]$headers[$c]};$rowMap=@{};for($r=0;$r -lt $sids.Count;$r++){$sid=$sids[$r];$rowMap[$sid]=$r+1;$a[($r+1),0]=[string]$sid;$a[($r+1),1]=[string]$M.students[$sid].name;$a[($r+1),2]=[string]$sid}
   foreach($rec in $group.Group){$c=[array]::IndexOf($headers,$rec.item_id);$r=$rowMap[$rec.student_id];$v=$rec.value;if($null -eq $v){$a[$r,$c]=$null}elseif($v -is [ValueType]){$a[$r,$c]=[double]$v}else{$a[$r,$c]="'"+[string]$v};$M.labels["$id|$($rec.item_id)"]=$rec.item_label;$rec.layer_id=$id;$rec.created_at=$now;$rec.modified_at=$now;$rec.history=@(@{event='import';at=$now;value=$rec.value})}
   $stage='새 파일';$b=$XL.Workbooks.Add();try{
    $s=$b.Worksheets.Item(1);$s.Name='Data';$stage='행렬 기록'
    $matrixInfo=@{rows=$a.GetLength(0);columns=$a.GetLength(1);types=@($a|ForEach-Object {if($null -eq $_){'null'}else{$_.GetType().FullName}}|Sort-Object -Unique);first=$a[0,0]};$matrixInfo|ConvertTo-Json|Set-Content (Join-Path $Root 'matrix-diagnostic.json') -Encoding utf8
    for($ri=0;$ri -lt $a.GetLength(0);$ri++){for($ci=0;$ci -lt $a.GetLength(1);$ci++){if($null -eq $a[$ri,$ci]){continue};if($a[$ri,$ci] -is [ValueType]){$s.Cells.Item(($ri+1),($ci+1)).Value2=[double]$a[$ri,$ci]}else{$s.Cells.Item(($ri+1),($ci+1)).Value2=[string]$a[$ri,$ci]}}}
    $stage='서식';$s.UsedRange.Columns.ColumnWidth=22
    $stage='정체성';$identity=$b.Worksheets.Add();$identity.Name='_identity';$identity.Cells.Item(1,2).Value2=$id;$identity.Visible=2
    $stage='신규 저장 '+$path;$b.SaveAs($path,51)
   }finally{$b.Close($false)}
   $M.layers[$id]=@{path=$rel;parent=$Parent;items=$items;name=('반입 '+[IO.Path]::GetFileNameWithoutExtension($File)+' / 제출 '+$group.Name);importBatch=$p.batch_id};$newLayers.Add($id)
  }
  foreach($rec in $p.records){
   $canonical="$($rec.layer_id)|$($rec.student_id)|$($rec.item_id)"
   $M.importRecords[$canonical]=@{status=$rec.status;record_id=$rec.record_id;batch_id=$p.batch_id;source_cell=$rec.source_cell}
   if($rec.resolution){$M.manualOverrides[$canonical]=@{student_id=$rec.student_id;item_id=$rec.item_id;record_id=$rec.record_id;import_batch_id=$p.batch_id;value=$rec.value;raw_value=$rec.raw_value;status=$rec.status;note=$rec.review_note;review_note=$rec.review_note;explicit=$true;source_file=$rec.source_file;source_sheet=$rec.source_sheet;source_row=$rec.source_row;source_column=$rec.source_column;modified_at=$now;history=@(@{event='teacher_resolution';at=$now;value=$rec.value;note=$rec.review_note;source_record_id=$rec.resolution.source_record_id})}}
  }
  $p.read_only=$false;$p.created_at=$now;$p.layers=$newLayers.ToArray();$p.active=$true
  $p|ConvertTo-Json -Depth 25|Set-Content (Join-Path $dir 'exchange.json') -Encoding utf8
  $M.importBatches[$p.batch_id]=@{layers=$newLayers.ToArray();active=$true;source_hash=$p.source_hash;created_at=$now;exchange='imports/'+$p.batch_id+'/exchange.json'};Save-Meta
  Move-Item -LiteralPath (Join-Path $Root 'pending-import.json') -Destination (Join-Path $dir 'completed.json');return $p
 }catch{throw ('반입 중단. 단계 '+$stage+' / 원본·상위 보존. 복구 필요: '+$_.Exception.Message+' '+$_.ScriptStackTrace)}finally{Close-Excel}
}
function NWithdraw($batch){
 if(!$M.importBatches -or !$M.importBatches.ContainsKey($batch)){throw '알 수 없는 배치'};$info=$M.importBatches[$batch];if(!$info.active){throw '이미 철회한 배치'}
 $ids=@($info.layers);foreach($ln in $M.links){if($ln.source -in $ids -or $ln.target -in $ids){throw '연결이 존재하므로 철회할 수 없습니다.'}}
 foreach($id in $ids){if(@($M.layers.Values|Where-Object {$_.parent -eq $id}).Count){throw '하위 작업층이 있어 철회할 수 없습니다.'}}
 foreach($id in $ids){$M.layers.Remove($id)};$info.active=$false;$info.withdrawn_at=Get-Date -Format o;Save-Meta
 # Books, source, exchange and teacher edits remain physically preserved for audit/recovery.
 return @{batch_id=$batch;active=$false;preserved_files=$true}
}
function NRecover {
 $pending=Join-Path $Root 'pending-import.json';if(!(Test-Path $pending)){throw '중단된 반입이 없습니다.'};$q=Get-Content $pending -Raw -Encoding utf8|ConvertFrom-Json
 if($q.batch -notmatch '^B[0-9a-f]{20}$'){throw '복구 배치ID 오류'};$before=Join-Path $Root ('imports/'+$q.batch+'/before-relations.json');$old=To-Map (Get-Content $before -Raw -Encoding utf8|ConvertFrom-Json)
 $dest=Join-Path $Root ('recovered-import-'+[guid]::NewGuid().ToString('N'));New-Item -ItemType Directory $dest,(Join-Path $dest 'books'),(Join-Path $dest 'backups'),(Join-Path $dest 'logs')|Out-Null
 foreach($info in $old.layers.Values){$src=Resolve-Local $info.path;$target=Join-Path $dest $info.path;New-Item -ItemType Directory (Split-Path $target) -Force|Out-Null;Copy-Item -LiteralPath $src -Destination $target};Copy-Item -LiteralPath $before -Destination (Join-Path $dest 'relations.json')
 # Earlier active batches must retain exchange/source records in the restored copy.
 foreach($bid in $old.importBatches.Keys){if($bid -notmatch '^B[0-9a-f]{20}$'){throw '복구 배치ID 오류'};New-Item -ItemType Directory (Join-Path $dest 'imports') -Force|Out-Null;Copy-Item -LiteralPath (Join-Path $Root ('imports/'+$bid)) -Destination (Join-Path $dest ('imports/'+$bid)) -Recurse}
 return @{workspace=$dest;original_preserved=$true;metadata_restored_to_before_batch=$true}
}

