# Read local source bytes only. Source subject keys are never roster identities.
function SHeader($row){
 $labels=@{};$subjects=@();$seen=@{};$duplicates=@()
 foreach($cn in ($row.Keys|Sort-Object)){$raw=$row[$cn].Value;$v=[string]$raw;if(!$v.Trim()){continue};if($raw -isnot [string] -or $v.Length -gt 120){return $null};$labels[$cn]=$v;$key=NNorm $v;if($seen.ContainsKey($key)){$duplicates+=,$cn};$seen[$key]=$true
  if($key -in @('식별번호','식별번호·판독','익명코드','학생id','student_id','학생코드','고유번호')){$subjects+=,$cn}
 }
 if(!$subjects.Count){$nameColumns=@($labels.Keys|Where-Object {(NRole $labels[$_]) -eq 'name'});$numberColumns=@($labels.Keys|Where-Object {(NRole $labels[$_]) -eq 'number'});if($nameColumns.Count -eq 1 -and $numberColumns.Count -eq 1){$subjects=@($nameColumns[0])}}
 if($subjects.Count -eq 1 -and $labels.Count -ge 2){return @{subject=$subjects[0];labels=$labels;duplicates=$duplicates}};return $null
}
function SKey($preview,$record){if($record.origin_key){return (NHash ('external|'+$record.origin_key))};return (NHash ($preview.source_hash+'|'+$record.source_record_id))}
function SPreview{
 if(!(Test-Path -LiteralPath $File) -or (Get-Item -LiteralPath $File).Length -gt 16MB){throw '16MB 이하 로컬 파일을 선택하세요.'}
 [void](Layer-Info $Parent);$hash=Shared-Hash $File;$records=[Collections.Generic.List[object]]::new();$issues=[Collections.Generic.List[object]]::new();$findings=@();$contexts=@{};$ext=[IO.Path]::GetExtension($File).ToLowerInvariant()
 if($ext -eq '.xlsx'){
  $book=[NaturalXlsx]::Read([IO.Path]::GetFullPath($File));$findings=@($book.Findings)
  foreach($sheet in $book.Sheets){
   $header=$null;$tables=0
   foreach($rn in ($sheet.Rows.Keys|Sort-Object)){
    $row=$sheet.Rows[$rn];$candidate=SHeader $row
    if($candidate){$header=$candidate;$tables++;continue}
    if(!$header){continue};if(!@($row.Values|Where-Object {$null -ne $_.Value -and [string]$_.Value -ne ''}).Count){continue}
    $subject=NCell $row $header.subject
    if($null -eq $subject -or [string]$subject -eq ''){$issues.Add(@{source_sheet=$sheet.Name;source_row=$rn;reason='식별번호가 없는 행: 자동 연결 안 함';cells=@($row.Keys|Sort-Object|ForEach-Object {@{column=$_;value=$row[$_].Value}})});continue}
    $contextId=NHash ($sheet.Name+'|'+$rn);$context=@($row.Keys|Sort-Object|ForEach-Object {@{column=[int]$_;label=$header.labels[$_];value=$row[$_].Value;formula=$row[$_].Formula}})
    $contexts[$contextId]=$context
    foreach($cn in ($row.Keys|Sort-Object)){
     if($cn -eq $header.subject){continue};$cell=$row[$cn];if($null -eq $cell.Value -or [string]$cell.Value -eq ''){continue};$label=$header.labels[$cn];$reason='학생·평가항목을 교사가 지정해야 함'
     if(!$label){$reason='빈 머리글: 항목 지정 필요'}elseif($cn -in $header.duplicates){$reason='중복 머리글: 항목 지정 필요'}
     $loc=$sheet.Name+'!'+$rn+':'+$cn;$rid='S'+(NHash ($hash+'|'+$loc)).Substring(0,24)
     $records.Add(@{source_record_id=$rid;source_subject_key=[string]$subject;source_sheet=$sheet.Name;source_row=$rn;source_column=$cn;source_cell=$loc;item_label=$label;raw_value=$cell.Value;value=$cell.Value;formula=$cell.Formula;reason=$reason;context_id=$contextId;table=$tables})
    }
   }
   if(!$tables){$issues.Add(@{source_sheet=$sheet.Name;source_row=0;reason='대상 표 머리글 없음: 원본 보존';cells=@()})}
  }
 }elseif($ext -eq '.json'){
  $j=To-Map (Get-Content -LiteralPath $File -Raw -Encoding utf8|ConvertFrom-Json)
  if($j.schema_version -ne '1.0' -or $j.records -isnot [array] -or !$j.records -or @($j.records).Count -gt 10000){throw '교환규격 1.0 records 배열(1~10000건)이 필요합니다.'}
  $ids=@{};$locations=@{};foreach($r in $j.records){
   foreach($k in @('record_id','source_file_hash','source_sheet','source_row','source_column','item_label')){if(!$r.ContainsKey($k) -or $null -eq $r[$k] -or [string]$r[$k] -eq ''){throw ('출처 필수 항목 누락: '+$k)}}
   foreach($k in @('record_id','source_file_hash','source_sheet','item_label')){if($r[$k] -isnot [string]){throw ('문자열 필수: '+$k)}}
   if($r.source_file_hash -notmatch '^[a-fA-F0-9]{64}$' -or $r.source_row -isnot [ValueType] -or $r.source_column -isnot [ValueType] -or $r.source_row -is [bool] -or $r.source_column -is [bool] -or $r.source_row -lt 1 -or $r.source_row -gt 20000 -or $r.source_column -lt 1 -or $r.source_column -gt 256 -or [math]::Floor([double]$r.source_row) -ne $r.source_row -or [math]::Floor([double]$r.source_column) -ne $r.source_column){throw '출처 해시·행·열 오류'}
   $rid=[string]$r.record_id;if($ids.ContainsKey($rid)){throw '중복 record_id'};$ids[$rid]=$true
   $location=NHash ($r.source_file_hash.ToUpperInvariant()+'|'+$r.source_sheet+'|'+$r.source_row+'|'+$r.source_column);if($locations.ContainsKey($location)){throw '동일 원문 셀의 중복 변환 기록'};$locations[$location]=$true
   if(!$r.ContainsKey('raw_value') -or !$r.ContainsKey('value')){throw '원문값과 변환값 필수'}
   foreach($v in @($r.value,$r.raw_value)){if($null -ne $v -and $v -isnot [string] -and $v -isnot [ValueType]){throw '배열·객체 입력값 거부'}}
   $contexts[$rid]=@($r);$records.Add(@{source_record_id=$rid;origin_key=($r.source_file_hash.ToUpperInvariant()+'|'+$rid);source_subject_key=[string]$r.student_id;source_sheet=$r.source_sheet;source_row=$r.source_row;source_column=$r.source_column;source_cell=($r.source_sheet+'!'+$r.source_row+':'+$r.source_column);item_label=$r.item_label;raw_value=$r.raw_value;value=$r.value;formula=$false;reason='외부 변환값: 학생·항목·값 교사 확인 필요';context_id=$rid;table=1})
  }
 }else{throw 'xlsx 또는 교환규격 JSON을 선택하세요.'}
 if((Shared-Hash $File) -ne $hash){throw '검사 중 원본 변경'}
 $fingerprint=NHash ($hash+'|'+(Shared-Hash $MetaPath)+'|'+$Parent+'|'+$AsOf+'|staging-1')
 foreach($r in $records){$key=SKey @{source_hash=$hash} $r;$r.resolved=($M.resolutions -and $M.resolutions.ContainsKey($key))}
 return @{schema_version='1.0';source_hash=$hash;fingerprint=$fingerprint;read_only=$true;records=$records.ToArray();quarantine=$issues.ToArray();contexts=$contexts;privacy_findings=$findings;external_transmission_allowed=$false;identity_mapping_automatic=$false}
}
function SCommit($request){
 $p=SPreview;if(!$request.expected_fingerprint -or $request.expected_fingerprint -ne $p.fingerprint){throw '원본·작업공간·기준일이 바뀌었습니다. 다시 검토하세요.'}
 $matches=@($p.records|Where-Object {$_.source_record_id -ceq $request.source_record_id});if($matches.Count -ne 1){throw '원문 기록을 다시 선택하세요.'};$src=$matches[0]
 if($src.resolved){throw '이미 처리한 원문 기록입니다. 저장된 작업층에서 수정하세요.'}
 if(!$request.student_id -or !$M.students.ContainsKey($request.student_id)){throw '등록된 학생을 명시적으로 선택하세요.'}
 if(!$request.item_label -or !$request.item_label.Trim()){throw '항목명을 입력하세요.'};$request.status=if($null -eq $request.value){'미입력'}else{'입력됨'};$request.review_note=''
 $v=$request.value;if($null -ne $v -and $v -isnot [string] -and $v -isnot [ValueType]){throw '단일 입력값 필요'};if($v -is [string] -and $v.Length -gt 32767){throw '셀 입력값은 32767자 이하여야 합니다.'}
 if($v -is [bool]){throw '논리값은 숫자로 확정할 수 없습니다.'}
 if($v -is [double] -and ([double]::IsNaN($v) -or [double]::IsInfinity($v))){throw '유한한 숫자 필요'}
 $key=SKey $p $src;$batch='B'+$key.Substring(0,20);$item='A.raw'+(NHash (NNorm $request.item_label)).Substring(0,12);$rid='R'+$key.Substring(0,24)
 $record=@{student_id=[string]$request.student_id;item_id=$item;record_id=$rid;import_batch_id=$batch;value=$v;raw_value=$src.raw_value;source_file=[IO.Path]::GetFileName($File);source_file_hash=$p.source_hash;source_sheet=$src.source_sheet;source_row=$src.source_row;source_column=$src.source_column;source_cell=$src.source_cell;status=$request.status;review_note=$request.review_note;item_label=$request.item_label;submission_slot=1;submission_label='교사 검토';created_at=$null;modified_at=$null;history=@();resolution=@{source_record_id=$src.source_record_id;source_subject_key=$src.source_subject_key;source_value=$src.value;context=$p.contexts[$src.context_id];teacher_mapping=$true;at=(Get-Date -Format o);note=$request.review_note}}
 if(!$M.resolutions){$M.resolutions=@{}};$M.resolutions[$key]=@{batch=$batch;record=$rid;at=(Get-Date -Format o)}
 $normal=@{batch_id=$batch;fingerprint=$p.fingerprint;source_hash=$p.source_hash;records=@($record);quarantine=@();duplicate_batch=($M.importBatches -and $M.importBatches.ContainsKey($batch));read_only=$true;parent=$Parent;as_of=$AsOf;external_transmission_allowed=$false}
 $script:PreviewFingerprint=$p.fingerprint
 $saved=NCommit $normal;foreach($id in $saved.layers){$M.layers[$id].name='검토 완료 · '+$request.item_label};Save-Meta;return $saved
}
