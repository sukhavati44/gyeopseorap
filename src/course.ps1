function CProfile {
 if(!$M.courses -or !$Name -or !$M.courses.ContainsKey($Name)){throw '교육과정 연결을 먼저 설정하세요.'};return $M.courses[$Name]
}
function CFields($p){
 return WFields $p
}
function CItemLabel($p,$item){if($Layer -eq $p.root){$a=@($p.areas|Where-Object root_item -EQ $item);if($a.Count -eq 1){return $a[0].label}};return (Label-For $Layer $item)}
function CConfigure($r){
 if(!$r.subject -or $r.year -notmatch '^20\d\d$' -or $r.term -notin @('1','2') -or $r.grade -notmatch '^(1[0-2]|[1-9])$'){throw '학년도·학기·학년·과목을 확인하세요.'}
 $rootInfo=Layer-Info $r.root;if($rootInfo.parent){throw '최상위 종합표를 선택하세요.'};$seen=@{}
 foreach($a in $r.areas){if(!$a.label -or $a.root_item -notin @($rootInfo.items) -or $seen.ContainsKey($a.root_item)){throw '영역명·종합표 항목 중복/누락'};$seen[$a.root_item]=$true
  if($a.work_layer){$li=Layer-Info $a.work_layer;if($li.parent -ne $r.root){throw '영역 작업표는 종합표 바로 아래여야 합니다.'};foreach($i in $a.work_items){if($i -notin @($li.items)){throw '작업표 항목이 없습니다.'}};if(!@($a.work_items).Count){throw '영역 세부 항목이 필요합니다.'}}
 }
 if(!$seen.Count){throw '평가영역을 하나 이상 선택하세요.'}
 if(!$M.courses){$M.courses=@{}};foreach($c in $M.courses.Values){if($c.root -eq $r.root){throw '이 종합표는 이미 교육과정에 연결됐습니다.'};if($c.year -eq $r.year -and $c.term -eq $r.term -and $c.grade -eq $r.grade -and $c.subject -eq $r.subject){throw '같은 학년도·학기·학년·과목이 이미 연결됐습니다.'}}
 $id='C'+[guid]::NewGuid().ToString('N');$r.id=$id;$M.courses[$id]=$r;Save-Meta;return @{id=$id;profile=$r;original_books_unchanged=$true}
}
function CWork($req){
 $p=CProfile;$matches=@($p.areas|Where-Object root_item -EQ $Area);if($matches.Count -ne 1){throw '평가영역을 선택하세요.'};$a=$matches[0];$prefix=$a.root_item.Split('.')[0]
 if(!$a.work_layer){
  if(@($M.links|Where-Object {$_.target -eq $p.root -and $_.targetItem -eq $a.root_item}).Count){throw '기존 연결을 먼저 확인하세요.'}
  $id='W'+[guid]::NewGuid().ToString('N').Substring(0,12);$item=$prefix+'.score';Excel;try{New-Layer $id $p.root @($item,($prefix+'.status'))}finally{Close-Excel}
  $M.layers[$id].name=$a.label+' 평가 작업표';$M.labels["$id|$item"]=$a.label+' 영역 점수';$a.work_layer=$id;$a.work_items=@($item);$p.numeric_items+=,$item
  $M.links+=@{source=$id;sourceItem=$item;target=$p.root;targetItem=$a.root_item};Save-Meta;return @{layer=$id;item=$item;created=$true;automatic_score_sync=$false}
 }
 if(!$req.label -or !$req.label.Trim()){throw '세부 항목명을 입력하세요.'};$id=$a.work_layer;$item=$prefix+'.i'+[guid]::NewGuid().ToString('N').Substring(0,10);Excel
 try{$b=Open-Book $id;try{$t=Read-Table $b;if($t.warnings.Count){throw '식별자 경고를 먼저 확인하세요.'};$cpRel='backups/item-'+[guid]::NewGuid().ToString('N');$cp=Resolve-Local $cpRel;New-Item -ItemType Directory $cp|Out-Null;Copy-Item $MetaPath (Join-Path $cp 'relations.json')
  foreach($lid in $M.layers.Keys){$rel=$M.layers[$lid].path;$dest=Join-Path $cp $rel;New-Item -ItemType Directory (Split-Path $dest) -Force|Out-Null;if($lid -eq $id){$b.SaveCopyAs($dest)}else{Copy-Item -LiteralPath (Resolve-Local $rel) -Destination $dest}}
  @{checkpoint=$cpRel;action='add_detail';layer=$id}|ConvertTo-Json|Set-Content (Join-Path $Root 'pending-sync.json') -Encoding utf8
  $c=($t.cols.Values|Measure-Object -Maximum).Maximum+1;$t.sheet.Cells.Item(1,$c).Value2=[string]$item;$b.Save();$M.layers[$id].items+=,$item;$M.labels["$id|$item"]=$req.label;$a.work_items+=,$item;if($req.numeric){$p.numeric_items+=,$item};Save-Meta
  Move-Item (Join-Path $Root 'pending-sync.json') (Join-Path $cp 'completed.json');return @{layer=$id;item=$item;created=$false;automatic_score_sync=$false}
 }finally{$b.Close($false)}}finally{Close-Excel}
}
function CSnapshot {
 $p=CProfile;$fields=CFields $p;$path=Resolve-Local (Layer-Info $Layer).path;$before=NHash ((Shared-Hash $MetaPath)+'|'+(Shared-Hash $path));$rows=[Collections.Generic.List[object]]::new()
 Excel;try{$b=Open-Book $Layer $true;try{$t=Read-Table $b
  foreach($sid in ($t.rows.Keys|Sort-Object)){$s=$M.students[$sid];$member=@($M.memberships|Where-Object {$_.student -eq $sid -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)})
   if($Student -and $sid -ne $Student){continue};if($Class -and $Class -notin @($member.class)){continue};if($p.classes -and !@($member|Where-Object {$_.class -in @($p.classes)}).Count){continue}
   foreach($item in $fields){if(!$t.cols.ContainsKey($item)){throw '표의 평가항목이 변경됐습니다.'};$r=$t.rows[$sid];$c=$t.cols[$item];$v=$t.values[$r,$c];$key="$Layer|$sid|$item";$manual=$M.manualOverrides[$key];$import=$M.importRecords[$key];$origin=$M.provenance[$key];$state=Source-State $t $sid $item $v
    $state=if($null -eq $v){'미입력'}else{'입력됨'}
    $rows.Add(@{student=$sid;name=$s.name;number=$s.number;class=($member.class -join ',');item=$item;label=(CItemLabel $p $item);value=$v;state=$state;note=$(if($manual){$manual.note}else{''});formula=([string]$t.formulas[$r,$c]).StartsWith('=');numeric=($item -in @($p.numeric_items));origin=$(if($manual){'교사 직접입력'}elseif($import){'자료 반입'}elseif($origin -and (Same-Value $v $origin.value)){'하위 작업표 참조'}elseif($origin){'직접수정·검토 필요'}else{'직접입력'});source=$origin.source;sourceItem=$origin.item;sourceFile=(Layer-Info $Layer).path;recordId=$(if($manual){$manual.record_id}elseif($import){$import.record_id}else{''});sourceCell=$import.source_cell})
   }
  };$warnings=@($t.warnings)
 }finally{$b.Close($false)}}finally{Close-Excel}
 $after=NHash ((Shared-Hash $MetaPath)+'|'+(Shared-Hash $path));if($before -ne $after){throw '조회 도중 자료가 변경됐습니다.'}
 return @{data=$rows.ToArray();fields=$fields;fingerprint=$before;warnings=$warnings;profile=$p;layer=$Layer;area=$Area;class=$Class;as_of=$AsOf}
}
function CValidateValue($v,$state,$numeric){
 if($null -ne $v){if($v -is [bool] -or ($v -isnot [string] -and $v -isnot [ValueType])){throw '숫자 또는 문자 값을 입력하세요.'};if($numeric -and $v -isnot [ValueType]){throw '점수는 숫자로 입력하세요.'};if($v -is [double] -and ([double]::IsNaN($v) -or [double]::IsInfinity($v))){throw '유한한 숫자로 입력하세요.'};if($v -is [string] -and $v.Length -gt 32767){throw '문자는 32767자 이하로 입력하세요.'}}
}
function CCheckSourceHashes($req){
 if(!$req.source_fingerprints){return}
 foreach($id in $req.source_fingerprints.Keys){if(!$M.layers.ContainsKey($id) -or (Shared-Hash (Resolve-Local $M.layers[$id].path)) -cne $req.source_fingerprints[$id]){throw '하위 값이 바뀌었습니다. 다시 가져오세요.'}}
}
function CCommit($req){
 $p=CProfile;$fields=CFields $p;$path=Resolve-Local (Layer-Info $Layer).path;$fp=NHash ((Shared-Hash $MetaPath)+'|'+(Shared-Hash $path))
 if(!$req.expected_fingerprint -or $req.expected_fingerprint -ne $fp){throw '조회 후 파일·관계가 변경됐습니다. 다시 조회하세요.'}
 if(!$req.changes){throw '반영할 값이 없습니다.'}
 CCheckSourceHashes $req
 $newFields=@($req.new_fields|Where-Object {$_})
 if($newFields.Count -and !(CIsLeaf)){throw '하위 연결이 있는 표에는 새 열을 자동 추가하지 않습니다.'}
 $newSeen=@{};foreach($nf in $newFields){if($nf.item -notmatch '^IMP\.[a-f0-9]{20}$' -or !$nf.label -or $nf.item -in $fields -or $newSeen[$nf.item]){throw '새 열 정보를 다시 확인하세요.'};$newSeen[$nf.item]=$true;$fields+=,$nf.item}
 $seen=@{};$accepted=[Collections.Generic.List[object]]::new();Excel
 try{$b=Open-Book $Layer;try{$t=Read-Table $b
  foreach($change in $req.changes){$sid=[string]$change.student;$item=[string]$change.item;$key="$Layer|$sid|$item"
   if(($Student -and $sid -ne $Student) -or $seen.ContainsKey($key) -or !$t.rows.ContainsKey($sid) -or $item -notin $fields){throw '중복/미등록 학생·항목 또는 현재 영역 밖 변경'};$seen[$key]=$true
   $members=@($M.memberships|Where-Object {$_.student -eq $sid -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)})
   if(($Class -and $Class -notin @($members.class)) -or ($p.classes -and !@($members|Where-Object {$_.class -in @($p.classes)}).Count)){throw '조회 학급 밖 학생 변경'}
   CValidateValue $change.value $change.status ($item -in @($p.numeric_items) -or @($newFields|Where-Object {$_.item -eq $item -and $_.numeric}).Count -gt 0)
   $r=$t.rows[$sid];$c=$t.cols[$item];$change.status=if($null -eq $change.value){'미입력'}else{'입력됨'};$change.note=''
   $accepted.Add(@{student=$sid;item=$item;row=$r;column=$c;before=$(if($c){$t.values[$r,$c]}else{$null});value=$change.value;status=$change.status;note=$change.note;canonical=$key})
  }
  if((NHash ((Shared-Hash $MetaPath)+'|'+(Shared-Hash $path))) -ne $fp){throw '검증 중 파일 변경'}
  $cpRel='backups/table-'+[guid]::NewGuid().ToString('N');$cp=Resolve-Local $cpRel;New-Item -ItemType Directory $cp|Out-Null;Copy-Item -LiteralPath $MetaPath -Destination (Join-Path $cp 'relations.json')
  foreach($id in $M.layers.Keys){$rel=$M.layers[$id].path;$dest=Join-Path $cp $rel;New-Item -ItemType Directory (Split-Path $dest) -Force|Out-Null;if($id -eq $Layer){$b.SaveCopyAs($dest)}else{Copy-Item -LiteralPath (Resolve-Local $rel) -Destination $dest}}
  $uploadName=if([IO.Path]::GetExtension($File) -eq '.csv'){'upload.csv'}else{'upload.xlsx'}
  if($req.import_hash){Copy-Item -LiteralPath $File -Destination (Join-Path $cp $uploadName);if((Shared-Hash (Join-Path $cp $uploadName)) -ne $req.import_hash){throw '업로드 파일이 저장 전 변경됐습니다.'}}
  CCheckSourceHashes $req
  @{checkpoint=$cpRel;action='table';layer=$Layer}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $Root 'pending-sync.json') -Encoding utf8
  foreach($nf in $newFields){$c=($t.cols.Values|Measure-Object -Maximum).Maximum+1;$t.sheet.Cells.Item(1,$c).Value2=[string]$nf.item;$t.cols[$nf.item]=[int]$c;$M.layers[$Layer].items+=,$nf.item;$M.labels["$Layer|$($nf.item)"]=$nf.label;if($nf.numeric){$p.numeric_items+=,$nf.item};foreach($a in $p.areas){if($a.work_layer -eq $Layer){$a.work_items+=,$nf.item}}}
  if($newFields.Count -and !$M.workspaceScopes[$Layer]){$M.workspaceScopes[$Layer]=@{course=$Name;area=$Area}}
  foreach($entry in $accepted){$entry.column=$t.cols[$entry.item]}
  $now=Get-Date -Format o
  foreach($c in $accepted){$v=$c.value;if($null -eq $v){[void]$t.sheet.Cells.Item($c.row,$c.column).ClearContents()}elseif($v -is [ValueType]){$t.sheet.Cells.Item($c.row,$c.column).Value2=[double]$v}else{$t.sheet.Cells.Item($c.row,$c.column).Value2="'"+[string]$v}
   $key=$c.canonical;$previous=$M.manualOverrides[$key];$record='M'+(NHash $key).Substring(0,24);$batch=$null;if($M.importRecords[$key]){$record=$M.importRecords[$key].record_id;$batch=$M.importRecords[$key].batch_id}
   $history=@();if($previous){$history=@($previous.history)};$history+=@{record_id=$record;at=$now;before=$c.before;after=$v;status=$c.status;note=$c.note;event='table_edit';upload_hash=$req.import_hash;upload_checkpoint=$(if($req.import_hash){$cpRel+'/'+$uploadName}else{$null})}
   $reference=if($req.source_refs){$req.source_refs[$c.student+'|'+$c.item]}else{$null};if($reference){$M.provenance[$key]=@{source=$reference.source;item=$reference.item;value=$v;refreshed=$now}}
   $M.manualOverrides[$key]=@{student_id=$c.student;item_id=$c.item;record_id=$record;import_batch_id=$batch;value=$v;raw_value=$c.before;status=$c.status;note=$c.note;review_note=$c.note;source_file=(Layer-Info $Layer).path;source_sheet='Data';source_row=$c.row;source_column=$c.column;modified_at=$now;history=$history;explicit=$true}
  }
  $b.Save();Save-Meta;Move-Item -LiteralPath (Join-Path $Root 'pending-sync.json') -Destination (Join-Path $cp 'completed.json')
  return @{saved=$accepted.Count;checkpoint=$cpRel;changes=$accepted.ToArray();manual_protection=$true;new_columns=$newFields.Count}
 }finally{$b.Close($false)}}finally{Close-Excel}
}
function CExport{
 try{
 if([IO.Path]::GetExtension($File).ToLowerInvariant() -ne '.xlsx' -or (Test-Path -LiteralPath $File)){throw '새 xlsx 파일 이름을 선택하세요. 기존 파일을 덮어쓰지 않습니다.'}
 $q=CSnapshot;$students=@($q.data|Group-Object {$_.student}|ForEach-Object {$_.Group[0]}|Sort-Object @{Expression={$_.class}},@{Expression={[int]$_.number}}|ForEach-Object {$_.student});if(!$students.Count){throw '조회한 학생이 없습니다.'};$fields=@($q.fields);$leading=if($Layer -eq $q.profile.root){4}else{5};$cols=$leading+$fields.Count;$matrix=New-Object 'object[,]' ($students.Count+3),$cols
 $matrix[0,0]=[string]('수행평가 / '+$q.profile.year+'학년도 '+$q.profile.term+'학기 '+$q.profile.grade+'학년 '+$q.profile.subject+' / '+$Class)
 $headers=@('반','번호','성명','학생코드');$keys=@('@class','@number','@name','@student');if($leading -eq 5){$headers=@('반','번호','성명','학생코드','합계 (상위 반영)');$keys=@('@class','@number','@name','@student','@total')}
 foreach($item in $fields){$label=CItemLabel $q.profile $item;$headers+=@($label);$keys+=@('value|'+$item)}
 for($c=0;$c -lt $cols;$c++){$matrix[2,$c]="'"+[string]$headers[$c];$matrix[1,$c]="'"+[string]$keys[$c]}
 for($r=0;$r -lt $students.Count;$r++){$sid=$students[$r];$records=@($q.data|Where-Object student -EQ $sid);$matrix[($r+3),0]="'"+[string]$records[0].class;$matrix[($r+3),1]=[double]$records[0].number;$matrix[($r+3),2]="'"+[string]$records[0].name;$matrix[($r+3),3]="'"+(StudentCode $Root $sid)
  for($i=0;$i -lt $fields.Count;$i++){$x=@($records|Where-Object item -EQ $fields[$i])[0];$v=$x.value;$col=$leading+$i;if($null -ne $v){$matrix[($r+3),$col]=if($v -is [ValueType]){[double]$v}else{"'"+[string]$v}}}
 }
 Excel;try{$b=$XL.Workbooks.Add();try{$s=$b.Worksheets.Item(1);$s.Name='성적표';while($b.Worksheets.Count -gt 1){$b.Worksheets.Item($b.Worksheets.Count).Delete()};for($ri=0;$ri -lt ($students.Count+3);$ri++){for($ci=0;$ci -lt $cols;$ci++){if($null -ne $matrix[$ri,$ci]){if($matrix[$ri,$ci] -is [double]){$s.Cells.Item(($ri+1),($ci+1)).Value2=[double]$matrix[$ri,$ci]}else{$s.Cells.Item(($ri+1),($ci+1)).Value2=[string]$matrix[$ri,$ci]}}}}
  if($leading -eq 5){
   for($ri=4;$ri -le ($students.Count+3);$ri++){$refs=@();for($i=0;$i -lt $fields.Count;$i++){if($fields[$i] -in @($q.profile.numeric_items)){$refs+=('RC'+($leading+$i+1))}};$formula=if($refs.Count){'=IF(COUNT('+($refs -join ',')+')=0,"",SUM('+($refs -join ',')+'))'}else{'=""'};$s.Cells.Item($ri,5).FormulaR1C1=$formula}
  }
  $s.UsedRange.Font.Name='맑은 고딕';$s.UsedRange.Font.Size=10;$s.UsedRange.Columns.ColumnWidth=18;$s.Rows.Item(3).Font.Bold=$true;$s.Rows.Item(3).Interior.Color=15790320;$s.Rows.Item(3).RowHeight=36;$s.Rows.Item(3).WrapText=$true;$s.Rows.Item(2).Hidden=$true;$s.Columns.Item(2).NumberFormat=$(if($Layer -eq $q.profile.root){'0'}else{'00'});$s.Range($s.Cells.Item(3,1),$s.Cells.Item(($students.Count+3),$cols)).AutoFilter()|Out-Null
  $s.Rows.Item(1).Font.Size=14;$s.Rows.Item(1).RowHeight=28;$s.Columns.Item(3).ColumnWidth=20;$s.Columns.Item(4).ColumnWidth=42

  $identity=$b.Worksheets.Add();$identity.Name='_exchange';$values=@('assessment-table-1.1.1',$Name,$Layer,$q.fingerprint,$AsOf,($fields|ConvertTo-Json -Compress),($students|ConvertTo-Json -Compress),$Class,$Area,$Student)
  for($i=0;$i -lt $values.Count;$i++){$identity.Cells.Item(($i+1),1).Value2="'"+[string]$values[$i]};$identity.Visible=2;$s.Activate();$XL.ActiveWindow.SplitRow=3;$XL.ActiveWindow.SplitColumn=$leading;$XL.ActiveWindow.FreezePanes=$true;$b.SaveAs([IO.Path]::GetFullPath($File),51)
 }finally{$b.Close($false)}}finally{Close-Excel}
 return @{file=$File;students=$students.Count;items=$fields.Count;workspace_unchanged=$true;fingerprint=$q.fingerprint}
 }catch{@{message=$_.Exception.ToString();stack=$_.ScriptStackTrace;position=$_.InvocationInfo.PositionMessage}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence/export-error.json') -Encoding utf8;throw}
}
function CTablePreview{
 if(!(Test-Path -LiteralPath $File) -or (Get-Item -LiteralPath $File).Length -gt 16MB){throw '16MB 이하 Excel 파일을 선택하세요.'}
 $hash=Shared-Hash $File;$book=[NaturalXlsx]::Read([IO.Path]::GetFullPath($File));$identity=@($book.Sheets|Where-Object Name -CEQ '_exchange');$sheets=@($book.Sheets|Where-Object Name -CEQ '성적표')
 if(CIsLeaf){return (CLeafPreview $book $hash)}
 if($identity.Count -ne 1 -or $sheets.Count -ne 1){throw '식별코드가 담긴 다운로드 Excel을 선택하세요.'}
 $rows=$identity[0].Rows;function XMeta($n){if(!$rows.ContainsKey([int]$n)){return ''};return [string](NCell $rows[[int]$n] 1)}
 if((XMeta 1) -notin @('assessment-table-1.0','assessment-table-1.1.1') -or (XMeta 2) -cne $Name -or (XMeta 3) -cne $Layer){throw '다른 프로젝트 또는 작업표의 Excel입니다.'}
 # Codes identify destinations. Export age, display filters and old scores do not block updates.
 $script:Class='';$script:Student='';$q=CSnapshot;$fields=@($q.fields);$studentIds=@($q.data|ForEach-Object {$_.student}|Sort-Object -Unique)
 $codeMap=StudentCodeMap $Root $M.students;$sheet=$sheets[0];$keyRow=if($sheet.Rows.ContainsKey(2) -and @($sheet.Rows[2].Values|Where-Object {$_.Value -eq '@student'}).Count){2}else{4};$firstData=if($keyRow -eq 2){4}else{5};$columns=@{};foreach($c in $sheet.Rows[$keyRow].Keys){$key=[string](NCell $sheet.Rows[$keyRow] $c);if(!$key){continue};if($columns.ContainsKey($key)){throw '중복 항목 식별코드'};$columns[$key]=[int]$c}
 if(!$columns.ContainsKey('@student')){throw '학생 식별코드 열이 없습니다.'}
 $importFields=@($fields|Where-Object {$columns.ContainsKey('value|'+$_)});if(!$importFields.Count){throw '일치하는 평가항목 식별코드가 없습니다.'}
 foreach($k in $columns.Keys){if($k.StartsWith('value|') -and $k.Substring(6) -notin $fields){throw ('알 수 없는 평가항목 코드: '+$k)}}
 $seen=@{};$changes=[Collections.Generic.List[object]]::new();$lookup=@{};foreach($r in $q.data){$lookup[$r.student+'|'+$r.item]=$r}
 foreach($rn in ($sheet.Rows.Keys|Sort-Object|Where-Object {$_ -ge $firstData})){
  $row=$sheet.Rows[$rn];if(!@($row.Values|Where-Object {$null -ne $_.Value -and [string]$_.Value -ne ''}).Count){continue}
  $sid=$codeMap[[string](NCell $row $columns['@student'])];if(!$sid -or $sid -notin $studentIds -or $seen.ContainsKey($sid)){throw '학생 식별코드가 없거나 중복·미등록 코드입니다.'};$seen[$sid]=$true
  foreach($item in $importFields){$old=$lookup[$sid+'|'+$item];$v=NCell $row $columns['value|'+$item];if($v -is [string] -and $v -eq ''){$v=$null}
   if($old.numeric -and $v -is [string]){$d=0.0;if([double]::TryParse($v,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){$v=$d}}
   if(!(Same-Value $v $old.value)){CValidateValue $v '' $old.numeric;$changes.Add(@{student=$sid;name=$old.name;item=$item;label=$old.label;before=$old.value;value=$v;beforeStatus='';status=$(if($null -eq $v){'미입력'}else{'입력됨'});note='';source_sheet='성적표';source_row=$rn;source_column=$columns['value|'+$item]})}
  }
 }
 if((Shared-Hash $File) -ne $hash){throw '읽는 도중 파일이 바뀌었습니다. 다시 반영하세요.'}
 return @{fingerprint=(NHash ($hash+'|'+$q.fingerprint+'|'+$Name+'|'+$Layer));expected_fingerprint=$q.fingerprint;changes=$changes.ToArray();source_hash=$hash;read_only=$true}
}
