. (Join-Path $PSScriptRoot 'import-layout.ps1')
function CIsLeaf {if($Layer -eq (CProfile).root){return $false};return !@($M.layers.Values|Where-Object {$_.parent -eq $Layer}).Count -and !@($M.links|Where-Object {$_.target -eq $Layer}).Count}
function CLeafSnapshot {
 $profile=CProfile;$fields=CFields $profile;$path=Resolve-Local (Layer-Info $Layer).path;$fp=NHash ((Shared-Hash $MetaPath)+'|'+(Shared-Hash $path))
 $b=[NaturalXlsx]::Read($path);$identity=@($b.Sheets|Where-Object Name -CEQ '_identity');$data=@($b.Sheets|Where-Object Name -CEQ 'Data')
 if($identity.Count -ne 1 -or (ImportCell $identity[0].Rows[1] 2) -cne $Layer -or $data.Count -ne 1){throw '현재 작업표를 다시 열어 주세요.'}
 $s=$data[0];$cols=@{};foreach($cn in $s.Rows[1].Keys){$cols[[string](ImportCell $s.Rows[1] $cn)]=[int]$cn};$rows=@{};$values=@{}
 foreach($rn in $s.Rows.Keys){if($rn -eq 1){continue};$sid=[string](ImportCell $s.Rows[$rn] $cols.student_id);if(!$sid -or !$M.students.ContainsKey($sid)){continue};if($rows.ContainsKey($sid)){throw '현재 작업표에 중복 학생 행이 있습니다.'};$rows[$sid]=[int]$rn;foreach($item in $fields){$values[$sid+'|'+$item]=ImportCell $s.Rows[$rn] $cols[$item]}}
 if((NHash ((Shared-Hash $MetaPath)+'|'+(Shared-Hash $path))) -cne $fp){throw '작업표가 변경되었습니다. 미리보기를 다시 확인하세요.'}
 return @{profile=$profile;fields=$fields;rows=$rows;values=$values;fingerprint=$fp}
}
function CLeafPreview($book,$hash){
 $script:Class='';$script:Student='';$q=CLeafSnapshot;$p=$q.profile;$fields=@($q.fields);$settings=@{}
 if($PayloadFile){$settings=To-Map (Get-Content -LiteralPath $PayloadFile -Raw -Encoding UTF8|ConvertFrom-Json)}
 $layout=ImportLayout $book $settings;$sheet=$layout.sheet;$roles=$layout.roles;$header=$layout.header
 $issues=[Collections.Generic.List[object]]::new();$columns=@();$newFields=@();$used=@{};$mapping=@()
 foreach($source in $layout.columns){
  $cn=$source.column;if($cn -in @($roles.Values)){continue};$label=$source.label;$item='';$kind=$source.kind
  $explicit=@($settings.columns|Where-Object {$_.column -eq $cn})
  if($settings.columns){if(!$explicit.Count -or $explicit[0].kind -eq 'skip'){continue};$kind=$explicit[0].kind;$item=[string]$explicit[0].target;if($explicit[0].label){$label=[string]$explicit[0].label}}
  elseif($source.role -or $source.label -match '^(합계|총점|총합|total|sum)(\s|\(|$)' -or $source.key -eq '@total'){continue}
  if($item -and $item -notin $fields){throw ('연결할 기존 열을 다시 선택하세요: '+$label)}
  if(!$item -and $source.key.StartsWith('value|') -and $source.key.Substring(6) -in $fields){$item=$source.key.Substring(6)}
  if(!$item){$same=@($fields|Where-Object {(CItemLabel $p $_) -ceq $label});if($same.Count -eq 1){$item=$same[0]}}
  $isNew=!$item;if($isNew){$item='IMP.'+(NHash ($Layer+'|'+$label)).Substring(0,20)}
  if($used[$item]){throw ('같은 작업표 열에 두 번 연결했습니다: '+$label+' · 연결 대상을 구분하세요.')};$used[$item]=$true
  $numeric=if($isNew){$kind -eq 'number'}else{$item -in @($p.numeric_items)}
  $targetLabel=if($isNew){$label}else{CItemLabel $p $item}
  $columns+=@{column=[int]$cn;item=$item;label=$targetLabel;numeric=$numeric;is_new=$isNew}
  $mapping+=@{column=[int]$cn;source=$source.label;target=$targetLabel;item=$item;is_new=$isNew;numeric=$numeric}
  if($isNew){$newFields+=@{item=$item;label=$label;numeric=$numeric}}
 }
 $codeMap=StudentCodeMap $Root $M.students;$resolved=@();$counts=@{};$matches=[Collections.Generic.List[object]]::new()
 foreach($rn in @($sheet.Rows.Keys|Sort-Object|Where-Object {$_ -ge $layout.first})){
  $row=$sheet.Rows[$rn];if(!@($row.Values|Where-Object {$null -ne $_.Value -and [string]$_.Value -ne ''}).Count){continue}
  $code=([string](ImportCell $row $roles.id)).Trim();$sid='';$reason='';$name=([string](ImportCell $row $roles.name)).Trim();$num=([string](ImportCell $row $roles.number)).Trim();$cls=([string](ImportCell $row $roles.class)).Trim()
  if($code){$sid=$codeMap[$code];if(!$sid){$reason='이 프로젝트에 없는 학생코드: '+$code}}
  else{
   if($cls -match '^\d+$'){$cls+='반'};$number=0
   if(!$cls -or !$name -or ![int]::TryParse($num,[ref]$number)){$reason='학생코드 또는 반·번호·이름이 필요합니다.'}
   else{$ids=@($M.students.Keys|Where-Object {$id=$_;$M.students[$id].name -ceq $name -and [int]$M.students[$id].number -eq $number -and @($M.memberships|Where-Object {$_.student -eq $id -and $_.class -eq $cls -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)}).Count});if($ids.Count -eq 1){$sid=$ids[0]}else{$reason='반·번호·이름에 맞는 학생을 한 명으로 찾지 못했습니다.'}}
  }
  if($sid -and !$q.rows.ContainsKey($sid)){$reason='현재 작업표 명부에 없는 학생입니다.'}
  $memberLabel=if($sid){@($M.memberships|Where-Object {$_.student -eq $sid -and $_.from -le $AsOf -and (!$_.until -or $_.until -ge $AsOf)}|ForEach-Object class) -join ','}else{''}
  $entry=@{class=$memberLabel;row=[int]$rn;code=$code;source_name=$name;student=$sid;name=$(if($sid){$M.students[$sid].name}else{''});number=$(if($sid){$M.students[$sid].number}else{''});result=$(if($reason){$reason}else{'연결됨'});method=$(if($code){'학생코드'}else{'반·번호·이름'})}
  $matches.Add($entry);if($reason){$issues.Add(@{row=[int]$rn;reason=$reason;code=$code;source_sheet=$sheet.Name});continue}
  $counts[$sid]=1+$counts[$sid];$resolved+=@{row=[int]$rn;student=$sid;entry=$entry}
 }
 $changes=[Collections.Generic.List[object]]::new();$matched=0;$blankSkipped=0
 foreach($rr in $resolved){
  if($counts[$rr.student] -gt 1){$rr.entry.result='중복 학생 행 · 이 학생은 제외';$issues.Add(@{row=$rr.row;reason=$rr.entry.result;source_sheet=$sheet.Name});continue};$matched++
  foreach($col in $columns){
   $cell=$sheet.Rows[$rr.row][$col.column];$v=$cell.Value
   if($cell.Formula -and $null -eq $v){$issues.Add(@{row=$rr.row;column=$col.column;reason='계산 결과가 없는 수식입니다. Excel에서 저장하거나 값 붙여넣기 후 다시 가져오세요.';source_sheet=$sheet.Name});continue}
   if($null -eq $v -or [string]$v -eq ''){if(!$settings.replace_blanks){$blankSkipped++;continue};$v=$null}
   elseif($col.numeric){$number=0.0;if($v -is [bool] -or ![double]::TryParse([string]$v,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)){$issues.Add(@{row=$rr.row;column=$col.column;reason=('숫자 열에 문자: '+[string]$v);source_sheet=$sheet.Name});continue};$v=$number}
   else{$v=[string]$v}
   $old=$q.values[$rr.student+'|'+$col.item];try{CValidateValue $v '' $col.numeric}catch{$issues.Add(@{row=$rr.row;column=$col.column;reason=$_.Exception.Message});continue}
   if(!(Same-Value $v $old)){$changes.Add(@{student=$rr.student;name=$M.students[$rr.student].name;item=$col.item;label=$col.label;before=$old;value=$v;status='';note='';source_sheet=$sheet.Name;source_row=$rr.row;source_column=$col.column})}
  }
 }
 $keep=@{};foreach($change in $changes){$keep[$change.item]=$true};$newFields=@($newFields|Where-Object {$keep[$_.item]})
 if((Shared-Hash $File) -cne $hash){throw '읽는 동안 파일이 바뀌었습니다. 저장 후 다시 선택하세요.'}
 $configHash=if($PayloadFile){Shared-Hash $PayloadFile}else{''}
 return @{fingerprint=(NHash ($hash+'|'+$q.fingerprint+'|'+$Name+'|'+$Layer+'|'+$configHash));expected_fingerprint=$q.fingerprint;source_hash=$hash;changes=$changes.ToArray();new_fields=$newFields;skipped=$issues.ToArray();leaf_import=$true;read_only=$true;matched_students=$matched;student_rows=$matches.ToArray();column_mapping=$mapping;blank_skipped=$blankSkipped;sheet=$sheet.Name;header=$header}
}
