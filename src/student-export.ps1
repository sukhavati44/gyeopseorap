# Read-only extraction. Never calls code issuance, recovery, workbook save, or Save-Meta.
function EMap($o){if($null -eq $o){return $null};if($o -is [Management.Automation.PSCustomObject]){$h=@{};foreach($p in $o.PSObject.Properties){$h[$p.Name]=EMap $p.Value};return $h};if($o -is [array]){return ,@($o|ForEach-Object {EMap $_})};return $o}
function EHash($path){$s=[IO.File]::Open($path,'Open','Read','ReadWrite');$h=[Security.Cryptography.SHA256]::Create();try{return [BitConverter]::ToString($h.ComputeHash($s))}finally{$h.Dispose();$s.Dispose()}}
function EPath($root,$rel){$p=[IO.Path]::GetFullPath((Join-Path $root $rel));if(!$p.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){throw '프로젝트 밖 연결 자료는 추출하지 않았습니다.'};$walk=$p;while($walk.Length -gt $root.Length){if((Get-Item -LiteralPath $walk -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw '연결 폴더는 허용하지 않습니다.'};$walk=Split-Path $walk};return $p}
function EReadJson($path){$script:eHashes[$path]=EHash $path;return EMap ([IO.File]::ReadAllText($path)|ConvertFrom-Json)}
function EVerify{foreach($path in $script:eHashes.Keys){if((EHash $path) -cne $script:eHashes[$path]){throw '내보내는 동안 원본이 바뀌었습니다. 생성 결과를 사용하지 말고 다시 실행하세요.'}}}
function ESheet($name,$headers){return @{name=$name;headers=$headers;rows=[Collections.Generic.List[object]]::new()}}
function EAdd($sheet,$prefix,$text){
 $value=if($null -eq $text){''}elseif($text -is [bool]){if($text){'true'}else{'false'}}else{[Convert]::ToString($text,[Globalization.CultureInfo]::InvariantCulture)}
 $chunks=[Collections.Generic.List[string]]::new();$i=0
 do{$len=[math]::Min(30000,$value.Length-$i);if($len -gt 0 -and $i+$len -lt $value.Length -and [char]::IsHighSurrogate($value[$i+$len-1])){$len--};$chunks.Add($value.Substring($i,$len));$i+=$len}while($i -lt $value.Length)
 $part=0;foreach($chunk in $chunks){$part++;$sheet.rows.Add([object[]](@($prefix)+@([string]$part,[string]$chunks.Count,$chunk)))}
}
function EFind($obj,$sid,$path,$sheet){
 if($null -eq $obj){return}
 if($obj -is [Collections.IDictionary]){
  if($obj.student -eq $sid -or $obj.student_id -eq $sid){EAdd $sheet @($path,'학생 연결 기록') ($obj|ConvertTo-Json -Depth 100 -Compress);return}
  foreach($k in @($obj.Keys|Sort-Object)){if($k -eq $sid -or ([string]$k).Contains('|'+$sid+'|')){EAdd $sheet @($path+'/'+$k,'학생 연결 기록') ($obj[$k]|ConvertTo-Json -Depth 100 -Compress)}else{EFind $obj[$k] $sid ($path+'/'+$k) $sheet}}
 }elseif($obj -is [array]){for($n=0;$n -lt $obj.Count;$n++){EFind $obj[$n] $sid ($path+'/'+$n) $sheet}}
}
function ELayerPath($id,$layers){$parts=@();$seen=@{};while($id){if($seen[$id]){throw '작업표 관계 순환'};if(!$layers[$id]){throw '상위 작업표 누락'};$seen[$id]=$true;$parts=@([string]$layers[$id].name)+$parts;$id=[string]$layers[$id].parent};return ($parts -join ' > ')}
function ExportStudents($root,$req){
 $ErrorActionPreference='Stop';$root=[IO.Path]::GetFullPath($root).TrimEnd('\');$script:eHashes=@{}
 if($req.scope -notin @('one','class') -or $req.identity -notin @('name','both','code')){throw '생성 범위와 식별정보를 선택하세요.'}
 foreach($marker in @('pending-sync.json','pending-import.json')){if(Test-Path -LiteralPath (Join-Path $root $marker)){throw '미완료 저장 자료가 있습니다. 원본은 변경하지 않았습니다.'}}
 $m=EReadJson (Join-Path $root 'relations.json');$project=EReadJson (Join-Path $root 'project.json')
 $codePath=Join-Path $root 'student-codes.json';$registry=if(Test-Path -LiteralPath $codePath){EReadJson $codePath}else{$null}
 $ids=if($req.scope -eq 'one'){@([string]$req.student)}else{@($m.memberships|Where-Object {$_.class -eq $req.class -and $_.from -le $req.as_of -and (!$_.until -or $_.until -ge $req.as_of)}|ForEach-Object {$_.student}|Sort-Object -Unique)}
 if(!$ids.Count -or @($ids|Where-Object {!$m.students.ContainsKey($_)}).Count){throw '내보낼 학생 또는 학급을 확인하세요.'}
 if($req.identity -ne 'name' -and @($ids|Where-Object {!$registry -or !$registry.codes[$_]}).Count){throw '먼저 프로젝트 화면에서 학생 코드를 발급하세요. 내보내기에서는 코드를 새로 발급하지 않습니다.'}
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 if(!('NaturalXlsx' -as [type])){Add-Type -Path (Join-Path $PSScriptRoot 'XlsxReader.cs') -ReferencedAssemblies System.Xml.Linq,System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem}
 if(!('StudentArchive' -as [type])){Add-Type -Path (Join-Path $PSScriptRoot 'StudentArchive.cs') -ReferencedAssemblies System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem}
 $books=@{};foreach($lid in @($m.layers.Keys|Sort-Object)){$path=EPath $root $m.layers[$lid].path;$script:eHashes[$path]=EHash $path;$b=[NaturalXlsx]::Read($path);$data=@($b.Sheets|Where-Object Name -EQ 'Data');if($data.Count -ne 1){throw ('Data 시트 누락: '+$lid)};$sheet=$data[0];$heads=@{};foreach($c in $sheet.Rows[1].Keys){$heads[[string]$sheet.Rows[1][$c].Value]=[int]$c};if(!$heads.student_id){throw ('학생 연결 열 누락: '+$lid)};$rows=@{};foreach($r in $sheet.Rows.Keys){if($r -eq 1){continue};$sid=[string]$sheet.Rows[$r][$heads.student_id].Value;if($sid){if($rows.ContainsKey($sid)){throw ('중복 학생 행: '+$lid)};$rows[$sid]=[int]$r}};$books[$lid]=@{sheet=$sheet;heads=$heads;rows=$rows;path=$path;formulas=[StudentArchive]::Formulas($path)}}
 $notes=@{};foreach($file in @('cell-memos.json','score-notes.json')){$path=Join-Path $root $file;if(Test-Path -LiteralPath $path){$notes[$file]=EReadJson $path}}
 $extras=@{};foreach($entry in @($m.trash.Values)){$path=EPath $root ($entry.archive+'/records.json');$extras[('휴지통/'+$entry.id)]=EReadJson $path}
 foreach($dir in @('imports','logs','transactions')){if(Test-Path -LiteralPath (Join-Path $root $dir)){foreach($file in Get-ChildItem -LiteralPath (Join-Path $root $dir) -Recurse -File -Filter '*.json'){$rel=$file.FullName.Substring($root.Length+1);$extras[$rel]=EReadJson (EPath $root $rel)}}}
 foreach($dir in @('imports','logs','transactions')){if(Test-Path -LiteralPath (Join-Path $root $dir)){foreach($file in Get-ChildItem -LiteralPath (Join-Path $root $dir) -Recurse -File|Where-Object {$_.Extension -in @('.jsonl','.ndjson')}){$rel=$file.FullName.Substring($root.Length+1);$path=EPath $root $rel;$script:eHashes[$path]=EHash $path;$extras[$rel]=@([IO.File]::ReadAllLines($path)|Where-Object {![string]::IsNullOrWhiteSpace($_)}|ForEach-Object {EMap ($_|ConvertFrom-Json)})}}}
 $dest=[IO.Path]::GetFullPath([string]$req.output);if(!(Test-Path -LiteralPath $dest -PathType Container)){throw '출력 폴더를 선택하세요.'};if($dest -eq $root -or $dest.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){throw '원본 프로젝트 밖의 출력 폴더를 선택하세요.'}
 $batch=Join-Path $dest ('학생기록_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,8));[void][IO.Directory]::CreateDirectory($batch)
 $result=[Collections.Generic.List[object]]::new()
 foreach($sid in $ids){$temp='';try{
  $person=$m.students[$sid];$code=if($registry -and $registry.codes[$sid]){$registry.prefix+'-'+$registry.codes[$sid].suffix}else{''}
  $identity=if($req.identity -eq 'code'){$code}elseif($req.identity -eq 'both'){[string]$person.name+'_'+$code}else{[string]$person.name}
  $classes=@($m.memberships|Where-Object {$_.student -eq $sid -and $_.from -le $req.as_of -and (!$_.until -or $_.until -ge $req.as_of)}|ForEach-Object {[string]$_.class}|Sort-Object -Unique)
  $classLabel=if($classes.Count){$classes -join ', '}else{'학급 미지정'}
  $info=ESheet '학생정보' @('구분','값');$info.rows.Add([object[]]@('식별정보',$identity));$info.rows.Add([object[]]@('조회 기준일',[string]$req.as_of));$info.rows.Add([object[]]@('범위','전체 작업표 · 조회 선택/깊이와 무관'))
  $info.rows.Add([object[]]@('학급',$classLabel))
  if($req.identity -ne 'code'){$info.rows.Add([object[]]@('이름',[string]$person.name));$info.rows.Add([object[]]@('번호',[string]$person.number))};if($req.identity -ne 'name'){$info.rows.Add([object[]]@('코드',$code))}
  $info.rows.Add([object[]]@('주의','원문 내용·메모에 포함된 실명은 자동 비식별화하지 않습니다. 긴 내용은 바로 다음 행에 이어집니다.'))
  $values=ESheet '전체항목' @('작업표 경로','작업표 ID','상위 작업표 ID','상위 연결 항목','항목 ID','항목 이름','값 종류','출처 파일','출처 셀','수식 여부','원본 수식','분할','분할 수','원본 내용')
  $memo=ESheet '메모·배점' @('종류','작업표 경로','작업표 ID','항목 ID','수정일','분할','분할 수','원본 내용')
  $links=ESheet '연결' @('종류','출발 작업표','출발 항목','도착 작업표','도착 항목','분할','분할 수','원본 내용')
  $other=ESheet '기타기록' @('원본 위치','구분','분할','분할 수','원본 내용')
  $manifest=ESheet '추출명세' @('구분','대상','결과')
  foreach($lid in @($books.Keys|Sort-Object)){
   $book=$books[$lid];$li=$m.layers[$lid];$path=ELayerPath $lid $m.layers;$scope=$m.workspaceScopes[$lid]
   $manifest.rows.Add([object[]]@('작업표',$lid,$(if($book.rows.ContainsKey($sid)){'학생 행 포함'}else{'이 학생의 행 없음'})))
   if(!$book.rows.ContainsKey($sid)){continue};$r=$book.rows[$sid]
   foreach($item in @($book.heads.Keys|Where-Object {$_ -notin @('student_id','virtual_name','code')}|Sort-Object {$book.heads[$_]})){
    $c=$book.heads[$item];$cell=$book.sheet.Rows[$r][$c];$label=[string]$m.labels[$lid+'|'+$item];if(!$label){$label=$item};$v=$cell.Value;$exportValue=$v;if($v -is [double] -and $book.formulas.ContainsKey('raw:Data!'+$cell.Address)){$exportValue=[string]$book.formulas['raw:Data!'+$cell.Address]};$type=if($null -eq $v){'빈값'}else{$v.GetType().Name}
    EAdd $values @($path,$lid,[string]$li.parent,[string]$scope.parent_item,$item,$label,$type,[string]$li.path,[string]$cell.Address,[string][bool]$cell.Formula,[string]$book.formulas['Data!'+$cell.Address]) $exportValue
   }
   $studentScope=@{};if($scope){foreach($k in $scope.Keys){if($k -eq 'target_students'){$studentScope[$k]=@($scope[$k]|Where-Object {$_ -eq $sid})}else{$studentScope[$k]=$scope[$k]}}}
   EAdd $links @('작업표 연결',$lid,'',[string]$li.parent,[string]$scope.parent_item) (@{layer=$li;scope=$studentScope}|ConvertTo-Json -Depth 100 -Compress)
  }
  $studentLayers=@{};foreach($lid in $books.Keys){if($books[$lid].rows.ContainsKey($sid)){$studentLayers[$lid]=$true}}
  foreach($archiveEntry in @($m.trash.Values)){
   if(@($extras['휴지통/'+$archiveEntry.id]|Where-Object {$_.student -eq $sid -or $_.student_id -eq $sid}).Count){
    $studentLayers[[string]$archiveEntry.layer]=$true
    EAdd $links @('보관 자료 연결',('휴지통/'+$archiveEntry.id),[string]$archiveEntry.item,[string]$archiveEntry.layer,'') ($archiveEntry|ConvertTo-Json -Depth 100 -Compress)
   }
  }
  foreach($file in $notes.Keys){foreach($entry in @($notes[$file].entries)){if($entry.student -eq $sid -or ($file -eq 'score-notes.json' -and $studentLayers[[string]$entry.layer])){$path=if($m.layers[$entry.layer]){ELayerPath $entry.layer $m.layers}else{'보관된 항목'};EAdd $memo @($(if($file -eq 'score-notes.json'){'열 배점'}else{'학생 셀 메모'}),$path,[string]$entry.layer,[string]$entry.item,[string]$entry.updated) $entry.text}}}
  foreach($link in @($m.links)){EAdd $links @('항목 연결',[string]$link.source,[string]$link.sourceItem,[string]$link.target,[string]$link.targetItem) ($link|ConvertTo-Json -Depth 100 -Compress)}
  foreach($k in @($m.Keys|Where-Object {$_ -notin @('students','layers','labels','links','courses','neisSources','trash','workspaceScopes')})){EFind $m[$k] $sid ('relations.json/'+$k) $other}
  foreach($k in $extras.Keys){EFind $extras[$k] $sid $k $other}
  foreach($k in @('courses','labels','neisSources','detachedAreas')){if($m[$k]){EAdd $other @(('relations.json/'+$k),'공통 연결 정보') ($m[$k]|ConvertTo-Json -Depth 100 -Compress)}}
  $manifest.rows.Add([object[]]@('추출','전체항목',[string]$values.rows.Count));$manifest.rows.Add([object[]]@('추출','메모·배점',[string]$memo.rows.Count));$manifest.rows.Add([object[]]@('추출','기타기록',[string]$other.rows.Count));$manifest.rows.Add([object[]]@('보존','원본 파일','읽기 전용 · 변경하지 않음'));$manifest.rows.Add([object[]]@('범위','백업 폴더','중복 백업 파일 자체는 복제하지 않음 · 현행·보관 기록과 기존 이력은 포함'))
  $safe=[regex]::Replace(($classLabel+'_'+$identity),'[<>:"/\\|?*\x00-\x1f]','_').Trim().TrimEnd('.');if(!$safe){$safe='학생'};if($safe.Length -gt 100){$safe=$safe.Substring(0,100)}
  $file=Join-Path $batch ($safe+'_'+[guid]::NewGuid().ToString('N').Substring(0,8)+'.xlsx');$temp=$file+'.partial'
  $readable=ESheet '학생 기록' @('영역 · 작업표','평가항목','점수','내용 · 판단 근거','메모','배점')
  $groups=@{};$order=[Collections.Generic.List[string]]::new()
  foreach($row in $values.rows){
   if([string]$row[4] -match '(^|[._])status$'){continue}
   $key=[string]$row[1]+'|'+[string]$row[4]
   $label=[string]$row[5];if($label -eq [string]$row[4]){$label='항목 '+[string]$books[[string]$row[1]].heads[[string]$row[4]]}
   $numeric=([string]$row[6] -in @('Double','Decimal','Int32','Int64','Single'))
   if(!$groups.ContainsKey($key)){$groups[$key]=@{path=[string]$row[0];label=$label;numeric=$numeric;values=[Collections.Generic.List[string]]::new();memos=[Collections.Generic.List[string]]::new();scores=[Collections.Generic.List[string]]::new()};$order.Add($key)}
   $groups[$key].values.Add([string]$row[13])
  }
  foreach($row in $memo.rows){
   $key=[string]$row[2]+'|'+[string]$row[3]
   if(!$groups.ContainsKey($key)){$label=[string]$m.labels[$key];if(!$label){$label=if([string]$row[3] -eq '__total__'){'합계'}else{'이름 없는 항목'}};$groups[$key]=@{path=[string]$row[1];label=$label;numeric=$false;values=[Collections.Generic.List[string]]::new();memos=[Collections.Generic.List[string]]::new();scores=[Collections.Generic.List[string]]::new()};$order.Add($key)}
   if([string]$row[0] -eq '열 배점'){$groups[$key].scores.Add([string]$row[7])}else{$groups[$key].memos.Add([string]$row[7])}
  }
  foreach($key in $order){$g=$groups[$key];$count=[math]::Max($g.values.Count,[math]::Max($g.memos.Count,$g.scores.Count));for($i=0;$i -lt $count;$i++){$value=if($i -lt $g.values.Count){$g.values[$i]}else{''};$label=$g.label+$(if($i -gt 0){' (계속)'}else{''});$readable.rows.Add([object[]]@($g.path,$label,$(if($g.numeric){$value}else{''}),$(if($g.numeric){''}else{$value}),$(if($i -lt $g.memos.Count){$g.memos[$i]}else{''}),$(if($i -lt $g.scores.Count){$g.scores[$i]}else{''})))}}
  [StudentArchive]::Write($temp,[object[]]@($info,$readable));EVerify;[IO.File]::Move($temp,$file);$result.Add(@{target=$identity;file=$file;status='완료';items=$readable.rows.Count;memos=$memo.rows.Count})
 }catch{$result.Add(@{target=$identity;status='실패';error=$_.Exception.Message;partial=$temp})}}
 $manifestFile=Join-Path $batch '생성결과.json';$out=@{directory=$batch;results=$result.ToArray();success=@($result|Where-Object status -EQ '완료').Count;failed=@($result|Where-Object status -EQ '실패').Count;scope=$req.scope;identity=$req.identity;source_unchanged=$false}
 try{EVerify;$out.source_unchanged=$true}catch{$out.error=$_.Exception.Message;$out.failed=$ids.Count;$out.success=0}
 [IO.File]::WriteAllText($manifestFile,($out|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($true));return $out
}
