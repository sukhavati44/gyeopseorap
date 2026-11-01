Add-Type -AssemblyName System.IO.Compression.FileSystem
if(!('ProjectBook' -as [type])){Add-Type -Path (Join-Path $PSScriptRoot 'ProjectBook.cs') -ReferencedAssemblies System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem}
function PExcelRoster($file){
 if([IO.Path]::GetExtension($file).ToLowerInvariant() -ne '.xlsx'){throw '명부 Excel(.xlsx)을 선택하세요.'}
 if(!('NaturalXlsx' -as [type])){Add-Type -Path (Join-Path $PSScriptRoot 'XlsxReader.cs') -ReferencedAssemblies System.Xml.Linq,System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem}
 $book=[NaturalXlsx]::Read([IO.Path]::GetFullPath($file));$roster=@()
 foreach($sheet in $book.Sheets){$header=0;$cols=@{};foreach($rn in @($sheet.Rows.Keys|Sort-Object|Select-Object -First 30)){$candidate=@{};foreach($cn in $sheet.Rows[$rn].Keys){$v=([string]$sheet.Rows[$rn][$cn].Value).Trim().ToLowerInvariant();$role=switch($v){'이름'{'name'};'성명'{'name'};'학생명'{'name'};'name'{'name'};'반'{'class'};'학급'{'class'};'class'{'class'};'번호'{'number'};'출석번호'{'number'};'number'{'number'}};if($role){if($candidate[$role]){throw '명부 머리글이 중복됩니다.'};$candidate[$role]=$cn}};if($candidate.name -and $candidate.class -and $candidate.number){$header=$rn;$cols=$candidate;break}}
 if(!$header){continue};foreach($rn in @($sheet.Rows.Keys|Sort-Object|Where-Object {$_ -gt $header})){$row=$sheet.Rows[$rn];$values=@{};foreach($k in @('name','class','number')){$cell=$row[[int]$cols[$k]];if($cell.Formula){throw '명부 식별칸의 수식 대신 값을 입력하세요.'};$values[$k]=([string]$cell.Value).Trim()};if(!$values.name -and !$values.class -and !$values.number){continue};if($values.class -match '^\d+$'){$values.class+='반'};$roster+=,$values}
 }
 if(!$roster.Count){throw '이름·반·번호 머리글과 학생 행을 찾지 못했습니다.'};return ,$roster
}
if(!('PHashUtility' -as [type])){Add-Type -TypeDefinition 'using System;using System.IO;using System.Security.Cryptography;public static class PHashUtility{public static string Get(string p){using(var s=File.OpenRead(p))using(var h=SHA256.Create())return BitConverter.ToString(h.ComputeHash(s)).Replace("-","");}}'}
function PHash($p){return [PHashUtility]::Get($p)}
function PPath($basePath,$rel){
 if([IO.Path]::IsPathRooted($rel) -or $rel -match '(^|[\\/])\.\.([\\/]|$)'){throw '프로젝트 밖을 가리키는 경로입니다.'}
 $full=[IO.Path]::GetFullPath((Join-Path $basePath $rel));if(!$full.StartsWith($basePath.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw '프로젝트 경계 밖입니다.'}
 $walk=$full;while($walk){if((Test-Path -LiteralPath $walk) -and ((Get-Item -LiteralPath $walk -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw '연결 폴더는 프로젝트 자료로 사용하지 않습니다.'};$next=Split-Path $walk;if($next -eq $walk){break};$walk=$next};return $full
}
function PInspect($projectRoot){
 $projectRoot=[IO.Path]::GetFullPath($projectRoot).TrimEnd('\');[void](PPath $projectRoot 'relations.json');$r=To-Map (Get-Content -LiteralPath (Join-Path $projectRoot 'relations.json') -Raw -Encoding UTF8|ConvertFrom-Json)
 foreach($id in $r.layers.Keys){$p=PPath $projectRoot ([string]$r.layers[$id].path);if(!(Test-Path -LiteralPath $p -PathType Leaf)){throw "데이터 파일이 없습니다: $($r.layers[$id].path)"}}
 foreach($id in $r.importBatches.Keys){if($id -notmatch '^B[0-9a-f]{20}$' -or !(Test-Path -LiteralPath (PPath $projectRoot ('imports/'+$id)))){throw '반입 원본 폴더가 누락되었습니다.'}}
 $info=if(Test-Path (Join-Path $projectRoot 'project.json')){To-Map (Get-Content (Join-Path $projectRoot 'project.json') -Raw -Encoding UTF8|ConvertFrom-Json)}else{@{id='';label=(Split-Path $projectRoot -Leaf);legacy=$true}}
 return @{root=$projectRoot;project=$info;students=$r.students.Count;layers=$r.layers.Count;pending=((Test-Path (Join-Path $projectRoot 'pending-sync.json')) -or (Test-Path (Join-Path $projectRoot 'pending-import.json')))}
}
function PCreate($req){
 if(Test-Path -LiteralPath $Root){throw '이미 있는 폴더는 바꾸지 않습니다. 새 프로젝트 폴더 이름을 지정하세요.'}
 $parentPath=Split-Path $Root;if(!(Test-Path $parentPath -PathType Container)){throw '프로젝트를 둘 상위 폴더가 없습니다.'};[void](PPath $parentPath (Split-Path $Root -Leaf))
 if(!$req.label -or $req.label.Length -gt 80 -or $req.year -notmatch '^20\d\d$' -or $req.term -notin @('1','2') -or $req.grade -notmatch '^(1[0-2]|[1-9])$' -or !$req.subject){throw '프로젝트명·학년도·학기·학년·과목을 확인하세요.'}
 $areaNames=@($req.areas|ForEach-Object {([string]$_) -split '[,;，；\r\n]'}|ForEach-Object {$_.Trim()}|Where-Object {$_});if(!$areaNames.Count -or $areaNames.Count -gt 12 -or @($areaNames|Sort-Object -Unique).Count -ne $areaNames.Count){throw '서로 다른 평가영역 이름을 1~12개 입력하세요.'}
 $roster=@();if($req.roster -eq 'synthetic120'){
  foreach($n in 1..120){$roster+=@{name=('가상학생{0:000}' -f $n);number=[string](1+(($n-1)%20));class=([string](1+[int][math]::Floor(($n-1)/20))+'반')}}
 }elseif($req.roster -eq 'xlsx'){
  $roster=PExcelRoster $req.roster_file
 }elseif($req.roster -eq 'csv'){
  if(!$req.roster_file -or [IO.Path]::GetExtension($req.roster_file) -ne '.csv'){throw '명부 CSV를 선택하세요.'};$roster=@(Import-Csv -LiteralPath $req.roster_file -Encoding UTF8);if(!$roster.Count -or (@($roster[0].PSObject.Properties.Name) -join ',') -ne 'name,class,number'){throw '명부 CSV 열은 name,class,number 순서여야 합니다.'}
 }elseif($req.roster -ne 'empty'){throw '빈 프로젝트 또는 명부 종류를 선택하세요.'}
 $seen=@{};foreach($s in $roster){$key=[string]$s.class+'|'+[string]$s.number;if(!$s.name -or !$s.class -or [string]$s.number -notmatch '^[1-9]\d*$' -or $seen[$key]){throw '명부의 성명·반·번호 누락 또는 같은 반 번호 중복을 확인하세요.'};$seen[$key]=$true}
 $finalRoot=$Root;$temp=Join-Path $parentPath ('.creating-'+[guid]::NewGuid().ToString('N'));$script:Root=$temp;$script:MetaPath=Join-Path $temp 'relations.json'
 foreach($d in @('','books','imports','originals','backups','logs')){New-Item -ItemType Directory (Join-Path $temp $d) -Force|Out-Null}
 $script:M=@{version=1;students=@{};memberships=@();layers=@{};links=@();provenance=@{};labels=@{};participation=@{};manualOverrides=@{};importRecords=@{};importBatches=@{};reviewItems=@{};courses=@{};workspaceScopes=@{}}
 foreach($s in $roster){$sid='S'+[guid]::NewGuid().ToString('N');$M.students[$sid]=@{name=[string]$s.name;number=[string]$s.number;code=$sid};$M.memberships+=@{student=$sid;class=[string]$s.class;from=([string]$req.year+'-01-01');until=$null}}
 $projectId='P'+[guid]::NewGuid().ToString('N');$courseId='C'+[guid]::NewGuid().ToString('N');$areas=@();$fields=@();$numeric=@();for($i=0;$i -lt $areaNames.Count;$i++){$prefix='E'+($i+1);$item=$prefix+'.score';$fields+=@($item,($prefix+'.status'));$areas+=@{label=$areaNames[$i];root_item=$item;work_layer='';work_items=@()};$numeric+=,$item;$M.labels["L1|$item"]=$areaNames[$i]}
 Save-Meta;[ProjectBook]::Create((Join-Path $Root 'books/L1.xlsx'),'L1',[string[]]$fields,$M.students);$M.layers.L1=@{path='books/L1.xlsx';parent='';items=@($fields);name='L1'};$M.layers.L1.name='수행평가 종합표';$M.courses[$courseId]=@{id=$courseId;year=[string]$req.year;term=[string]$req.term;grade=[string]$req.grade;subject=[string]$req.subject;root='L1';classes=@($roster|ForEach-Object {$_.class}|Sort-Object -Unique);areas=$areas;numeric_items=$numeric};Save-Meta
 $project=@{schema=1;id=$projectId;label=[string]$req.label;created_at=(Get-Date -Format o);synthetic=($req.roster -eq 'synthetic120');roster_mode=$req.roster;paths=@{data='books';relations='relations.json';originals='originals';imports='imports';backups='backups';logs='logs'}}
 [IO.File]::WriteAllText((Join-Path $temp 'project.json'),($project|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($true));if($req.roster -in @('csv','xlsx')){Copy-Item -LiteralPath $req.roster_file -Destination (Join-Path $temp ('originals/roster.'+$req.roster))}
 [void](PInspect $temp);[IO.Directory]::Move($temp,$finalRoot);$script:Root=$finalRoot;$script:MetaPath=Join-Path $Root 'relations.json';return @{root=$Root;project_id=$projectId;course_id=$courseId;students=$roster.Count;values_assigned=0;participation_assigned=0}
}
function PFiles($basePath){return @(Get-ChildItem $basePath -File -Recurse -Force|Where-Object {$rel=$_.FullName.Substring($basePath.Length+1);$rel -notmatch '^(backups|recovered-[^\\]+)[\\/]' -and $rel -ne '.manager.lock'})}
function PBackup{
 $info=PInspect $Root;if($info.pending){throw '저장 중단 표시가 있습니다. 기존 체크포인트 복구를 먼저 확인하세요.'};if(@(Get-ChildItem $Root -Filter '~$*.xlsx' -Recurse -Force).Count){throw 'Excel을 저장하고 닫은 뒤 백업하세요.'}
 $lock=$null;try{$lock=[IO.File]::Open((Join-Path $Root '.manager.lock'),'OpenOrCreate','ReadWrite','None');$files=PFiles $Root;$hashes=@{};foreach($f in $files){$rel=$f.FullName.Substring($Root.Length+1);[void](PPath $Root $rel);$hashes[$rel]=PHash $f.FullName}
  $dest=Join-Path $Root ('backups/project-'+[guid]::NewGuid().ToString('N'));New-Item -ItemType Directory $dest -Force|Out-Null
  foreach($rel in $hashes.Keys){$d=PPath $dest $rel;New-Item -ItemType Directory (Split-Path $d) -Force|Out-Null;Copy-Item -LiteralPath (PPath $Root $rel) -Destination $d;if((PHash $d) -ne $hashes[$rel] -or (PHash (PPath $Root $rel)) -ne $hashes[$rel]){throw '백업 중 원본이 바뀌었습니다. 이 백업은 미완료입니다.'}}
  if(@(PFiles $Root).Count -ne $files.Count){throw '백업 중 파일 구성이 바뀌었습니다.'}
  @{schema=1;project_id=$info.project.id;created_at=(Get-Date -Format o);files=$hashes}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $dest 'project-backup.json') -Encoding UTF8;return @{backup=$dest;files=$hashes.Count;verified=$true}
 }finally{if($lock){$lock.Dispose()}}
}
function PRestore{
 if(Test-Path -LiteralPath $Root){throw '복원 대상은 아직 없는 새 프로젝트 폴더여야 합니다.'};$src=[IO.Path]::GetFullPath($File).TrimEnd('\');$manifest=To-Map (Get-Content (Join-Path $src 'project-backup.json') -Raw -Encoding UTF8|ConvertFrom-Json)
 if(!$manifest.files -or $manifest.schema -ne 1){throw '완료된 프로젝트 백업 목록이 없습니다.'};[void](PPath (Split-Path $Root) (Split-Path $Root -Leaf));foreach($rel in $manifest.files.Keys){$p=PPath $src $rel;if((PHash $p) -ne $manifest.files[$rel]){throw "백업 파일이 변경되었습니다: $rel"}}
 $tmp=Join-Path (Split-Path $Root) ('.restoring-'+[guid]::NewGuid().ToString('N'));New-Item -ItemType Directory $tmp|Out-Null
 foreach($rel in $manifest.files.Keys){$dst=PPath $tmp $rel;New-Item -ItemType Directory (Split-Path $dst) -Force|Out-Null;Copy-Item -LiteralPath (PPath $src $rel) -Destination $dst;if((PHash $dst) -ne $manifest.files[$rel]){throw '복원 사본 해시가 다릅니다.'}}
 [void](PInspect $tmp);foreach($d in @('backups','logs','originals','imports')){New-Item -ItemType Directory (Join-Path $tmp $d) -Force|Out-Null};[IO.Directory]::Move($tmp,$Root);return @{root=$Root;files=$manifest.files.Count;verified=$true;original_unchanged=$true}
}
function PArchive($req){
 if(!$req.confirmed){throw '프로젝트 삭제 경고를 확인하세요.'}
 $target=[IO.Path]::GetFullPath($Root).TrimEnd('\');$parent=Split-Path $target
 if(!$parent -or $target -eq [IO.Path]::GetPathRoot($target).TrimEnd('\') -or $PSScriptRoot.StartsWith($target+'\',[StringComparison]::OrdinalIgnoreCase)){throw '프로그램 또는 드라이브 폴더는 삭제할 수 없습니다.'}
 if($req.selected_path){if(![string]::Equals([IO.Path]::GetFullPath($req.selected_path).TrimEnd('\'),$target,[StringComparison]::OrdinalIgnoreCase)){throw '선택한 경로가 다릅니다.'}}
 function PRecentMatch($entries){foreach($e in $entries){if($e.path -and [string]::Equals([IO.Path]::GetFullPath([string]$e.path).TrimEnd('\'),$target,[StringComparison]::OrdinalIgnoreCase)){return $true};if($e.value -and (PRecentMatch $e.value)){return $true}};return $false}
 $recentFile=Join-Path $PSScriptRoot 'recent-projects.json';$registered=$false
 if(Test-Path -LiteralPath $recentFile){$registered=PRecentMatch (Get-Content -LiteralPath $recentFile -Raw -Encoding UTF8|ConvertFrom-Json)}
 if(!$registered -and !(Test-Path -LiteralPath (Join-Path $target 'project.json'))){throw '최근 목록에 등록되었거나 프로젝트 표시 파일이 있는 폴더만 보관할 수 있습니다.'}
 if(!(Test-Path -LiteralPath $target)){return @{removed_from_active=$true;permanently_deleted=$false;original_path=$target;retained_path=''}}
 if(!(Test-Path -LiteralPath $target -PathType Container)){throw '프로젝트 폴더를 선택하세요.'}
 foreach($entry in @(Get-Item -LiteralPath $target -Force)+@(Get-ChildItem -LiteralPath $target -Force -Recurse)){if($entry.Attributes -band [IO.FileAttributes]::ReparsePoint){throw '연결된 폴더가 있어 이동하지 않았습니다.'}}
 if(@(Get-ChildItem -LiteralPath $target -Filter '~$*.xlsx' -Recurse -Force).Count){throw 'Excel 파일을 닫은 뒤 삭제하세요.'}
 $lock=$null;try{$lock=[IO.File]::Open((Join-Path $target '.manager.lock'),'OpenOrCreate','ReadWrite','None')}catch{throw '다른 작업이 저장 중입니다. 완료 후 다시 시도하세요.'}finally{if($lock){$lock.Dispose()}}
 $trash=Join-Path $parent '.project-trash';$destination=Join-Path $trash ((Split-Path $target -Leaf)+'-'+[guid]::NewGuid().ToString('N'))
 if(!([IO.Path]::GetFullPath($destination).StartsWith($parent.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase))){throw '보관 위치를 확인하세요.'}
 if(Test-Path -LiteralPath $trash){if((Get-Item -LiteralPath $trash -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw '보관 위치가 연결 폴더입니다.'}}else{[void][IO.Directory]::CreateDirectory($trash)}
 [IO.Directory]::Move($target,$destination)
 return @{removed_from_active=$true;permanently_deleted=$false;original_path=$target;retained_path=$destination;project_id=''}
}
