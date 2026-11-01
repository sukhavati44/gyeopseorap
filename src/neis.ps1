# Narrow verified NEIS template adapter; local only. Original package parts preserved on export.
function NeisRead {
 if([IO.Path]::GetExtension($File) -ne '.xlsx'){throw 'xlsx 원본을 선택하세요.'}
 $fn=[IO.Path]::GetFileName($File);if($fn -notmatch '(?<year>20\d\d)학년도\s+(?<term>[12])학기\s+주간\s+(?<grade>\d+)학년\s+(?<subject>.+)_전체_(?<class>\d+)강의실'){throw '학년도·학기·학년·과목·학급을 파일명에서 확인할 수 없습니다.'};$info=@{};foreach($k in @('year','term','grade','subject','class')){$info[$k]=$Matches[$k]}
 $book=[NaturalXlsx]::Read([IO.Path]::GetFullPath($File));if($book.Sheets.Count -ne 1){throw '현재 지원하는 단일 시트 나이스 서식이 아닙니다.'};$s=$book.Sheets[0]
 foreach($pair in @(@(1,'과목'),@(2,'반'),@(3,'번호'),@(4,'성명'))){if((NCell $s.Rows[1] $pair[0]) -cne $pair[1]){throw '나이스 학생 머리글을 확인하세요.'}}
 $areas=@();foreach($c in 5..7){$label=[string](NCell $s.Rows[1] $c);if(!$label -or (NCell $s.Rows[3] $c) -ne '점수'){throw '3개 평가영역 점수 서식만 지원합니다.'};$areas+=@{label=$label;column=$c}}
 $people=@();$seen=@{};foreach($rn in @($s.Rows.Keys|Sort-Object|Where-Object {$_ -ge 4})){$row=$s.Rows[$rn];$cls=[string](NCell $row 2);$num=[string](NCell $row 3);$name=[string](NCell $row 4);if(!$cls -and !$num -and !$name){continue};if($cls -ne $info.class -or $num -notmatch '^[1-9]\d*$' -or !$name -or $seen[$num]){throw '학급·번호·이름 또는 중복 번호를 확인하세요.'};$seen[$num]=$true;$values=@();foreach($c in 5..7){if($row[$c].Formula){throw '원본 점수 수식은 자동 반입하지 않습니다.'};$v=NCell $row $c;if($null -ne $v -and [string]$v -ne ''){$n=0.0;if(![double]::TryParse([string]$v,[ref]$n)){throw '점수칸의 숫자를 확인하세요.'};$v=$n}else{$v=$null};$values+=,$v};$people+=@{number=$num;name=$name;row=[int]$rn;values=$values}}
 if(!$people.Count){throw '명렬이 없습니다.'};return @{info=$info;sheet=$s.Name;areas=$areas;people=$people;hash=(Shared-Hash $File)}
}
function NeisAreaItems($profile,$source,$request){
 if($request -and $request.area_items){$items=@($request.area_items);if($items.Count -ne $source.areas.Count -or @($items|Sort-Object -Unique).Count -ne $items.Count -or @($items|Where-Object {$_ -notin @($profile.areas.root_item)}).Count){throw '서로 다른 반입 대상 영역을 원본 열마다 선택하세요.'};return ,$items}
 $items=@();foreach($a in $source.areas){$match=@($profile.areas|Where-Object {$_.label -ceq $a.label});if($match.Count -ne 1){throw '반입할 평가영역을 선택하세요.'};$items+=,$match[0].root_item};return ,$items
}
function NeisImport {
 $n=NeisRead;$p=CProfile;$i=$n.info;foreach($k in @('year','term','grade','subject')){if([string]$p[$k] -cne [string]$i[$k]){throw '프로젝트와 원본의 학년도·학기·학년·과목이 다릅니다.'}}
 if($M.layers.Count -ne 1 -or $Layer -ne $p.root){throw '나이스 최초 명렬 반입은 하위 작업표 생성 전에 종합표에서 진행하세요.'}
 $request=if($PayloadFile){To-Map (Get-Content -LiteralPath $PayloadFile -Raw -Encoding UTF8|ConvertFrom-Json)}else{@{}};$areaItems=NeisAreaItems $p $n $request
 $cls=$i.class+'반';if(!$M.neisSources){$M.neisSources=@{}};if($M.neisSources[$cls]){throw '이미 반입한 반입니다. 기존 반·코드·서식은 유지했습니다.'}
 if(@($M.memberships|Where-Object class -EQ $cls).Count){throw '같은 학급의 기존 명렬이 있습니다. 이름만으로 합치지 않습니다.'}
 $cpRel='backups/neis-'+[guid]::NewGuid().ToString('N');$cp=Resolve-Local $cpRel;New-Item -ItemType Directory $cp|Out-Null;Copy-Item $MetaPath (Join-Path $cp 'relations.json');$bookPath=Resolve-Local $M.layers[$p.root].path;Copy-Item $bookPath (Join-Path $cp 'root.xlsx')
 $copyRel='originals/neis-'+$i.class+'-'+$n.hash.Substring(0,12)+'.xlsx';$copy=Resolve-Local $copyRel;Copy-Item -LiteralPath $File -Destination $copy;if((Shared-Hash $copy) -ne $n.hash){throw '원본 사본이 다릅니다.'}
 @{checkpoint=$cpRel;action='neis_import'}|ConvertTo-Json|Set-Content (Join-Path $Root 'pending-sync.json') -Encoding UTF8
 $mapping=@();$newPeople=@();foreach($person in $n.people){$sid='S'+[guid]::NewGuid().ToString('N');$M.students[$sid]=@{name=$person.name;number=$person.number;code=$sid};$M.memberships+=@{student=$sid;class=$cls;from=($i.year+'-01-01');until=$null};$mapping+=@{student=$sid;row=$person.row;number=$person.number};$newPeople+=@{student=$sid;person=$person}}
 Excel;try{$b=$XL.Workbooks.Open($bookPath);try{$s=$b.Worksheets.Item('Data');$cols=@{};for($c=1;$c -le $s.UsedRange.Columns.Count;$c++){$cols[[string]$s.Cells.Item(1,$c).Value2]=$c};$r=$s.UsedRange.Rows.Count+1;foreach($x in $newPeople){$s.Cells.Item($r,1).Value2=$x.student;$s.Cells.Item($r,2).Value2=[string]$x.person.name;$s.Cells.Item($r,3).Value2=$x.student;for($a=0;$a -lt $n.areas.Count;$a++){$v=$x.person.values[$a];if($null -ne $v){$s.Cells.Item($r,$cols[$areaItems[$a]]).Value2=[double]$v}};$r++};$b.Save()}finally{$b.Close($false)}}finally{Close-Excel}
 $p.classes=@(@($p.classes)+$cls|Sort-Object -Unique);$M.neisSources[$cls]=@{path=$copyRel;hash=$n.hash;sheet=$n.sheet;mapping=$mapping;columns=@($n.areas.column);items=@($areaItems);source_name=[IO.Path]::GetFileName($File)};Save-Meta;Move-Item (Join-Path $Root 'pending-sync.json') (Join-Path $cp 'completed.json')
 return @{class=$cls;new=$newPeople.Count;matched=0;total=$M.students.Count;duplicates=0;source_preserved=$true;checkpoint=$cpRel}
}
function NeisExport {
 $p=CProfile;if(!$M.neisSources -or !$M.neisSources.Count){throw '나이스 원본을 먼저 반입하세요.'};if(Test-Path -LiteralPath $File){throw '새 출력 폴더를 선택하세요.'};$classes=@($M.neisSources.Keys|Sort-Object);if($Class){if($Class -notin $classes){throw '해당 반의 원본이 없습니다.'};$classes=@($Class)}
 $book=[NaturalXlsx]::Read((Resolve-Local $M.layers[$p.root].path));$s=@($book.Sheets|Where-Object Name -EQ 'Data')[0];$cols=@{};foreach($c in $s.Rows[1].Keys){$cols[[string](NCell $s.Rows[1] $c)]=$c};$rows=@{};foreach($r in $s.Rows.Keys){if($r -gt 1){$sid=[string](NCell $s.Rows[$r] $cols.student_id);if(!$sid -or $rows[$sid]){throw '종합표 학생 ID를 확인하세요.'};$rows[$sid]=$s.Rows[$r]}}
 $plans=@();foreach($cls in $classes){$n=$M.neisSources[$cls];$src=Resolve-Local $n.path;if((Shared-Hash $src) -ne $n.hash){throw '보관한 나이스 원본이 바뀌었습니다.'};$changes=@{};foreach($mapRow in $n.mapping){if(!$rows[$mapRow.student]){throw '종합표의 학생이 누락됐습니다.'};for($a=0;$a -lt $n.items.Count;$a++){$v=NCell $rows[$mapRow.student] $cols[$n.items[$a]];if($null -ne $v -and [string]$v -ne ''){$d=0.0;if(![double]::TryParse([string]$v,[ref]$d)){throw '나이스 점수에는 숫자만 출력합니다.'};$changes[([char](64+[int]$n.columns[$a])).ToString()+$mapRow.row]=[string]$d.ToString('R',[Globalization.CultureInfo]::InvariantCulture)}else{$changes[([char](64+[int]$n.columns[$a])).ToString()+$mapRow.row]=''}}};$plans+=@{class=$cls;source=$src;name=$n.source_name;changes=$changes;filter=("A3:G"+(($n.mapping.row|Measure-Object -Maximum).Maximum))}}
 New-Item -ItemType Directory $File|Out-Null;$result=@();foreach($plan in $plans){$dest=Join-Path $File $plan.name;[NeisPackage]::Write($plan.source,$dest,$plan.changes,$plan.filter);$result+=@{class=$plan.class;file=$dest;allowed_cells=$plan.changes.Count}}
 return @{files=$result;originals_unchanged=$true;operational_upload_tested=$false}
}
if(!('NeisPackage' -as [type])){Add-Type -Path (Join-Path $PSScriptRoot 'NeisPackage.cs') -ReferencedAssemblies System.Xml.Linq,System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem}
