# Shared by the local dialog and engine; no Excel process is needed.
function ImportRole($v){
 $text=([string]$v -replace '\s','').ToLowerInvariant()
 switch -Regex ($text){
  '^(@student|학생id|student_id|학생코드|학생식별자|학생식별코드|식별코드|고유번호|code)$' {return 'id'}
  '^(@name|이름|성명|학생명|virtual_name)$' {return 'name'}
  '^(@class|학급|반)$' {return 'class'}
  '^(@number|번호|출석번호)$' {return 'number'}
 };return ''
}
function ImportCell($row,$col){if($row -and $col -and $row.ContainsKey([int]$col)){return $row[[int]$col].Value};return $null}
function ImportColumnName([int]$n){$name='';while($n -gt 0){$n--;$name=[char](65+($n%26))+$name;$n=[math]::Floor($n/26)};return $name}
function ImportLayout($book,$settings=@{}){
 $sheets=@($book.Sheets|Where-Object {$_.Name -notlike '_*'});$sheet=$null;$header=0;$roles=@{}
 foreach($candidate in $sheets){
  if($settings.sheet -and $candidate.Name -cne $settings.sheet){continue}
  foreach($rn in @($candidate.Rows.Keys|Sort-Object|Select-Object -First 50)){
   if($settings.header -and [int]$settings.header -ne $rn){continue}
   if(@($candidate.Rows[$rn].Values|Where-Object {$_.Value -eq '@student'}).Count){continue}
   $rr=@{};foreach($cn in @($candidate.Rows[$rn].Keys|Sort-Object)){$role=ImportRole (ImportCell $candidate.Rows[$rn] $cn);if($role -and !$rr.ContainsKey($role)){$rr[$role]=[int]$cn}}
   if($settings.roles){$rr=$settings.roles}
   if($rr.id -or ($rr.class -and $rr.number -and $rr.name) -or $settings.allow_unmatched){$sheet=$candidate;$header=[int]$rn;$roles=$rr;break}
  };if($sheet){break}
 }
 if(!$sheet){throw '학생코드 또는 반·번호·이름 열을 찾지 못했습니다. 시트·제목 행과 학생 연결 열을 선택하세요.'}
 $keyRow=$null;$first=$header+1
 if($sheet.Rows.ContainsKey($header-1) -and @($sheet.Rows[$header-1].Values|Where-Object {$_.Value -eq '@student'}).Count){$keyRow=$sheet.Rows[$header-1]}
 elseif($sheet.Rows.ContainsKey($first) -and @($sheet.Rows[$first].Values|Where-Object {$_.Value -eq '@student'}).Count){$keyRow=$sheet.Rows[$first];$first++}
 $cols=@();foreach($cn in @($sheet.Rows[$header].Keys|Sort-Object)){
  $label=([string](ImportCell $sheet.Rows[$header] $cn)).Trim();if(!$label -or $label.StartsWith('@')){continue}
  $role=ImportRole $label;$samples=@($sheet.Rows.Keys|Sort-Object|Where-Object {$_ -ge $first}|ForEach-Object {ImportCell $sheet.Rows[$_] $cn}|Where-Object {$null -ne $_ -and [string]$_ -ne ''})
  $textValue=$false;foreach($v in $samples){$number=0.0;if(![double]::TryParse([string]$v,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)){$textValue=$true}}
  $kind=if($textValue -or $label -match '(합계|총점|총합|최종|판정|검수|보류|승인|특이|근거|메모)'){'text'}else{'number'}
  $cols+=@{column=[int]$cn;letter=(ImportColumnName $cn);label=$label;key=[string](ImportCell $keyRow $cn);role=$role;kind=$kind;sample=($samples|Select-Object -First 2) -join ' / '}
 }
 return @{sheet=$sheet;header=$header;first=$first;roles=$roles;columns=$cols;sheets=@($sheets|ForEach-Object Name)}
}
