function CodeTime($value){if(!$value){return '미발급'};return ([datetime]$value).ToString('MMddHHmm')}
function CodeRegistry($root,$regenerate=$false,$required=''){
 if(!$regenerate -and $script:studentCodeProjects[$root] -and (!$required -or $script:studentCodeProjects[$root].codes.ContainsKey($required))){return $script:studentCodeProjects[$root]}
 $data=Get-Content -LiteralPath (Join-Path $root 'relations.json') -Raw -Encoding UTF8|ConvertFrom-Json;$ids=@($data.students.PSObject.Properties.Name)
 if($required -and $required -notin $ids){throw '학생 코드 대상이 명부에 없습니다.'}
 $project=Get-Content -LiteralPath (Join-Path $root 'project.json') -Raw -Encoding UTF8|ConvertFrom-Json;$profile=@($data.courses.PSObject.Properties)[0].Value
 $year=[string]$profile.year;if($year.Length -gt 2){$year=$year.Substring($year.Length-2)};$subject=([string]$profile.subject).Trim();if(!$subject){$subject='평가'};$prefix=$year+'-'+$profile.term+'-'+$profile.grade+'-'+$subject.Substring(0,1)
 $legacyPrefix=([string]$project.label -replace '[^\p{L}\p{N}_-]','');if($legacyPrefix.Length -gt 24){$legacyPrefix=$legacyPrefix.Substring(0,24)}
 $path=Join-Path $root 'student-codes.json';$lock=$null;$temp=$null
 try{
  $lock=[IO.File]::Open((Join-Path $root '.student-codes.lock'),'OpenOrCreate','ReadWrite','None')
  $codes=@{};$aliases=@{};$issued='';$changed='';$batch='';$width=2
  if(Test-Path -LiteralPath $path){$old=Get-Content -LiteralPath $path -Raw -Encoding UTF8|ConvertFrom-Json;$prefix=[string]$old.prefix;$issued=[string]$old.issued_at;$changed=[string]$old.changed_at;$batch=[string]$old.batch_at;$width=[math]::Max(2,[int]$old.width)
   foreach($p in $old.codes.PSObject.Properties){$codes[$p.Name]=@{suffix=[string]$p.Value.suffix;created_at=[string]$p.Value.created_at;changed_at=[string]$p.Value.changed_at}}
   foreach($p in $old.aliases.PSObject.Properties){$aliases[$p.Name]=[string]$p.Value}
  }
  if($regenerate){foreach($sid in $codes.Keys){$aliases[$prefix+'-'+$codes[$sid].suffix]=$sid};$codes=@{}}
  $used=@{};foreach($key in $aliases.Keys){$used[$key]=$true};foreach($entry in $codes.Values){$key=$prefix+'-'+$entry.suffix;if($entry.suffix -cnotmatch '^[a-z0-9]{2,5}$' -or $used.ContainsKey($key)){throw '학생 코드 중복 또는 형식 오류'};$used[$key]=$true}
  $missing=@($ids|Where-Object {!$codes.ContainsKey($_)})
  while(($used.Keys|Where-Object {$_.Length -eq $prefix.Length+1+$width}).Count+$missing.Count -gt [math]::Pow(36,$width)){$width++;if($width -gt 5){throw '사용 가능한 학생 코드 조합이 부족합니다.'}}
  $now=(Get-Date).ToString('o');$rng=[Security.Cryptography.RandomNumberGenerator]::Create();$alphabet='abcdefghijklmnopqrstuvwxyz0123456789'
  try{foreach($sid in $missing){do{$bytes=New-Object byte[] $width;$rng.GetBytes($bytes);$suffix='';foreach($b in $bytes){while($b -ge 252){$one=New-Object byte[] 1;$rng.GetBytes($one);$b=$one[0]};$suffix+=$alphabet[[int]($b%36)]};$key=$prefix+'-'+$suffix}while($used.ContainsKey($key));$used[$key]=$true;$codes[$sid]=@{suffix=$suffix;created_at=$now;changed_at=$now}}}finally{$rng.Dispose()}
  if($missing.Count -or !(Test-Path -LiteralPath $path)){
   if(!$issued){$issued=$now};$changed=$now;$batch=$now;$registry=@{version=2;prefix=$prefix;width=$width;issued_at=$issued;changed_at=$changed;batch_at=$batch;codes=$codes;aliases=$aliases}
   $temp=Join-Path $root ('.student-codes-'+[guid]::NewGuid().ToString('N')+'.tmp');[IO.File]::WriteAllText($temp,($registry|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($true))
   if(Test-Path -LiteralPath $path){[IO.File]::Replace($temp,$path,[Management.Automation.Language.NullString]::Value)}else{[IO.File]::Move($temp,$path)}
  }
  $result=@{prefix=$prefix;width=$width;issued_at=$issued;changed_at=$changed;batch_at=$batch;codes=$codes;aliases=$aliases;legacyPrefix=$legacyPrefix;seed=[string]$project.id};$script:studentCodeProjects[$root]=$result;return $result
 }finally{if($temp -and (Test-Path -LiteralPath $temp)){[IO.File]::Delete($temp)};if($lock){$lock.Dispose()}}
}
function StudentCode($root,$sid){$r=CodeRegistry $root $false $sid;return $r.prefix+'-'+$r.codes[$sid].suffix}
function StudentCodeMap($root,$students){
 $r=CodeRegistry $root;$map=@{};foreach($key in $r.aliases.Keys){$map[$key]=$r.aliases[$key]}
 foreach($sid in $students.Keys){$map[$sid]=$sid;$map[(StudentCode $root $sid)]=$sid;if($students[$sid].code){$map[[string]$students[$sid].code]=$sid};$sha=[Security.Cryptography.SHA256]::Create();try{$tail=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($r.seed+'|'+$sid)))).Replace('-','').Substring(0,12).ToLowerInvariant()}finally{$sha.Dispose()};$map[$r.legacyPrefix+'-'+$tail]=$sid};return $map
}
