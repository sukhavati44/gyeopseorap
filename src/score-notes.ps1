# Optional column scoring notes; separate from student cell annotations.
$script:scoreNoteRoot='';$script:scoreNotes=@{};$script:scoreRules=@{}
function ScoreParse([string]$text){
 $totals=New-Object 'Collections.Generic.HashSet[decimal]';[void]$totals.Add([decimal]0);$lines=0
 foreach($line in ($text -split '\r?\n')){
  if([string]::IsNullOrWhiteSpace($line)){continue};$lines++;$choices=New-Object 'Collections.Generic.HashSet[decimal]'
  foreach($token in ($line -split '[,;，；]')){
   $n=[decimal]0
   if([string]::IsNullOrWhiteSpace($token) -or ![decimal]::TryParse($token.Trim(),[Globalization.NumberStyles]::AllowLeadingSign -bor [Globalization.NumberStyles]::AllowDecimalPoint,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){throw '각 줄에 가능한 숫자를 쉼표 또는 세미콜론으로 구분하세요. 예: 5,4,3'}
   [void]$choices.Add($n)
  }
  $next=New-Object 'Collections.Generic.HashSet[decimal]';$operations=0
  foreach($a in $totals){foreach($b in $choices){$operations++;if($operations -gt 1000000){throw '배점 조합 계산량이 너무 많습니다. 배점 목록을 줄여 주세요.'};[void]$next.Add(($a+$b));if($next.Count -gt 100000){throw '가능한 점수 조합이 너무 많습니다. 배점 목록을 줄여 주세요.'}}};$totals=$next
 }
 if(!$lines){return $null};return [pscustomobject]@{scores=$totals;elements=$lines;display=(@($totals|Sort-Object|Select-Object -First 30|ForEach-Object {$_.ToString([Globalization.CultureInfo]::InvariantCulture)}) -join ', ');count=$totals.Count}
}
function ScoreEnsure($root){
 if($script:scoreNoteRoot -ceq $root){return};$notes=@{};$rules=@{};$path=Join-Path $root 'score-notes.json'
 if(Test-Path -LiteralPath $path){$data=[IO.File]::ReadAllText($path)|ConvertFrom-Json;foreach($entry in @($data.entries)){if($entry.layer -and $entry.item){$key=$entry.layer+'|'+$entry.item;$notes[$key]=$entry;try{$rules[$key]=ScoreParse $entry.text}catch{$rules[$key]=[pscustomobject]@{error=$_.Exception.Message}}}}}
 $script:scoreNotes=$notes;$script:scoreRules=$rules;$script:scoreNoteRoot=$root
}
function ScoreSet($root,$layer,$item,[string]$text){
 ScoreEnsure $root;$rule=ScoreParse $text;$key=$layer+'|'+$item;$next=@{};foreach($k in $script:scoreNotes.Keys){$next[$k]=$script:scoreNotes[$k]}
 if(!$rule){$next.Remove($key)}else{$next[$key]=[pscustomobject]@{layer=$layer;item=$item;text=$text;updated=(Get-Date).ToString('o')}}
 $path=Join-Path $root 'score-notes.json';$temp=Join-Path $root ('.score-notes-'+[guid]::NewGuid().ToString('N')+'.tmp')
 try{[IO.File]::WriteAllText($temp,(@{version=1;entries=@($next.Values)}|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($true));if(Test-Path -LiteralPath $path){[IO.File]::Replace($temp,$path,[System.Management.Automation.Language.NullString]::Value)}else{[IO.File]::Move($temp,$path)};$script:scoreNotes=$next;if($rule){$script:scoreRules[$key]=$rule}else{$script:scoreRules.Remove($key)}}finally{if(Test-Path -LiteralPath $temp){[IO.File]::Delete($temp)}}
}
function ScoreMismatch($rule,$value){
 if(!$rule -or $rule.error -or $null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)){return $false}
 $n=[decimal]0;if(![decimal]::TryParse([string]$value,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){return $false}
 return !$rule.scores.Contains($n)
}
