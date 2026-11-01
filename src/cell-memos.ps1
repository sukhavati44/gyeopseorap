# Cell annotations are independent of score values and workbook calculations.
$script:memoRoot='';$script:memoData=@{}
function MemoEnsure($root){
 if($script:memoRoot -ceq $root){return}
 $data=@{};$path=Join-Path $root 'cell-memos.json'
 if(Test-Path -LiteralPath $path){
  $json=[IO.File]::ReadAllText($path)|ConvertFrom-Json
  foreach($entry in @($json.entries)){if($entry.layer -and $entry.student -and $entry.item){$data[($entry.layer+'|'+$entry.student+'|'+$entry.item)]=$entry}}
 }
 $script:memoData=$data;$script:memoRoot=$root
}
function MemoGet($root,$layer,$student,$item){MemoEnsure $root;return $script:memoData[($layer+'|'+$student+'|'+$item)]}
function MemoSet($root,$layer,$student,$item,$text){
 MemoEnsure $root
 $key=$layer+'|'+$student+'|'+$item;$next=@{};foreach($k in $script:memoData.Keys){$next[$k]=$script:memoData[$k]}
 if([string]::IsNullOrWhiteSpace($text)){$next.Remove($key)}else{$next[$key]=[pscustomobject]@{layer=$layer;student=$student;item=$item;text=$text;updated=(Get-Date).ToString('o')}}
 $path=Join-Path $root 'cell-memos.json';$temp=Join-Path $root ('.cell-memos-'+[guid]::NewGuid().ToString('N')+'.tmp')
 try{
  [IO.File]::WriteAllText($temp,(@{version=1;entries=@($next.Values)}|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($true))
  if(Test-Path -LiteralPath $path){[IO.File]::Replace($temp,$path,[System.Management.Automation.Language.NullString]::Value)}else{[IO.File]::Move($temp,$path)}
  $script:memoData=$next
 }finally{if(Test-Path -LiteralPath $temp){[IO.File]::Delete($temp)}}
}
