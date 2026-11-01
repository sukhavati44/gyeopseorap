# Explicit local trash cleanup. Never traverse outside the displayed trash directory.
function CleanupDirectory($path){
 $full=[IO.Path]::GetFullPath($path).TrimEnd('\');if(!$full -or $full -eq [IO.Path]::GetPathRoot($full).TrimEnd('\')){throw '휴지통 경로가 아닙니다.'}
 $walk=$full;while($walk){if((Test-Path -LiteralPath $walk) -and ((Get-Item -LiteralPath $walk -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw '연결 폴더는 비우지 않습니다.'};$next=Split-Path $walk;if($next -eq $walk){break};$walk=$next}
 return $full
}
function CleanupPlan($req){
 $root=CleanupDirectory $Root;$paths=@();$pending=$null
 if($req.mode -eq 'projects'){
  if((Split-Path $root -Leaf) -cne '.project-trash'){throw '프로젝트 휴지통(.project-trash)만 비울 수 있습니다.'}
  if(Test-Path -LiteralPath $root){$paths=@(Get-ChildItem -LiteralPath $root -Directory -Force|ForEach-Object {$_.FullName})}
 }elseif($req.mode -eq 'workspace'){
  $metaFile=Join-Path $root 'relations.json';$data=To-Map (Get-Content -LiteralPath $metaFile -Raw -Encoding UTF8|ConvertFrom-Json)
  $pendingFile=Join-Path $root 'pending-trash.json';if(Test-Path -LiteralPath $pendingFile){$pending=To-Map (Get-Content -LiteralPath $pendingFile -Raw -Encoding UTF8|ConvertFrom-Json);if($pending.stage -ne 'completed'){throw '이동 중인 휴지통 요청이 있습니다. 해당 이동을 마친 뒤 비우세요.'}}
  foreach($id in $data.trash.Keys){if($id -notmatch '^[a-f0-9]{32}$'){throw '휴지통 항목 경로를 확인하세요.'};$paths+=,(Join-Path $root ('trash/'+$id));$paths+=,(Join-Path $root ('backups/trash-'+$id))}
 }else{throw '휴지통 종류를 선택하세요.'}
 $items=@();foreach($path in $paths){$full=CleanupDirectory $path;if(!$full.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){throw '휴지통 밖 경로입니다.'};if(!(Test-Path -LiteralPath $full)){continue};$files=@(Get-ChildItem -LiteralPath $full -File -Recurse -Force);$links=@(Get-ChildItem -LiteralPath $full -Recurse -Force|Where-Object {$_.Attributes -band [IO.FileAttributes]::ReparsePoint});if($links.Count){throw '휴지통에 연결 폴더가 있어 비우지 않습니다.'};$items+=@{path=$full;files=$files.Count;bytes=[long](($files|Measure-Object Length -Sum).Sum)}}
 return @{mode=$req.mode;root=$root;items=$items;files=[int](($items|ForEach-Object {$_.files}|Measure-Object -Sum).Sum);bytes=[long](($items|ForEach-Object {$_.bytes}|Measure-Object -Sum).Sum);entries=$(if($req.mode -eq 'workspace'){$data.trash.Count}else{$items.Count})}
}
function CleanupCommit($req){
 if(!$req.confirmed){throw '영구 삭제 경고 확인이 필요합니다.'};$plan=CleanupPlan $req
 $expected=@($req.paths|Sort-Object);$actual=@($plan.items|ForEach-Object {$_.path}|Sort-Object);if(($expected -join "`n") -cne ($actual -join "`n")){throw '휴지통 내용이 바뀌었습니다. 다시 확인하세요.'}
 $errors=@();foreach($item in $plan.items){try{$resolved=CleanupDirectory $item.path;if(!$resolved.StartsWith($plan.root+'\',[StringComparison]::OrdinalIgnoreCase)){throw '삭제 경계 밖입니다.'};Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction Stop}catch{$errors+=($_.Exception.Message)}}
 if($req.mode -eq 'workspace'){
  $metaFile=Join-Path $plan.root 'relations.json';$data=To-Map (Get-Content -LiteralPath $metaFile -Raw -Encoding UTF8|ConvertFrom-Json)
  foreach($id in @($data.trash.Keys)){if(!(Test-Path -LiteralPath (Join-Path $plan.root ('trash/'+$id))) -and !(Test-Path -LiteralPath (Join-Path $plan.root ('backups/trash-'+$id)))){$data.trash.Remove($id)}}
  $tmp=$metaFile+'.cleanup.tmp';[IO.File]::WriteAllText($tmp,($data|ConvertTo-Json -Depth 45),[Text.UTF8Encoding]::new($true));[IO.File]::Replace($tmp,$metaFile,[NullString]::Value)
 }
 return @{completed=(!$errors.Count);errors=$errors;removed=$plan.items.Count;remaining=(CleanupPlan $req).entries}
}