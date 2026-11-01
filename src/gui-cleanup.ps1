function UCleanupDialog($mode='projects'){
 $d=New-Object Windows.Forms.Form;$d.Text='휴지통 비우기 · 영구 삭제';$d.Size=New-Object Drawing.Size(940,650);$d.StartPosition='CenterParent';$d.Font=$f.Font
 $target=if($mode -eq 'workspace'){$ws}elseif($ws){Join-Path (Split-Path $ws) '.project-trash'}elseif($nRecent.CurrentRow -and $nRecent.CurrentRow.Tag){Join-Path (Split-Path ([string]$nRecent.CurrentRow.Tag)) '.project-trash'}else{''}
 $state=@{target=$target;plan=$null}
 [void](Label $d $(if($mode -eq 'workspace'){'현재 프로젝트의 항목·작업공간 휴지통과 해당 이동의 보관 사본을 비웁니다.'}else{'선택한 위치의 .project-trash만 비웁니다. 다른 폴더와 사용 중인 프로젝트는 대상이 아닙니다.'}) 20 15 880 45)
 $where=TextBox $d 20 65 880;$where.ReadOnly=$true;$where.Text=$target
 $summary=Label $d '내용 확인을 누르면 항목 수와 용량을 표시합니다.' 20 110 880 35
 $grid=Grid $d 20 155 880 265 @('보관 폴더','파일 수','용량(MB)');$grid.Columns[0].FillWeight=300
 $empty=DButton $d '휴지통 완전히 비우기' 620 480 280 {DSafe {
  if(!$state.plan){return};$warning='다음 휴지통을 영구 삭제합니다.'+"`r`n"+$state.target+"`r`n"+$state.plan.entries+'항목 / '+$state.plan.files+'파일 / '+[math]::Round($state.plan.bytes/1MB,2)+" MB`r`n`r`n삭제 후 이 보관함에서 열람·복구할 수 없습니다. 계속하시겠습니까?"
  if([Windows.Forms.MessageBox]::Show($d,$warning,'휴지통 비우기 · 복구 불가','YesNo','Warning','Button2') -ne 'Yes'){return}
  $empty.Enabled=$false;$req=@{mode=$mode;confirmed=$true;paths=@($state.plan.items|ForEach-Object {$_.path})}
  try{$r=(Run 'cleanupcommit' @{PayloadFile=(NConfigFile $req)} $state.target)|ConvertFrom-Json;$grid.Rows.Clear();$state.plan=$null;$summary.Text='비우기 완료 · 남은 항목 '+$r.remaining+'개';if($r.errors.Count){$summary.Text='일부 파일을 지우지 못했습니다. 내용을 다시 확인하세요.';[Windows.Forms.MessageBox]::Show($d,($r.errors -join "`r`n"),'비우기 결과')|Out-Null};if($mode -eq 'workspace'){MemoryClear;Read-Meta}}finally{$empty.Enabled=($null -ne $state.plan)}
 }};$empty.Enabled=$false
 DButton $d '내용 확인' 20 480 220 {DSafe {
  if(!$state.target){$summary.Text='먼저 휴지통이 있는 상위 폴더를 선택하세요.';return}
  $state.plan=(Run 'cleanupscan' @{PayloadFile=(NConfigFile @{mode=$mode})} $state.target)|ConvertFrom-Json;$grid.Rows.Clear();foreach($entry in $state.plan.items){[void]$grid.Rows.Add([object[]]@($entry.path,$entry.files,[math]::Round($entry.bytes/1MB,2)))}
  $summary.Text=('{0}항목 / {1}파일 / {2:N2} MB' -f $state.plan.entries,$state.plan.files,($state.plan.bytes/1MB));$empty.Enabled=($state.plan.items.Count -gt 0)
 }}|Out-Null
 if($mode -eq 'projects'){DButton $d '보관 위치 선택' 260 480 220 {$pick=New-Object Windows.Forms.FolderBrowserDialog;$pick.Description='프로젝트가 있었던 상위 폴더 또는 .project-trash를 선택하세요.';try{if($pick.ShowDialog($d) -eq 'OK'){$state.target=if((Split-Path $pick.SelectedPath -Leaf) -eq '.project-trash'){$pick.SelectedPath}else{Join-Path $pick.SelectedPath '.project-trash'};$where.Text=$state.target;$state.plan=$null;$grid.Rows.Clear();$empty.Enabled=$false;$summary.Text='내용 확인을 눌러 대상과 용량을 확인하세요.'}}finally{$pick.Dispose()}}|Out-Null}
 [void](Label $d '일반 백업·미연결 작업공간·Windows 휴지통은 비우지 않습니다. 이 기능은 보안 완전삭제 기능이 아닙니다.' 20 430 880 45)
 DButton $d '닫기' 740 545 160 {$d.Close()}|Out-Null;try{[void]$(Ui12Dialog $d;$d.ShowDialog($f))}finally{$d.Dispose()}
}
$nTrashEmpty=DButton $nHome '프로젝트 휴지통 비우기' 800 150 350 {DSafe {UCleanupDialog 'projects'}}