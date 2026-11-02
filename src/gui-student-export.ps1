function StudentExportDialog{
 DClean;if(!$ws -or !$meta){throw '프로젝트를 먼저 여세요.'}
 $d=New-Object Windows.Forms.Form;$d.Text='기록 내보내기 · 학생별 전체 자료';$d.Size=New-Object Drawing.Size(760,430);$d.StartPosition='CenterParent'
 [void](Label $d '조회 항목·깊이와 무관하게 전체 자료를 학생 1명당 XLSX 1개로 저장합니다.' 20 20 700 35)
 [void](Label $d '생성 범위' 20 75 100 25);$scope=Combo $d 130 72 220;Choice $scope 'one' '1인';Choice $scope 'class' '학급 단위';$scope.SelectedIndex=0
 [void](Label $d '대상' 20 120 100 25);$target=Combo $d 130 117 570
 [void](Label $d '식별정보' 20 165 100 25);$identity=Combo $d 130 162 220;Choice $identity 'name' '이름';Choice $identity 'both' '둘 다';Choice $identity 'code' '코드';$identity.SelectedIndex=$wIdentityMode
 $exportRegistry=$null;$registryPath=Join-Path $ws 'student-codes.json';if(Test-Path -LiteralPath $registryPath){$exportRegistry=Map ([IO.File]::ReadAllText($registryPath)|ConvertFrom-Json)}
 $exportClasses=@{};$exportDate=$ddate.Value.ToString('yyyy-MM-dd');foreach($member in $meta.memberships){if($member.from -le $exportDate -and (!$member.until -or $member.until -ge $exportDate)){$exportClasses[[string]$member.student]=(@($exportClasses[[string]$member.student],[string]$member.class)|Where-Object {$_}) -join ', '}}
 $targetFill={ $target.Items.Clear();if($scope.SelectedItem.Id -eq 'one'){foreach($sid in @($meta.students.Keys|Sort-Object { $meta.students[$_].number })){$person=$meta.students[$sid];$label=[string]$exportClasses[$sid]+' / '+[string]$person.number+'번 '+$person.name;$code=if($exportRegistry -and $exportRegistry.codes[$sid]){$exportRegistry.prefix+'-'+$exportRegistry.codes[$sid].suffix}else{'코드 미발급'};if($identity.SelectedItem.Id -eq 'both'){$label+=' / '+$code}elseif($identity.SelectedItem.Id -eq 'code'){$label=$code};Choice $target $sid $label};$sid=if($dashGrid.CurrentRow){[string]$dashGrid.CurrentRow.Tag}else{''};Choose $target $sid}else{foreach($cl in @($meta.memberships|Where-Object {$_.from -le $ddate.Value.ToString('yyyy-MM-dd') -and (!$_.until -or $_.until -ge $ddate.Value.ToString('yyyy-MM-dd'))}|ForEach-Object {$_.class}|Sort-Object -Unique)){Choice $target $cl $cl};if($dc.SelectedItem.Id){Choose $target $dc.SelectedItem.Id}};if($target.SelectedIndex -lt 0 -and $target.Items.Count){$target.SelectedIndex=0} }
 $scope.Add_SelectedIndexChanged($targetFill);$identity.Add_SelectedIndexChanged($targetFill);& $targetFill
 [void](Label $d '코드만 출력해도 원문 내용·메모의 실명은 자동으로 지우지 않습니다.' 20 215 700 35)
 $result=Label $d '' 20 252 700 40
 $go=DButton $d '내보내기' 440 320 130 {DSafe {
  if(!$target.SelectedItem){throw '대상 학생 또는 학급을 선택하세요.'};$pick=New-Object Windows.Forms.FolderBrowserDialog;$pick.Description='학생별 XLSX를 저장할 폴더 선택 (프로젝트 밖)';$pick.SelectedPath=$base;if($pick.ShowDialog($d) -ne 'OK'){return}
  $go.Enabled=$false;$close.Enabled=$false;$scope.Enabled=$false;$target.Enabled=$false;$identity.Enabled=$false;$d.ControlBox=$false
  try{$r=(Run 'studentexport' @{PayloadFile=(Write-Request @{scope=[string]$scope.SelectedItem.Id;student=$(if($scope.SelectedItem.Id -eq 'one'){$target.SelectedItem.Id}else{''});class=$(if($scope.SelectedItem.Id -eq 'class'){$target.SelectedItem.Id}else{''});identity=[string]$identity.SelectedItem.Id;as_of=$ddate.Value.ToString('yyyy-MM-dd');output=$pick.SelectedPath})})|ConvertFrom-Json}finally{$go.Enabled=$true;$close.Enabled=$true;$scope.Enabled=$true;$target.Enabled=$true;$identity.Enabled=$true;$d.ControlBox=$true}
  $result.Text='완료 '+$r.success+'명 / 실패 '+$r.failed+'명';$dashMessage.Text=$result.Text+' · '+$r.directory
  $details=($r.results|ForEach-Object { $_.target+' : '+$_.status+$(if($_.error){' / '+$_.error}else{''}) }) -join "`r`n"
  [void][Windows.Forms.MessageBox]::Show($d,($result.Text+"`r`n"+$r.directory+"`r`n"+$details),'기록 내보내기 결과','OK','Information')
 }};$close=DButton $d '닫기' 580 320 120 {$d.Close()};$d.CancelButton=$close
 try{Ui12Dialog $d;[void]$d.ShowDialog($f)}finally{$d.Dispose()}
}
