function NChooseNeisAreas($files){
 $preview=$null;$example='';$lastError=''
 foreach($file in @($files|Sort-Object -Unique)){try{$preview=(Run 'neispreview' (DArgs @{File=$file}))|ConvertFrom-Json;$example=[IO.Path]::GetFileName($file);break}catch{$lastError=$_.Exception.Message}}
 if(!$preview){throw ('영역을 읽을 수 있는 파일이 없습니다. '+$lastError)}
 $profile=$meta.courses[$dashProfileId];$d=New-Object Windows.Forms.Form;$d.Text='반입할 평가영역 선택 · 전체 파일에 적용';$d.Size=New-Object Drawing.Size(850,440);$d.StartPosition='CenterParent';$d.Font=$f.Font
 [void](Label $d ('기준 파일: '+$example) 20 15 790 45)
 [void](Label $d '원본의 각 점수 열을 어느 영역에 넣을지 선택하세요. 이 연결을 선택한 모든 학급 파일의 같은 열에 적용합니다.' 20 65 790 50)
 $choices=@();for($i=0;$i -lt $preview.areas.Count;$i++){
  [void](Label $d (($i+1).ToString()+'번째 점수 열: '+$preview.areas[$i].label+' →') 20 (125+45*$i) 390 35)
  $box=Combo $d 420 (125+45*$i) 380;foreach($a in $profile.areas){Choice $box $a.root_item $a.label};$match=@($profile.areas|Where-Object {$_.label -ceq $preview.areas[$i].label});if($match.Count -eq 1){Choose $box $match[0].root_item}elseif($i -lt $box.Items.Count){$box.SelectedIndex=$i};$choices+=,$box
 }
 $errorLabel=Label $d '' 20 275 780 40;$errorLabel.ForeColor=[Drawing.Color]::Firebrick
 DButton $d '선택한 영역으로 전체 반입' 340 325 280 {$items=@($choices|ForEach-Object {if($_.SelectedItem){[string]$_.SelectedItem.Id}});if($items.Count -ne $choices.Count -or @($items|Sort-Object -Unique).Count -ne $items.Count){$errorLabel.Text='각 원본 열에 서로 다른 대상 영역을 선택하세요.';return};$d.Tag=[pscustomobject]@{items=$items};$d.DialogResult='OK';$d.Close()}|Out-Null
 $cancel=DButton $d '취소' 640 325 160 {$d.DialogResult='Cancel';$d.Close()};$d.CancelButton=$cancel
 try{if($(Ui12Dialog $d;$d.ShowDialog($f)) -eq 'OK'){return $d.Tag};return $null}finally{$d.Dispose()}
}
# NEIS original input and original-format summary output.
function NImportClasses($files,$areaItems=@()){
 $target=$ws;$results=[Collections.Generic.List[object]]::new();$paths=@($files|Sort-Object -Unique);$index=0;$stopped=$false
 foreach($path in $paths){$index++;if($stopped){$results.Add(@{file=$path;ok=$false;message='앞선 저장 중단으로 미실행'});continue}
  $nCurrent.Text='학급 일괄 반입 중 '+$index+'/'+$paths.Count+' · '+[IO.Path]::GetFileName($path)
  try{$args2=DArgs @{File=$path};if($areaItems.Count){$args2.PayloadFile=NConfigFile @{area_items=@($areaItems)}};$r=(Run 'neisimport' $args2 $target)|ConvertFrom-Json;$results.Add(@{file=$path;ok=$true;class=$r.class;new=$r.new;total=$r.total;message='반입 완료'})}
  catch{$results.Add(@{file=$path;ok=$false;message=($_.Exception.Message -split "`r?`n"|Where-Object {$_}|Select-Object -First 1)});if(Test-Path -LiteralPath (Join-Path $target 'pending-sync.json')){$stopped=$true}}
 }
 return $results.ToArray()
}
$neisImportButton=DButton $nHome '나이스 학급 일괄 반입' 40 790 360 {DSafe {
 DClean;if(!$ws){throw '빈 프로젝트를 먼저 만드세요.'}
 $pick=New-Object Windows.Forms.OpenFileDialog;$pick.Title='나이스 학급 파일 여러 개 선택 (Ctrl+A / Ctrl / Shift)';$pick.Filter='Excel|*.xlsx';$pick.Multiselect=$true
 try{if($pick.ShowDialog($f) -ne 'OK'){return};$files=@($pick.FileNames)}finally{$pick.Dispose()}
 $areaSelection=NChooseNeisAreas $files;if(!$areaSelection){return};$target=$ws;$results=@(NImportClasses $files @($areaSelection.items));$ok=@($results|Where-Object {$_.ok});$failed=@($results|Where-Object {!$_.ok});$newCount=($ok|ForEach-Object {$_.new}|Measure-Object -Sum).Sum
 $report='학급 반입 완료: '+$ok.Count+'/'+$files.Count+'파일 성공 · 학생 '+[int]$newCount+'명 추가 · 실패/미실행 '+$failed.Count+'파일'
 $detail=($results|ForEach-Object {[IO.Path]::GetFileName($_.file)+' : '+$(if($_.ok){$_.class+' '+$_.new+'명'}else{$_.message})}) -join "`r`n"
 MemoryClear;NOpen $target;$dashMessage.Text=$report;$nCurrent.Text=$report+"`r`n"+$detail;Log ($report+"`r`n"+$detail)
 if($failed.Count){[Windows.Forms.MessageBox]::Show($f,$report+"`r`n`r`n"+$detail,'학급 일괄 반입 결과','OK','Information')|Out-Null}
}}
$neisDownload=DButton $dashboard '총괄표 다운로드' 770 175 150 {DSafe {DClean;if(!$meta.neisSources){throw '나이스 원본을 먼저 반입하세요.'};$dialog=New-Object Windows.Forms.SaveFileDialog;$dialog.Title='새 출력 폴더 이름 (선택 반 또는 전체)';$dialog.FileName='나이스총괄표-'+(Get-Date -Format 'yyyyMMdd-HHmmss');$dialog.Filter='새 폴더 이름|*';if($dialog.ShowDialog($f) -ne 'OK'){return};$r=(Run 'neisexport' (DArgs @{File=$dialog.FileName}))|ConvertFrom-Json;$dashMessage.Text='원본 서식 총괄표 '+$r.files.Count+'개 출력: '+$dialog.FileName}}
$neisDownload.Visible=$false
$f.Add_Activated({if($dashMode -eq 'summary' -and $meta.neisSources){$neisDownload.Visible=$true;$dashDown.Visible=$false}else{$neisDownload.Visible=$false;$dashDown.Visible=$true};NRefreshActions})

if(Get-Command NRefreshActions -ErrorAction SilentlyContinue){NRefreshActions}
