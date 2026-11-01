$codePanel=New-Object Windows.Forms.GroupBox;$codePanel.Text='학생 코드 관리';$codePanel.Font=$uiFontBold;$codePanel.BackColor=$pale;$nHome.Controls.Add($codePanel)
$codeInfo=Label $codePanel '프로젝트를 불러오세요.' 16 28 280 150;$codeInfo.Font=$uiFontRegular
$codeNotice=Label $codePanel '프로젝트 접두어 + 중복 없는 소문자·숫자 난수. 발급한 코드는 다시 열어도 유지됩니다.' 16 180 280 72;$codeNotice.Font=$uiFontRegular
function CodePanelRefresh{
 $open=($ws -and $meta -and $meta.students.Count);foreach($button in @($codeIssue,$codeRegenerate,$codeDownload)){$button.Enabled=[bool]$open;Ui12Button $button ([bool]$open) ($button -eq $codeRegenerate) $false}
 if(!$open){$codeInfo.Text='명부가 있는 프로젝트를 불러오세요.';return}
 $path=Join-Path $ws 'student-codes.json'
 if(!(Test-Path -LiteralPath $path)){$codeInfo.Text='학생 '+$meta.students.Count+"명`r`n학생 코드 미발급";return}
 $r=CodeRegistry $ws;$codeInfo.Text='현재 프로젝트 코드: '+$r.prefix+"`r`n최초 발급: "+(CodeTime $r.issued_at)+"`r`n최근 변경: "+(CodeTime $r.changed_at)+"`r`n발급 인원: "+$r.codes.Count+"명`r`n개별 난수: "+$r.width+'자리'
}
function CodeTableDownload{
 DClean;$r=CodeRegistry $ws;$dialog=New-Object Windows.Forms.SaveFileDialog;$dialog.Filter='Excel 대조표|*.xlsx';$dialog.FileName=$r.prefix+'_학생코드대조표_변경'+(CodeTime $r.changed_at)+'.xlsx'
 if($dialog.ShowDialog($f) -ne 'OK'){return};if(Test-Path -LiteralPath $dialog.FileName){throw '기존 파일 대신 새 파일 이름을 선택하세요.'}
 if(!('StudentCodeTable' -as [type])){Add-Type -Path (Join-Path $base 'StudentCodeTable.cs') -ReferencedAssemblies System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem}
 $rows=New-Object 'Collections.Generic.List[string[]]';$rows.Add([string[]]@('반','번호','이름','학생코드','코드 생성일시','코드 변경일시'))
 $people=@($meta.students.Keys|ForEach-Object {$sid=$_;$membership=@($meta.memberships|Where-Object student -EQ $sid|Sort-Object from -Descending)[0];[pscustomobject]@{sid=$sid;class=$membership.class;number=$meta.students[$sid].number;name=$meta.students[$sid].name}}|Sort-Object class,@{Expression={[int]$_.number}})
 foreach($p in $people){$entry=$r.codes[$p.sid];$rows.Add([string[]]@([string]$p.class,([int]$p.number).ToString('00'),[string]$p.name,(StudentCode $ws $p.sid),(CodeTime $entry.created_at),(CodeTime $entry.changed_at)))}
 [StudentCodeTable]::Write($dialog.FileName,$rows.ToArray());$codeNotice.Text='대조표 다운로드 완료 · 코드 변경일시 '+(CodeTime $r.changed_at);CodePanelRefresh
}
$codeIssue=DButton $codePanel '학생 코드 발급' 16 270 280 {DSafe {DClean;if(!$ws -or !$meta.students.Count){return};$script:studentCodeProjects.Remove($ws);[void](CodeRegistry $ws);$script:wCodes=@{};if($dashSnapshot){WIdentityApply};CodePanelRefresh;$codeNotice.Text='기존 코드는 유지하고 미발급 학생만 추가했습니다.'}}
$codeRegenerate=DButton $codePanel '코드 재생성 · 반영' 16 312 280 {DSafe {DClean;if(!$ws -or !$meta.students.Count){return};$text="전체 학생의 표시 코드를 새 난수로 바꿉니다.`r`n기존 OCR/PDF에 인쇄된 코드와 달라집니다.`r`n이전 코드는 연결 이력으로 보관하며 점수·메모는 유지합니다.";if([Windows.Forms.MessageBox]::Show($f,$text,'학생 코드 전체 재생성','OKCancel','Warning') -ne 'OK'){return};[void](CodeRegistry $ws $true);$script:wCodes=@{};if($dashSnapshot){WIdentityApply};CodePanelRefresh;$codeNotice.Text='새 코드 반영 완료 · '+(CodeTime (CodeRegistry $ws).changed_at)}}
$codeDownload=DButton $codePanel '학생–코드 대조표 다운로드' 16 354 280 {DSafe {CodeTableDownload}}
function CodePanelLayout{
 $w=$nHome.ClientSize.Width;$h=$nHome.ClientSize.Height;$cw=310;$leftWidth=[math]::Max(400,$w-48-$cw-16)
 $nRecent.SetBounds(24,244,$leftWidth,[math]::Max(300,$h-390));$codePanel.SetBounds((24+$leftWidth+16),244,$cw,$nRecent.Height)
 $codeInfo.SetBounds(16,28,278,116);$codeNotice.SetBounds(16,150,278,[math]::Max(24,$codePanel.Height-286));$codeNotice.AutoEllipsis=$true
 $y=$codePanel.Height-126;foreach($button in @($codeIssue,$codeRegenerate,$codeDownload)){$button.SetBounds(16,$y,278,32);$y+=40}
}
$script:codeHomeLayout=${function:Ui12HomeLayout};function Ui12HomeLayout{& $codeHomeLayout;CodePanelLayout}
$script:codeRefreshActions=${function:NRefreshActions};function NRefreshActions{& $codeRefreshActions;if($nView -eq 'project'){CodePanelRefresh}}
$script:codeIdentityApply=${function:WIdentityApply};function WIdentityApply{
 & $codeIdentityApply
 if($ws -and $dashGrid.Columns.Contains('view_code') -and $dashGrid.Rows.Count){$r=CodeRegistry $ws;$dashGrid.Columns['view_code'].Width=180;$dashGrid.Columns['view_code'].HeaderText="학생코드`n생성 "+(CodeTime $r.batch_at);foreach($row in $dashGrid.Rows){$row.Cells['view_code'].ToolTipText='코드 생성 '+(CodeTime $r.codes[[string]$row.Tag].created_at)}}
}
Ui12Style $codePanel;Ui12HomeLayout;CodePanelRefresh
