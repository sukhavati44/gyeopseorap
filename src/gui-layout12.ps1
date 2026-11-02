# TRACKA-R1: geometry only. No data or engine changes.
$script:uiTreeWidth=220;$script:uiLayoutBusy=$false
$uiTools=New-Object Windows.Forms.FlowLayoutPanel;$uiTools.WrapContents=$true;$uiTools.BackColor=[Drawing.Color]::White;$uiTools.Padding=New-Object Windows.Forms.Padding(0);$dashboard.Controls.Add($uiTools)
foreach($spec in @(@($attachExisting,186),@($dashNew,76),@($dashCancel,88),@($dashDown,120),@($neisDownload,120),@($nRosterPrompt,120),@($dashUp,116),@($dashSave,88),@($wOne,105),@($wAll,82))){$button=$spec[0];$button.Parent=$uiTools;$button.Width=$spec[1];$button.Height=30;$button.Margin=New-Object Windows.Forms.Padding(0,0,6,4)}
$uiStatus=New-Object Windows.Forms.Panel;$dashboard.Controls.Add($uiStatus);$dashSelected.Parent=$uiStatus;$dashSelected.AutoEllipsis=$true
$uiLegends=@($dashboard.Controls|Where-Object {$_ -is [Windows.Forms.Label] -and $_.Text -in @('상위 값과 다름','하위 값과 다름','상·하위 모두 다름')});foreach($legend in $uiLegends){$legend.Parent=$uiStatus}
$uiLegendChips=@();$uiLegendFont=New-Object Drawing.Font('맑은 고딕',8.5)
foreach($legend in $uiLegends){$chip=New-Object Windows.Forms.Label;$chip.BackColor=$legend.BackColor;$uiStatus.Controls.Add($chip);$uiLegendChips+=,$chip;$legend.BackColor=[Drawing.Color]::Transparent;$legend.ForeColor=$navy;$legend.Font=$uiLegendFont;$legend.TextAlign='MiddleLeft'}
$uiDetail=DButton $dashboard '상세' 0 0 46 {[Windows.Forms.MessageBox]::Show($f,$dashMessage.Text,'작업 상태 상세','OK','Information')|Out-Null};$uiDetail.Height=26
$uiStatus.BorderStyle='FixedSingle';$uiStatus.BackColor=[Drawing.Color]::FromArgb(247,249,252)
# Reuse the existing detail action as the compact left-hand status box.
$dashMessage.Parent=$uiStatus;$uiDetail.Parent=$uiStatus;$dashSelected.BackColor=[Drawing.Color]::Transparent;$dashMessage.BackColor=[Drawing.Color]::Transparent
$uiDivider=New-Object Windows.Forms.Panel;$uiDivider.BackColor=$pale;$uiDivider.Width=5;$uiDivider.Cursor=[Windows.Forms.Cursors]::SizeWE;$dashboard.Controls.Add($uiDivider)
$script:uiDragging=$false;$uiDivider.Add_MouseDown({$script:uiDragging=$true;$uiDivider.Capture=$true});$uiDivider.Add_MouseUp({$script:uiDragging=$false;$uiDivider.Capture=$false});$uiDivider.Add_MouseMove({param($sender,$e)if($uiDragging){$script:uiTreeWidth=[math]::Max(165,[math]::Min(285,$uiTreeWidth+$e.X-2));DLayout}})
function DLayout {
 if($uiLayoutBusy){return};$script:uiLayoutBusy=$true
 try{
  $w=$dashboard.ClientSize.Width;$h=$dashboard.ClientSize.Height;$left=$uiTreeWidth+24;$width=[math]::Max(500,$w-$left-12);$bottom=$h-48
  $dashTitle.SetBounds(48,6,235,28);$dashTitle.Font=New-Object Drawing.Font('맑은 고딕',12,[Drawing.FontStyle]::Bold)
  $nProjectLabel.SetBounds(12,35,$uiTreeWidth,48);$nProjectLabel.AutoEllipsis=$true
  $wModeButton.SetBounds(($w-286),4,150,28);$nProjectHeader.SetBounds(($w-128),4,116,28)
  $dashFilters.SetBounds($left,36,$width,54)
  $specs=@(@($dy,44,78),@($dt,160,74),@($dg,271,70),@($dc,365,84),@($ds,491,106))
  foreach($spec in $specs){$spec[0].SetBounds($spec[1],2,$spec[2],26)}
  $labels=@($dashFilters.Controls|Where-Object {$_ -is [Windows.Forms.Label] -and $_ -ne $dashContext})
  foreach($label in $labels){switch($label.Text){'학년도'{$label.SetBounds(0,5,44,24)};'학기'{$label.SetBounds(128,5,32,24)};'학년'{$label.SetBounds(239,5,32,24)};'반'{$label.SetBounds(347,5,18,24)};'과목'{$label.SetBounds(455,5,36,24)};'기준일'{$label.SetBounds(603,5,45,24)}}}
  foreach($label in $labels){if($label.Text -eq '기준일'){$label.Width=54}}
  $ddate.SetBounds(658,2,122,26);$dashQuery.SetBounds(($width-72),1,72,28);$dashContext.SetBounds(0,31,$width,22)
  $headerOffset=0;if($width -lt 860){$dashFilters.Height=60;$headerOffset=6;foreach($label in $labels){if($label.Text -eq '기준일'){$label.SetBounds(0,34,54,24)}};$ddate.SetBounds(54,30,122,26);$dashContext.SetBounds(186,33,($width-270),24);$dashQuery.Top=29}
  $dashHeading.SetBounds($left,(94+$headerOffset),$width,25);$dashHeading.Font=New-Object Drawing.Font('맑은 고딕',11,[Drawing.FontStyle]::Bold);$dashHeading.AutoEllipsis=$true;if($nActionTips){$nActionTips.SetToolTip($dashHeading,$dashHeading.Text);$nActionTips.SetToolTip($nProjectLabel,$nProjectLabel.Text)}
  $uiTools.SetBounds($left,(122+$headerOffset),$width,68);$uiTools.PerformLayout();$toolBottom=0;foreach($control in $uiTools.Controls){if($control.Visible){$toolBottom=[math]::Max($toolBottom,$control.Bottom)}};$toolHeight=[math]::Max(30,$toolBottom);$uiTools.Height=$toolHeight
  $gridTop=128+$headerOffset+$toolHeight;$dashGrid.SetBounds($left,$gridTop,$width,[math]::Max(260,$bottom-84-$gridTop));$dashGrid.ColumnHeadersHeight=48
  $uiStatus.SetBounds($left,($dashGrid.Bottom-1),$width,84);$dashSelected.SetBounds(64,6,[math]::Max(120,$width-244),24)
  $i=0;foreach($legend in $uiLegends){$legend.SetBounds(($width-136),(4+$i*19),128,18);$uiLegendChips[$i].SetBounds(($width-158),(7+$i*19),16,12);$i++}
  $dashMessage.SetBounds(10,34,($width-180),42);$dashMessage.AutoEllipsis=$true;$uiDetail.SetBounds(10,6,46,24);$dashCount.Visible=$false
  $side.SetBounds(8,96,$uiTreeWidth,[math]::Max(400,$bottom-96));foreach($label in @($side.Controls|Where-Object {$_ -is [Windows.Forms.Label] -and $_.Visible})){$label.SetBounds(0,0,$uiTreeWidth,24)}
  $buttons=@($side.Controls|Where-Object {$_ -is [Windows.Forms.Button] -and $_.Visible}|Sort-Object Top);$treeHeight=[math]::Max(180,$side.Height-28-$buttons.Count*34);$wTree.SetBounds(0,26,$uiTreeWidth,$treeHeight);$j=0;foreach($button in $buttons){$button.SetBounds(0,(32+$treeHeight+$j*34),$uiTreeWidth,28);$j++}
  $uiDivider.SetBounds(($left-10),96,5,$side.Height)
  $nBar.Height=38;foreach($button in $nTabs.Values){$button.Width=[math]::Max(100,[int](($f.ClientSize.Width-42)/7));$button.Height=30;$button.Margin=New-Object Windows.Forms.Padding(3,4,3,0)}
 }finally{$script:uiLayoutBusy=$false}
}
function Ui12HomeLayout{
 $w=$nHome.ClientSize.Width;$h=$nHome.ClientSize.Height;$nHomeTitle.SetBounds(80,12,300,38);$nHomeDescription.SetBounds(390,24,($w-420),24);$nFlow.SetBounds(24,54,($w-48),40)
 $top=@($nHome.Controls|Where-Object {$_ -is [Windows.Forms.Button] -and $_.Text -in @('새 프로젝트 만들기','프로젝트 불러오기','프로젝트 휴지통 비우기')}|Sort-Object Left);$bw=[int](($w-64)/3);for($i=0;$i -lt $top.Count;$i++){$top[$i].SetBounds((24+$i*($bw+8)),100,$bw,34)}
 $nCurrent.SetBounds(24,144,($w-48),62);foreach($label in $nHome.Controls|Where-Object {$_ -is [Windows.Forms.Label] -and $_.Text -like '최근 프로젝트*'}){$label.SetBounds(24,216,700,24)}
 $nRecent.SetBounds(24,244,($w-48),[math]::Max(300,$h-390))
 $bottom=@($nHome.Controls|Where-Object {$_ -is [Windows.Forms.Button] -and $_ -notin $top}|Sort-Object Top,Left);$bw=[int](($w-48-8*($bottom.Count-1))/$bottom.Count);for($i=0;$i -lt $bottom.Count;$i++){$bottom[$i].SetBounds((24+$i*($bw+8)),($h-132),$bw,34)}
 foreach($label in $nHome.Controls|Where-Object {$_ -is [Windows.Forms.Label] -and $_.Text -like 'Drive*'}){$label.SetBounds(24,($h-84),($w-48),36)}
}
function Ui12Dialog($dialog){
 if($dialog -isnot [Windows.Forms.Form]){return}
 if(($dialog.Tag -is [string]) -and ($dialog.Tag -eq '__layout_done')){return}
 if($dialog -eq $f -or $dialog.AccessibleDescription -eq 'ui12-layout-done'){return}
 $dialog.AccessibleDescription='ui12-layout-done';$old=$dialog.ClientSize;$width=$old.Width;$height=$old.Height
 switch($dialog.Text){
 '새 프로젝트 만들기'{$width=740;$height=580}
 '열 이름 바꾸기'{$width=460;$height=175}
 '작업공간 이름'{$width=460;$height=175}
 '열 여러 개 추가'{$width=620;$height=390}
 '새 열의 사용 방식'{$width=560;$height=235}
 '열 종류'{$width=540;$height=210}
 '기존 열에 하위 공간 만들기'{$width=620;$height=300}
 default{if($old.Width -gt 900){$width=[math]::Min(1050,[int]($f.ClientSize.Width*.78));$height=[math]::Min(680,[int]($f.ClientSize.Height*.72))}}
 }
 $sx=$width/[double]$old.Width;$sy=$height/[double]$old.Height
 $dialog.ClientSize=New-Object Drawing.Size($width,$height)
 foreach($control in $dialog.Controls){$control.SetBounds([int]($control.Left*$sx),[int]($control.Top*$sy),[int]($control.Width*$sx),$(if($control -is [Windows.Forms.Button]){34}elseif($control -is [Windows.Forms.ComboBox]){28}else{[math]::Max(24,[int]($control.Height*$sy))}));if($control -is [Windows.Forms.Button]){$control.Anchor='Bottom,Right'}elseif($control -is [Windows.Forms.DataGridView] -or $control -is [Windows.Forms.TextBox] -and $control.Multiline){$control.Anchor='Top,Bottom,Left,Right'}else{$control.Anchor='Top,Left,Right'}}
 Ui12DialogDetails $dialog
 if(Get-Command Ui12Style -ErrorAction SilentlyContinue){Ui12Style $dialog}
 $dialog.StartPosition='CenterParent';$dialog.MinimumSize=$dialog.Size
}
function Ui12DialogDetails($dialog){
 $w=$dialog.ClientSize.Width;$h=$dialog.ClientSize.Height
 $buttons=@($dialog.Controls|Where-Object {$_ -is [Windows.Forms.Button]}|Sort-Object Left)
 $grids=@($dialog.Controls|Where-Object {$_ -is [Windows.Forms.DataGridView]})
 switch($dialog.Text){
  '기존 열에 하위 공간 만들기'{
   $labels=@($dialog.Controls|Where-Object {$_ -is [Windows.Forms.Label]}|Sort-Object Top)
   for($i=0;$i -lt $labels.Count;$i++){$labels[$i].SetBounds(18,(14+$i*34),($w-36),32);$labels[$i].AutoEllipsis=$true;$nActionTips.SetToolTip($labels[$i],$labels[$i].Text)}
   foreach($box in $dialog.Controls|Where-Object {$_ -is [Windows.Forms.TextBox]}){$box.SetBounds(18,126,($w-36),30)}
   $buttons[0].SetBounds(($w-400),($h-52),260,34);$buttons[1].SetBounds(($w-128),($h-52),110,34)
  }
  '선택 작업공간에 비정형 자료 반입'{
   foreach($label in $dialog.Controls|Where-Object {$_ -is [Windows.Forms.Label]}){$label.AutoEllipsis=$true;$nActionTips.SetToolTip($label,$label.Text)}
   foreach($grid in $grids){$grid.SetBounds(16,150,($w-32),($h-220))}
   foreach($button in $buttons){if($button.Text -eq '정상 기록 반입'){$button.SetBounds(($w-236),($h-52),220,34)}elseif($button.Text -eq '예외 행 해결'){$button.SetBounds(16,($h-52),190,34)}else{$button.Anchor='Top,Right';$button.SetBounds(($w-236),58,220,34)}}
   foreach($box in $dialog.Controls|Where-Object {$_ -is [Windows.Forms.TextBox]}){$box.SetBounds(16,60,($w-268),28);$box.Anchor='Top,Left,Right'}
  }
  '선택 학생 · 필요한 깊이만 조회'{
   $dialog.ClientSize=New-Object Drawing.Size(800,500);$w=800;$h=500
   foreach($label in $dialog.Controls|Where-Object {$_ -is [Windows.Forms.Label]}){$label.SetBounds(16,12,768,36);$label.AutoEllipsis=$true}
   foreach($combo in $dialog.Controls|Where-Object {$_ -is [Windows.Forms.ComboBox]}){$combo.SetBounds(16,58,430,28)}
   foreach($button in $buttons){$button.SetBounds(458,56,220,34);$button.Anchor='Top,Right'}
   foreach($grid in $grids){$grid.SetBounds(16,102,768,380)}
  }
 }
}
function Ui12Help{
 $d=New-Object Windows.Forms.Form;$d.Text='오류·복구 안내';$d.Font=$f.Font;$d.ClientSize=New-Object Drawing.Size(620,320);$d.MinimumSize=New-Object Drawing.Size(500,300);$d.StartPosition='CenterParent';$d.ShowInTaskbar=$false
 $text=New-Object Windows.Forms.TextBox;$text.Multiline=$true;$text.ReadOnly=$true;$text.ScrollBars='Vertical';$text.BorderStyle='None';$text.BackColor=[Drawing.Color]::White;$text.SetBounds(20,20,580,238);$text.Anchor='Top,Bottom,Left,Right'
 $text.Text="상위·하위 값이 다르면 표 아래 범례의 색으로 표시합니다.`r`n`r`n자료 가져오기: 최하위 작업표를 권장합니다. 학생 연결 확인 → 열 연결 → 현재 표에 저장 순서로 진행하세요. 제외 항목 탭에서 원본 행·열과 이유를 확인할 수 있습니다.`r`n`r`n저장 중단: 추가 저장을 멈추고 기존 체크포인트 복구를 확인하세요.`r`n`r`n프로젝트 백업·새 폴더 복원은 프로젝트 탭에서 수행합니다.`r`n원본 Excel은 저장·닫기 후 조회하세요."
 $d.Controls.Add($text);$close=DButton $d '닫기' 490 270 110 {$d.Close()};$close.Anchor='Bottom,Right';$close.Height=34;$d.CancelButton=$close;$d.AcceptButton=$close
 if(Get-Command Ui12Style -ErrorAction SilentlyContinue){Ui12Style $d};$d.Add_Shown({$text.Select(0,0);$close.Focus()})
 try{[void]$d.ShowDialog($f)}finally{$d.Dispose()}
}
$wTree.ShowNodeToolTips=$true
$wTree.Add_AfterSelect({param($sender,$e)if($e.Node){$e.Node.ToolTipText=$e.Node.FullPath}})
$dashGrid.Add_Scroll({$dashGrid.Invalidate()});$dashGrid.Add_ColumnWidthChanged({$dashGrid.Invalidate()});$dashGrid.Add_ColumnStateChanged({$dashGrid.Invalidate()});$dashGrid.Add_ColumnDisplayIndexChanged({$dashGrid.Invalidate()})
$f.MinimumSize=New-Object Drawing.Size(1080,740)
$nHome.Add_Resize({Ui12HomeLayout});$f.Add_Shown({DLayout;Ui12HomeLayout});DLayout;Ui12HomeLayout
$script:layoutOnLoad=${function:NOnLoad};function NOnLoad{& $layoutOnLoad;DLayout}
