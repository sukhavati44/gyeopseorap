# Track B: lightweight visual styling; no data or workflow changes.
function Ui12Color($hex){[Drawing.ColorTranslator]::FromHtml($hex)}
$script:uiFontRegular=New-Object Drawing.Font('맑은 고딕',9.5)
$script:uiFontBold=New-Object Drawing.Font('맑은 고딕',9.5,[Drawing.FontStyle]::Bold)
$script:uiFontSmall=New-Object Drawing.Font('맑은 고딕',9)
function Ui12Button($button,$ready=$true,$caution=$false,$locked=$false){
 $primary=$button.Text -in @('변경 저장','선택 항목에 하위 공간 만들기','새 프로젝트 만들기','프로젝트 생성','다음','확인','기존 기록 유지하고 만들기','정상 기록 반입','③ 현재 표에 저장')
 $danger=$button.Text -match '삭제|비우기|휴지통으로'
 $background='#EDF2F7';$foreground='#1A202C';$border='#CBD5E0';$font=if($button.Parent -eq $uiTools){$uiFontBold}else{$uiFontRegular}
 if($button.Parent -eq $nBar){if($button.Tag -eq $nView){$background='#1A365D';$foreground='#FFFFFF';$font=$uiFontBold}else{$background='#F7FAFC'}}
 elseif($primary -or $button -eq $wModeButton){$background='#1A365D';$foreground='#FFFFFF';$border='#0D1B2A';$font=$uiFontBold}
 elseif($button.Parent -eq $side){$background='#F7FAFC'}
 if($danger){$background='#FFF5F5';$foreground='#9B2C2C';$border='#FEB2B2'}
 if(($caution -or !$ready) -and !$danger){$background='#FEFCBF';$foreground='#744210';$border='#D69E2E'}
 if($locked -or !$button.Enabled){$background='#E2E8F0';$foreground='#718096';$border='#CBD5E0'}
 $button.FlatStyle='Flat';$button.FlatAppearance.BorderSize=1;$button.FlatAppearance.BorderColor=Ui12Color $border
 $button.BackColor=Ui12Color $background;$button.ForeColor=Ui12Color $foreground;$button.Font=$font
 $button.FlatAppearance.MouseOverBackColor=Ui12Color $(if($background -eq '#1A365D'){'#2C5282'}elseif($danger){'#FED7D7'}else{'#E2E8F0'})
}
function Ui12Style($control){
 if($control -is [Windows.Forms.Form]){$control.BackColor=[Drawing.Color]::White}
 foreach($child in $control.Controls){
  if($child -is [Windows.Forms.Button]){Ui12Button $child}
  elseif($child -is [Windows.Forms.DataGridView]){
   $child.GridColor=Ui12Color '#E2E8F0';$child.BackgroundColor=[Drawing.Color]::White
   $child.ColumnHeadersDefaultCellStyle.Font=$uiFontBold;$child.ColumnHeadersDefaultCellStyle.ForeColor=Ui12Color '#1A202C'
   $child.ColumnHeadersDefaultCellStyle.BackColor=Ui12Color '#EDF2F7';$child.EnableHeadersVisualStyles=$false
  }
  elseif($child -is [Windows.Forms.TextBox]){if(!$child.ReadOnly){$child.BackColor=[Drawing.Color]::White;$child.BorderStyle='FixedSingle'}}
  if($child.Controls.Count){Ui12Style $child}
 }
}
Ui12Style $f
$script:uiStatusStrongFont=New-Object Drawing.Font('맑은 고딕',10,[Drawing.FontStyle]::Bold)
function Ui12StatusMessageStyle{
 $important=$dashMessage.Text -match '배점 조합으로|연결된 열:|하위값 반영:|저장|반영했습니다|반영을 취소|완료|오류|실패|처리 중'
 $dashMessage.Font=if($important){$uiStatusStrongFont}else{$uiFontRegular}
 $dashMessage.ForeColor=Ui12Color $(if($dashMessage.Text -like '배점 조합으로*'){'#AE2630'}elseif($important){'#1A365D'}else{'#4A5568'})
}
$dashMessage.Add_TextChanged({Ui12StatusMessageStyle});Ui12StatusMessageStyle
$uiStatus.Add_Paint({param($sender,$e)$pen=New-Object Drawing.Pen((Ui12Color '#CBD5E0'),1);try{$e.Graphics.DrawLine($pen,0,0,$sender.Width,0)}finally{$pen.Dispose()}})
$columnMenu.BackColor=[Drawing.Color]::White;$columnMenu.ForeColor=Ui12Color '#1A202C';$columnMenu.Font=$uiFontRegular;$columnMenu.Padding=New-Object Windows.Forms.Padding(2)
$columnMenu.Renderer=New-Object Windows.Forms.ToolStripProfessionalRenderer
foreach($item in $columnMenu.Items){$item.Padding=New-Object Windows.Forms.Padding(3,4,3,4)}
$script:themeRefreshActions=${function:NRefreshActions}
function NRefreshActions{& $themeRefreshActions;if($nFlow){$nFlow.Text=$nFlow.Text.Replace('파랑: 실행 가능 · 노랑: 선행 작업 확인/주의 (클릭 가능)','진한색: 주요 실행 · 옅은색: 보조 실행 · 노랑: 선행 작업 확인/주의 (클릭 가능)')}}
NRefreshActions
$nRecent.DefaultCellStyle.SelectionBackColor=Ui12Color '#2B6CB0';$nRecent.DefaultCellStyle.SelectionForeColor=[Drawing.Color]::White;$nRecent.DefaultCellStyle.Font=$uiFontRegular
$columnMenu.Add_Paint({param($sender,$e)$pen=New-Object Drawing.Pen((Ui12Color '#718096'),1);try{$e.Graphics.DrawRectangle($pen,0,0,($sender.Width-1),($sender.Height-1))}finally{$pen.Dispose()}})
