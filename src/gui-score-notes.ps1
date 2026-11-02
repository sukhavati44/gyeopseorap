. (Join-Path $base 'score-notes.ps1')
$script:scoreColor=[Drawing.Color]::FromArgb(174,38,48)
$script:scoreCellNormalFont=$dashGrid.Font
$script:scoreCellBoldFont=New-Object Drawing.Font($scoreCellNormalFont,[Drawing.FontStyle]::Bold)
function ScoreTarget{
 if(!$ws -or !$dashSnapshot){return $null};$col=if($script:scoreHeaderTarget){$script:scoreHeaderTarget}elseif($dashGrid.CurrentCell){$dashGrid.CurrentCell.OwningColumn}else{$null};if(!$col){return $null}
 if(!$col.Tag.item -and $col.Name -ne 'total'){return $null}
 $item=if($col.Name -eq 'total'){'__sum__'}else{[string]$col.Tag.item}
 return @{root=$ws;layer=$dashLayer;item=$item;label=($col.HeaderText -replace '\r?\n',' ')}
}
function ScoreEdit{
 $target=ScoreTarget;if(!$target){return};ScoreEnsure $target.root;$key=$target.layer+'|'+$target.item
 $d=New-Object Windows.Forms.Form;$d.Text='배점 · 1행 열 메모';$d.Size=New-Object Drawing.Size(650,500);$d.StartPosition='CenterParent'
 [void](Label $d ($target.label+' · 배점 (선택)') 18 16 595 30)
 [void](Label $d '한 줄은 한 평가요소입니다. 같은 줄의 가능 점수는 쉼표·세미콜론으로 구분합니다.' 18 50 595 44)
 $text=New-Object Windows.Forms.TextBox;$text.Multiline=$true;$text.AcceptsReturn=$true;$text.ScrollBars='Vertical';$text.AccessibleName='배점 메모 · 줄마다 평가요소, 쉼표 또는 세미콜론으로 점수 구분';$text.SetBounds(18,98,595,200);$text.Text=[string]$scoreNotes[$key].text;$d.Controls.Add($text)
 $preview=Label $d '' 18 308 595 90;$preview.ForeColor=[Drawing.Color]::FromArgb(109,66,145)
 $text.Add_TextChanged({try{$rule=ScoreParse $this.Text;$preview.Text=if($rule){'가능한 점수 ('+$rule.count+'개): '+$rule.display+$(if($rule.count -gt 30){' …'}else{''})}else{'배점 없음 · 점수 대조를 하지 않습니다.'}}catch{$preview.Text=$_.Exception.Message}})
 $ok=Button $d '배점 저장' 350 410 125 {try{ScoreSet $target.root $target.layer $target.item $text.Text;$d.DialogResult='OK';$d.Close()}catch{$preview.Text=$_.Exception.Message}}
 $cancel=Button $d '취소' 485 410 125 {$d.DialogResult='Cancel';$d.Close()};$d.CancelButton=$cancel
 Ui12Dialog $d;$d.Add_Shown({try{$rule=ScoreParse $text.Text;$preview.Text=if($rule){'가능한 점수: '+$rule.display}else{'예: 5,4,3 [엔터] 3;1 [엔터] 2,0'}}catch{$preview.Text=$_.Exception.Message};$text.Focus()})
 try{if($d.ShowDialog($f) -eq 'OK'){$dashGrid.Invalidate();ScoreHeaderTips;$dashMessage.Text='배점 저장 완료 · 가능한 점수를 글자색으로 안내합니다. 점수는 변경하지 않았습니다.'}}finally{$d.Dispose()}
}
function ScoreHeaderTips{
 if(!$ws){return};ScoreEnsure $ws;foreach($col in $dashGrid.Columns){$item=if($col.Name -eq 'total'){'__sum__'}else{$col.Tag.item};if(!$item){continue};$key=$dashLayer+'|'+$item
 if($scoreNotes[$key]){$col.HeaderCell.ToolTipText='[배점] '+$scoreNotes[$key].text+"`r`n머리글 우클릭 → 배점 추가 / 수정"}
 }
}
$script:scoreHeaderTarget=$null
$dashGrid.Add_CellMouseDown({param($sender,$e) $script:scoreHeaderTarget=if($e.RowIndex -eq -1 -and $e.ColumnIndex -ge 0){$sender.Columns[$e.ColumnIndex]}else{$null}})
$scoreMenu=$columnMenu.Items.Add('배점 추가 / 수정 · 1행 메모');$scoreMenu.Add_Click({DSafe {ScoreEdit}})
$scoreHeaderMenu=$columnTools.Items.Add('배점 추가 / 수정 · 1행 메모');$scoreHeaderMenu.Add_Click({DSafe {ScoreEdit}})
$columnMenu.Add_Opening({$scoreMenu.Enabled=(!$busy -and !$dashLoading -and $null -ne (ScoreTarget))})
$columnTools.Add_Opening({$scoreHeaderMenu.Enabled=(!$busy -and !$dashLoading -and $null -ne (ScoreTarget))})
$dashGrid.Add_CellMouseDown({param($sender,$e)
 if($e.RowIndex -eq -1 -and $e.ColumnIndex -ge 0 -and $e.Button -eq [Windows.Forms.MouseButtons]::Right -and !$busy -and !$dashLoading){$script:columnToolTarget=$sender.Columns[$e.ColumnIndex];$columnTools.Show($sender,[Drawing.Point]::new($sender.GetCellDisplayRectangle($e.ColumnIndex,-1,$true).Left,$sender.ColumnHeadersHeight))}
})
$dashGrid.Add_CellFormatting({param($sender,$e)
 if($dashLoading -or !$ws -or $e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0){return};$col=$sender.Columns[$e.ColumnIndex];$item=if($col.Name -eq 'total'){'__sum__'}else{$col.Tag.item};if(!$item){return}
 ScoreEnsure $ws;$rule=$scoreRules[$dashLayer+'|'+$item];$bad=ScoreMismatch $rule $sender.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
 if($bad){$e.CellStyle.Font=$scoreCellBoldFont;$e.CellStyle.ForeColor=$scoreColor;$e.CellStyle.SelectionForeColor=$scoreColor;$sender.Rows[$e.RowIndex].Cells[$e.ColumnIndex].ToolTipText+="`r`n배점 조합으로 나올 수 없는 점수입니다."}else{$e.CellStyle.Font=$scoreCellNormalFont;$e.CellStyle.ForeColor=$navy;$e.CellStyle.SelectionForeColor=$navy}
})
$dashGrid.Add_CellPainting({param($sender,$e)
 if($e.RowIndex -ne -1 -or $e.ColumnIndex -lt 0 -or !$ws -or $dashLoading){return};$col=$sender.Columns[$e.ColumnIndex];$item=if($col.Name -eq 'total'){'__sum__'}else{$col.Tag.item};if(!$item){return};ScoreEnsure $ws
 if(!$scoreNotes[$dashLayer+'|'+$item]){return};if(!$e.Handled){$e.Paint($e.ClipBounds,[Windows.Forms.DataGridViewPaintParts]::All)}
 # Small outlined note icon, clear of the filter arrow and linked-header fill.
 $r=$e.CellBounds;$x=$r.Right-18;$y=$r.Top+4
 $iconPen=New-Object Drawing.Pen([Drawing.Color]::FromArgb(145,137,161),1)
 $linePen=New-Object Drawing.Pen([Drawing.Color]::FromArgb(207,197,225),1);$linePen.DashStyle=[Drawing.Drawing2D.DashStyle]::Dot
 try{
  $e.Graphics.DrawRectangle($iconPen,$x,$y,10,11)
  $e.Graphics.DrawLine($iconPen,($x+2),($y+4),($x+8),($y+4));$e.Graphics.DrawLine($iconPen,($x+2),($y+7),($x+6),($y+7))
  $e.Graphics.DrawLine($linePen,($r.Left+2),($r.Bottom-2),($r.Right-3),($r.Bottom-2))
 }finally{$iconPen.Dispose();$linePen.Dispose()};$e.Handled=$true
})
$script:scorePreviousSelection=${function:DSelection}
function DSelection{
 & $scorePreviousSelection
 if($dashLoading -or !$ws -or !$dashGrid.CurrentCell){return};$col=$dashGrid.CurrentCell.OwningColumn;$item=if($col.Name -eq 'total'){'__sum__'}else{$col.Tag.item};if(!$item){if($dashMessage.Text -like '배점 조합으로*'){$dashMessage.Text='학생정보 열은 배점 대조 대상이 아닙니다.'};return};$target=@{item=$item};ScoreEnsure $ws;$rule=$scoreRules[$dashLayer+'|'+$target.item]
 if($rule.error){$dashMessage.Text='배점 메모 형식을 확인하세요. 머리글 우클릭 → 배점 추가 / 수정';return}
 if(ScoreMismatch $rule $dashGrid.CurrentCell.Value){$dashMessage.Text='배점 조합으로 나올 수 없는 점수입니다. 가능한 점수: '+$rule.display}
 elseif($dashMessage.Text -like '배점 조합으로*'){$dashMessage.Text=if($rule -and ![string]::IsNullOrWhiteSpace([string]$dashGrid.CurrentCell.Value)){'배점에 맞는 점수입니다. 저장·상위 반영은 기존 방식으로 진행합니다.'}else{'값을 입력하거나 Excel 값을 반영하세요. 배점이 있는 열만 점수를 대조합니다.'}}
}
$dashGrid.Add_CellValueChanged({if(!$dashLoading){DSelection}})
$script:scorePreviousHeader=${function:WSelectedHeader};function WSelectedHeader{& $scorePreviousHeader;ScoreHeaderTips}
$script:scorePreviousLoad=${function:NOnLoad};function NOnLoad{& $scorePreviousLoad;ScoreHeaderTips}
$scoreLegend=New-Object Windows.Forms.Label;$scoreLegend.Text='배점 불일치';$scoreLegend.ForeColor=[Drawing.Color]::Black;$scoreLegend.Font=$uiLegendFont;$scoreLegend.TextAlign='MiddleLeft';$uiStatus.Controls.Add($scoreLegend)
$scoreLegendSample=New-Object Windows.Forms.Label;$scoreLegendSample.Text='12';$scoreLegendSample.ForeColor=$scoreColor;$scoreLegendSample.Font=New-Object Drawing.Font($uiLegendFont,[Drawing.FontStyle]::Bold);$scoreLegendSample.TextAlign='MiddleCenter';$uiStatus.Controls.Add($scoreLegendSample)
function ScoreLegendLayout{
 $width=$uiStatus.ClientSize.Width
 $scoreLegend.SetBounds(($width-136),61,128,18);$scoreLegendSample.SetBounds(($width-160),61,20,18)
}
$uiStatus.Add_Resize({ScoreLegendLayout});ScoreLegendLayout
