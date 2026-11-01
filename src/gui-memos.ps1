. (Join-Path $base 'cell-memos.ps1')
function MemoCellItem($cell){if(!$cell){return ''};if($cell.OwningColumn.Name -eq 'total'){return '__sum__'};return [string]$cell.OwningColumn.Tag.item}
function MemoTooltip($layer,$student,$item,$tip){$memo=MemoGet $ws $layer $student $item;if($memo){return (@($tip,('메모: '+$memo.text))|Where-Object {$_}) -join "`r`n"};return $tip}
function MemoCurrent{
 $cell=$dashGrid.CurrentCell;$item=MemoCellItem $cell
 if(!$ws -or !$dashSnapshot -or !$cell -or !$item){return $null}
 return @{root=$ws;layer=$dashLayer;student=[string]$cell.OwningRow.Tag;item=$item;name=[string]$cell.OwningRow.Cells['name'].Value;label=($cell.OwningColumn.HeaderText -replace "`r?`n",' ')}
}
function MemoEdit{
 $target=MemoCurrent;if(!$target){return};$old=MemoGet $target.root $target.layer $target.student $target.item
 $d=New-Object Windows.Forms.Form;$d.Text='셀 메모';$d.Size=New-Object Drawing.Size(620,390);$d.StartPosition='CenterParent'
 [void](Label $d ($target.name+' / '+$target.label) 18 18 565 32)
 $text=New-Object Windows.Forms.TextBox;$text.Multiline=$true;$text.AcceptsReturn=$true;$text.ScrollBars='Vertical';$text.SetBounds(18,58,565,220);$text.Text=[string]$old.text;$d.Controls.Add($text)
 [void](Label $d '메모는 즉시 저장됩니다. 점수·상위 반영에는 영향을 주지 않습니다.' 18 280 565 24)
 $ok=Button $d '메모 저장' 347 312 115 {$d.DialogResult='OK';$d.Close()};$cancel=Button $d '취소' 473 312 110 {$d.DialogResult='Cancel';$d.Close()};$d.CancelButton=$cancel
 Ui12Dialog $d;$d.Add_Shown({$text.Focus();$text.SelectionStart=$text.TextLength})
 try{if($d.ShowDialog($f) -eq 'OK'){MemoSet $target.root $target.layer $target.student $target.item $text.Text;$dashGrid.Invalidate();MemoRefreshList;$dashMessage.Text='셀 메모 저장 완료 · 점수는 변경하지 않았습니다.'}}finally{$d.Dispose()}
}
$memoEdit=$columnMenu.Items.Add('셀 메모 추가 / 수정');$memoEdit.Padding=New-Object Windows.Forms.Padding(3,4,3,4);$memoEdit.Add_Click({DSafe {MemoEdit}})
$memoDelete=$columnMenu.Items.Add('셀 메모 삭제');$memoDelete.Padding=New-Object Windows.Forms.Padding(3,4,3,4);$memoDelete.Add_Click({DSafe {$target=MemoCurrent;if($target){MemoSet $target.root $target.layer $target.student $target.item '';$dashGrid.Invalidate();MemoRefreshList;$dashMessage.Text='셀 메모를 삭제했습니다. 점수는 유지됩니다.'}}})
$columnMenu.Add_Opening({
 $target=MemoCurrent;$memoEdit.Enabled=(!$busy -and !$dashLoading -and $null -ne $target);$memoDelete.Enabled=$false
 if($memoEdit.Enabled){$memoDelete.Enabled=($null -ne (MemoGet $target.root $target.layer $target.student $target.item))}
})
$dashGrid.Add_CellPainting({param($sender,$e)
 if($dashLoading -or !$ws -or $e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0){return}
 $cell=$sender.Rows[$e.RowIndex].Cells[$e.ColumnIndex];$item=MemoCellItem $cell;if(!$item){return}
 $memo=MemoGet $ws $dashLayer ([string]$cell.OwningRow.Tag) $item;if(!$memo){return}
 if(!$e.Handled){$e.Paint($e.ClipBounds,[Windows.Forms.DataGridViewPaintParts]::All)}
 $r=$e.CellBounds;$points=[Drawing.Point[]]@([Drawing.Point]::new($r.Right-9,$r.Top+2),[Drawing.Point]::new($r.Right-2,$r.Top+2),[Drawing.Point]::new($r.Right-2,$r.Top+9))
 $brush=New-Object Drawing.SolidBrush([Drawing.Color]::FromArgb(155,44,44));try{$e.Graphics.FillPolygon($brush,$points)}finally{$brush.Dispose()};$e.Handled=$true
})

# A compact switch reveals a scrollable list without adding controls to every cell.
$memoToggle=New-Object Windows.Forms.CheckBox;$memoToggle.Text='메모 전체 보기';$memoToggle.AutoSize=$true;$memoToggle.SetBounds(102,2,120,24);$side.Controls.Add($memoToggle);$memoToggle.BringToFront()
$memoPanel=New-Object Windows.Forms.Panel;$memoPanel.BackColor=[Drawing.Color]::FromArgb(255,252,237);$memoPanel.BorderStyle='FixedSingle';$memoPanel.Visible=$false;$dashboard.Controls.Add($memoPanel)
$memoList=New-Object Windows.Forms.TextBox;$memoList.Multiline=$true;$memoList.ReadOnly=$true;$memoList.ScrollBars='Vertical';$memoList.Dock='Fill';$memoList.BackColor=$memoPanel.BackColor;$memoList.BorderStyle='None';$memoPanel.Padding=New-Object Windows.Forms.Padding(10);$memoPanel.Controls.Add($memoList)
function MemoListLayout{if($memoPanel.Visible){$memoPanel.SetBounds([math]::Max($dashGrid.Left,($dashGrid.Right-340)),$dashGrid.Top,330,[math]::Min(440,$dashGrid.Height));$memoPanel.BringToFront()}}
function MemoRefreshList{
 $memoPanel.Visible=($memoToggle.Checked -and $null -ne $dashSnapshot)
 if(!$memoPanel.Visible){return};MemoEnsure $ws;$lines=New-Object 'Collections.Generic.List[string]'
 foreach($row in @($dashGrid.Rows|Where-Object Visible)){foreach($col in @($dashGrid.Columns|Sort-Object DisplayIndex)){$item=MemoCellItem $row.Cells[$col.Index];if(!$item){continue};$m=MemoGet $ws $dashLayer ([string]$row.Tag) $item;if($m){$lines.Add(([string]$row.Cells['class'].Value+' '+$row.Cells['number'].Value+'번 '+$row.Cells['name'].Value+' / '+($col.HeaderText -replace "`r?`n",' ')));$lines.Add([string]$m.text);$lines.Add('')}}}
 $memoList.Text=if($lines.Count){$lines -join "`r`n"}else{'현재 작업표에 메모가 없습니다.'};$memoList.Select(0,0);MemoListLayout
}
$memoToggle.Add_CheckedChanged({DSafe {MemoRefreshList}});$dashboard.Add_Resize({MemoListLayout})
$memoShowAll=$columnMenu.Items.Add('메모 전체 보기');$memoShowAll.Add_Click({$memoToggle.Checked=$true})
$memoHideAll=$columnMenu.Items.Add('메모 전체 접기');$memoHideAll.Add_Click({$memoToggle.Checked=$false})
foreach($menu in @($memoShowAll,$memoHideAll)){$menu.Padding=New-Object Windows.Forms.Padding(3,4,3,4)}
$script:memoPreviousOnLoad=${function:NOnLoad};function NOnLoad{& $memoPreviousOnLoad;MemoRefreshList}
