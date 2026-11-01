# Selection uses a border; fill is reserved for linked-score discrepancies.
$dashGrid.EditMode='EditOnKeystrokeOrF2'
$dashGrid.MultiSelect=$true
$dashGrid.EditableTabNavigation=$true
$dashGrid.DefaultCellStyle.SelectionBackColor=[Drawing.Color]::White
$dashGrid.DefaultCellStyle.SelectionForeColor=$navy
$dashSelected.BackColor=[Drawing.Color]::White
$dashSelected.Font=New-Object Drawing.Font($f.Font,[Drawing.FontStyle]::Bold)
$dashGrid.Add_CellPainting({param($sender,$e)
 if($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0 -or !$dashGrid.CurrentCell){return}
 $cell=$dashGrid.Rows[$e.RowIndex].Cells[$e.ColumnIndex];if(!$cell.Selected){return}
 $e.Paint($e.ClipBounds,[Windows.Forms.DataGridViewPaintParts]::All);$pen=New-Object Drawing.Pen($blue,1)
 try{$r=$e.CellBounds;$ri=$e.RowIndex;$ci=$e.ColumnIndex
  if($ri -eq 0 -or !$dashGrid.Rows[$ri-1].Cells[$ci].Selected){$e.Graphics.DrawLine($pen,$r.Left,$r.Top,$r.Right,$r.Top)}
  if($ri -eq $dashGrid.Rows.Count-1 -or !$dashGrid.Rows[$ri+1].Cells[$ci].Selected){$e.Graphics.DrawLine($pen,$r.Left,($r.Bottom-1),$r.Right,($r.Bottom-1))}
  $visible=@($dashGrid.Columns|Where-Object Visible|Sort-Object DisplayIndex);$position=[array]::IndexOf($visible,$cell.OwningColumn)
  if($position -le 0 -or !$dashGrid.Rows[$ri].Cells[$visible[$position-1].Index].Selected){$e.Graphics.DrawLine($pen,$r.Left,$r.Top,$r.Left,$r.Bottom)}
  if($position -eq $visible.Count-1 -or !$dashGrid.Rows[$ri].Cells[$visible[$position+1].Index].Selected){$e.Graphics.DrawLine($pen,($r.Right-1),$r.Top,($r.Right-1),$r.Bottom)}
  if($cell -eq $dashGrid.CurrentCell){$pen.Width=2;$e.Graphics.DrawRectangle($pen,($r.X+1),($r.Y+1),($r.Width-3),($r.Height-3))}
 }finally{$pen.Dispose()};$e.Handled=$true
})
$dashGrid.Add_ColumnHeaderMouseClick({param($sender,$e)if(!$dashLoading -and $e.ColumnIndex -ge 0 -and $dashGrid.Rows.Count){$row=if($dashGrid.CurrentCell){$dashGrid.CurrentCell.RowIndex}else{0};$dashGrid.CurrentCell=$dashGrid.Rows[$row].Cells[$e.ColumnIndex];WSelectedHeader;DSelection}})
$dashGrid.Add_Leave({$dashGrid.Invalidate()})
$dashGrid.ShowCellToolTips=$true
$renameColumn.Text='선택 항목 이름 바꾸기 (F2)'
$wRename.Text='이름 바꾸기 (F2)'
$dashGrid.Add_KeyDown({param($sender,$e)
 if($busy -or $dashLoading){return}
 if($e.KeyCode -eq [Windows.Forms.Keys]::F2){
  $e.Handled=$true;$e.SuppressKeyPress=$true
  if($dashGrid.IsCurrentCellInEditMode){[void]$dashGrid.EndEdit()}
  if($dashGrid.CurrentCell -and $dashGrid.CurrentCell.OwningColumn.Tag.item){$renameColumn.PerformClick()}else{$dashMessage.Text='이름을 바꿀 항목의 점수·기록 셀을 선택한 뒤 F2를 누르세요. 학생 성명·번호는 이 기능으로 바꾸지 않습니다.'}
 }elseif($e.KeyCode -eq [Windows.Forms.Keys]::Enter -and !$dashGrid.IsCurrentCellInEditMode){
  $e.Handled=$true;$e.SuppressKeyPress=$true;[void]$dashGrid.BeginEdit($true)
 }
})
$wTree.Add_KeyDown({param($sender,$e)if(!$busy -and $e.KeyCode -eq [Windows.Forms.Keys]::F2){$e.Handled=$true;$e.SuppressKeyPress=$true;$wRename.PerformClick()}})

$dashGrid.Add_CellFormatting({param($sender,$e)
 if($dashLoading -or $e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0){return}
 $col=$dashGrid.Columns[$e.ColumnIndex];$item=if($col.Name -eq 'total'){'__sum__'}else{$col.Tag.item}
 $key=$dashLayer+'|'+$dashGrid.Rows[$e.RowIndex].Tag+'|'+$item;$different=if($item){$script:lightMismatch[$key]}else{''}
 $color=LightMismatchColor $key
 if($col.Name -in @('class','number','name','view_code') -and $dashGrid.CurrentCell -and $dashGrid.CurrentCell.RowIndex -eq $e.RowIndex){$color=[Drawing.Color]::FromArgb(239,242,246)}
 $e.CellStyle.BackColor=$color;$e.CellStyle.SelectionBackColor=$color
 if($item){$tip=if($different){$different}else{''};if(Test-Path Function:\MemoTooltip){$tip=MemoTooltip $dashLayer ([string]$dashGrid.Rows[$e.RowIndex].Tag) $item $tip};$dashGrid.Rows[$e.RowIndex].Cells[$e.ColumnIndex].ToolTipText=$tip}
})

$script:identityHighlightRow=-1
$dashGrid.Add_CurrentCellChanged({
 $next=if($dashGrid.CurrentCell){$dashGrid.CurrentCell.RowIndex}else{-1}
 if($identityHighlightRow -ne $next){if($identityHighlightRow -ge 0 -and $identityHighlightRow -lt $dashGrid.Rows.Count){$dashGrid.InvalidateRow($identityHighlightRow)};if($next -ge 0 -and $next -lt $dashGrid.Rows.Count){$dashGrid.InvalidateRow($next)};$script:identityHighlightRow=$next}
})
