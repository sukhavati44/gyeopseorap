$script:columnSearches=@{};$script:columnToolTarget=$null
function ColumnApplySearch{
 $dashGrid.CurrentCell=$null
 foreach($row in $dashGrid.Rows){$show=$true;foreach($col in $dashGrid.Columns){$query=$columnSearches[$dashLayer+'|'+$col.Name];$text=[string]$row.Cells[$col.Index].Value;if($col.Name -eq 'number' -and $text -match '^\d+$'){$text=([int]$text).ToString('00')};if($query -and $text.IndexOf($query,[StringComparison]::OrdinalIgnoreCase) -lt 0){$show=$false;break}};$row.Visible=$show}
 $dashMessage.Text='검색 결과 '+@($dashGrid.Rows|Where-Object Visible).Count+' / '+$dashGrid.Rows.Count+'명 · 열 제목 ▼에서 정렬·검색 해제';MemoRefreshList
}
function ColumnToolsPrepare{
 foreach($col in $dashGrid.Columns){$col.SortMode='Programmatic';$col.HeaderCell.Style.Padding=New-Object Windows.Forms.Padding(0,0,22,0)}
 if(@($columnSearches.Keys|Where-Object {$_.StartsWith($dashLayer+'|')})){ColumnApplySearch};WSelectedHeader
}
$columnTools=New-Object Windows.Forms.ContextMenuStrip
$ascending=$columnTools.Items.Add('오름차순');$ascending.Add_Click({DSafe {$dashGrid.Sort($columnToolTarget,[ComponentModel.ListSortDirection]::Ascending);$dashGrid.Invalidate()}})
$descending=$columnTools.Items.Add('내림차순');$descending.Add_Click({DSafe {$dashGrid.Sort($columnToolTarget,[ComponentModel.ListSortDirection]::Descending);$dashGrid.Invalidate()}})
$search=$columnTools.Items.Add('이 열 검색');$search.Add_Click({DSafe {$key=$dashLayer+'|'+$columnToolTarget.Name;$q=Input-Name '열 검색' '포함할 내용을 입력하세요. 빈칸으로 확인하면 이 열 검색을 해제합니다.' ([string]$columnSearches[$key]);if($q){$columnSearches[$key]=$q}else{$columnSearches.Remove($key)};ColumnApplySearch}})
$clear=$columnTools.Items.Add('이 열 검색 해제');$clear.Add_Click({$columnSearches.Remove($dashLayer+'|'+$columnToolTarget.Name);ColumnApplySearch})
$clearAll=$columnTools.Items.Add('모든 열 검색 해제');$clearAll.Add_Click({$script:columnSearches=@{};ColumnApplySearch})
$dashGrid.Add_ColumnHeaderMouseClick({param($sender,$e)
 if($busy -or $dashLoading -or $e.ColumnIndex -lt 0){return};$script:columnToolTarget=$sender.Columns[$e.ColumnIndex]
 if($e.X -lt ($columnToolTarget.Width-24)){
  if($columnToolTarget.Tag.item -and @(WColumnChildren $columnToolTarget.Tag.item).Count){$dashMessage.Text='연결된 열: 머리글 더블클릭 → 하위 작업공간 이동 / ▼ → 정렬·검색'}
  else{$dashMessage.Text='현재 표에서 직접 입력하는 열입니다. ▼를 누르면 정렬·검색할 수 있습니다.'}
  return
 }
 $columnTools.Show($sender,[Drawing.Point]::new($e.X+$sender.GetCellDisplayRectangle($e.ColumnIndex,-1,$true).X,$sender.ColumnHeadersHeight))
})
$dashGrid.Add_SortCompare({param($sender,$e)
 $numeric=($e.Column.Name -eq 'number' -or $e.Column.Name -eq 'total' -or $e.Column.Tag.item -in @(LightNumericItems $dashLayer));$a=0.0;$b=0.0
 if($numeric -and [double]::TryParse([string]$e.CellValue1,[ref]$a) -and [double]::TryParse([string]$e.CellValue2,[ref]$b)){$e.SortResult=$a.CompareTo($b)}else{$e.SortResult=[string]::Compare([string]$e.CellValue1,[string]$e.CellValue2,[StringComparison]::OrdinalIgnoreCase)};$e.Handled=$true
})
$dashGrid.Add_CellFormatting({param($sender,$e)
 if($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0 -and $sender.Columns[$e.ColumnIndex].Name -eq 'number'){$n=0;if([int]::TryParse([string]$e.Value,[ref]$n)){$e.Value=$n.ToString('00');$e.FormattingApplied=$true}}
})
$script:columnToolsPreviousLoad=${function:NOnLoad};function NOnLoad{& $columnToolsPreviousLoad;ColumnToolsPrepare}
$wModeButton.Add_Click({ColumnToolsPrepare})

$dashGrid.Add_CellPainting({param($sender,$e)
 if($e.RowIndex -ne -1 -or $e.ColumnIndex -lt 0){return};$e.Paint($e.ClipBounds,[Windows.Forms.DataGridViewPaintParts]::All);$r=$e.CellBounds;$x=$r.Right-14;$y=$r.Top+([int]($r.Height/2));$brush=New-Object Drawing.SolidBrush($navy)
 try{$e.Graphics.FillPolygon($brush,[Drawing.Point[]]@([Drawing.Point]::new($x-4,$y-2),[Drawing.Point]::new($x+4,$y-2),[Drawing.Point]::new($x,$y+3)))}finally{$brush.Dispose()};$e.Handled=$true
})
