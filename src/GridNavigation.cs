using System;using System.Linq;using System.Windows.Forms;
public class AssessmentGrid : DataGridView {
 public bool EditableTabNavigation {get;set;}
 public bool NavigateEditableTab(bool reverse){
  if(!EditableTabNavigation || CurrentCell==null) return false;
  if(IsCurrentCellInEditMode && !EndEdit()) return true;
  var cols=Columns.Cast<DataGridViewColumn>().Where(c=>c.Visible).OrderBy(c=>c.DisplayIndex).ToArray();
  if(cols.Length==0 || Rows.Count==0) return true;
  int start=CurrentCell.RowIndex*cols.Length+Array.IndexOf(cols,CurrentCell.OwningColumn), count=Rows.Count*cols.Length;
  for(int step=1;step<=count;step++){
   int pos=(start+(reverse?-step:step)+count)%count;
   var row=Rows[pos/cols.Length];var cell=row.Cells[cols[pos%cols.Length].Index];
   if(row.Visible && !row.IsNewRow && !cell.ReadOnly){CurrentCell=cell;ClearSelection();cell.Selected=true;return true;}
  }return true;
 }
 bool Tab(Keys keys){return EditableTabNavigation && (keys&Keys.KeyCode)==Keys.Tab && (keys&(Keys.Control|Keys.Alt))==0 && NavigateEditableTab((keys&Keys.Shift)!=0);}
 protected override bool ProcessDialogKey(Keys keys){return Tab(keys)||base.ProcessDialogKey(keys);}
 protected override bool ProcessDataGridViewKey(KeyEventArgs e){return Tab(e.KeyData)||base.ProcessDataGridViewKey(e);}
}