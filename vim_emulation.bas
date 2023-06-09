Rem Attribute VBA_ModuleType=VBAModule
Option Explicit

' up, down, left, right
Public Sub go_up()
  Application.SendKeys "{UP}", True
End Sub
Public Sub go_down()
  Application.SendKeys "{DOWN}", True
End Sub
Public Sub go_left()
  Application.SendKeys "{LEFT}", True
End Sub
Public Sub go_right()
  Application.SendKeys "{RIGHT}", True
End Sub
' visual up, down, left, right
Public Sub visual_up()
  Application.SendKeys "+{UP}", True
End Sub
Public Sub visual_down()
  Application.SendKeys "+{DOWN}", True
End Sub
Public Sub visual_left()
  Application.SendKeys "+{LEFT}", True
End Sub
Public Sub visual_right()
  Application.SendKeys "+{RIGHT}", True
End Sub

' editing
Public Sub edit_cell()
  Application.SendKeys "{F2}", True
End Sub
Public Sub edit_begin()
  ' TODO: don't sequencially call sendkeys
  Call go_begin_of_row_values
  Selection.Offset(0, -1).Select
  Application.SendKeys "{F2}", True
  Application.SendKeys "{HOME}", True
End Sub
Public Sub edit_end()
  ' TODO: don't sequencially call sendkeys
  Call go_end_of_row_values
  Selection.Offset(0, 1).Select
  Application.SendKeys "{F2}", True
  Application.SendKeys "{HOME}", True
End Sub
Public Sub overwrite_cell()
  Selection.Clear
  Application.SendKeys "{F2}", True
End Sub

' contiguous left, right
Public Sub go_contiguous_left()
  Application.SendKeys "^{LEFT}", True
End Sub
Public Sub go_contiguous_right()
  Application.SendKeys "^{RIGHT}", True
End Sub
Public Sub visual_contiguous_left()
  Application.SendKeys "^+{LEFT}", True
End Sub
Public Sub visual_contiguous_right()
  Application.SendKeys "^+{RIGHT}", True
End Sub

' insert rows 
Public Sub insert_row_above()
  Dim row As Long
  row = Selection.row
  Rows(row & ":" & row).EntireRow.Insert
  Cells(row, Selection.Column).Select
  Application.SendKeys "{F2}", True
End Sub
Public Sub insert_row_below()
  Dim row As Long
  row = Selection.row
  Rows(row + 1 & ":" & row + 1).EntireRow.Insert
  Cells(row + 1, Selection.Column).Select
  Application.SendKeys "{F2}", True
End Sub

' delete rows, cells
Public Sub delete_row()
  Dim row As Long
  row = Selection.row
  Rows(row & ":" & row).EntireRow.Delete
  Cells(row, Selection.Column).Select
  Application.SendKeys "{F2}", True
End Sub
Public Sub delete_selected()
  Selection.Clear
End Sub
Public Sub vis_delete_selected()
  Selection.Clear
  Call teardown_v_mode_shortcuts
End Sub

' rotate Selection's active cell/anchor clockwise 1 corner
Public Sub rotate_anchor()
  If ActiveCell.Row = Selection.Row And ActiveCell.Column < Selection.Column + Selection.Columns.Count - 1 Then
    Selection.Cells(1, Selection.Columns.Count).Activate
  ElseIf ActiveCell.Row = Selection.Row Then
    Selection.Cells(Selection.Rows.Count, Selection.Columns.Count).Activate
  ElseIf (ActiveCell.Row = (Selection.Row + Selection.Rows.Count - 1)) And (Selection.Column < ActiveCell.Column) Then
    Selection.Cells(Selection.Rows.Count, 1).Activate
  Else
    Selection.Cells(1, 1).Activate
  End If
End Sub

' auto pivot anchor back to pre-action corner
Public Sub auto_rotate_anchor(anchor_row As Long, anchor_col As Long, left_col As Long, right_col As Long, top_row As Long, bottom_row As Long)
  If top_row < anchor_row Then
    Call rotate_anchor
    If right_col > anchor_col Then
      Call rotate_anchor
    End If
    If left_col = anchor_col And left_col <> right_col Then
      Call rotate_anchor
    End If
  End If
  If left_col < anchor_col Then
    Call rotate_anchor
  End If
End Sub

'
' big movements
'
Public Sub go_top_of_viewport()
  Dim w As Window: Set w = ActiveWindow
  Cells(w.ScrollRow, Selection.Column).Select
End Sub

Public Sub go_begin_of_row()
  Dim row As Long: row = Selection.row
  Cells(row, 1).Select
End Sub

Public Sub screen_scroller(anchor_col As Long, end_col As Long)
  Dim vis_left As Long: Dim vis_width As Long: Dim vis_right As Long
  vis_left = ActiveWindow.VisibleRange.Column
  vis_width = ActiveWindow.VisibleRange.Columns.Count - 1
  vis_right = vis_left + vis_width
  If Not (vis_left < end_col And end_col < vis_right) Then
    If end_col > anchor_col Then ' to right
      ActiveWindow.ScrollColumn = end_col - vis_width + 2
    Else                         ' to left
      ActiveWindow.ScrollColumn = end_col
    End If
  End If
End Sub

Public Sub visual_begin_of_row(anchor_row As Long, anchor_col As Long)
Application.ScreenUpdating = False
  Dim anchor_range As Range: Set anchor_range = Cells(anchor_row, anchor_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim end_row As Long: Dim end_col As Long: Dim end_range As Range
  end_col = 1
  If top_row < anchor_row Then
    Set end_range = Cells(top_row, end_col) 
  Else
    Set end_range = Cells(bottom_row, end_col)
  End If
  Range(anchor_range, end_range).Select
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  Call auto_rotate_anchor(anchor_row, anchor_col, left_col, right_col, top_row, bottom_row)
  Call screen_scroller(anchor_col, end_col)
Application.ScreenUpdating = True
End Sub

Public Sub go_begin_of_row_values()
If IsEmpty(Cells(Selection.Row, 1)) Then
  If IsEmpty(Cells(Selection.Row, 1).End(xlToRight)) Then
    Cells(Selection.Row, 1).Select 
  Else
    Cells(Selection.Row, 1).End(xlToRight).Select
  End If
Else
  Cells(Selection.Row, 1).Select 
End If
End Sub

Public Sub visual_begin_of_row_values(anchor_row As Long, anchor_col As Long)
Application.ScreenUpdating = False
  Dim anchor_range As Range: Set anchor_range = Cells(anchor_row, anchor_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim end_row As Long: Dim end_col As Long: Dim end_range As Range
  Dim sel_right_col As Long
  sel_right_col = Cells(anchor_row, 1).End(xlToRight).Column
  If IsEmpty(Cells(anchor_row, sel_right_col)) Then ' nothing in row
    end_col = 1
  Else
    end_col = sel_right_col
  End If
  If top_row < anchor_row Then
    Set end_range = Cells(top_row, end_col) 
  Else
    Set end_range = Cells(bottom_row, end_col)
  End If
  Range(anchor_range, end_range).Select
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  Call auto_rotate_anchor(anchor_row, anchor_col, left_col, right_col, top_row, bottom_row)
  Call screen_scroller(anchor_col, end_col)
Application.ScreenUpdating = True
End Sub

Public Sub go_end_of_row_values()
Application.ScreenUpdating = False
  Dim row As Long: Dim col As Long
  row = Selection.row
  col = Cells(row, 16384).End(xlToLeft).Column
  Cells(row, col).Select ' Excel doesn't scroll to end but it is nice to see, so I do below
  Dim vis_left As Long: Dim vis_width As Long: Dim vis_right As Long
  vis_left = ActiveWindow.VisibleRange.Column
  vis_width = ActiveWindow.VisibleRange.Columns.Count - 1
  vis_right = vis_left + vis_width
  If Not (vis_left < col And col < vis_right) Then
    If col > vis_width Then
      ActiveWindow.ScrollColumn = col - vis_width + 2
    Else
      ActiveWindow.ScrollColumn = col
    End If
  End If
Application.ScreenUpdating = True
End Sub

Public Sub visual_end_of_row_values(anchor_row As Long, anchor_col As Long)
Application.ScreenUpdating = False
  Dim anchor_range As Range: Set anchor_range = Cells(anchor_row, anchor_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim end_row As Long: Dim end_col As Long: Dim end_range As Range
  end_col = Cells(anchor_row, 16384).End(xlToLeft).Column
  If top_row < anchor_row Then
    Set end_range = Cells(top_row, end_col) 
  Else
    Set end_range = Cells(bottom_row, end_col)
  End If
  Range(anchor_range, end_range).Select
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  Call auto_rotate_anchor(anchor_row, anchor_col, left_col, right_col, top_row, bottom_row)
  Call screen_scroller(anchor_col, end_col)
Application.ScreenUpdating = True
End Sub

' TODO, just del everything right of selected?
Public Sub del_end_of_row_values()
  Dim anchor_row As Long: Dim anchor_col As Long: Dim end_col As Long
  anchor_row = Selection.Row: anchor_col = Selection.Column
  end_col = Cells(anchor_row, 16384).End(xlToLeft).Column
  If end_col >= anchor_col Then 
    Call visual_end_of_row_values(anchor_row, anchor_col)
    Selection.Clear
  End If
  Cells(anchor_row, anchor_col).Select
End Sub

Public Sub go_bottom_of_viewport()
Application.ScreenUpdating = False
  Dim new_view_row As Long, old_view_row As Long
  Dim w As Window: Set w = ActiveWindow
  old_view_row = w.ScrollRow
  w.LargeScroll Down:=1
  new_view_row = w.ScrollRow: w.ScrollRow = old_view_row
  Cells(new_view_row - 1, Selection.Column).Select
Application.ScreenUpdating = True
End Sub

' page up, down
' cannot use simple Application.SendKeys "{PGUP}"
' else you annoyingly have to keep un/pressing <CONTROL> key
Public Sub page_up()
Application.ScreenUpdating = False
  Dim row As Long: Dim col As Long: Dim rows_down As Long
  row = Selection.row: col = Selection.Column
  rows_down = row - ActiveWindow.ScrollRow
  ActiveWindow.LargeScroll Down:=-1
  Cells(ActiveWindow.ScrollRow + rows_down, col).Select
Application.ScreenUpdating = True
End Sub
Public Sub page_down()
Application.ScreenUpdating = False
  Dim row As Long: Dim col As Long: Dim rows_down As Long
  row = Selection.row: col = Selection.Column
  rows_down = row - ActiveWindow.ScrollRow
  ActiveWindow.LargeScroll Down:=1
  Cells(ActiveWindow.ScrollRow + rows_down, col).Select
Application.ScreenUpdating = True
End Sub
' visual page up, down
' TODO: replace annoying "+{PGUP}", easy to accidentally delete
Public Sub visual_page_up()
  Application.SendKeys "+{PGUP}", True
End Sub
Public Sub visual_page_down()
  Application.SendKeys "+{PGDN}", True
End Sub

' copy
Public Sub vis_copy_selected()
  Selection.Copy
  Call teardown_v_mode_shortcuts
End Sub

' cut
Public Sub cut_selected()
  Selection.Cut
  Call teardown_v_mode_shortcuts
End Sub

' paste
Public Sub paste_values()
  If Application.CutCopyMode Then
    On Error Resume Next
    Selection.PasteSpecial Paste:=xlPasteValues
  End If
  Call teardown_v_mode_shortcuts
End Sub
Public Sub paste()
  If Application.CutCopyMode Then
    ActiveSheet.paste
  End If
  Call teardown_v_mode_shortcuts
End Sub

' undo & redo
Public Sub undo()
  Application.SendKeys "^z", True
End Sub
Public Sub redo()
  Application.SendKeys "^y", True
End Sub

' searching
Sub find_string()
  Dim obj As Object: Dim search_str As String
  search_str = InputBox("/", "search string", "")
  If search_str = "" Then
    Exit Sub
  End If
  Set obj = ActiveSheet.cells.find(what:=search_str, lookat:=xlPart)
  If Not obj Is Nothing Then
    obj.Activate
  Else
    MsgBox "not found"
  End If
End Sub
Function find_next()
  Dim match As Range
  Set match = cells.findNext(After:=ActiveCell)
  If Not match Is Nothing Then
    match.Activate
  End If
End Function
Function find_prev()
  Dim match As Range
  Set match = cells.findPrevious(After:=ActiveCell)
  If Not match Is Nothing Then
    match.Activate
  End If
End Function
