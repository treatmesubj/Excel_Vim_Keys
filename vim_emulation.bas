Rem Attribute VBA_ModuleType=VBAModule
Option Explicit

' up, down, left, right
Public Sub go_up()
  Application.SendKeys "{UP}"
End Sub
Public Sub go_down()
  Application.SendKeys "{DOWN}"
End Sub
Public Sub go_left()
  Application.SendKeys "{LEFT}"
End Sub
Public Sub go_right()
  Application.SendKeys "{RIGHT}"
End Sub
' visual up, down, left, right
Public Sub visual_up()
  Application.SendKeys "+{UP}"
End Sub
Public Sub visual_down()
  Application.SendKeys "+{DOWN}"
End Sub
Public Sub visual_left()
  Application.SendKeys "+{LEFT}"
End Sub
Public Sub visual_right()
  Application.SendKeys "+{RIGHT}"
End Sub

' editing
Public Sub edit()
  Application.SendKeys "{F2}"
End Sub
Public Sub edit_begin()
  Call go_begin_of_row_values
  Call edit
  Application.SendKeys "{HOME}"
End Sub
Public Sub edit_end()
Application.ScreenUpdating = False
  Dim row As Long: Dim col As Long
  row = Selection.row
  Cells(row, 16384).Select
  col = Selection.End(xlToLeft).Column
  Cells(row, col).Select
  Call edit
  Application.SendKeys "{END}"
Application.ScreenUpdating = True
End Sub

' contiguous left, right
Public Sub go_contiguous_left()
  Application.SendKeys "^{LEFT}"
End Sub
Public Sub go_contiguous_right()
  Application.SendKeys "^{RIGHT}"
End Sub
Public Sub visual_contiguous_left()
  Application.SendKeys "^+{LEFT}"
End Sub
Public Sub visual_contiguous_right()
  Application.SendKeys "^+{RIGHT}"
End Sub

' insert rows 
Public Sub insert_row_above()
  Dim row As Long
  row = Selection.row
  Rows(row & ":" & row).EntireRow.Insert
  Cells(row, Selection.Column).Select
  Application.SendKeys "{F2}"
End Sub
Public Sub insert_row_below()
  Dim row As Long
  row = Selection.row
  Rows(row + 1 & ":" & row + 1).EntireRow.Insert
  Cells(row + 1, Selection.Column).Select
  Application.SendKeys "{F2}"
End Sub

' delete rows, cells
Public Sub delete_row()
  Dim row As Long
  row = Selection.row
  Rows(row & ":" & row).EntireRow.Delete
  Cells(row, Selection.Column).Select
  Application.SendKeys "{F2}"
End Sub
Public Sub delete_selected()
  Application.SendKeys "{DEL}"
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

Public Sub visual_begin_of_row(start_row As Long, start_col As Long)
Application.ScreenUpdating = False
  Dim start_range As Range: Set start_range = Cells(start_row, start_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim end_row As Long: Dim end_col As Long: Dim end_range As Range
  Cells(start_row, 1).Select
  end_col = Selection.End(xlToLeft).Column
  If top_row < start_row Then
    Set end_range = Cells(top_row, end_col) 
  Else
    Set end_range = Cells(bottom_row, end_col)
  End If
  Range(start_range, end_range).Select
  ' need to pivot anchor back to start
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  If top_row < start_row Then
    Application.SendKeys "^."
    If right_col > start_col Then
      Application.SendKeys "^."
    End If
    If left_col = start_col And left_col <> right_col Then
      Application.SendKeys "^."
    End If
  End If
  If left_col < start_col Then
    Application.SendKeys "^."
  End If
Application.ScreenUpdating = True
End Sub

Public Sub go_begin_of_row_values()
Application.ScreenUpdating = False
  Call go_begin_of_row
  Call go_contiguous_right
  If IsEmpty(Selection) Then
    Cells(Selection.row, 1).Select
  End If
Application.ScreenUpdating = True
End Sub

Public Sub go_end_of_row_values()
Application.ScreenUpdating = False
  Dim row As Long: Dim col As Long
  row = Selection.row
  Cells(row, 16384).Select
  col = Selection.End(xlToLeft).Column
  Cells(row, col).Select ' idk why Excel doesn't scroll. I have to below
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

Public Sub visual_end_of_row_values(start_row As Long, start_col As Long)
Application.ScreenUpdating = False
  Dim start_range As Range: Set start_range = Cells(start_row, start_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim end_row As Long: Dim end_col As Long: Dim end_range As Range
  Cells(start_row, 16384).Select
  end_col = Selection.End(xlToLeft).Column
  If top_row < start_row Then
    Set end_range = Cells(top_row, end_col) 
  Else
    Set end_range = Cells(bottom_row, end_col)
  End If
  Range(start_range, end_range).Select
  ' need to pivot anchor back to start
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  If top_row < start_row Then
    Application.SendKeys "^."
    If right_col > start_col Then
      Application.SendKeys "^."
    End If
    If left_col = start_col And left_col <> right_col Then
      Application.SendKeys "^."
    End If
  End If
  If left_col < start_col Then
    Application.SendKeys "^."
  End If
Application.ScreenUpdating = True
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
' because you annoyingly have to keep un/pressing <CONTROL> key
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
' will settle for annoying "+{PGUP}" for now
Public Sub visual_page_up()
  Application.SendKeys "+{PGUP}"
End Sub
Public Sub visual_page_down()
  Application.SendKeys "+{PGDN}"
End Sub

' search
Public Sub do_search()
  Application.SendKeys "^f"
End Sub

' copy
Public Sub copy_selected()
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
  Application.SendKeys "^z"
End Sub
Public Sub redo()
  Application.SendKeys "^y"
End Sub
