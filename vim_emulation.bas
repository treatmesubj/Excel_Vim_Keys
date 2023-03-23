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
  Call go_begin_of_row
  Call go_contiguous_right
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

' big movements
Public Sub go_top_of_viewport()
  Dim w As Window: Set w = ActiveWindow
  Cells(w.ScrollRow, Selection.Column).Select
End Sub
Public Sub go_begin_of_row()
  Dim row As Long: row = Selection.row
  Cells(row, 1).Select
End Sub
Public Sub go_begin_of_row_values()
Application.ScreenUpdating = False 
  Call go_begin_of_row
  Dim row As Long: Dim col As Long
  row = Selection.row: col = Selection.End(xlToRight).Column
  Cells(row, col).Select
  If IsEmpty(Selection) Then
    Cells(row, 1).Select
  End If
Application.ScreenUpdating = True
End Sub
Public Sub go_end_of_row_values()
Application.ScreenUpdating = False
  Dim row As Long: Dim col As Long
  row = Selection.row
  Cells(row, 16384).Select
  col = Selection.End(xlToLeft).Column
  Cells(row, col).Select
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
Public Sub page_up()
  Application.SendKeys "{PGUP}"
End Sub
Public Sub page_down()
  Application.SendKeys "{PGDN}"
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
