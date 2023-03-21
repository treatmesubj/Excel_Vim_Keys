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
  Dim top_row As Long: Dim left_col As Long
  top_row = Selection.row: left_col = Selection.Column
  Dim start_range as Range: Set start_range = Selection
  If top_row > 1 Then
    Dim end_range as Range: Set end_range = Cells(top_row - 1, left_col)
    Range(start_range, end_range).Select
  End If
End Sub
Public Sub visual_down()
  Dim top_row As Long: Dim left_col As Long
  top_row = Selection.row: left_col = Selection.Column
  Dim bottom_row As Long
  bottom_row = top_row + Selection.Rows.Count - 1 
  Dim start_range as Range: Set start_range = Selection
  If bottom_row < 1048576 Then
    Dim end_range as Range: Set end_range = Cells(bottom_row + 1, left_col)
    Range(start_range, end_range).Select
  End If
End Sub
Public Sub visual_left()
  Dim top_row As Long: Dim left_col As Long
  top_row = Selection.row: left_col = Selection.Column
  Dim start_range as Range: Set start_range = Selection
  If left_col > 1 Then
    Dim end_range as Range: Set end_range = Cells(top_row, left_col - 1)
    Range(start_range, end_range).Select
  End If
End Sub
Public Sub visual_right()
  Dim top_row As Long: Dim left_col As Long
  top_row = Selection.row: left_col = Selection.Column
  Dim right_col As Long
  right_col = Selection.Columns.Count + left_col - 1
  Dim start_range as Range: Set start_range = Selection
  If right_col < 116384 Then
    Dim end_range as Range: Set end_range = Cells(top_row, right_col + 1)
    Range(start_range, end_range).Select
  End If
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
  Dim row As Long: Dim col As Long
  row = Selection.row
  Cells(row, 16384).Select
  col = Selection.End(xlToLeft).Column
  Cells(row, col).Select
  Call edit
  Application.SendKeys "{END}"
End Sub

' contiguous left, right
Public Sub go_contiguous_left()
  Dim row As Long: Dim col As Long
  row = Selection.row: col = Selection.End(xlToLeft).Column ' this row, next contig cell to left
  Cells(row, col).Select: col = Selection.Column
  If (IsEmpty(Selection) Or col = 1) And row > 1 Then ' if nothing
    Cells(row - 1, 16384).Select ' go up row to right edge
    If IsEmpty(Selection) Then ' next contig cell to left
      row = Selection.row: col = Selection.End(xlToLeft).Column
      Cells(row, col).Select
    End If
  End If
End Sub
Public Sub go_contiguous_right()
  Dim row As Long: Dim col As Long
  row = Selection.row: col = Selection.End(xlToRight).Column ' this row, next contig cell to right
  Cells(row, col).Select: col = Selection.Column
  If (IsEmpty(Selection) Or col = 16384) Then ' if nothing
    Cells(row+1, 1).Select ' go down row to left edge
    If IsEmpty(Selection) Then ' next contig cell to right
      row = Selection.row: col = Selection.End(xlToRight).Column
      Cells(row, col).Select
      If IsEmpty(Selection) Then ' if row empty, stay at left edge
        Cells(row, 1).Select
      End If
    End If
  End If
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
Public Sub delete_cell()
  Application.SendKeys "{DEL}"
End Sub

' big movements
Public Sub go_top_of_viewport()
  Dim w As Window
  Set w = ActiveWindow
  Cells(w.ScrollRow, Selection.Column).Select
End Sub
Public Sub go_begin_of_row()
  Dim row As Long
  row = Selection.row
  Cells(row, 1).Select
End Sub
Public Sub go_begin_of_row_values()
  Call go_begin_of_row
  Dim row As Long: Dim col As Long
  row = Selection.row: col = Selection.End(xlToRight).Column
  Cells(row, col).Select
  If IsEmpty(Selection) Then
    Cells(row, 1).Select
  End If
End Sub
Public Sub go_end_of_row_values()
  Dim row As Long: Dim col As Long
  row = Selection.row
  Cells(row, 16384).Select
  col = Selection.End(xlToLeft).Column
  Cells(row, col).Select
End Sub
Public Sub go_bottom_of_viewport()
  Dim new_view_row As Long, old_view_row As Long
  Dim w As Window
  Set w = ActiveWindow
  old_view_row = w.ScrollRow
  Application.ScreenUpdating = False
  w.LargeScroll Down:=1
  new_view_row = w.ScrollRow
  w.ScrollRow = old_view_row
  Application.ScreenUpdating = True
  Cells(new_view_row - 1, Selection.Column).Select
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

' paste
Public Sub paste_values()
  If Application.CutCopyMode Then
    Selection.PasteSpecial Paste:=xlPasteValues
  End If
End Sub
