' visual up
Public Sub visual_up(start_row As Long, start_col As Long)
  Dim start_range As Range: Set start_range = Cells(start_row, start_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  Dim end_range As Range
  If top_row > 1 And top_row < start_row Then
    If left_col < start_col Then
      Set end_range = Cells(top_row - 1, left_col)
    Else
      Set end_range = Cells(top_row - 1, right_col)
    End If
    Range(start_range, end_range).Select
  ElseIf top_row = 1 Then
    'pass
  Else
    If left_col < start_col Then
      Set end_range = Cells(bottom_row - 1, left_col)
    Else
      Set end_range = Cells(bottom_row - 1, right_col)
    End If
    Range(start_range, end_range).Select
  End If
End Sub

' visual down
Public Sub visual_down(start_row As Long, start_col As Long)
  Dim start_range As Range: Set start_range = Cells(start_row, start_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  Dim end_range As Range
  If bottom_row < 1048576 And bottom_row > start_row Then
    If left_col < start_col Then
      Set end_range = Cells(bottom_row + 1, left_col)
    Else
      Set end_range = Cells(bottom_row + 1, right_col)
    End If
    Range(start_range, end_range).Select
  ElseIf bottom_row = 1048576 Then
    'pass
  Else
    If left_col < start_col Then
      Set end_range = Cells(top_row + 1, left_col)
    Else
      Set end_range = Cells(top_row + 1, right_col)
    End If
    Range(start_range, end_range).Select
  End If
End Sub

' visual left
Public Sub visual_left(start_row As Long, start_col As Long)
  Dim start_range As Range: Set start_range = Cells(start_row, start_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  Dim end_range As Range
  If left_col > 1 And left_col < start_col Then
    If top_row < start_row Then
      Set end_range = Cells(top_row, left_col - 1)
    Else
      Set end_range = Cells(bottom_row, left_col - 1)
    End If
    Range(start_range, end_range).Select
  ElseIf left_col = 1 Then
    'pass
  Else
    If top_row < start_row Then
      Set end_range = Cells(top_row, right_col - 1)
    Else
      Set end_range = Cells(bottom_row, right_col - 1)
    End If
    Range(start_range, end_range).Select
  End If
End Sub

' visual right
Public Sub visual_right(start_row As Long, start_col As Long)
  Dim start_range As Range: Set start_range = Cells(start_row, start_col)
  Dim top_row As Long: top_row = Selection.row
  Dim bottom_row As Long: bottom_row = top_row + Selection.Rows.Count - 1
  Dim left_col As Long: left_col = Selection.Column
  Dim right_col As Long: right_col = Selection.Columns.Count + left_col - 1
  Dim end_range As Range
  If right_col < 116384 And right_col > start_col Then
    If top_row < start_row Then
      Set end_range = Cells(top_row, right_col + 1)
    Else
      Set end_range = Cells(bottom_row, right_col + 1)
    End If
    Range(start_range, end_range).Select
  ElseIf right_col = 116384 Then
    'pass
  Else
    If top_row < start_row Then
      Set end_range = Cells(top_row, left_col + 1)
    Else
      Set end_range = Cells(bottom_row, left_col + 1)
    End If
    Range(start_range, end_range).Select
  End If
End Sub

Application.OnKey "h", "'visual_left """ & start_row & """, " & start_col & " '"
Application.OnKey "{BS}", "'visual_left """ & start_row & """, " & start_col & " '"
Application.OnKey "j", "'visual_down """ & start_row & """, " & start_col & " '"
Application.OnKey "k", "simple_visual_up"
Application.OnKey "l", "'visual_right """ & start_row & """, " & start_col & " '"
Application.OnKey " ", "'visual_right """ & start_row & """, " & start_col & " '"

