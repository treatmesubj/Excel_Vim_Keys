Rem Attribute VBA_ModuleType=VBAModule
Sub setup_shortcuts()
  Application.OnKey "h", "go_left"
  Application.OnKey "j", "go_down"
  Application.OnKey "k", "go_up"
  Application.OnKey "l", "go_right"
  Application.OnKey "{BS}", "go_left"
  Application.OnKey " ", "go_right"

  Application.OnKey "i", "edit"
  Application.OnKey "a", "edit"
  Application.OnKey "+a", "edit_end"
  Application.OnKey "+i", "edit_begin"

  Application.OnKey "o", "insert_row_below"
  Application.OnKey "+o", "insert_row_above"

  'Application.OnKey "dd", "delete_row"
  Application.OnKey "x", "delete_selected"

  Application.OnKey "b", "go_contiguous_left"
  Application.OnKey "w", "go_contiguous_right"
  Application.OnKey "e", "go_contiguous_right"

  Application.OnKey "+h", "go_top_of_viewport"
  Application.OnKey "^u", "page_up" 
  Application.OnKey "+l", "go_bottom_of_viewport"
  Application.OnKey "^d", "page_down"
  Application.OnKey "+4", "go_end_of_row_values" '$
  Application.OnKey "0", "go_begin_of_row"
  Application.OnKey "+-", "go_begin_of_row_values" '_

  Application.OnKey "v", "setup_v_mode_shortcuts" 

  Application.OnKey "P", "paste_values"
  Application.OnKey "p", "paste"

  Application.OnKey "u", "undo"
  Application.OnKey "^r", "redo"

  Application.OnKey "/", "do_search"
End Sub

Sub setup_v_mode_shortcuts()
  Call teardown_shortcuts
  Dim start_row As Long: Dim start_col As Long
  start_row = Selection.Row: start_col = Selection.Column

  Application.OnKey "h", "'visual_left """ & start_row & """, " & start_col & " '"
  Application.OnKey "{BS}", "'visual_left """ & start_row & """, " & start_col & " '"
  Application.OnKey "j", "'visual_down """ & start_row & """, " & start_col & " '"
  Application.OnKey "k", "'visual_up """ & start_row & """, " & start_col & " '"
  Application.OnKey "l", "'visual_right """ & start_row & """, " & start_col & " '"
  Application.OnKey " ", "'visual_right """ & start_row & """, " & start_col & " '"

  'Application.OnKey "b", "visual_contiguous_left"
  'Application.OnKey "w", "visual_contiguous_right"
  'Application.OnKey "e", "visual_contiguous_right"

  'Application.OnKey "+4", "visual_end_of_row_values" '$
  'Application.OnKey "0", "visual_begin_of_row"
  'Application.OnKey "+-", "visual_begin_of_row_values" '_

  Application.OnKey "x", "delete_selected"
  Application.OnKey "d", "cut_selected"
  Application.OnKey "y", "copy_selected"
  Application.OnKey "P", "paste_values"
  Application.OnKey "p", "paste"

  Application.OnKey "v", "teardown_v_mode_shortcuts"
  Application.OnKey "{ESC}", "teardown_v_mode_shortcuts"
End Sub

Sub teardown_v_mode_shortcuts()
  Application.OnKey "y"
  Application.OnKey "{ESC}"
  Application.OnKey "d"
  Call setup_shortcuts
End Sub

Sub teardown_shortcuts()
  Application.OnKey "h"
  Application.OnKey "j"
  Application.OnKey "k"
  Application.OnKey "l"
  Application.OnKey "{BS}"
  Application.OnKey " "

  Application.OnKey "i"
  Application.OnKey "a"
  Application.OnKey "+a"
  Application.OnKey "+i"

  Application.OnKey "o"
  Application.OnKey "+o"

  'Application.OnKey "dd"
  'Application.OnKey "dw"
  Application.OnKey "x"
  Application.OnKey "d"

  Application.OnKey "b"
  Application.OnKey "w"
  Application.OnKey "e"

  Application.OnKey "+h"
  Application.OnKey "^u"
  Application.OnKey "+l"
  Application.OnKey "^d"
  Application.OnKey "+4"
  Application.OnKey "0"
  Application.OnKey "+-"
  
  Application.OnKey "v"

  Application.OnKey "p"

  Application.OnKey "u"
  Application.OnKey "^r"

  Application.OnKey "/"

  Application.OnKey "y"
  Application.OnKey "{ESC}"
 
End Sub
