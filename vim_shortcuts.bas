Rem Attribute VBA_ModuleType=VBAModule
Sub setup_shortcuts()
  Application.OnKey "h", "go_left"
  Application.OnKey "j", "go_down"
  Application.OnKey "k", "go_up"
  Application.OnKey "l", "go_right"
  Application.OnKey "{BS}", "go_left"
  Application.OnKey " ", "go_right"

  Application.OnKey "i", "edit_cell"
  Application.OnKey "a", "edit_cell"
  Application.OnKey "+a", "edit_end"
  Application.OnKey "+i", "edit_begin"

  Application.OnKey "o", "insert_row_below"
  Application.OnKey "+o", "insert_row_above"

  'Application.OnKey "dd", "delete_row"
  Application.OnKey "x", "delete_selected"
  Application.OnKey "d", "cut_selected"
  Application.OnKey "r", "overwrite_cell"
  Application.OnKey "R", "overwrite_cell"

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
  Application.OnKey "+6", "go_begin_of_row_values" '^

  Application.OnKey "v", "setup_v_mode_shortcuts" 

  Application.OnKey "P", "paste_values"
  Application.OnKey "p", "paste"

  Application.OnKey "u", "undo"
  Application.OnKey "^r", "redo"

  Application.OnKey "/", "find_string"
  Application.OnKey "n", "find_next"
  Application.OnKey "+n", "find_prev"
End Sub

'Sub setup_i_mode_shortcuts()
'  Call teardown_shortcuts()
'  Application.OnKey "~", "enter_for_edit"
'  Application.OnKey "{TAB}", "tab_for_edit"
'  Application.OnKey "{ESC}", "esc_for_edit"
'End Sub
'
'Sub teardown_i_mode_shortcuts()
'  Application.OnKey "~"  'ENTER
'  Application.OnKey "{TAB}"
'  Application.OnKey "{ESC}"
'  Call setup_shortcuts
'End Sub

Sub setup_v_mode_shortcuts()
  Call teardown_shortcuts
  Dim anchor_row As Long: Dim anchor_col As Long
  anchor_row = Selection.Row: anchor_col = Selection.Column

  Application.OnKey "h", "visual_left"
  Application.OnKey "{BS}", "visual_left"
  Application.OnKey "j", "visual_down"
  Application.OnKey "k", "visual_up"
  Application.OnKey "l", "visual_right"
  Application.OnKey " ", "visual_right"

  Application.OnKey "b", "visual_contiguous_left"
  Application.OnKey "w", "visual_contiguous_right"
  Application.OnKey "e", "visual_contiguous_right"

  Application.OnKey "^u", "visual_page_up"
  Application.OnKey "^d", "visual_page_down"
  Application.OnKey "+4", "'visual_end_of_row_values """ & anchor_row & """, " & anchor_col & " '" 
  Application.OnKey "0", "'visual_begin_of_row """ & anchor_row & """, " & anchor_col & " '"
  Application.OnKey "+-", "'visual_begin_of_row_values """ & anchor_row & """, " & anchor_col & " '"
  Application.OnKey "+6", "'visual_begin_of_row_values """ & anchor_row & """, " & anchor_col & " '"

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
  Application.OnKey "x"
  Application.OnKey "d"
  Application.OnKey "r"
  Application.OnKey "R"

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
  Application.OnKey "+6"

  Application.OnKey "v"

  Application.OnKey "p"

  Application.OnKey "u"
  Application.OnKey "^r"

  Application.OnKey "/"
  Application.OnKey "n"
  Application.OnKey "+n"

  Application.OnKey "y"
  Application.OnKey "{ESC}"
 
End Sub
