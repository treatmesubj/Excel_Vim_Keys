Attribute VB_Name = "keycheck"
Option Explicit
#If VBA7 And Win64 Then '64bit
  Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer
#Else '32bit
  Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
#End If
Private Const VK_CONTROL = &H11
Private Const VK_SHIFT = &H10
Private Const VK_MENU = &H12

Sub keycheck()
'Ctrl
  If GetAsyncKeyState(VK_CONTROL) < 0 Then
    MsgBox "Control is pressed."
  End If
'Shift
  If GetAsyncKeyState(VK_SHIFT) < 0 Then
    MsgBox "Shift is pressed."
  End If
End Sub
