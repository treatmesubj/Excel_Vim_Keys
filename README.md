# Docs
- [Application.OnKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey)
- [Application.SendKeys](https://learn.microsoft.com/en-us/office/vba/api/excel.application.sendkeys)

## Set Up
- [Setting Up an Excel Add-In](https://trumpexcel.com/excel-add-in/)

## Disable Lotus Compatibility
`/` is mapped to `Alt` to remain compatible with ancient Lotus 123 for some reason. Disable via `File -> Options -> Advanced -> Lotus Compatibility`

## Reference
- [ExcelLikeVim](https://github.com/kjnh10/ExcelLikeVim)
- [xlpro.tips](https://xlpro.tips/posts/excel-and-vim/)

### Listen Keys
- [GetAsyncKeyState](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getasynckeystate)
    - [Snake](https://github.com/Vitosh/VBA_personal/blob/f5d8fb859e824df998fb2c21383f94c6d57f7242/Algorithms/Games/SnakeAttempt.vb)
```
Private Sub ReadKey()
    Select Case True
        Case GetAsyncKeyState(vbKeyUp):
            movingDirection = GoUp
        Case GetAsyncKeyState(vbKeyRight):
            movingDirection = GoRight
        Case GetAsyncKeyState(vbKeyDown):
            movingDirection = GoDown
        Case GetAsyncKeyState(vbKeyLeft):
            movingDirection = GoLeft
    End Select
End Sub
```
