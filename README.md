<div align="center">

## Drag form without titlebar


</div>

### Description

Drag a form that has no titlebar! Add the routine listed below, and call it in the 'MouseDown' event of the form (or a control on the form): MoveWindow Me.Hwnd
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kamilche](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kamilche.md)
**Level**          |Beginner
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kamilche-drag-form-without-titlebar__1-30622/archive/master.zip)

### API Declarations

```
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
```


### Source Code

```
Public Sub MoveWindow(TheHwnd As Long)
  'Drag the form with the mouse
  ReleaseCapture
  SendMessage TheHwnd, &HA1, 2, 0&
End Sub
```

