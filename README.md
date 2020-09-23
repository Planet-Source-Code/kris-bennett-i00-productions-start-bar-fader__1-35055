<div align="center">

## Start Bar Fader

<img src="PIC2002524236587243.jpg">
</div>

### Description

This fades the start bar so it is semitransparent. Only works in Windows 2000+ - Don't know why I made this - Please post comments and rate this!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-05-24 20:51:48
**By**             |[Kris Bennett \(i00 Productions\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kris-bennett-i00-productions.md)
**Level**          |Intermediate
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Start\_Bar\_864955242002\.zip](https://github.com/Planet-Source-Code/kris-bennett-i00-productions-start-bar-fader__1-35055/archive/master.zip)

### API Declarations

```
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
```





