VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Function EnableTransparancy(ByVal hwnd As Long, Perc As Integer) As Long

Dim msg As Long

On Error Resume Next

If Perc < 0 Or Perc > 255 Then
    EnableTransparancy = 1
Else


    msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    msg = msg Or WS_EX_LAYERED
    
    SetWindowLong hwnd, GWL_EXSTYLE, msg
    
    SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
    
    EnableTransparancy = 0
End If

If Err Then
    EnableTransparancy = 2
End If
    
End Function

Private Sub Form_Load()
EnableTransparancy FindWindow("Shell_TrayWnd", vbNullString), 200
Unload Me
End Sub
