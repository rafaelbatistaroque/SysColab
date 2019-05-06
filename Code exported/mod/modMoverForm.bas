Attribute VB_Name = "modMoverForm"
Option Compare Database
Option Explicit

'MOVER FORM
Public Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SYSCOMMAND = &H112

