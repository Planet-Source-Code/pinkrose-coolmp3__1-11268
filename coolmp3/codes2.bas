Attribute VB_Name = "codes2"
Option Explicit
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&
                                        
                                        
                                       
Public Const WM_SYSCOMMAND = &H112
