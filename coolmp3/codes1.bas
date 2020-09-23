Attribute VB_Name = "codes1"

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&     ' This setting is not in your API viewer, not sure why.
                                        ' If you use SC_MOVE then the mouse moves to the title bar
                                        ' and then moves the form, which makes forms with no title bar
                                        ' to not work.
Public Const WM_SYSCOMMAND = &H112

Dim CurRgn, TempRgn As Long  ' Region variables
Public Function AutoFormShape(bg As Form, transColor)
Dim X, Y As Integer



CurRgn = CreateRectRgn(0, 0, bg.Width / 15, bg.Height / 15)

While Y <= bg.Height / 15
    While X <= bg.Width / 15
        If GetPixel(bg.hdc, X, Y) = transColor Then
            TempRgn = CreateRectRgn(X, Y, X + 1, Y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
        End If
        X = X + 1
        
    Wend
        Y = Y + 1
        X = 0
        
Wend
success = SetWindowRgn(bg.hWnd, CurRgn, True)
DeleteObject (CurRgn)


 
End Function


