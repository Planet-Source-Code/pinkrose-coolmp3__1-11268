VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "CoolMp3"
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   Picture         =   "main.frx":0000
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tm2 
      Interval        =   1
      Left            =   3720
      Top             =   360
   End
   Begin VB.Timer tm1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   840
   End
   Begin VB.PictureBox pallbuttons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      Picture         =   "main.frx":349E
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   5400
   End
   Begin VB.PictureBox pbuttons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   1230
      Width           =   2700
   End
   Begin VB.Image ipos 
      Height          =   120
      Left            =   465
      Picture         =   "main.frx":7428
      Top             =   1020
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Line ln2 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      X1              =   6
      X2              =   20
      Y1              =   34
      Y2              =   34
   End
   Begin VB.Label lbvolumex 
      BackStyle       =   0  'Transparent
      Height          =   885
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   210
   End
   Begin VB.Line ln1 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      X1              =   271
      X2              =   285
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label lbvolume 
      BackStyle       =   0  'Transparent
      Height          =   885
      Left            =   4065
      TabIndex        =   4
      Top             =   90
      Width           =   210
   End
   Begin MediaPlayerCtl.MediaPlayer mp1 
      Height          =   1215
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   615
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -2
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lbtitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   210
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3195
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim volrc, vollc
Dim v1, v2, v3, v4
Dim f1, f2, f3, f4
Dim t, b
Dim c1, c2, c3
Dim yc, yc1
Dim vol
Dim vol1
Dim pos
Dim cp
Dim h
Dim title


Public Sub openf()

Dim filebox As OPENFILENAME
Dim fname As String
Dim retval As Long


' Configure how the dialog box will look
filebox.lStructSize = Len(filebox)
filebox.hwndOwner = Me.hWnd
filebox.lpstrTitle = "Open File"
' The next line sets up the file types drop-box
filebox.lpstrFilter = "MP3 Files" & vbNullChar & "*.mp3" & vbNullChar & vbNullChar
filebox.lpstrFile = Space(255)
filebox.nMaxFile = 255
filebox.lpstrFileTitle = Space(255)
filebox.nMaxFileTitle = 255
' Allow only existing files and hide the read-only check box
filebox.flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY


retval = GetOpenFileName(filebox)

If retval <> 0 Then
  fname = Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
End If
filename = fname
title = fname
mp1.filename = fname
mp1.Stop
ipos.Visible = True

End Sub

Private Sub Form_Load()

Dim dummy As Long

AutoFormShape main, RGB(255, 0, 255)

dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)


mp1.AutoRewind = True
mp1.AutoStart = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim result As Long

If Button = 1 Then
  ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
  result& = SendMessage(Me.hWnd, &H112, &HF012, 0)  ' This tells the OS to pick up the form to be moved
End If

End Sub


Private Sub ipos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

c3 = 1
pos = X / 15
tm1.Enabled = False
If filename = "" Then Exit Sub

End Sub


Private Sub ipos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


If c3 = 1 Then
  If ipos.Left >= 230 Then
    ipos.Left = 230
    Exit Sub
  End If
  If ipos.Left <= 30 Then
    ipos.Left = 30
    Exit Sub
  End If
  ipos.Left = ipos.Left + ((X / 15) - pos)
  mp1.CurrentPosition = (mp1.Duration * ipos.Left) / 240
End If

'v1 = ((mp1.CurrentPosition / mp1.Duration)) * 100
'v2 = (Int(v1 * 199) / 100)
'ipos.Left = v2 + 31


End Sub


Private Sub ipos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

c3 = 0
tm1.Enabled = True
If filename = "" Then Exit Sub

End Sub


Private Sub lbvolume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

c1 = 1
yc = Y
ln1.Y1 = lbvolume.Top + yc / 15
ln1.Y2 = lbvolume.Top + yc / 15
End Sub


Private Sub lbvolume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If c1 = 1 Then
  yc = Y
  If Y <= 0 Then
    yc = 0
  End If
  If Y >= 60 * 15 Then
    yc = (60 * 15)
  End If
  vol = -((yc / 15) * Int(6000 / 60))
  mp1.Volume = vol
  ln1.Y1 = lbvolume.Top + yc / 15
  ln1.Y2 = lbvolume.Top + yc / 15
  
  lbtitle = vol
  
  
End If
End Sub


Private Sub lbvolume_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

c1 = 0

End Sub


Private Sub lbvolumex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

c2 = 1
yc1 = Y
ln2.Y1 = lbvolumex.Top + yc1 / 15
ln2.Y2 = lbvolumex.Top + yc1 / 15

End Sub


Private Sub lbvolumex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim dummy

If c2 = 1 Then
  'movet = Y
  yc1 = Y
  If Y <= 4 * 15 Then
    yc1 = 0
  End If
  If Y >= 60 * 15 Then
    yc1 = (60 * 15)
  End If
  
  volrc = Hex((yc1 / 15) * Int(65535 / 60))
  vollc = Hex((yc1 / 15) * Int(65535 / 60))
  vol1 = "&h" & Trim((volrc)) & Trim((vollc))
  'mp1.Volume = vol1
  dummy = waveOutSetVolume(0, vol1)
  ln2.Y1 = lbvolumex.Top + yc1 / 15
  ln2.Y2 = lbvolumex.Top + yc1 / 15
      
End If
End Sub


Private Sub lbvolumex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

c2 = 0

End Sub


Private Sub pallbuttons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lbtimer = X

End Sub


Private Sub pbuttons_Click()

If ctr = 1 Then
  mp1.CurrentPosition = 0
End If

If ctr = 2 Then
  If filename = "" Then Exit Sub
  mp1.Play
  
  tm1.Enabled = True
End If

If ctr = 3 Then
  If mp1.CurrentPosition <= 0 Then Exit Sub
  mp1.Pause
End If

If ctr = 4 Then
  mp1.Stop
  mp1.CurrentPosition = 0
  ipos.Left = 31
  v1 = 0: v2 = 0
  tm1.Enabled = False
End If

If ctr = 5 Then
  mp1.CurrentPosition = mp1.Duration
End If
If ctr = 6 Then
  openf
End If

End Sub

Private Sub pbuttons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim dummy As Long

If Button = 1 Then
  If X >= 0 And X <= 30 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 29, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)
   ctr = 1
  End If
  If X >= 31 And X <= 60 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 90, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)
   ctr = 2
  End If
  If X >= 61 And X <= 90 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 150, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)
   ctr = 3
  End If
  If X >= 91 And X <= 120 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 210, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)
   ctr = 4
  End If
  If X >= 121 And X <= 150 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 270, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)
   ctr = 5
  End If
  If X >= 151 And X <= 180 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 330, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)
   ctr = 6
  End If






End If

End Sub

Private Sub pbuttons_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim dummy As Long

If Button = 1 Then
  If ctr = 1 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)

  End If
  If ctr = 2 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)

  End If
  If ctr = 3 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)

  End If
  If ctr = 4 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)

  End If
  If ctr = 5 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)

  End If
  If ctr = 6 Then
    pbuttons.Cls
    dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 0, 0, 30, 15, pallbuttons.hdc, 0, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 31, 0, 30, 15, pallbuttons.hdc, 61, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 61, 0, 30, 15, pallbuttons.hdc, 121, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 91, 0, 30, 15, pallbuttons.hdc, 181, 0, SRCCOPY)
           dummy = BitBlt(pbuttons.hdc, 121, 0, 30, 15, pallbuttons.hdc, 241, 0, SRCCOPY)
           'dummy = BitBlt(pbuttons.hdc, 151, 0, 30, 15, pallbuttons.hdc, 301, 0, SRCCOPY)

  End If
  
End If

End Sub


Private Sub tm1_Timer()

If mp1.CurrentPosition <= 0 Then
  Exit Sub
End If

v1 = ((mp1.CurrentPosition / mp1.Duration)) * 100
v2 = (Int(v1 * 199) / 100)
ipos.Left = v2 + 31


End Sub


Private Sub tm2_Timer()

'simple thing,so do your best and change the big wow thing : )
'and dont forget to call me

h = h + 1
lbtitle.Caption = Mid(title, h, 45)
If h >= Len(title) Then
  h = 0
End If




End Sub
