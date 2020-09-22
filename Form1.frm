VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2280
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   2760
      Top             =   2640
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2265
      Left            =   600
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   0
      Top             =   75
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dhwnd As Long, dhdc As Long
Dim mouse As PointAPI
Dim ZF As Integer
Private Sub Form_Load()
Dim lReigon&, lResult&
If Me.Picture <> 0 Then Call SetAutoRgn(Me)
lReigon& = CreateEllipticRgn(0, 0, 157, 153)
lResult& = SetWindowRgn(Picture1.hwnd, lReigon&, True)
dhwnd = GetDesktopWindow
dhdc = GetDC(dhwnd)
AlwaysOnTop Me, True ' Use this as the call to this fuction.
ZF = 200
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
  Timer1.Enabled = False
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&
  Timer1.Enabled = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
  Timer1.Enabled = False
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&
  Timer1.Enabled = True
End If
End Sub


Private Sub Timer1_Timer()
Dim w As Long, h As Long, sw As Long, sh As Long, x As Long, y As Long
If Form1.Visible = True Then
  GetCursorPos mouse        ' capture mouse-position
  w = Picture1.ScaleWidth   ' destination width
  h = Picture1.ScaleHeight  ' destination height
  sw = w * (1 / (ZF / 100))
  sh = h * (1 / (ZF / 100))
  x = mouse.x - sw / 2       ' x source position (center to destination)
  y = mouse.y - sh / 2       ' y source position (center to destination)
  Picture1.Cls                ' clean picturebox
  StretchBlt Picture1.hdc, 0, 0, w, h, dhdc, x, y, sw, sh, SRCCOPY
  ' copy desktop (source) and strech to picturebox (destination)
End If
End Sub


Private Sub Timer2_Timer()
Dim p As Long
If GetAsyncKeyState(VK_F09) Then p = ShowWindow(Me.hwnd, SW_HIDE)
If GetAsyncKeyState(VK_F10) Then p = ShowWindow(Me.hwnd, SW_NORMAL)
If Form1.Visible Then If GetAsyncKeyState(&H26) Then If ZF < 500 Then ZF = ZF + 1
If Form1.Visible Then If GetAsyncKeyState(&H28) Then If ZF > 1 Then ZF = ZF - 1
End Sub


