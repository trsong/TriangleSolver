VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1560
   ClientLeft      =   4665
   ClientTop       =   3585
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   1560
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   300
      Left            =   840
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   240
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   2640
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   840
      ScaleHeight     =   1185
      ScaleWidth      =   5505
      TabIndex        =   3
      Top             =   120
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1560
      Picture         =   "frmSplash.frx":AE86
      ScaleHeight     =   1185
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "指导老师：刘健"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   3840
      Width           =   2625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "程序流程图：文博"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "帮助文件：江辰"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   2625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":23888
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "使用详情请见帮助"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   4320
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V1.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   1140
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Color As Integer
Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Picture2.Left = frmSplash.Width / 2 - Picture2.Width / 2
Label1(1).Left = frmSplash.Width / 2 - Label1(1).Width / 2
Picture2.Cls
nWidth = 50
Stripes = Picture1.Width / nWidth
P2 = Picture1.Height
P1 = nWidth
For i = 0 To Picture1.Width + nWidth Step nWidth
p3 = i
For K = Picture1.Width To i Step -nWidth
p4 = K
Picture2.Cls
R% = BitBlt(Picture2.hdc, 0, 0, i, P2, Picture1.hdc, 0, 0, &HCC0020)
R% = BitBlt(Picture2.hdc, p4, 0, P1, P2, Picture1.hdc, p3, 0, &HCC0020)
For j = 1 To 5000
Next j
Next K
Next i

Me.Move Me.Left, Me.Top, 7455, 4935
Picture2.Left = frmSplash.Width / 2 - Picture2.Width / 2
Label1(1).Left = frmSplash.Width / 2 - Label1(1).Width / 2

End Sub

Private Sub Form_Click()
Unload Me
Form1.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Form1.Show
End Sub

Private Sub Form_Load()
Dim XX As Long
XX = SetWindowPos(Me.Hwnd, -1, 0, 0, 0, 0, 3)
Picture2.Left = frmSplash.Width / 2 - Picture2.Width / 2
Label1(1).Left = frmSplash.Width / 2 - Label1(1).Width / 2
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
End Sub

Private Sub LogoSol_Click()
Unload Me
Form1.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label1(1).ForeColor = vbWhite
For i% = 2 To 5
Label1(i).ForeColor = vbBlack
Next i
Color = 0
End Sub

Private Sub Label1_Click(Index As Integer)
 Shell "hh.exe 帮助文件.chm", 1
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Color = Index
End Sub

Private Sub Timer1_Timer()

Form1.Show
Unload Me


End Sub

Private Sub Timer2_Timer()
Dim s, j, w, l As Integer
Dim a(12) As Integer
Static z
If z > 100 Then
Timer2.Enabled = False
For i% = 2 To 5
Label1(i).ForeColor = vbBlack
Next i
z = 0
End If
s = 43:
j = 38
w = 29
l = 16
Label1(2).Left = Label1(2).Left + s
Label1(3).Left = Label1(3).Left + j
Label1(4).Left = Label1(4).Left + w
Label1(5).Left = Label1(5).Left + l
z = z + 1
For i = 0 To 12
    a(i) = Int(Rnd * 256)
Next i
Label1(2).ForeColor = RGB(a(0), a(1), a(2))
Label1(3).ForeColor = RGB(a(3), a(4), a(5))
Label1(4).ForeColor = RGB(a(6), a(7), a(8))
Label1(5).ForeColor = RGB(a(9), a(10), a(11))
End Sub

Private Sub Timer3_Timer()
Dim a(3) As Integer
If Color <> 0 And Color <> 1 Then
    For i% = 0 To 2
        a(i) = Int(Rnd * 256)
    Next i
    Randomize
    Label1(Color).ForeColor = RGB(a(0), a(1), a(2))
Else
    If Color = 0 Then
    For i% = 2 To 5
    Label1(i).ForeColor = vbBlack
    Next i
    Label1(1).ForeColor = vbWhite
    End If
    If Color = 1 Then
    Label1(1).ForeColor = vbRed
    End If
End If
End Sub
