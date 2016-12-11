VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "解三角形专器 V1.0 上大附中特别版"
   ClientHeight    =   7950
   ClientLeft      =   3780
   ClientTop       =   2265
   ClientWidth     =   8970
   DrawStyle       =   6  'Inside Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":114DA
   ScaleHeight     =   7950
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   5880
      Top             =   960
   End
   Begin VB.Frame Judge2 
      Caption         =   "三角形的构成"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   98
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton JgCmd 
         Caption         =   "立即判断"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   109
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox JgTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   960
         TabIndex        =   105
         Text            =   "1.414"
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox JgTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   104
         Text            =   "1"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox JgTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   960
         TabIndex        =   103
         Text            =   "1"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label JgHlp 
         AutoSize        =   -1  'True
         Caption         =   "存在两解，点我显示另一解"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   480
         TabIndex        =   112
         Top             =   5880
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.Label JgHlp 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   111
         Top             =   4680
         Width           =   180
      End
      Begin VB.Label JgHlp 
         AutoSize        =   -1  'True
         Caption         =   "该三角形是："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   110
         Top             =   4080
         Width           =   1980
      End
      Begin VB.Label JgHlp 
         AutoSize        =   -1  'True
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   360
         TabIndex        =   108
         Top             =   3360
         Width           =   570
      End
      Begin VB.Label JgHlp 
         AutoSize        =   -1  'True
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   360
         TabIndex        =   107
         Top             =   2640
         Width           =   570
      End
      Begin VB.Label JgHlp 
         AutoSize        =   -1  'True
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   360
         TabIndex        =   106
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label JgLb 
         AutoSize        =   -1  'True
         Caption         =   "已知两边及其一边的对角"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   102
         Top             =   1440
         Width           =   2475
      End
      Begin VB.Label JgLb 
         AutoSize        =   -1  'True
         Caption         =   "已知两角及其夹边"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   101
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label JgLb 
         AutoSize        =   -1  'True
         Caption         =   "已知两边及其夹角"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   100
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label JgLb 
         AutoSize        =   -1  'True
         Caption         =   "已知三边"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   99
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame MinMax2 
      Caption         =   "求最值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   84
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
      Begin VB.PictureBox MMPic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DrawWidth       =   3
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   4320
         ScaleHeight     =   3945
         ScaleWidth      =   3465
         TabIndex        =   94
         Top             =   2160
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton MMCmd 
         Caption         =   "开    始"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3720
         TabIndex        =   93
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox MMLt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2670
         ItemData        =   "Form1.frx":24ABE
         Left            =   240
         List            =   "Form1.frx":24AC0
         TabIndex        =   91
         Top             =   3000
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox MMTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   90
         Text            =   "8888"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox MMTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1200
         TabIndex        =   89
         Text            =   "8888"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label MMLb 
         AutoSize        =   -1  'True
         Caption         =   "点击图像框可以查看函数图像"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   97
         Top             =   5880
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label MMLb 
         AutoSize        =   -1  'True
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4320
         TabIndex        =   96
         Top             =   1680
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label MMLb 
         AutoSize        =   -1  'True
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4320
         TabIndex        =   95
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label 数据记录 
         Caption         =   "数据记录"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   92
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label MMLb 
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   88
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label MMLb 
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   87
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label MinMaxHlp 
         AutoSize        =   -1  'True
         Caption         =   "已知一边及其对角求面积的最大值"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   86
         Top             =   720
         Width           =   4725
      End
      Begin VB.Label MinMaxHlp 
         AutoSize        =   -1  'True
         Caption         =   "已知两边求面积的最大值"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   85
         Top             =   360
         Width           =   3465
      End
   End
   Begin VB.Frame SCH2 
      Caption         =   "三角公式查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   71
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      Begin VB.PictureBox SCHP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   3240
         Picture         =   "Form1.frx":24AC2
         ScaleHeight     =   465
         ScaleWidth      =   2385
         TabIndex        =   83
         Top             =   5280
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.PictureBox SCHP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   3840
         Picture         =   "Form1.frx":3F954
         ScaleHeight     =   345
         ScaleWidth      =   2745
         TabIndex        =   80
         Top             =   4680
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.PictureBox SCHP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   3720
         Picture         =   "Form1.frx":4C266
         ScaleHeight     =   345
         ScaleWidth      =   2265
         TabIndex        =   78
         Top             =   3480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.PictureBox SCHP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   3840
         Picture         =   "Form1.frx":6E764
         ScaleHeight     =   345
         ScaleWidth      =   2985
         TabIndex        =   77
         Top             =   2160
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.PictureBox SCHP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   3360
         Picture         =   "Form1.frx":8EE8E
         ScaleHeight     =   345
         ScaleWidth      =   2385
         TabIndex        =   73
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.PictureBox SCHP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   3240
         Picture         =   "Form1.frx":C1FB4
         ScaleHeight     =   585
         ScaleWidth      =   2145
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label SCHHlp 
         Caption         =   "半角公式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   82
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label SCHHlp 
         Caption         =   "三倍角公式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   81
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label SCHHlp 
         Caption         =   "倍角公式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   79
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label SCHHlp 
         Caption         =   "万能置换公式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   76
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label SCHHlp 
         Caption         =   "积化和差公式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   75
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label SCHHlp 
         Caption         =   "和差化积公式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   74
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame GG2 
      Caption         =   "勾股数组的确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   59
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox GGCov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "帮我计算"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox GGTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   720
         TabIndex        =   67
         Top             =   5160
         Width           =   3015
      End
      Begin VB.TextBox GGTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   2640
         TabIndex        =   66
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox GGTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   3840
         TabIndex        =   65
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Timer GGPT 
         Enabled         =   0   'False
         Interval        =   15
         Left            =   120
         Top             =   4560
      End
      Begin VB.PictureBox GGHp 
         BackColor       =   &H00FFFFFF&
         FillStyle       =   6  'Cross
         Height          =   3555
         Left            =   120
         Picture         =   "Form1.frx":1032BE
         ScaleHeight     =   3495
         ScaleWidth      =   2355
         TabIndex        =   61
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label GGHlp 
         Caption         =   "   请保留3位小数后输入。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   5640
         TabIndex        =   70
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label GGHlp 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   4560
         TabIndex        =   69
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label GGHlp 
         Caption         =   "弦边（斜边c）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   64
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label GGHlp 
         Caption         =   "股边（直角边b）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   63
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label GGHlp 
         Caption         =   "勾边（直角边a）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   62
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label GGHlp 
         Caption         =   $"Form1.frx":12561C
         Height          =   1935
         Index           =   0
         Left            =   2640
         TabIndex        =   60
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.PictureBox PicDRG 
      Height          =   375
      Left            =   720
      Picture         =   "Form1.frx":1257AD
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   32
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame RTD2 
      Caption         =   "弧度制转角度制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox RTDTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   51
         Top             =   3360
         Width           =   3255
      End
      Begin VB.CheckBox RTDCov 
         Caption         =   "以∏弧度为单位"
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label RTDHlp 
         Caption         =   "此处∏取3.14159265"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   240
         TabIndex        =   58
         Top             =   6120
         Width           =   3735
      End
      Begin VB.Label RTDHlp 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   57
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label RTDSol 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   360
         TabIndex        =   56
         Top             =   5400
         Width           =   3735
      End
      Begin VB.Label RTDHlp 
         Caption         =   "度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   3600
         TabIndex        =   55
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label RTDSol 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   480
         TabIndex        =   54
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label RTDHlp 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   53
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label RTDHlp 
         Caption         =   "弧"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   3600
         TabIndex        =   52
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label RTDHlp 
         Caption         =   $"Form1.frx":199D47
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame DTR2 
      Caption         =   "角度制转弧度制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   30
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox DTRTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   37
         Top             =   4800
         Width           =   3255
      End
      Begin VB.TextBox DTRTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   36
         Top             =   3960
         Width           =   3255
      End
      Begin VB.CheckBox DTRCov 
         Caption         =   "以“度”为单位"
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox DTRTx 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   3120
         Width           =   3255
      End
      Begin VB.Label DTRSol 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   45
         Top             =   5640
         Width           =   3255
      End
      Begin VB.Label DTRHlp 
         Caption         =   $"Form1.frx":199E8B
         Height          =   2055
         Index           =   0
         Left            =   360
         TabIndex        =   34
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label DTRSol 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Index           =   3
         Left            =   -480
         TabIndex        =   48
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label DTRSol 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   4080
         TabIndex        =   47
         Top             =   240
         Width           =   15
      End
      Begin VB.Label DTRHlp 
         Caption         =   "弧度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   3120
         TabIndex        =   43
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label DTRSol 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   46
         Top             =   6360
         Width           =   2655
      End
      Begin VB.Label DTRLb 
         Caption         =   "度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   3600
         TabIndex        =   44
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label DTRHlp 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   6360
         Width           =   375
      End
      Begin VB.Label DTRHlp 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label DTRLb 
         Caption         =   "秒"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   3600
         TabIndex        =   40
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label DTRLb 
         Caption         =   "分"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   3600
         TabIndex        =   39
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label DTRLb 
         Caption         =   "度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   3600
         TabIndex        =   38
         Top             =   3120
         Width           =   495
      End
   End
   Begin VB.Frame 计算结果 
      Caption         =   "计算结果"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Label sol1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   720
         TabIndex        =   24
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label sol1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   720
         TabIndex        =   23
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label sol1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   720
         TabIndex        =   22
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label sol1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   21
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label sol1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label SSSA 
         Caption         =   "S:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   2560
         Width           =   615
      End
      Begin VB.Label SSSA 
         Caption         =   "M:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label SSSA 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1340
         Width           =   495
      End
      Begin VB.Label SSSA 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   760
         Width           =   495
      End
      Begin VB.Label SSSA 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame s3 
      Caption         =   "解三角形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
      Begin VB.TextBox S3T 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   840
         TabIndex        =   11
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox S3T 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox S3T 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label S3L 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label S3L 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label S3L 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Pnt 
      Caption         =   "绘图"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CheckBox 微调 
         Caption         =   "微调"
         Height          =   255
         Left            =   3480
         TabIndex        =   29
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "图像保存"
         CausesValidation=   0   'False
         Height          =   105
         Left            =   0
         TabIndex        =   28
         Top             =   3960
         Width           =   0
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DrawWidth       =   2
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3600
         Left            =   240
         ScaleHeight     =   3570
         ScaleWidth      =   3570
         TabIndex        =   4
         Top             =   240
         Width           =   3600
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   30
         Left            =   240
         Max             =   1001
         Min             =   1
         SmallChange     =   10
         TabIndex        =   3
         Top             =   4320
         Value           =   1
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "自动缩放"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   3895
         Width           =   1095
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3615
         LargeChange     =   5
         Left            =   3960
         Max             =   360
         SmallChange     =   3
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
      Begin VB.Label 绘图提示 
         Caption         =   "    水平滚动条用于放大或缩小图像，竖直滚动条用于旋转图像。"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   4680
         Width           =   3975
      End
      Begin VB.Label 比例计算 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   3945
         Width           =   1935
      End
   End
   Begin VB.PictureBox Pic 
      Height          =   615
      Left            =   3360
      Picture         =   "Form1.frx":199F5A
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label solhlp 
      Caption         =   "    注：解三角形全部采用的是角度制解答。所有三角形都是即时解答，如有不完善之处请见谅。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label LOGO 
      AutoSize        =   -1  'True
      Caption         =   "解三角形专器   "
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   25
      Top             =   6360
      Width           =   3735
   End
   Begin VB.Menu SCH 
      Caption         =   "三角公式查询"
   End
   Begin VB.Menu MiMax 
      Caption         =   "求最值"
      Begin VB.Menu FReal 
         Caption         =   "全体实数范围内求最值"
      End
      Begin VB.Menu FNat 
         Caption         =   "全体自然数范围求最值"
      End
   End
   Begin VB.Menu Judge 
      Caption         =   "三角形的构成"
   End
   Begin VB.Menu GG 
      Caption         =   "勾股数组的确定"
   End
   Begin VB.Menu DRG 
      Caption         =   "角度与弧度的转化"
      Begin VB.Menu RTD1 
         Caption         =   "弧度制转角度制"
      End
      Begin VB.Menu DTR1 
         Caption         =   "角度制转弧度制"
      End
   End
   Begin VB.Menu Sol 
      Caption         =   "解三角形"
      Begin VB.Menu SSS 
         Caption         =   "已知三边（S.S.S型）"
         Index           =   1
      End
      Begin VB.Menu SAS 
         Caption         =   "已知两边及其夹角（S.A.S型）"
         Index           =   2
      End
      Begin VB.Menu SSA 
         Caption         =   "已知两边及其一边的对角（S.S.A型）两解"
         Index           =   3
      End
      Begin VB.Menu ASA 
         Caption         =   "已知两角及其夹边（A.S.A型）"
         Index           =   4
      End
      Begin VB.Menu Paint 
         Caption         =   "画图"
         Index           =   5
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助"
      Begin VB.Menu helper 
         Caption         =   "帮助文件"
         Shortcut        =   {F1}
      End
      Begin VB.Menu picture 
         Caption         =   "查看本程序流程图"
      End
      Begin VB.Menu Abtus 
         Caption         =   "关于我们"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HL As Boolean '决定着SSA型双解问题，显示哪一解。
Public i, Cc As Integer 'i 用于循环时控制
Public MMTF As Boolean '决定最大值最小值解决哪类问题
Public MMConst, JgConst As Integer '决定最大值最小值解决哪类问题
Dim data() As Tri
Dim yfx() As Fx
Public Smax As Currency

Private Sub Abtus_Click()
frmSplash.Show
frmSplash.Timer1.Enabled = False
End Sub

'Picture1.Scale (-10 - HScroll1.Value, 10 + HScroll1.Value)-(10 + HScroll1.Value, -10 -HScroll1.Value)
Private Sub ASA_Click(Index As Integer)
For i = 0 To 4
sol1(i).Caption = ""
Next i
For i = 0 To 2
S3T(i).Text = ""
Next i
s3.Visible = True
s3.Move 12, 12, 3375, 3375
计算结果.Visible = True
计算结果.Move 0, 3600, 3500, 3255
SSSA(0).Caption = "a:": SSSA(1).Caption = "B:": SSSA(2).Caption = "c:"

    SSSA(0).ForeColor = vbBlack: SSSA(1).ForeColor = vbRed: SSSA(2).ForeColor = vbBlack
    S3L(0).Caption = "A:": S3L(1).Caption = "b:": S3L(2).Caption = "C:"
    S3L(0).ForeColor = vbRed: S3L(1).ForeColor = vbBlack: S3L(2).ForeColor = vbRed
    s3.Caption = "已知两角及其夹边（右键清空）"
End Sub


Private Sub Command1_Click()
'Picture1.Scale (-10 * (HScroll1.Value / 10), 10 * (HScroll1.Value / 10))-(10 * (HScroll1.Value / 10), -10 * (HScroll1.Value / 10))
Call PntFsh


End Sub

Private Sub DTRTD_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub DRG_Click()
Call AllHide
Pic.Visible = False
PicDRG.Visible = True
PicDRG.Move 4560, 120, 3615, 6015
solhlp.Visible = False
s3.Visible = False
计算结果.Visible = False
Pnt.Visible = False
Paint(5).Checked = False
Call MMHide
End Sub


Private Sub DTR1_Click()
DTR2.Visible = True
RTD2.Visible = False
DTR2.Move 240, 0, 4215, 7095
End Sub

Private Sub DTRCov_Click()
Dim i As Integer
If DTRCov.Value = 0 Then
    DTRCov.Caption = "以“度，分，秒”为单位"
    DTRTx(1).Visible = True: DTRTx(2).Visible = True
    DTRLb(1).Move 3600, 3960: DTRLb(2).Move 3600, 4800: DTRLb(3).Move 3600, 5640
    DTRSol(0).Move 360, 5640, 3255, 615
    DTRHlp(1).Move 120, 5640
    DTRSol(0).Move 360, 5640
    For i = 0 To 3
        DTRSol(i).Caption = ""
    Next i
    For i = 0 To 2
        DTRTx(i).Text = ""
    Next i
    
End If

If DTRCov.Value = 1 Then
    DTRCov.Caption = "以“度”为单位"
    DTRTx(1).Visible = False: DTRTx(2).Visible = False
    DTRSol(0).Move 360, 3960, 3255, 615
    DTRSol(2).Move 360, 4800, 3255, 615
    DTRSol(3).Move 360, 5640, 3255, 615
    DTRHlp(1).Move 120, 3960
    DTRLb(3).Move 3600, 3960: DTRLb(1).Move 3600, 4800: DTRLb(2).Move 3600, 5640
    For i = 0 To 3
    DTRSol(i).Caption = ""
    Next i
    For i = 0 To 2
    DTRTx(i).Text = ""
    Next i
    
End If
End Sub


Private Sub DTRTx_Change(Index As Integer)
Dim K, l, M As Single
Dim total As Single
K = Val(DTRTx(0).Text): l = Val(DTRTx(1).Text): M = Val(DTRTx(2).Text)

If DTRCov.Value = 0 And K >= 0 And K < 1000000000 And l >= 0 And l < 60 And M >= 0 And M < 60 Then
    total = Format((K * 3600 + l * 60 + M) / 3600, "0.#####")
    DTRSol(0).Caption = total
    DTRSol(1).Caption = DTR(total)

Else
        If DTRCov.Value = 1 And K < 10000000 Then
             total = K
                K = Fix(total)
            l = Fix(60 * (total - K))
            M = total * 3600 - K * 3600 - l * 60
            DTRSol(0).Caption = K
          DTRSol(2).Caption = l
            DTRSol(3).Caption = Format(M, "0.##")
             DTRSol(1).Caption = DTR(total)
     
Else
    MsgBox "抱歉，您的输入有误，请从新输入", vbOKOnly + vbQuestion, "请从新输入"
    K = 0: M = 0: l = 0
      DTRTx(0).Text = "": DTRTx(1).Text = "": DTRTx(2).Text = ""
      End If
End If
End Sub

Private Sub DTRTx_KeyPress(Index As Integer, KeyAscii As Integer)

If (KeyAscii < 48 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Or (KeyAscii > 57 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Then
KeyAscii = 0
End If

End Sub

Private Sub FNat_Click()
Call AllHide
Call MMHide
MinMax2.Visible = True
MinMax2.Caption = "全体自然数范围求最值"
MinMax2.Move 120, 0, 8055, 6255

End Sub

Private Sub Form_Load()

Pnt.Move 3600, 120, 4335, 5175
Form1.Width = 8370
Form1.Height = 7620
HL = 1
For i = 1 To 5
SCHHlp(i).AutoSize = True
Next i
MMTF = 0

End Sub

Private Sub FReal_Click()
Call AllHide
Call MMHide
MinMax2.Visible = True
MinMax2.Caption = "全体实数范围内求最值"
MinMax2.Move 120, 0, 8055, 6255
End Sub

Private Sub GG_Click()
Call AllHide
GG2.Visible = True
GG2.Move 360, 240, 7695, 6015
GGPT.Enabled = True
Call MMHide
Call MMHide
End Sub

Private Sub GGCov_Click()
For i = 0 To 2
    GGTx(i).Text = ""
Next i
If GGCov.Value = 1 Then
    GGCov.Caption = "帮我验证"
    GGTx(2).Enabled = False
End If
If GGCov.Value = 0 Then
    GGCov.Caption = "帮我计算"
    GGTx(2).Enabled = True
End If
End Sub

Private Sub GGHlp_Click(Index As Integer)
For i = 0 To 2
    GGTx(i).Text = ""
Next i
End Sub

Private Sub GGHp_Click()
GGPT.Enabled = Not (GGPT.Enabled)
End Sub


Private Sub GGPT_Timer()
Dim X, y, z, w As Single
Dim N As Single, X1, Y1, X2, Y2, X3, Y3 As Currency
Static M As Long
Dim t(8), i As Integer
Dim Bx, cx, cy As Currency
X = 1
y = (Sqr(5) - 1) / 2
z = Sqr(X ^ 2 + y ^ 2)
M = M + 1

Bx = Module1.Paint(X, y, z).Bx
cx = Module1.Paint(X, y, z).cx
cy = Module1.Paint(X, y, z).cy
N = DTR(M)
X1 = -(Bx + cx) / 3: Y1 = -cy / 3
X2 = 2 / 3 * Bx - cx / 3: Y2 = -cy / 3
X3 = 2 / 3 * cx - Bx / 3: Y3 = 2 / 3 * cy
w = (Sqr(X ^ 2 + (y / 2) ^ 2)) / 3 * 2

Randomize
For i = 0 To 8
t(i) = Int(Rnd * 256)
Next i
GGHp.Cls
GGHp.Scale (-w, w)-(w, -w)
GGHp.DrawWidth = 5
GGHp.Line (X1 * Cos(N) + Y1 * Sin(N), Y1 * Cos(N) - X1 * Sin(N))-(X2 * Cos(N) + Y2 * Sin(N), Y2 * Cos(N) - X2 * Sin(N)), RGB(t(0), t(1), t(2))
GGHp.Line (X3 * Cos(N) + Y3 * Sin(N), Y3 * Cos(N) - X3 * Sin(N))-(X1 * Cos(N) + Y1 * Sin(N), Y1 * Cos(N) - X1 * Sin(N)), RGB(t(3), t(4), t(5))
GGHp.Line (X3 * Cos(N) + Y3 * Sin(N), Y3 * Cos(N) - X3 * Sin(N))-(X2 * Cos(N) + Y2 * Sin(N), Y2 * Cos(N) - X2 * Sin(N)), RGB(t(6), t(7), t(8))

If M >= 360 Then M = 0
End Sub

Private Sub GGTx_Change(Index As Integer)
Dim a, B, C As Currency
a = Val(GGTx(0).Text): B = Val(GGTx(1).Text): C = Val(GGTx(2).Text)
If a > 1000000 Or B > 1000000 Or C > 100000 Then
For i = 0 To 2
    GGTx(i).Text = ""
Next i
MsgBox "数据过大，请从新输入", vbCritical + vbOKOnly, "请从新输入"
End If
If a <> 0 And B <> 0 And C <> 0 And CCur(Format(a ^ 2 + B ^ 2, "0.###")) = CCur(Format(C ^ 2, "0.###")) Then
    GGHlp(4).Caption = "√"
Else
    If GGTx(2).Text <> "" And GGTx(1).Text <> "" And GGTx(2).Text <> "" Then
         GGHlp(4).Caption = "×"
    End If
End If

'If GGCov.Value = 0 And (Not (A + C > B) Or Not (A + B > C) Or Not (B + C > A)) And (GGTx(2).Text <> "" And GGTx(1).Text <> "" And GGTx(0).Text <> "") Then
''    MsgBox "两边之和大于第三边，请从新输入!", vbCritical + vbOKOnly, "请从新输入!"
'    For i = 0 To 2
'        GGTx(i).Text = ""
'    Next i
'End If

If GGCov.Value = 0 Then
    GGTx(2).Enabled = True
    GGHlp(4).Visible = True
End If
If GGCov.Value = 1 Then
    GGTx(2).Enabled = False
    GGHlp(4).Visible = False
    If (GGTx(1).Text <> "" And GGTx(0).Text <> "") Then
        GGTx(2).Text = Format(Sqr(a ^ 2 + B ^ 2), "0.###")
    End If
End If
End Sub

Private Sub GGTx_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Or (KeyAscii > 57 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Then
KeyAscii = 0
End If
End Sub

Private Sub GGTx_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    For i = 0 To 2
        GGTx(i).Text = ""
    Next i
End If
End Sub

Private Sub helper_Click()
Shell "hh.exe 帮助文件.chm", 1
End Sub

Private Sub HScroll1_Change()
Picture1.Cls
Picture1.Scale (-10 * (HScroll1.Value / 10), 10 * (HScroll1.Value / 10))-(10 * (HScroll1.Value / 10), -10 * (HScroll1.Value / 10))
Call Pnt1(a, B, C)
If 10 * (HScroll1.Value / 10) < 10 Then 比例计算.Caption = "等比例放大到" & Format((10 / (10 * (HScroll1.Value / 10))), "0.###") & "倍"
If 10 * (HScroll1.Value / 10) > 10 Then 比例计算.Caption = "等比例缩小到" & Format((10 / (10 * (HScroll1.Value / 10))), "0.###") & "倍"
If 10 * (HScroll1.Value / 10) = CCur(10) Then 比例计算.Caption = "等比例显示"
End Sub

Private Sub JgCmd_Click()
JgHlp(3).Caption = ""
JgHlp(4).Caption = ""
Select Case JgConst
Case 0
    a = Val(JgTx(0).Text): B = Val(JgTx(1).Text): C = Val(JgTx(2).Text)
    If (a + B > C) And (B + C > a) And (a + C > B) Then
        'module1.A,B,C已解出
    Else
        JgHlp(3).Caption = "两边之和大于第三边！"
        
        a = 0: B = 0: C = 0
    End If
    
Case 1
     a = Val(JgTx(0).Text): B1 = Val(JgTx(1).Text): C = Val(JgTx(2).Text)
    If a > 0 And C > 0 And B1 > 0 And B1 < 180 Then
        B = Format(Sqr(a ^ 2 + C ^ 2 - 2 * a * C * Cos(DTR(B1))), "0.###")
        A1 = Format(RTD(ACos((B ^ 2 + C ^ 2 - a ^ 2) / (2 * B * C))), "0.###")
        C1 = Format(RTD(ACos((a ^ 2 + B ^ 2 - C ^ 2) / (2 * a * B))), "0.###")
    Else
        JgHlp(3).Caption = "角B应大于0小于180"
        
         a = 0: B = 0: C = 0
    End If
Case 2
    A1 = Val(JgTx(0).Text): B = Val(JgTx(1).Text): C1 = Val(JgTx(2).Text)
    If A1 > 0 And A1 < 180 And C1 > 0 And C1 < 180 And B > 0 And A1 + C1 < 180 Then
        B1 = 180 - A1 - C1
            C = Format((Sin(DTR(C1))) / (Sin(DTR(B1))) * B, "0.###")
            a = Format((Sin(DTR(A1))) / (Sin(DTR(B1))) * B, "0.###")
    Else
        If A1 >= 180 Then JgHlp(3).Caption = "角A应大于0小于180"
        If C1 >= 180 Then JgHlp(3).Caption = "角C应大于0小于180"
        If A1 + C1 >= 180 Then JgHlp(3).Caption = "角B应大于0小于180"
        
         a = 0: B = 0: C = 0
    End If
Case 3
    a = Val(JgTx(0).Text): A1 = Val(JgTx(1).Text): B = Val(JgTx(2).Text)
    If A1 > 0 And A1 < 180 And a >= B * CCur(Sin(DTR(A1))) And B > 0 Then
     '___________________________________________________
                                If a >= B Then
                                        B1 = RTD(ASin(B / a * Sin(DTR(A1))))
                                        C = Format(a * Cos(DTR(B1)) + B * Cos(DTR(A1)), "0.###")
                                        
                                Else
                                      If a = B * CCur(Sin(DTR(A1))) Then
                                        B1 = 90
                                        C1 = Format(180 - B1 - A1, "0.###")
                                        C = Format(Sqr(B ^ 2 - a ^ 2), "0.###")
                                        
                                        
                                        Else
                                            
                                            If a < B And a > B * CCur(Sin(DTR(A1))) And HL = True Then
                                            Call HL1
                                            JgHlp(5).Visible = True
                                            Else
                                                If a < B And a > B * CCur(Sin(DTR(A1))) And HL = False Then
                                                Call HL2
                                                JgHlp(5).Visible = True
                                                End If
                                        
                                            End If
                                           End If
                                End If
     
     
     '___________________________________________________
     
    Else
    If A1 >= 180 Then JgHlp(3).Caption = "角A应大于0小于180"
    If a < B * CCur(Sin(DTR(A1))) Then JgHlp(3).Caption = "a边太小或b边太大"
    a = 0: B = 0: C = 0
    
    End If


End Select
If JgConst <> 1 Then
    i = Module1.Style(a, B, C)
Else
  If a = 0 And B = 0 And C = 0 Then
  i = 0
  Else
    i = 5
    If A1 = B1 Or B1 = C1 Or A1 = C1 Then i = 3
    If (A1 = B1 And B1 <> C1) Or (B1 = C1 And A1 <> B1) Or (A1 = C1 And C1 <> B1) Then i = 4
    If (A1 = 90) Or (B1 = 90) Or (C1 = 90) Then
        If A1 = B1 Or B1 = C1 Or A1 = C1 Then
            i = 1
        Else
            i = 2
        End If
    End If
  End If
End If
    If a <> 0 And B <> 0 And C <> 0 Then
        Select Case i
        Case 0
          JgHlp(3).Caption = "!!!!"
            JgHlp(4).Caption = "!!!!"
        Case 1
            JgHlp(3).Caption = "该三角形是："
            JgHlp(4).Caption = "等腰直角三角形"
        Case 2
            JgHlp(3).Caption = "该三角形是："
            JgHlp(4).Caption = "普通直角三角形"
        Case 3
            JgHlp(3).Caption = "该三角形是："
            JgHlp(4).Caption = "等边三角形"
        Case 4
            JgHlp(3).Caption = "该三角形是："
            JgHlp(4).Caption = "普通等腰三角形"
        Case 5
            JgHlp(3).Caption = "该三角形是："
            JgHlp(4).Caption = "非特殊三角形"
        Case Else
            JgHlp(3).Caption = "!!!!2"
            JgHlp(4).Caption = "!!!!2"
        End Select

        Picture1.Cls
        Call PntFsh
    Else
        If JgTx(0) = "" Or JgTx(1) = "" Or JgTx(2) = "" Then
            Picture1.Cls
            JgHlp(3).Caption = "请您输入数据"
            JgHlp(4).Caption = ""
            
        End If
        For i = 0 To 2
            JgTx(i).Text = ""
        Next i
    End If
End Sub

Private Sub JgHlp_Click(Index As Integer)
HL = Not (HL)
If Index = 5 Then
    If HL = True Then
    Call HL1
    Else
    Call HL2
    End If
End If
End Sub


Private Sub JgLb_Click(Index As Integer)
For i = 0 To 3
    JgLb(i).BorderStyle = 0
Next i
JgLb(Index).BorderStyle = 1
JgConst = Index
For i = 0 To 2
    JgHlp(i).ForeColor = vbBlack
    JgTx(i).Text = ""
Next i
Call JgShow
JgHlp(3).Caption = ""
JgHlp(4).Caption = ""
JgHlp(5).Visible = False
a = 0: B = 0: C = 0
Select Case Index
Case 0
    JgHlp(0).Caption = "a:"
    JgHlp(1).Caption = "b:"
    JgHlp(2).Caption = "c:"
Case 1
    JgHlp(0).Caption = "a:"
    JgHlp(1).Caption = "B:": JgHlp(1).ForeColor = vbRed
    JgHlp(2).Caption = "c:"
Case 2
    JgHlp(0).Caption = "A:": JgHlp(0).ForeColor = vbRed
    JgHlp(1).Caption = "b:"
    JgHlp(2).Caption = "C:": JgHlp(2).ForeColor = vbRed
Case 3
    JgHlp(0).Caption = "a:"
    JgHlp(1).Caption = "A:": JgHlp(1).ForeColor = vbRed
    JgHlp(2).Caption = "b:"
End Select
End Sub

Private Sub JgLb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
JgLb(Index).ForeColor = vbRed
Select Case Index
Case 0
    JgLb(1).ForeColor = vbBlack
    JgLb(2).ForeColor = vbBlack
    JgLb(3).ForeColor = vbBlack
Case 1
    JgLb(0).ForeColor = vbBlack
    JgLb(2).ForeColor = vbBlack
    JgLb(3).ForeColor = vbBlack
Case 2
    JgLb(0).ForeColor = vbBlack
    JgLb(1).ForeColor = vbBlack
    JgLb(3).ForeColor = vbBlack
Case 3
    JgLb(0).ForeColor = vbBlack
    JgLb(1).ForeColor = vbBlack
    JgLb(2).ForeColor = vbBlack
End Select
End Sub

Private Sub JgTx_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Or (KeyAscii > 57 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Then
KeyAscii = 0
End If
End Sub

Private Sub Judge_Click()
Call AllHide
Call MMHide
Call JgHide
Judge2.Visible = True
Judge2.Move 4440, 0, 3735, 6255
Pnt.Visible = True
Pnt.Move 0, 0, 4335, 5175
solhlp.Visible = True
solhlp.Move 0, 5280, 4335, 855
End Sub

Private Sub Judge2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
For i = 0 To 3
JgLb(i).ForeColor = vbBlack
Next i
End Sub

Private Sub Label1_Click()
Form1.Caption = Format(a ^ 2 + B ^ 2, "0.##") & "dddd" & Format(C ^ 2, "0.##")
End Sub

Private Sub MinMax2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
        MinMaxHlp(0).ForeColor = vbBlack
        MinMaxHlp(1).ForeColor = vbBlack
       
End Sub

Private Sub MinMaxHlp_Click(Index As Integer)

Call MMHide
Call MMShow
MMLt.Clear
Select Case Index
    Case 0
        MMTF = False
        MMLb(0).Caption = "边A": MMLb(1).Caption = "边B"
        If MinMaxHlp(0).Left = 240 Then
            MinMaxHlp(0).Left = MinMax2.Left + (MinMax2.Width / 2) - (MinMaxHlp(0).Width / 2)
            MinMaxHlp(1).Left = 240
            
        End If
        
    Case 1
         MMTF = True
         MMLb(0).Caption = "边A": MMLb(1).Caption = "角A"
        If MinMaxHlp(1).Left = 240 Then
            MinMaxHlp(1).Left = MinMax2.Left + (MinMax2.Width / 2) - (MinMaxHlp(1).Width / 2)
            MinMaxHlp(0).Left = 240
        End If
       
End Select
For i = 0 To 1
MinMaxHlp(i).BorderStyle = 0
Next i
MinMaxHlp(Index).BorderStyle = 1
If MinMax2.Caption = "全体实数范围内求最值" And MMTF = False Then MMConst = 1
If MinMax2.Caption = "全体实数范围内求最值" And MMTF = True Then MMConst = 2
If MinMax2.Caption = "全体自然数范围求最值" And MMTF = False Then MMConst = 3
If MinMax2.Caption = "全体自然数范围求最值" And MMTF = True Then MMConst = 4
End Sub

Private Sub MinMaxHlp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

Select Case Index
    Case 0
        MinMaxHlp(0).ForeColor = vbRed
        MinMaxHlp(1).ForeColor = vbBlack
        
    Case 1
        MinMaxHlp(0).ForeColor = vbBlack
        MinMaxHlp(1).ForeColor = vbRed
        
End Select
End Sub

Private Sub MMCmd_Click()
On Error GoTo Wrong
Dim a, B, A1, B1, C1 As Integer
Static C As Integer
Dim Ma, Mb, Mc
Dim K, l, M, K1, L1, M1, s, P, stp As Currency
Dim i, Mi As Integer 'Index
Dim N As Integer

If MMTx(0).Text = "" Or MMTx(1).Text = "" Then
MsgBox "请输入数据", vbOKOnly + vbInformation, "输入不为空"
MMTx(0).Text = "": MMTx(1).Text = ""
End If

Select Case MMConst
Case 1
     MMLt.Clear
    Smax = 0
    If MMTx(0).Text <> "" And MMTx(1).Text <> "" Then
        K = Val(MMTx(0).Text): l = Val(MMTx(1).Text)
        s = 0: i = 0
        ReDim data(179) As Tri
        ReDim yfx(179) As Fx
        For N = 1 To 179
            M = CCur(Sqr(K ^ 2 + l ^ 2 + 2 * K * l * Cos(DTR(N))))
            P = (K + l + M) / 2
            s = Format(Sqr(P * (P - K) * (P - l) * (P - M)), "0.###")
            If s > Smax Then
                Smax = s: Ma = K: Mb = l: Mc = M
            End If
            data(N).Bx = Module1.Paint(K, l, M).Bx
            data(N).cx = Module1.Paint(K, l, M).cx
            data(N).cy = Module1.Paint(K, l, M).cy
            yfx(N).X = M
            yfx(N).y = s
             MMLt.AddItem "[" & N & "]" & " " & "C=" & M & "   " & "S=" & s
             
        Next N
         MMLb(2).Visible = True: MMLb(3).Visible = True
        MMLb(2).Caption = "Smax=" & Smax
        MMLb(3).Caption = "A=" & Ma & "   B=" & Mb & "   C=" & Mc
        Cc = Mc
    
    End If
Case 2
    MMLt.Clear
    Smax = 0
    If MMTx(0).Text <> "" And MMTx(1).Text <> "" Then
        K = Val(MMTx(0).Text): K1 = Val(MMTx(1).Text)
        s = 0: i = 0
         'Abs(A / Sin(DTR(A1)) - 2)
         ReDim data(1000) As Tri
      ReDim yfx(1000) As Fx
        If K / Sin(DTR(K1)) - 1 > 2 Then
            For M = 1 To (K / Sin(DTR(K1)) - 1)
                M1 = RTD(ASin(M / K * Sin(DTR(K1))))
                l = M * Cos(DTR(K1)) + K * Cos(DTR(M1))
                P = (K + l + M) / 2
                s = Format(Sqr(P * (P - K) * (P - l) * (P - M)), "0.###")
                
                If s > Smax Then
                Smax = s: Ma = K: Mb = Format(CCur(l), "0.##"): Mc = M
                End If
                data(i).Bx = Module1.Paint(K, l, M).Bx
                data(i).cx = Module1.Paint(K, l, M).cx
                data(i).cy = Module1.Paint(K, l, M).cy
                yfx(i).X = M
                yfx(i).y = s
                 MMLt.AddItem "[" & i & "]" & " " & "C=" & M & "   " & "S=" & s
                 
                If M < K Then
                i = i + 1
                    M1 = 180 - M1
                    l = M * Cos(DTR(K1)) + K * Cos(DTR(M1))
                P = (K + l + M) / 2
                s = Format(Sqr(P * (P - K) * (P - l) * (P - M)), "0.###")
                
                If s > Smax Then
                Smax = s: Ma = K: Mb = Format(CCur(l), "0.##"): Mc = M
                End If
                data(i).Bx = Module1.Paint(K, l, M).Bx
                data(i).cx = Module1.Paint(K, l, M).cx
                data(i).cy = Module1.Paint(K, l, M).cy
                yfx(i).X = M
                yfx(i).y = s
                 MMLt.AddItem "[" & i & "]" & " " & "C=" & M & "   " & "S=" & s
                 
                 End If
                i = i + 1
            Next M
        Else
            MMLt.Clear
            MMLt.AddItem "不存在这样的自然数C，请从新输入"
        
        End If
    MMLb(2).Visible = True: MMLb(3).Visible = True
        MMLb(2).Caption = "Smax=" & Smax
        MMLb(3).Caption = "A=" & Ma & "   B=" & Mb & "   C=" & Mc
        Cc = Mc
    End If
Case 3
    MMLt.Clear
    Smax = 0
    If MMTx(0).Text <> "" And MMTx(1).Text <> "" Then
      a = Val(MMTx(0).Text): B = Val(MMTx(1).Text)
      s = 0: i = 0
      ReDim data(a + B) As Tri
      ReDim yfx(a + B) As Fx
      For C = Abs(a - B) + 1 To a + B - 1
            
            P = (a + B + C) / 2
            s = Format(Sqr(P * (P - K) * (P - l) * (P - M)), "0.###")
            If s > Smax Then
                Smax = s: Ma = a: Mb = B: Mc = C
            End If
            data(i).Bx = Module1.Paint(a, B, C).Bx
            data(i).cx = Module1.Paint(a, B, C).cx
            data(i).cy = Module1.Paint(a, B, C).cy
            yfx(i).X = C
            yfx(i).y = s
             MMLt.AddItem "[" & i & "]" & " " & "C=" & C & "   " & "S=" & s
             i = i + 1
        Next C
        MMLb(2).Visible = True: MMLb(3).Visible = True
        MMLb(2).Caption = "Smax=" & Smax
        MMLb(3).Caption = "A=" & Ma & "   B=" & Mb & "   C=" & Mc
        Cc = C
    
       
    End If
Case 4
    MMLt.Clear
    Smax = 0
    If MMTx(0).Text <> "" And MMTx(1).Text <> "" Then
        a = Val(MMTx(0).Text): A1 = Val(MMTx(1).Text)
        s = 0: i = 0
         'Abs(A / Sin(DTR(A1)) - 2)
         ReDim data(1000) As Tri
      ReDim yfx(1000) As Fx
        If a / Sin(DTR(A1)) - 1 > 2 Then
            For C = 1 To (a / Sin(DTR(A1)) - 1)
                C1 = RTD(ASin(C / a * Sin(DTR(A1))))
                B = C * Cos(DTR(A1)) + a * Cos(DTR(C1))
                P = (a + B + C) / 2
                s = Format(Sqr(P * (P - K) * (P - l) * (P - M)), "0.###")
                
                If s > Smax Then
                Smax = s: Ma = a: Mb = Int(B): Mc = C
                End If
                data(i).Bx = Module1.Paint(a, B, C).Bx
                data(i).cx = Module1.Paint(a, B, C).cx
                data(i).cy = Module1.Paint(a, B, C).cy
                yfx(i).X = C
                yfx(i).y = s
                 MMLt.AddItem "[" & i & "]" & " " & "C=" & C & "   " & "S=" & s
                 
                If C < a Then
                i = i + 1
                    C1 = 180 - C1
                    B = C * Cos(DTR(A1)) + a * Cos(DTR(C1))
                    P = (a + B + C) / 2
                    s = Format(Sqr(P * (P - K) * (P - l) * (P - M)), "0.###")
                
                    If s > Smax Then
                        Smax = s: Ma = a: Mb = Int(B): Mc = C
                    End If
                    data(i).Bx = Module1.Paint(a, B, C).Bx
                    data(i).cx = Module1.Paint(a, B, C).cx
                    data(i).cy = Module1.Paint(a, B, C).cy
                    yfx(i).X = C
                    yfx(i).y = s
                    MMLt.AddItem "[" & i & "]" & " " & "C=" & C & "   " & "S=" & s
                 End If
                i = i + 1
            Next C
        Else
            MMLt.Clear
            MMLt.AddItem "不存在这样的自然数C，请从新输入"
        
        End If
    MMLb(2).Visible = True: MMLb(3).Visible = True
        MMLb(2).Caption = "Smax=" & Smax
        MMLb(3).Caption = "A=" & Ma & "   B=" & Mb & "   C=" & Mc
        Cc = Mc
    End If
Exit Sub
Wrong:
MsgBox "您输入的数据过大，请从新输入", vbCritical + vbOKOnly, "请从新输入"
MMTx(0).Text = "": MMTx(1).Text = ""
ReDim data(1000) As Tri
ReDim yfx(1000) As Fx
MMLb(2).Caption = "": MMLb(3).Caption = ""
MMLt.Clear
MMPic.Cls
End Select
End Sub




Private Sub MMLt_Click()
On Error Resume Next
Dim Max As Currency
Select Case MMConst
Case 1
    MMPic.Cls
    Max = IIf(data(MMLt.ListIndex).Bx > (Sqr((data(MMLt.ListIndex).cx) ^ 2 + (data(MMLt.ListIndex).cy) ^ 2)), data(MMLt.ListIndex).Bx, Sqr((data(MMLt.ListIndex).cx) ^ 2 + (data(MMLt.ListIndex).cy) ^ 2))
    MMPic.Scale (-Max, Max)-(Max, -1)
    MMPic.Line (data(MMLt.ListIndex).Bx, 0)-(0, 0), vbRed
    MMPic.Line (data(MMLt.ListIndex).Bx, 0)-(data(MMLt.ListIndex).cx, data(MMLt.ListIndex).cy), vbGreen
    MMPic.Line (data(MMLt.ListIndex).cx, data(MMLt.ListIndex).cy)-(0, 0), vbBlue
Case 2
    MMPic.Cls
    Max = IIf(data(MMLt.ListIndex).Bx > (Sqr((data(MMLt.ListIndex).cx) ^ 2 + (data(MMLt.ListIndex).cy) ^ 2)), data(MMLt.ListIndex).Bx, Sqr((data(MMLt.ListIndex).cx) ^ 2 + (data(MMLt.ListIndex).cy) ^ 2))
    MMPic.Scale (-1, Max)-(Max, -Max)
    MMPic.Line (data(MMLt.ListIndex).Bx, 0)-(0, 0), vbRed
    MMPic.Line (data(MMLt.ListIndex).Bx, 0)-(data(MMLt.ListIndex).cx, data(MMLt.ListIndex).cy), vbGreen
    MMPic.Line (data(MMLt.ListIndex).cx, data(MMLt.ListIndex).cy)-(0, 0), vbBlue
Case 3
    MMPic.Cls
    Max = IIf(data(MMLt.ListIndex).Bx > (Sqr((data(MMLt.ListIndex).cx) ^ 2 + (data(MMLt.ListIndex).cy) ^ 2)), data(MMLt.ListIndex).Bx, Sqr((data(MMLt.ListIndex).cx) ^ 2 + (data(MMLt.ListIndex).cy) ^ 2))
    MMPic.Scale (-Max, Max)-(Max, -1)
    MMPic.Line (data(MMLt.ListIndex).Bx, 0)-(0, 0), vbRed
    MMPic.Line (data(MMLt.ListIndex).Bx, 0)-(data(MMLt.ListIndex).cx, data(MMLt.ListIndex).cy), vbGreen
    MMPic.Line (data(MMLt.ListIndex).cx, data(MMLt.ListIndex).cy)-(0, 0), vbBlue
Case 4
MMPic.Cls
    Max = IIf(data(MMLt.ListIndex).Bx > (Sqr((data(MMLt.ListIndex).cx) ^ 2 + (data(MMLt.ListIndex).cy) ^ 2)), data(MMLt.ListIndex).Bx, Sqr((data(MMLt.ListIndex).cx) ^ 2 + (data(MMLt.ListIndex).cy) ^ 2))
    MMPic.Scale (-1, Max)-(Max, -Max)
    MMPic.Line (data(MMLt.ListIndex).Bx, 0)-(0, 0), vbRed
    MMPic.Line (data(MMLt.ListIndex).Bx, 0)-(data(MMLt.ListIndex).cx, data(MMLt.ListIndex).cy), vbGreen
    MMPic.Line (data(MMLt.ListIndex).cx, data(MMLt.ListIndex).cy)-(0, 0), vbBlue
End Select
End Sub

Private Sub MMPic_Click()
On Error Resume Next
MMPic.Cls
MMPic.Scale (0, Smax)-(Cc, 0)
MMPic.Line (-1000, 0)-(1000, 0)
MMPic.Line (0, 1000)-(0, -1000)
For i = 0 To MMLt.ListCount - 2
MMPic.Line (yfx(i).X, yfx(i).y)-(yfx(i + 1).X, yfx(i + 1).y)
Next i
End Sub

Private Sub MMTx_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Or (KeyAscii > 57 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Then
KeyAscii = 0
End If
If KeyAscii = 46 And (MMConst = 3 Or MMConst = 4) Then
KeyAscii = 0
End If
End Sub

Private Sub Paint_Click(Index As Integer)
Pnt.Visible = Not (Pnt.Visible)
Paint(5).Checked = Not (Paint(5).Checked)
If Pnt.Visible = True Then Call PntFsh
End Sub



Private Sub Pic_Click()
Pnt.Visible = True
Paint(5).Checked = True
Pnt.Move 3600, 120, 4335, 5175
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub picture_Click()
On Error Resume Next
Shell "c:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE 程序流程图.doc", 1
Shell "d:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE 程序流程图.doc", 1
Shell "e:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE 程序流程图.doc", 1

Shell "f:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE 程序流程图.doc", 1

End Sub

Private Sub RTD1_Click()

RTD2.Visible = True
DTR2.Visible = False
RTD2.Move 240, 0, 4215, 7095

End Sub

Private Sub RTDCov_Click()
If RTDCov.Value = 0 Then
    RTDCov.Caption = "以∏弧度为单位"
    RTDHlp(1).Caption = "弧"
    End If
If RTDCov.Value = 1 Then
    RTDCov.Caption = "以弧度为单位"
    RTDHlp(1).Caption = "∏"
End If
RTDTx.Text = ""
RTDSol(0).Caption = "": RTDSol(1).Caption = ""
End Sub

Private Sub RTDTx_Change()
Dim R As Single 'Rad
Dim K, l, M As Single 'Deg
Dim total As Single
R = Val(RTDTx.Text)
If RTDCov.Value = 0 And R < 1000000 And R >= 0 Then
    total = Module1.RTD2(R)
    K = Fix(total)
    l = Fix(60 * (total - K))
    M = total * 3600 - K * 3600 - l * 60
    RTDSol(0).Caption = Format(total, "0.###")
    RTDSol(1).Caption = K & "°" & l & "＇" & Format(M, "0.##") & "＇＇"
    
    Else
        If RTDCov.Value = 1 And R < 1000000 Then
        total = Module1.RTD2(R * PI)
        K = Fix(total)
        l = Fix(60 * (total - K))
         M = total * 3600 - K * 3600 - l * 60
        RTDSol(0).Caption = Format(total, "0.###")
        RTDSol(1).Caption = K & "°" & l & "＇" & Format(M, "0.##") & "＇＇"
        Else
            MsgBox "抱歉，您的输入有误，请从新输入", vbOKOnly + vbQuestion, "请从新输入"
            R = 0: total = 0: K = 0: l = 0: M = 0: RTDTx.Text = "": RTDSol(0).Caption = "": RTDSol(1).Caption = ""
        End If
End If
End Sub


Private Sub RTDTx_KeyPress(KeyAscii As Integer)


If (KeyAscii < 48 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Or (KeyAscii > 57 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Then
KeyAscii = 0
End If

End Sub

Private Sub s3_Click()
Call MMHide
End Sub

Private Sub S3T_Change(Index As Integer)
On Error Resume Next
If s3.Caption = "已知三边（右键清空）" Then
    计算结果.Caption = "计算结果"
    a = Val(S3T(0).Text): B = Val(S3T(1).Text): C = Val(S3T(2).Text)
    If Module1.Judge(a, B, C) = True Then
       P = 0.5 * (a + B + C)
       s = Format(Sqr(P * (P - a) * (P - B) * (P - C)), "0.###")
       M = a + B + C
       A1 = Format(RTD(ACos((B ^ 2 + C ^ 2 - a ^ 2) / (2 * B * C))), "0.###")
       B1 = Format(RTD(ACos((a ^ 2 + C ^ 2 - B ^ 2) / (2 * a * C))), "0.###")
        C1 = Format(RTD(ACos((a ^ 2 + B ^ 2 - C ^ 2) / (2 * a * B))), "0.###")
        sol1(0).Caption = A1 & "度"
        sol1(1).Caption = B1 & "度"
        sol1(2).Caption = C1 & "度"
        sol1(3).Caption = M
        sol1(4).Caption = s
        计算结果.Caption = "计算结果"
    Picture1.Cls
    Call PntFsh
    Else
        If S3T(0) <> "" And S3T(1) <> "" And S3T(2) <> "" Then
        ' 'MsgBox  "抱歉，您的输入有误，请从新输入", vbOKOnly +vbquestion, "请从新输入"
        sol1(0).Caption = ""
        sol1(1).Caption = ""
        sol1(2).Caption = ""
        sol1(3).Caption = ""
        sol1(4).Caption = ""
        End If
    End If
'______________________________________________________________________________________________________
Else
    If s3.Caption = "已知l两边及其夹角（右键清空）" Then
        计算结果.Caption = "计算结果"
        a = Val(S3T(0).Text): B1 = Val(S3T(1).Text): C = Val(S3T(2).Text)
        If a > 0 And C > 0 And B1 > 0 And B1 < 180 Then
           B = Format(Sqr(a ^ 2 + C ^ 2 - 2 * a * C * Cos(DTR(B1))), "0.###")
            A1 = Format(RTD(ACos((B ^ 2 + C ^ 2 - a ^ 2) / (2 * B * C))), "0.###")
            C1 = Format(RTD(ACos((a ^ 2 + B ^ 2 - C ^ 2) / (2 * a * B))), "0.###")
            P = 0.5 * (a + B + C)
            s = Format(Sqr(P * (P - a) * (P - B) * (P - C)), "0.###")
            M = Format(a + B + C, "0.###")
            sol1(0).Caption = A1 & "度"
            sol1(1).Caption = B
            sol1(2).Caption = C1 & "度"
            sol1(3).Caption = M
            sol1(4).Caption = s
             计算结果.Caption = "计算结果"
             Picture1.Cls
           Call PntFsh
            Else
                If S3T(0) <> "" And S3T(1) <> "" And S3T(2) <> "" Then
                 ' 'MsgBox  "抱歉，您的输入有误，请从新输入", vbOKOnly +vbquestion, "请从新输入"
                 'MsgBox  "抱歉，您的输入有误，请从新输入", vbOKOnly +vbquestion, "请从新输入"
                 sol1(0).Caption = ""
                 sol1(1).Caption = ""
                sol1(2).Caption = ""
                sol1(3).Caption = ""
                sol1(4).Caption = ""
                End If
        End If
'________________________________________________________________________________________________________________________
    Else
    If s3.Caption = "已知两角及其夹边（右键清空）" Then
        计算结果.Caption = "计算结果"
        A1 = Val(S3T(0).Text): B = Val(S3T(1).Text): C1 = Val(S3T(2).Text)
        If A1 > 0 And A1 < 180 And C1 > 0 And C1 < 180 And B > 0 And A1 + C1 < 180 Then
            B1 = 180 - A1 - C1
            C = Format((Sin(DTR(C1))) / (Sin(DTR(B1))) * B, "0.###")
            a = Format((Sin(DTR(A1))) / (Sin(DTR(B1))) * B, "0.###")
            P = 0.5 * (a + B + C)
            s = Format(Sqr(P * (P - a) * (P - B) * (P - C)), "0.###")
            M = Format(a + B + C, "0.###")
            sol1(0).Caption = a
            sol1(1).Caption = B1 & "度"
            sol1(2).Caption = C
            sol1(3).Caption = M
            sol1(4).Caption = s
            Picture1.Cls
            Call PntFsh
            Else
                If S3T(0) <> "" And S3T(1) <> "" And S3T(2) <> "" Then
                 ' 'MsgBox  "抱歉，您的输入有误，请从新输入", vbOKOnly +vbquestion, "请从新输入"
                 sol1(0).Caption = ""
                 sol1(1).Caption = ""
                sol1(2).Caption = ""
                sol1(3).Caption = ""
                sol1(4).Caption = ""
                End If
        End If
    Else
'_______________________________________________________________________________________________________________
    If s3.Caption = "已知两边及其一边的对角（右键清空）" Then
       计算结果.Caption = "计算结果"
       a = Val(S3T(0).Text): A1 = Val(S3T(1).Text): B = Val(S3T(2).Text)
        If A1 > 0 And A1 < 180 And a >= B * CCur(Sin(DTR(A1))) And B > 0 Then
                                If a >= B Then
                                        B1 = RTD(ASin(B / a * Sin(DTR(A1))))
                                        C = Format(a * Cos(DTR(B1)) + B * Cos(DTR(A1)), "0.###")
                                        C1 = 180 - A1 - B1
                                        P = 0.5 * (a + B + C)
                                        s = Format(Sqr(P * (P - a) * (P - B) * (P - C)), "0.###")
                                        M = Format(a + B + C, "0.###")
                                        B1 = Format(B1, "0.###")
                                        C1 = Format(C1, "0.###")
                                        sol1(0).Caption = B1 & "度"
                                        sol1(1).Caption = C
                                        sol1(2).Caption = C1 & "度"
                                        sol1(3).Caption = M
                                        sol1(4).Caption = s
                                        计算结果.Caption = "计算结果:一解"
                                        Picture1.Cls
                                        Call PntFsh
                                Else
                                      If a = B * CCur(Sin(DTR(A1))) Then
                                        B1 = 90
                                        C1 = Format(180 - B1 - A1, "0.###")
                                        C = Format(Sqr(B ^ 2 - a ^ 2), "0.###")
                                        M = Format(a + B + C, "0.###")
                                        P = 0.5 * (a + B + C)
                                        s = Format(Sqr(P * (P - a) * (P - B) * (P - C)), "0.###")
                                        sol1(0).Caption = B1 & "度"
                                        sol1(1).Caption = C
                                        sol1(2).Caption = C1 & "度"
                                        sol1(3).Caption = M
                                        sol1(4).Caption = s
                                        计算结果.Caption = "计算结果:一解(直角三角形)"
                                        Picture1.Cls
                                        Call PntFsh
                                        
                                        Else
                                            'It's hard ,isn't it?
                                            'HL 是全剧变量，'决定着SSA型双解问题，显示哪一解。
                                            '1是第一解，锐角解，0是第二解钝角解
                                            If a < B And a > B * CCur(Sin(DTR(A1))) And HL = True Then
                                            Call HL1
                                            Else
                                                If a < B And a > B * CCur(Sin(DTR(A1))) And HL = False Then
                                                Call HL2
                                                End If
                                        
                                            End If
                                           End If
                                End If
       
       
       
        Else
                If S3T(0) <> "" And S3T(1) <> "" And S3T(2) <> "" Then
                 ' 'MsgBox  "抱歉，您的输入有误，请从新输入", vbOKOnly +vbquestion, "请从新输入"
                 sol1(0).Caption = ""
                 sol1(1).Caption = ""
                sol1(2).Caption = ""
                sol1(3).Caption = ""
                sol1(4).Caption = ""
                End If
       
        End If
    End If
    
    
    
    End If
    End If
End If
End Sub

Private Sub S3T_KeyPress(Index As Integer, KeyAscii As Integer)

If (KeyAscii < 48 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Or (KeyAscii > 57 And Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack))) Then
KeyAscii = 0
End If
End Sub

Private Sub S3T_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
S3T(0).Text = "": S3T(1).Text = "": S3T(2).Text = ""
End If
End Sub

Private Sub SAS_Click(Index As Integer)
For i = 0 To 4
sol1(i).Caption = ""
Next i
For i = 0 To 2
S3T(i).Text = ""
Next i
s3.Visible = True
s3.Move 12, 12, 3375, 3375
计算结果.Visible = True
计算结果.Move 0, 3600, 3500, 3255
SSSA(0).Caption = "A:": SSSA(1).Caption = "b:": SSSA(2).Caption = "C:"
    SSSA(0).ForeColor = vbRed: SSSA(1).ForeColor = vbBlack: SSSA(2).ForeColor = vbRed
    S3L(0).Caption = "a:": S3L(1).Caption = "B:": S3L(2).Caption = "c:"
    S3L(0).ForeColor = vbBlack: S3L(1).ForeColor = vbRed: S3L(2).ForeColor = vbBlack
    s3.Caption = "已知l两边及其夹角（右键清空）"
End Sub

Private Sub SCH_Click()
Call AllHide
SCH2.Visible = True
SCH2.Move 120, 0, 8055, 6255
Call MMHide
End Sub

Private Sub SCH2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
        SCHHlp(0).ForeColor = vbBlack
        SCHHlp(1).ForeColor = vbBlack
        SCHHlp(2).ForeColor = vbBlack
        SCHHlp(3).ForeColor = vbBlack
        SCHHlp(4).ForeColor = vbBlack
        SCHHlp(5).ForeColor = vbBlack
End Sub

Private Sub SCHHlp_Click(Index As Integer)
Call SCHHide

For i = 0 To 5
SCHHlp(i).BorderStyle = 0
Next i

SCHHlp(Index).BorderStyle = 1
Select Case Index
    Case 0
       SCHP(0).Move 1800, 600, 5895, 3375
        SCHP(0).Visible = True
    Case 1
       SCHP(1).Move 1800, 600, 6495, 2415
        SCHP(1).Visible = True
    Case 2
        SCHP(2).Move 1800, 600, 3255, 3135
        SCHP(2).Visible = True
    Case 3
        SCHP(3).Move 1800, 600, 4095, 2655
        SCHP(3).Visible = True
    Case 4
        SCHP(4).Move 1800, 600, 4095, 975
        SCHP(4).Visible = True
    Case 5
        SCHP(5).Move 1800, 600, 3255, 2535
        SCHP(5).Visible = True
End Select
End Sub

Private Sub SCHHlp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case Index
    Case 0
        SCHHlp(0).ForeColor = vbRed
        SCHHlp(1).ForeColor = vbBlack
        SCHHlp(2).ForeColor = vbBlack
        SCHHlp(3).ForeColor = vbBlack
        SCHHlp(4).ForeColor = vbBlack
        SCHHlp(5).ForeColor = vbBlack
    Case 1
        SCHHlp(0).ForeColor = vbBlack
        SCHHlp(1).ForeColor = vbRed
        SCHHlp(2).ForeColor = vbBlack
        SCHHlp(3).ForeColor = vbBlack
        SCHHlp(4).ForeColor = vbBlack
        SCHHlp(5).ForeColor = vbBlack
    
    Case 2
        SCHHlp(0).ForeColor = vbBlack
        SCHHlp(1).ForeColor = vbBlack
        SCHHlp(2).ForeColor = vbRed
        SCHHlp(3).ForeColor = vbBlack
        SCHHlp(4).ForeColor = vbBlack
        SCHHlp(5).ForeColor = vbBlack
    
    Case 3
        SCHHlp(0).ForeColor = vbBlack
        SCHHlp(1).ForeColor = vbBlack
        SCHHlp(2).ForeColor = vbBlack
        SCHHlp(3).ForeColor = vbRed
        SCHHlp(4).ForeColor = vbBlack
        SCHHlp(5).ForeColor = vbBlack
    
    Case 4
        SCHHlp(0).ForeColor = vbBlack
        SCHHlp(1).ForeColor = vbBlack
        SCHHlp(2).ForeColor = vbBlack
        SCHHlp(3).ForeColor = vbBlack
        SCHHlp(4).ForeColor = vbRed
        SCHHlp(5).ForeColor = vbBlack
    
    Case 5
        SCHHlp(0).ForeColor = vbBlack
        SCHHlp(1).ForeColor = vbBlack
        SCHHlp(2).ForeColor = vbBlack
        SCHHlp(3).ForeColor = vbBlack
        SCHHlp(4).ForeColor = vbBlack
        SCHHlp(5).ForeColor = vbRed
    
End Select
End Sub

Private Sub Sol_Click()
Call AllHide
solhlp.Visible = True
solhlp.Move 3600, 5400, 4455, 855
Pic.Visible = True
Pic.Move 3960, 240, 3735, 4935
PicDRG.Visible = False
DTR2.Visible = 0
RTD2.Visible = False

End Sub

Private Sub sol1_Click(Index As Integer)
HL = Not HL
If A1 > 0 And A1 < 180 And a >= B * CCur(Sin(DTR(A1))) And B > 0 Then
    If a < B And a > B * CCur(Sin(DTR(A1))) And HL = True Then
    Call HL1
    Else
        If a < B And a > B * CCur(Sin(DTR(A1))) And HL = False Then
        Call HL2
        End If
    End If
End If
End Sub

Private Sub SSA_Click(Index As Integer)
For i = 0 To 4
sol1(i).Caption = ""
Next i
For i = 0 To 2
S3T(i).Text = ""
Next i
s3.Visible = True
s3.Move 12, 12, 3375, 3375
计算结果.Visible = True
计算结果.Move 0, 3600, 3500, 3255
SSSA(0).Caption = "B:": SSSA(1).Caption = "c:": SSSA(2).Caption = "C:"

    SSSA(0).ForeColor = vbRed: SSSA(1).ForeColor = vbBlack: SSSA(2).ForeColor = vbRed
    S3L(0).Caption = "a:": S3L(1).Caption = "A:": S3L(2).Caption = "b:"
    S3L(0).ForeColor = vbBlack: S3L(1).ForeColor = vbRed: S3L(2).ForeColor = vbBlack
    s3.Caption = "已知两边及其一边的对角（右键清空）"
End Sub

Private Sub SSS_Click(Index As Integer)
For i = 0 To 4
sol1(i).Caption = ""
Next i
For i = 0 To 2
S3T(i).Text = ""
Next i
s3.Visible = True
s3.Move 12, 12, 3375, 3375
计算结果.Visible = True
计算结果.Move 0, 3600, 3500, 3255
SSSA(0).Caption = "A:": SSSA(1).Caption = "B:": SSSA(2).Caption = "C:"

    SSSA(0).ForeColor = vbRed: SSSA(1).ForeColor = vbRed: SSSA(2).ForeColor = vbRed
    S3L(0).Caption = "a:": S3L(1).Caption = "b:": S3L(2).Caption = "c:"
    S3L(0).ForeColor = vbBlack: S3L(1).ForeColor = vbBlack: S3L(2).ForeColor = vbBlack
    s3.Caption = "已知三边（右键清空）"
End Sub

Private Sub SSSA_Click(Index As Integer)
HL = Not HL
If A1 > 0 And A1 < 180 And a >= B * CCur(Sin(DTR(A1))) And B > 0 Then
    If a < B And a > B * CCur(Sin(DTR(A1))) And HL = True Then
    Call HL1
    Else
        If a < B And a > B * CCur(Sin(DTR(A1))) And HL = False Then
        Call HL2
        End If
    End If
End If

End Sub

Private Sub Style_Click()
Call MMHide
End Sub








Private Sub Timer1_Timer()
tmp = LOGO.Caption
LOGO.Caption = Right(tmp, Len(tmp) - 1) & Left(tmp, 1)
End Sub

Private Sub VScroll1_Change()
Dim N As Single, X1, Y1, X2, Y2, X3, Y3 As Currency
Dim Bx, cx, cy As Currency
N = VScroll1.Value / 180 * PI
Bx = Module1.Paint(a, B, C).Bx
cx = Module1.Paint(a, B, C).cx
cy = Module1.Paint(a, B, C).cy
''
X1 = -(Bx + cx) / 3: Y1 = -cy / 3
X2 = 2 / 3 * Bx - cx / 3: Y2 = -cy / 3
X3 = 2 / 3 * cx - Bx / 3: Y3 = 2 / 3 * cy


Picture1.Cls
Picture1.Line (-150, 0)-(150, 0)
Picture1.Line (0, 150)-(0, -150)
Picture1.Line (X1 * Cos(N) + Y1 * Sin(N), Y1 * Cos(N) - X1 * Sin(N))-(X2 * Cos(N) + Y2 * Sin(N), Y2 * Cos(N) - X2 * Sin(N)), vbRed
Picture1.Line (X3 * Cos(N) + Y3 * Sin(N), Y3 * Cos(N) - X3 * Sin(N))-(X1 * Cos(N) + Y1 * Sin(N), Y1 * Cos(N) - X1 * Sin(N)), vbGreen
Picture1.Line (X3 * Cos(N) + Y3 * Sin(N), Y3 * Cos(N) - X3 * Sin(N))-(X2 * Cos(N) + Y2 * Sin(N), Y2 * Cos(N) - X2 * Sin(N)), vbBlue


End Sub




Private Sub Pnt1(ByVal X As Currency, ByVal y As Currency, ByVal z As Currency)
Dim Bx, cx, cy As Currency
Bx = Module1.Paint(a, B, C).Bx
cx = Module1.Paint(a, B, C).cx
cy = Module1.Paint(a, B, C).cy

Picture1.Cls
Picture1.Line (-150, 0)-(150, 0)
Picture1.Line (0, 150)-(0, -150)
Picture1.Line (-(Bx + cx) / 3, -cy / 3)-((2 / 3 * Bx - cx / 3), -cy / 3), vbRed
Picture1.Line (2 / 3 * cx - Bx / 3, 2 / 3 * cy)-(-(Bx + cx) / 3, -cy / 3), vbGreen
Picture1.Line (2 / 3 * cx - Bx / 3, 2 / 3 * cy)-((2 / 3 * Bx - cx / 3), -cy / 3), vbBlue
End Sub

Private Sub 计算结果_Click()
HL = Not HL
If A1 > 0 And A1 < 180 And a >= B * CCur(Sin(DTR(A1))) And B > 0 Then
    If a < B And a > B * CCur(Sin(DTR(A1))) And HL = True Then
    Call HL1
    Else
        If a < B And a > B * CCur(Sin(DTR(A1))) And HL = False Then
        Call HL2
        End If
    End If
End If
End Sub
Private Sub HL1() '0<B1<90
B1 = RTD(ASin(B / a * Sin(DTR(A1))))
C = Format(a * Cos(DTR(B1)) + B * Cos(DTR(A1)), "0.###")
C1 = 180 - A1 - B1
P = 0.5 * (a + B + C)
s = Format(Sqr(P * (P - a) * (P - B) * (P - C)), "0.###")
M = Format(a + B + C, "0.###")
B1 = Format(B1, "0.###")
C1 = Format(C1, "0.###")
sol1(0).Caption = B1 & "度"
sol1(1).Caption = C
sol1(2).Caption = C1 & "度"
sol1(3).Caption = M
sol1(4).Caption = s
计算结果.Caption = "计算结果:两解，情况一，点我显示情况二"
Picture1.Cls
Call PntFsh

End Sub
Private Sub HL2() '90<B1<180
B1 = 180 - RTD(ASin(B / a * Sin(DTR(A1))))
C = Format(a * Cos(DTR(B1)) + B * Cos(DTR(A1)), "0.###")
C1 = 180 - A1 - B1
P = 0.5 * (a + B + C)
s = Format(Sqr(P * (P - a) * (P - B) * (P - C)), "0.###")
M = Format(a + B + C, "0.###")
B1 = Format(B1, "0.###")
C1 = Format(C1, "0.###")
sol1(0).Caption = B1 & "度"
sol1(1).Caption = C
sol1(2).Caption = C1 & "度"
sol1(3).Caption = M
sol1(4).Caption = s
计算结果.Caption = "计算结果:两解，情况二，点我显示情况一"
Picture1.Cls
Call PntFsh

End Sub

Private Sub 微调_Click()
If 微调.Value = 1 Then
    HScroll1.SmallChange = 1
    HScroll1.LargeChange = 3
    VScroll1.SmallChange = 1
    VScroll1.LargeChange = 3
Else
    If 微调.Value = 0 Then
    HScroll1.SmallChange = 10
    HScroll1.LargeChange = 30
    VScroll1.SmallChange = 3
    VScroll1.LargeChange = 5
    End If
End If
End Sub

Private Sub PntFsh() '自动刷新
On Error Resume Next
Dim Bx, cx, cy As Currency
Dim Ex, Ey, Fx, Fy, Gx, Gy, Max As Currency
Dim E, F, G, X As Currency

Bx = Module1.Paint(a, B, C).Bx
cx = Module1.Paint(a, B, C).cx
cy = Module1.Paint(a, B, C).cy
''
Ex = -(Bx + cx) / 3: Ey = -cy / 3
Fx = 2 / 3 * Bx - cx / 3: Fy = -cy / 3
Gx = 2 / 3 * cx - Bx / 3: Gy = 2 / 3 * cy

E = Sqr(Ex ^ 2 + Ey ^ 2)
F = Sqr(Fx ^ 2 + Fy ^ 2)
G = Sqr(Gx ^ 2 + Gy ^ 2)

Max = E
If Max < F Then Max = F
If Max < G Then Max = G

X = Max


Picture1.Scale (-X, X)-(X, -X)
If X < 10 Then 比例计算.Caption = "等比例放大到" & Format((10 / X), "0.###") & "倍"
If X > 10 Then 比例计算.Caption = "等比例缩小到" & Format((10 / X), "0.###") & "倍"
If X = CCur(10) Then 比例计算.Caption = "等比例显示"

Picture1.Cls
Call Pnt1(a, B, C)

End Sub


Private Sub AllHide()
GG2.Visible = False
PicDRG.Visible = False
Pnt.Visible = False
计算结果.Visible = False
Pic.Visible = False
s3.Visible = False
RTD2.Visible = False
DTR2.Visible = False
solhlp.Visible = False
Paint(5).Checked = 0
SCH2.Visible = False
GGPT.Enabled = False
MinMax2.Visible = False
MMLt.Clear
MinMaxHlp(0).Left = 240
MinMaxHlp(1).Left = 240
Judge2.Visible = False
JgHlp(5).Visible = False
For i = 0 To 3
    JgLb(i).BorderStyle = 0
Next i
End Sub

Private Sub SCHHide()
For i = 0 To 5
SCHP(i).Visible = False
Next i
End Sub
Private Sub MMShow()
MMLb(0).Visible = True: MMLb(1).Visible = True
MMTx(0).Visible = True: MMTx(1).Visible = True
数据记录.Visible = True
MMLt.Visible = True: MMPic.Visible = True
MMCmd.Visible = True
MMLb(4).Visible = True
End Sub
Private Sub MMHide()
MMLb(0).Visible = False: MMLb(1).Visible = False
MMTx(0).Visible = False: MMTx(1).Visible = False
数据记录.Visible = False
MMLt.Visible = False: MMPic.Visible = False
MMCmd.Visible = False
MinMaxHlp(0).Left = 240: MinMaxHlp(1).Left = 240
MinMaxHlp(0).BorderStyle = 0: MinMaxHlp(1).BorderStyle = 0
MMTx(0).Text = "": MMTx(1).Text = ""
MMPic.Cls
MMLb(2).Caption = "": MMLb(3).Caption = ""
MMLb(2).Visible = False: MMLb(3).Visible = False
MMLb(4).Visible = False
End Sub

Private Sub JgHide()
For i = 0 To 2
JgTx(i).Visible = False
JgHlp(i).Visible = False
JgTx(i).Text = ""
JgHlp(i).Caption = ""

Next i
JgHlp(3).Caption = "": JgHlp(4).Caption = ""
JgHlp(4).Visible = False
JgHlp(3).Visible = False
JgCmd.Visible = False
End Sub


Private Sub JgShow()

For i = 0 To 2
JgTx(i).Visible = True
JgHlp(i).Visible = True
JgHlp(4).Visible = True
JgHlp(3).Visible = True
JgCmd.Visible = True
Next i
End Sub
