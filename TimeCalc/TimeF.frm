VERSION 5.00
Begin VB.Form TimeF 
   Caption         =   "时间计算器"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "TimeF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   7695
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option1 
      Caption         =   "浮动输出"
      Height          =   495
      Index           =   1
      Left            =   6240
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "固定输出"
      Height          =   495
      Index           =   0
      Left            =   6240
      TabIndex        =   17
      Top             =   1560
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "即时模式"
      Height          =   495
      Left            =   5880
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   120
   End
   Begin VB.CommandButton Comand 
      Caption         =   "差值(2-1)"
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Comand 
      Caption         =   "延时(1+2)"
      Height          =   495
      Index           =   0
      Left            =   4200
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "获取系统时间"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "②"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "①"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   12
      Top             =   2640
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   10
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3840
      TabIndex        =   9
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3840
      TabIndex        =   4
      Top             =   600
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   315
   End
End
Attribute VB_Name = "TimeF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Firsttime As Date
Public Steptime As Date
Public Nowtime As Date
Public flag As Byte

Private Sub Form_Load()
  flag = 9
End Sub

Private Sub Command1_Click()
  Systime = Now()
  Text1 = Hour(Systime)
  Text2 = Minute(Systime)
  Text3 = Second(Systime)
  Firsttime = TimeSerial(Text1, Text2, Text3)
End Sub

Private Sub Comand_Click(Index As Integer)
On Error Resume Next
  Steptime = TimeSerial(Text4, Text5, Text6)
    If Err Then MsgBox "数据错误！", vbCritical, "严重警告": Exit Sub
  Nowtime = IIf(Index = 0, Firsttime + Steptime, Steptime - Firsttime)
  Label2 = Format(Nowtime, "HH:MM:SS")
  flag = Index
End Sub

Private Sub Check1_Click()
  Timer1.Enabled = Not Timer1.Enabled
  Option1(0).Visible = Not Option1(0).Visible
  Option1(1).Visible = Not Option1(1).Visible
End Sub

Private Sub Timer1_Timer()
  Command1_Click
    Select Case flag
       Case 0
         If Option1(0).Value = True Then
          If Text6 - 1 = -1 Then
            Text6 = 59
            If Text5 - 1 = -1 Then
              Text5 = 59
              Text4 = Text4 - 1
            Else
              Text5 = Text5 - 1
            End If
          Else
            Text6 = Text6 - 1
          End If
         End If
          Comand_Click (0)
       Case 1
         If Option1(1).Value = False Then
          If Text6 + 1 = 60 Then
            Text6 = 0
            If Text5 + 1 = 60 Then
              Text5 = 0
              Text4 = Text4 + 1
            Else
              Text5 = Text5 + 1
            End If
          Else
            Text6 = Text6 + 1
          End If
         End If
          Comand_Click (1)
    End Select
End Sub
