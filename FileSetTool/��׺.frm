VERSION 5.00
Begin VB.Form 后缀 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "后缀统改 (请在指定当前路径后操作)"
   ClientHeight    =   855
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   3735
   Icon            =   "后缀.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox T 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Text            =   "txt"
      ToolTipText     =   "默认当前路径就是程序所在的路径"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox T 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "htm"
      ToolTipText     =   "请先指定当前路径，以避免出现意想不到的后果"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "目标扩展名(&N)"
      Height          =   180
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原扩展名(&O):"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "后缀"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Top = Sett.Top - Sett.Height
Me.Left = Sett.Left
End Sub

Private Sub OKButton_Click()
Dim sign As String, ss As String
sign = "*." & T(0)
'类型字符串
ss = Dir(sign, vbSystem Or vbHidden)
'提取第一个文件
  While Len(ss) <> 0
    Name ss As Left(ss, Len(ss) - Len(T(0))) & T(1)
    '改名
      ss = Dir(sign, vbSystem Or vbHidden)
  Wend
  MsgBox "修改成功！", vbSystemModal
Unload Me
Sett.SetFocus
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub
