VERSION 5.00
Begin VB.Form 属性 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "属性设置"
   ClientHeight    =   1095
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   4470
   Icon            =   "属性.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "别名(&A)"
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "目录(&D)"
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "存档(&V)"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "系统(&H)"
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "隐藏(&H)"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只读(&R)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "属性"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private att As Byte, at0 As Byte

Private Sub Form_Load()
Me.Top = Sett.Top - Sett.Height - 100
Me.Left = Sett.Left

Me.Caption = "属性设置 - " & Dir(Sett.Te1, vbHidden Or vbSystem)
On Error Resume Next
att = GetAttr(Sett.Te1)
at0 = att
  If att >= vbAlias Then Check1(5).Value = 1: att = att - 64
  If att >= vbArchive Then Check1(3).Value = 1: att = att - 32
  If att >= vbDirectory Then Check1(4).Value = 1: att = att - 16
  If att >= vbSystem Then Check1(2).Value = 1: att = att - 4
  If att >= vbHidden Then Check1(1).Value = 1: att = att - 2
  If att >= vbReadOnly Then Check1(0).Value = 1
End Sub

Private Sub OKButton_Click()
 att = 0
On Error Resume Next
For o = 0 To 5
  If Check1(o).Value = 1 Then
   Select Case o
     Case 0
      att = att + 1
     Case 1
      att = att + 2
     Case 2
      att = att + 4
     Case 3
      att = att + 32
     Case 4
      att = att + 16
     Case 5
      att = att + 64
   End Select
  End If
Next o
If att = at0 Then GoTo en
SetAttr Sett.Te1, att
If Err Then
MsgBox "出现错误！请确定您的操作是否正确，然后再试一次。", vbSystemModal
Else: MsgBox "操作成功！", vbSystemModal
End If
en:
Unload Me
Sett.SetFocus
End Sub

Private Sub Check1_Click(Index As Integer)
 If Index <> 4 Then
  If Check1(4).Value = 1 Then Check1(4).Value = 0
 End If
 If Index = 5 Then MsgBox "这个属性一般不能设置。", vbSystemModal
End Sub
