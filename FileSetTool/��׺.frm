VERSION 5.00
Begin VB.Form ��׺ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��׺ͳ�� (����ָ����ǰ·�������)"
   ClientHeight    =   855
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   3735
   Icon            =   "��׺.frx":0000
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
      ToolTipText     =   "Ĭ�ϵ�ǰ·�����ǳ������ڵ�·��"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox T 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "htm"
      ToolTipText     =   "����ָ����ǰ·�����Ա���������벻���ĺ��"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ŀ����չ��(&N)"
      Height          =   180
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ԭ��չ��(&O):"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "��׺"
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
'�����ַ���
ss = Dir(sign, vbSystem Or vbHidden)
'��ȡ��һ���ļ�
  While Len(ss) <> 0
    Name ss As Left(ss, Len(ss) - Len(T(0))) & T(1)
    '����
      ss = Dir(sign, vbSystem Or vbHidden)
  Wend
  MsgBox "�޸ĳɹ���", vbSystemModal
Unload Me
Sett.SetFocus
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub
