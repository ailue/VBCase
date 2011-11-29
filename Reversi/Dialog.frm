VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "è´ËçºÚ°×Æå´æµµ"
   ClientHeight    =   3495
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   2640
      Pattern         =   "*.txt"
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2190
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "È¡Ïû"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "È·¶¨"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fls As String
Private Sub Dir1_Change()
   File1.Path = Dir1.Path
     App.Path = File1.Path
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
  fls = File1.FileName
End Sub
