VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Sett 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "文件设置小工具  璐绥居士编写"
   ClientHeight    =   975
   ClientLeft      =   6840
   ClientTop       =   2910
   ClientWidth     =   5280
   Icon            =   "Set.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   975
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CDg 
      Left            =   2400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.*"
      DialogTitle     =   "文件选择"
      Filter          =   "所有文件 (*.*)"
   End
   Begin VB.CommandButton Co4 
      Caption         =   "存在性"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "检查指定文件是否存在。快捷键 F4"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Co3 
      Caption         =   "复制绝对路径"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      ToolTipText     =   "将当前文件的绝对路径复制到剪贴板上。快捷键 F3"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Co2 
      Caption         =   "属性设置"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "设置指定文件的属性。快捷键 F2"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Co1 
      Caption         =   "后缀统改"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "将指定目录下同一扩展名的文件统统改为另一扩展名。快捷键 F1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Co0 
      Caption         =   "改名(&G)"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "给指定文件另起一名称。快捷键 F5"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Te1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "清空文本框请按ESC"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label La1 
      Caption         =   "请双击此处或将文件拖到这里"
      Height          =   420
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   "可以直接输入文件名进行当前目录下的操作"
      Top             =   120
      Width           =   1320
   End
End
Attribute VB_Name = "Sett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public add As String, fm As String
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H8
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub SetFormTopmost(TheForm As Form)

SetWindowPos TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE

End Sub

Private Sub Form_DblClick()
'关于信息
MsgBox "文件设置小工具" & Chr(13) & "版本: 1.0" & Chr(13) & "璐绥基地出品" & Chr(13) & Chr(13) & "本软件没什么技术含量，主要为了方便大家的文件设置操作。" & Chr(13) & "程序绿色免费，小巧玲珑。希望您推荐给你的朋友们。", vbInformation + vbSystemModal, "关于"
End Sub

Private Sub Form_Load()
SetFormTopmost Me
add = ""
If Not Command = "" Then Te1 = Command
End Sub

Private Sub La1_dblClick()
'通用对话框
 CDg.Action = 1
 fm = CDg.FileTitle
 add = CDg.FileName
 Te1 = add
' CurDir = Left(add, Len(add) - Len(fm))
End Sub

Private Sub La1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Call Te1_OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub Te1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'OLE拖动
 Dim TF As Variant
  For Each TF In Data.Files
    add = TF
    Te1 = add
     For op = Len(add) - 2 To 2 Step -1
       If Mid(add, op, 1) = "\" Then fm = Mid(add, op + 1): Exit For
     Next op
  Next
End Sub

Private Sub Co0_Click()
'改名
 If Te1 = "" Then Te1.SetFocus: Exit Sub
' Te1 = Dir(Te1 & "*")
  If Dir(Te1, vbHidden Or vbSystem) = "" Then MsgBox "没有选定文件！", vbSystemModal: Te1.SetFocus: Exit Sub
ttt = InputBox("你选择的文件是 " & Dir(Te1, vbHidden Or vbSystem) & Chr(13) & Chr(13) & "请输入新文件名:", "强力改名", Te1, 6795, 4595)
  If ttt = "" Then Te1.SetFocus: Exit Sub
On Error Resume Next
Name Te1 As ttt
 If Not Err Then
   Te1 = ttt: MsgBox "改名成功！", vbSystemModal
  Else: MsgBox "改名失败！请检查文件的使用情况。"
 End If
End Sub

Private Sub Co1_Click()
'扩展名
后缀.Show
End Sub

Private Sub Co2_Click()
'属性
 If Te1 = "" Then Te1.SetFocus: Exit Sub
' Te1 = Dir(Te1 & "*")
属性.Show
If Err = 53 Then MsgBox "文件未找见！请确定其存在性，然后再试一次。", vbSystemModal: 属性.Hide: Sett.Show
End Sub

Private Sub Co3_Click()
'复制绝对路径
 Clipboard.Clear
    If Len(Te1) <> 0 Then
      cud = IIf(Not Right(CurDir, 1) = "\", CurDir & "\", CurDir)
      Clipboard.SetText cud & Te1
    Else: Te1.SetFocus: Exit Sub
    End If
If Mid(Te1, 2, 1) = ":" Then Clipboard.Clear: Clipboard.SetText Te1
'If add <> "" Then Clipboard.SetText add
MsgBox "地址已复制到剪贴板中。", vbSystemModal
End Sub

Private Sub Co4_Click()
'存在性
 If Te1 = "" Then Te1.SetFocus: Exit Sub
On Error Resume Next
MkDir Te1
If Err Then
MsgBox "存在该文件。", vbSystemModal
Else: RmDir Te1: MsgBox "指定文件不存在！", vbSystemModal
End If
End Sub

Private Sub Te1_KeyUp(KeyCode As Integer, Shift As Integer)
'快捷键
Select Case KeyCode
  Case vbKeyF1
   Co1_Click
  Case vbKeyF2
   Co2_Click
  Case vbKeyF3
   Co3_Click
  Case vbKeyF4
   Co4_Click
  Case vbKeyF5
   Co0_Click
  Case vbKeyEscape
   Te1 = "": add = "": fm = ""
End Select
End Sub
