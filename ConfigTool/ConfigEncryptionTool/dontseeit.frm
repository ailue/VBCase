VERSION 5.00
Begin VB.Form dontst 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "配置式加密工具"
   ClientHeight    =   855
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   1815
   Icon            =   "dontseeit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox pW 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   14
      OLEDropMode     =   1  'Manual
      PasswordChar    =   "&"
      TabIndex        =   1
      ToolTipText     =   "输入正确的密码后回车即打开解密对话框。"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label bel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "解密密码输入："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "本程序的诸多快捷键能方便您的操作"
      Top             =   120
      Width           =   1260
   End
   Begin VB.Menu mu1 
      Caption         =   "保护(&P)"
      Begin VB.Menu plus 
         Caption         =   "加密(&P)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu unplus 
         Caption         =   "解密(&-)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu systm 
         Caption         =   "恢复保护(&+)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu desty 
         Caption         =   "完全解保(&D)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu g1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu g1 
         Caption         =   "设隐藏(&H)"
         Index           =   1
         Shortcut        =   {F7}
      End
      Begin VB.Menu g1 
         Caption         =   "使显示(&U)"
         Index           =   2
         Shortcut        =   {F8}
      End
      Begin VB.Menu g1 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu g2 
         Caption         =   "关于(&A)..."
         Index           =   1
         Shortcut        =   {F4}
      End
      Begin VB.Menu g2 
         Caption         =   "退出(&X)"
         Index           =   2
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mu2 
      Caption         =   "配置(&S)"
      Begin VB.Menu delt 
         Caption         =   "CLSID选择(&S)"
         Index           =   0
         Begin VB.Menu CLSID 
            Caption         =   "【手工输入】..."
            Index           =   0
            Shortcut        =   ^{INSERT}
         End
         Begin VB.Menu CLSID 
            Caption         =   "我的电脑"
            Index           =   1
            Shortcut        =   ^Q
         End
         Begin VB.Menu CLSID 
            Caption         =   "我的文档"
            Index           =   2
            Shortcut        =   ^W
         End
         Begin VB.Menu CLSID 
            Caption         =   "回收站"
            Index           =   3
            Shortcut        =   ^E
         End
         Begin VB.Menu CLSID 
            Caption         =   "网上邻居"
            Index           =   4
            Shortcut        =   ^R
         End
         Begin VB.Menu CLSID 
            Caption         =   "控制面板"
            Index           =   5
            Shortcut        =   ^T
         End
         Begin VB.Menu CLSID 
            Caption         =   "打印机"
            Index           =   6
            Shortcut        =   ^Y
         End
         Begin VB.Menu CLSID 
            Caption         =   "计划任务"
            Index           =   7
            Shortcut        =   ^U
         End
         Begin VB.Menu CLSID 
            Caption         =   "扫描仪和数码相机"
            Index           =   8
            Shortcut        =   ^I
         End
         Begin VB.Menu CLSID 
            Caption         =   "Internet Explorer"
            Index           =   9
            Shortcut        =   ^O
         End
         Begin VB.Menu CLSID 
            Caption         =   "&Office项目  →"
            Index           =   10
            Begin VB.Menu CLSIDo 
               Caption         =   "Word"
               Index           =   1
            End
            Begin VB.Menu CLSIDo 
               Caption         =   "Excel"
               Index           =   2
            End
            Begin VB.Menu CLSIDo 
               Caption         =   "PowerPoint"
               Index           =   3
            End
            Begin VB.Menu CLSIDo 
               Caption         =   "Access"
               Index           =   4
            End
            Begin VB.Menu CLSIDo 
               Caption         =   "Outlook"
               Index           =   5
            End
            Begin VB.Menu CLSIDo 
               Caption         =   "HTML文档"
               Index           =   6
            End
         End
         Begin VB.Menu CLSID 
            Caption         =   "更多  →"
            Index           =   11
            Begin VB.Menu CLSIDp 
               Caption         =   "帮助与支持"
               Index           =   1
               Shortcut        =   ^S
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "Windows安全性"
               Index           =   2
               Shortcut        =   ^D
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "运行"
               Index           =   3
               Shortcut        =   ^F
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "搜索"
               Index           =   4
               Shortcut        =   ^G
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "管理工具"
               Index           =   5
               Shortcut        =   ^H
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "网络连接"
               Index           =   6
               Shortcut        =   ^J
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "字体"
               Index           =   7
               Shortcut        =   ^K
            End
         End
      End
      Begin VB.Menu delt 
         Caption         =   "复制码(&C)"
         Index           =   1
      End
      Begin VB.Menu delt 
         Caption         =   "防删处理(&F)"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   {F9}
      End
      Begin VB.Menu delt 
         Caption         =   "运行&WinRAR"
         Index           =   3
         Shortcut        =   ^Z
      End
      Begin VB.Menu delt 
         Caption         =   "&Desktop.ini"
         Index           =   4
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "dontst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cd$, fuck$, sis$, pan$
Private Const 控制面板 = "{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Private Const pwd$ = "jiaxh"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub SetFormTopmost(TheForm As Form)
  SetWindowPos TheForm.hwnd, -1, 0, 0, 0, 0, &H8 + &H2 + &H1
End Sub

Private Sub uset(fir As String)
On Error Resume Next
If fir = "" Then Exit Sub
  ei = MsgBox("解密请选'是'；恢复请选'否'；执行其他操作请选'取消'。", vbInformation + vbSystemModal + vbYesNoCancel, "参数提示")
  Select Case ei
    Case vbYes
      pan = fir
    Case vbNo
      SetAttr fir, vbSystem: End
    Case vbCancel
      sis = fir
  End Select
fuck = fir
If Err Then MsgBox "参数错误！程序即将退出。", vbSystemModal + vbCritical: End
End Sub

Private Sub Form_Load()
SetFormTopmost Me
 cd = 控制面板
 fuck = "MySecretFiles"
 pan = ""
  Call uset(Command)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 pW.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 pW.SetFocus
End Sub

Private Sub mu1_Click()
 unplus.Enabled = IIf(pW = pwd, True, False)
 desty.Enabled = IIf(pW = pwd, True, False)
End Sub

Private Sub mu2_Click()
 delt(3).Enabled = IIf(pW = pwd, True, False): delt(4).Enabled = IIf(pW = pwd, True, False)
End Sub

Private Sub PLUS_Click()
On Error Resume Next
iss = InputBox("请输入目标文件夹名:", "生成文件夹配置文件并作CLSID伪装", fuck)
If iss = "" Then Exit Sub
MkDir iss
 SetAttr iss & "\desktop.ini", 0
 Kill iss & "\desktop.ini"
If delt(2).Checked = True Then
  MkDir iss & "/VISTA..\"
'  SetAttr iss & "/VISTA..\", vbSystem + vbHidden
'  Shell "cmd.exe /c attrib " & iss & "/VISTA..\" & " +h +s", vbHide
 Else
  RmDir iss & "/VISTA..\"
End If
 Open iss & "\desktop.ini" For Output As #1
      Print #1, "[.ShellClassInfo]"
      Print #1, "CLSID=" & cd
 Close #1
SetAttr iss & "\desktop.ini", vbSystem + vbHidden
SetAttr iss, vbSystem
MsgBox "加密成功！", vbSystemModal
sis = iss
End Sub

Private Sub unplus_Click()
On Error Resume Next
 If pan <> "" Then
  iss = pan
 Else
  iss = InputBox("请输入目标文件夹名:", "暂时解除伪装，使能正常存取文件", fuck)
 End If
If iss = "" Then Exit Sub
SetAttr iss, 0
sis = iss
  If pan <> "" Then End
End Sub

Private Sub systm_Click()
iss = InputBox("请输入目标文件夹名:", "恢复伪装状态，保护文件夹不被非法预览", fuck)
If iss = "" Then Exit Sub
SetAttr iss, vbSystem
sis = iss
End Sub

Private Sub desty_Click()
On Error Resume Next
iss = InputBox("请输入目标文件夹名:", "删除配置文件，清理本软件给您带来的不便", fuck)
If iss = "" Then Exit Sub
 SetAttr iss & "\desktop.ini", 0
 Kill iss & "\desktop.ini"
RmDir iss & "/VISTA..\"
SetAttr iss, 0
MsgBox "配置式加密已完全解除！", vbSystemModal
sis = iss
End Sub

Private Sub g1_Click(iNdex As Integer)
On Error Resume Next
Select Case iNdex
 Case 1
   SetAttr sis, vbSystem + vbHidden
 Case 2
   SetAttr sis, vbSystem
End Select
'If Error Then MsgBox "出现错误，请检查您的操作是否有误，然后再试一次。", vbSystemModal + vbCritical
End Sub

Private Sub pw_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
  If pW = "hide" Then Unload Me
  If pW = "exit" Then End
  mu1_Click
  If unplus.Enabled = True Then unplus_Click
 End If
End Sub

Private Sub CLSID_Click(iNdex As Integer)
Select Case iNdex
 Case 1
   cd = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
   '我的电脑
 Case 2
   cd = "{450D8FBA-AD25-11D0-98A8-0800361B1103}"
   '我的文档
 Case 3
   cd = "{645FF040-5081-101B-9F08-00AA002F954E}"
   '回收站
 Case 4
   cd = "{208D2C60-3AEA-1069-A2D7-08002B30309D}"
   '网上邻居
 Case 5
   cd = 控制面板
 Case 6
   cd = "{2227A280-3AEA-1069-A2DE-08002B30309D}"
   '打印机
 Case 7
   cd = "{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"
   '计划任务
 Case 8
   cd = "{E211B736-43FD-11D1-9EFB-0000F8757FCD}"
   '扫描仪
 Case 9
   cd = "{871C5380-42A0-1069-A2EA-08002B30309D}"
   'IE
 Case 0
   cod = InputBox("请在这里输入或粘贴CLSID：", "CLSID输入", cd)
   If cod = "" Then cod = cd
    If Mid(cod, 10, 1) <> "-" Or Mid(cod, 15, 1) <> "-" Or Mid(cod, 20, 1) <> "-" Or Mid(cod, 25, 1) <> "-" Or Len(cod) <> 38 Then MsgBox "格式不匹配 ^?^", vbSystemModal + vbCritical: CLSID_Click (10)
   cd = cod
End Select
End Sub

Private Sub CLSIDo_Click(iNdex As Integer)
Select Case iNdex
 Case 1
   cd = "{00020906-0000-0000-C000-000000000046}"
   'Word
 Case 2
   cd = "{00020810-0000-0000-C000-000000000046}"
   'Excel
 Case 3
   cd = "{64818D10-4F9B-11CF-86EA-00AA00B929E8}"
   'PowerPoint
 Case 4
   cd = "{73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9}"
   'Access
 Case 5
   cd = "{00020D75-0000-0000-C000-000000000046}"
   'Outlook
 Case 6
   cd = "{25336920-03F9-11CF-8FD0-00AA00686F13}"
   'HTML文档
End Select
End Sub

Private Sub CLSIDp_Click(iNdex As Integer)
Select Case iNdex
 Case 1
   cd = "{2559a1f1-21d7-11d4-bdaf-00c04f60b9f0}"
   '帮助与支持
 Case 2
   cd = "{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}"
   'Windows 安全性
 Case 3
   cd = "{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}"
   '运行
 Case 4
   cd = "{2559a1f0-21d7-11d4-bdaf-00c04f60b9f0}"
   '搜索
 Case 5
   cd = "{D20EA4E1-3957-11d2-A40B-0C5020524153}"
   '管理工具
 Case 6
   cd = "{7007ACC7-3202-11D1-AAD2-00805FC1270E}"
   '网络连接
 Case 7
   cd = "{D20EA4E1-3957-11d2-A40B-0C5020524152}"
   '字体
End Select
End Sub

Private Sub delt_Click(iNdex As Integer)
Select Case iNdex
  Case 1
    Clipboard.Clear
    Clipboard.SetText cd
    MsgBox "当前CLSID已复制到系统剪贴板。" & Chr(13) & Chr(13) & "在文件夹名称后加一个半角小数点，再将此信息粘贴，即可达到加密效果。", vbInformation + vbSystemModal, "当前CLSID：" & cd
  Case 2
    delt(iNdex).Checked = IIf(MsgBox("需要对加密文件夹进行防删处理吗？" & Chr(13) & Chr(13) & "设置后一般人就无法删除您的资料了。", vbYesNo + vbQuestion + vbSystemModal, "防删选项") = vbYes, True, False)
  Case 3
    If Dir("C:\Program Files\WinRAR\WinRAR.exe") = "" Then MsgBox "找不到WinRAR！", vbCritical + vbSystemModal, "出错": Exit Sub
    Call Shell("C:\Program Files\WinRAR\WinRAR.exe", vbNormalFocus)
  Case 4
    On Error Resume Next
    Call Shell("NOTEPAD.exe " & sis & "\desktop.ini", vbNormalFocus)
End Select
End Sub

Private Sub g2_Click(iNdex As Integer)
Select Case iNdex
 Case 1
  MsgBox "本程序是 璐绥居士 为 鹰勾鼻子 编写的，仅限私人使用，擅自传播者拖出去打！", vbExclamation + vbSystemModal, "关于"
 Case 2
  End
End Select
End Sub

Private Sub Form_DblClick()
g2_Click (1)
End Sub

Private Sub bel_dblClick()
 MsgBox "密码忘了？找璐绥居士吧，但愿他没忘:)", vbSystemModal + vbQuestion, "密码提示"
End Sub
