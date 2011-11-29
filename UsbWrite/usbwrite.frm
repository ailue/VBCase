VERSION 5.00
Begin VB.Form usbwrite 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "USBWrite"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2385
   Icon            =   "usbwrite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton a 
      Caption         =   "可写(&E)"
      Height          =   855
      Index           =   0
      Left            =   1320
      Picture         =   "usbwrite.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "关闭写保护"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton a 
      Caption         =   "只读(&R)"
      Height          =   855
      Index           =   1
      Left            =   120
      Picture         =   "usbwrite.frx":6814
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "启动写保护"
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "usbwrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click(Index As Integer)
Open "c:\q.bat" For Output As #1
   Print #1, "reg delete HKLM\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies /v WriteProtect /f"
   Print #1, "reg add HKLM\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies /v WriteProtect /t reg_dword /d " & Index & " /f"
   Print #1, "del %0"
Close #1
   Shell "c:\q.bat", vbHide
If Not Err Then
Select Case Index
 Case 1
   MsgBox "设置成功！您的系统现在处于移动设备写保护状态。" & Chr(13) & Chr(13) & "本设置在您插入新的设备时生效。退出程序并不影响写保护状态。" & Chr(13) & Chr(13) & "您可以关闭本程序，需要重新设置时再执行。", vbSystemModal + vbInformation, "写保护启动"
 Case 0
   MsgBox "设置成功！您的系统现在可以正常执行移动设备相关操作。" & Chr(13) & Chr(13) & "本设置在您插入新的设备时生效。" & Chr(13) & Chr(13) & "您可以关闭本程序，需要重新设置时再执行。", vbSystemModal + vbInformation, "写保护关闭"
'   Kill "c:\q.bat"
End Select
End If
End Sub

Private Sub Form_Click()
MsgBox "本程序是 璐绥居士 编写的。适合Windows XP SP2以上版本的用户。如果您觉得好用，请推荐给您的朋友们。", vbInformation + vbSystemModal, "绿色、迷你、免费小软件：移动设备写保护开关"
End Sub
