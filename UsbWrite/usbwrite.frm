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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton a 
      Caption         =   "��д(&E)"
      Height          =   855
      Index           =   0
      Left            =   1320
      Picture         =   "usbwrite.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "�ر�д����"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton a 
      Caption         =   "ֻ��(&R)"
      Height          =   855
      Index           =   1
      Left            =   120
      Picture         =   "usbwrite.frx":6814
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "����д����"
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
   MsgBox "���óɹ�������ϵͳ���ڴ����ƶ��豸д����״̬��" & Chr(13) & Chr(13) & "���������������µ��豸ʱ��Ч���˳����򲢲�Ӱ��д����״̬��" & Chr(13) & Chr(13) & "�����Թرձ�������Ҫ��������ʱ��ִ�С�", vbSystemModal + vbInformation, "д��������"
 Case 0
   MsgBox "���óɹ�������ϵͳ���ڿ�������ִ���ƶ��豸��ز�����" & Chr(13) & Chr(13) & "���������������µ��豸ʱ��Ч��" & Chr(13) & Chr(13) & "�����Թرձ�������Ҫ��������ʱ��ִ�С�", vbSystemModal + vbInformation, "д�����ر�"
'   Kill "c:\q.bat"
End Select
End If
End Sub

Private Sub Form_Click()
MsgBox "�������� ����ʿ ��д�ġ��ʺ�Windows XP SP2���ϰ汾���û�����������ú��ã����Ƽ������������ǡ�", vbInformation + vbSystemModal, "��ɫ�����㡢���С������ƶ��豸д��������"
End Sub
