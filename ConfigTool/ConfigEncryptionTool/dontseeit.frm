VERSION 5.00
Begin VB.Form dontst 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����ʽ���ܹ���"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox pW 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   14
      OLEDropMode     =   1  'Manual
      PasswordChar    =   "&"
      TabIndex        =   1
      ToolTipText     =   "������ȷ�������س����򿪽��ܶԻ���"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label bel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�����������룺"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "�����������ݼ��ܷ������Ĳ���"
      Top             =   120
      Width           =   1260
   End
   Begin VB.Menu mu1 
      Caption         =   "����(&P)"
      Begin VB.Menu plus 
         Caption         =   "����(&P)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu unplus 
         Caption         =   "����(&-)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu systm 
         Caption         =   "�ָ�����(&+)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu desty 
         Caption         =   "��ȫ�Ᵽ(&D)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu g1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu g1 
         Caption         =   "������(&H)"
         Index           =   1
         Shortcut        =   {F7}
      End
      Begin VB.Menu g1 
         Caption         =   "ʹ��ʾ(&U)"
         Index           =   2
         Shortcut        =   {F8}
      End
      Begin VB.Menu g1 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu g2 
         Caption         =   "����(&A)..."
         Index           =   1
         Shortcut        =   {F4}
      End
      Begin VB.Menu g2 
         Caption         =   "�˳�(&X)"
         Index           =   2
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mu2 
      Caption         =   "����(&S)"
      Begin VB.Menu delt 
         Caption         =   "CLSIDѡ��(&S)"
         Index           =   0
         Begin VB.Menu CLSID 
            Caption         =   "���ֹ����롿..."
            Index           =   0
            Shortcut        =   ^{INSERT}
         End
         Begin VB.Menu CLSID 
            Caption         =   "�ҵĵ���"
            Index           =   1
            Shortcut        =   ^Q
         End
         Begin VB.Menu CLSID 
            Caption         =   "�ҵ��ĵ�"
            Index           =   2
            Shortcut        =   ^W
         End
         Begin VB.Menu CLSID 
            Caption         =   "����վ"
            Index           =   3
            Shortcut        =   ^E
         End
         Begin VB.Menu CLSID 
            Caption         =   "�����ھ�"
            Index           =   4
            Shortcut        =   ^R
         End
         Begin VB.Menu CLSID 
            Caption         =   "�������"
            Index           =   5
            Shortcut        =   ^T
         End
         Begin VB.Menu CLSID 
            Caption         =   "��ӡ��"
            Index           =   6
            Shortcut        =   ^Y
         End
         Begin VB.Menu CLSID 
            Caption         =   "�ƻ�����"
            Index           =   7
            Shortcut        =   ^U
         End
         Begin VB.Menu CLSID 
            Caption         =   "ɨ���Ǻ��������"
            Index           =   8
            Shortcut        =   ^I
         End
         Begin VB.Menu CLSID 
            Caption         =   "Internet Explorer"
            Index           =   9
            Shortcut        =   ^O
         End
         Begin VB.Menu CLSID 
            Caption         =   "&Office��Ŀ  ��"
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
               Caption         =   "HTML�ĵ�"
               Index           =   6
            End
         End
         Begin VB.Menu CLSID 
            Caption         =   "����  ��"
            Index           =   11
            Begin VB.Menu CLSIDp 
               Caption         =   "������֧��"
               Index           =   1
               Shortcut        =   ^S
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "Windows��ȫ��"
               Index           =   2
               Shortcut        =   ^D
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "����"
               Index           =   3
               Shortcut        =   ^F
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "����"
               Index           =   4
               Shortcut        =   ^G
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "������"
               Index           =   5
               Shortcut        =   ^H
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "��������"
               Index           =   6
               Shortcut        =   ^J
            End
            Begin VB.Menu CLSIDp 
               Caption         =   "����"
               Index           =   7
               Shortcut        =   ^K
            End
         End
      End
      Begin VB.Menu delt 
         Caption         =   "������(&C)"
         Index           =   1
      End
      Begin VB.Menu delt 
         Caption         =   "��ɾ����(&F)"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   {F9}
      End
      Begin VB.Menu delt 
         Caption         =   "����&WinRAR"
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
Private Const ������� = "{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Private Const pwd$ = "jiaxh"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub SetFormTopmost(TheForm As Form)
  SetWindowPos TheForm.hwnd, -1, 0, 0, 0, 0, &H8 + &H2 + &H1
End Sub

Private Sub uset(fir As String)
On Error Resume Next
If fir = "" Then Exit Sub
  ei = MsgBox("������ѡ'��'���ָ���ѡ'��'��ִ������������ѡ'ȡ��'��", vbInformation + vbSystemModal + vbYesNoCancel, "������ʾ")
  Select Case ei
    Case vbYes
      pan = fir
    Case vbNo
      SetAttr fir, vbSystem: End
    Case vbCancel
      sis = fir
  End Select
fuck = fir
If Err Then MsgBox "�������󣡳��򼴽��˳���", vbSystemModal + vbCritical: End
End Sub

Private Sub Form_Load()
SetFormTopmost Me
 cd = �������
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
iss = InputBox("������Ŀ���ļ�����:", "�����ļ��������ļ�����CLSIDαװ", fuck)
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
MsgBox "���ܳɹ���", vbSystemModal
sis = iss
End Sub

Private Sub unplus_Click()
On Error Resume Next
 If pan <> "" Then
  iss = pan
 Else
  iss = InputBox("������Ŀ���ļ�����:", "��ʱ���αװ��ʹ��������ȡ�ļ�", fuck)
 End If
If iss = "" Then Exit Sub
SetAttr iss, 0
sis = iss
  If pan <> "" Then End
End Sub

Private Sub systm_Click()
iss = InputBox("������Ŀ���ļ�����:", "�ָ�αװ״̬�������ļ��в����Ƿ�Ԥ��", fuck)
If iss = "" Then Exit Sub
SetAttr iss, vbSystem
sis = iss
End Sub

Private Sub desty_Click()
On Error Resume Next
iss = InputBox("������Ŀ���ļ�����:", "ɾ�������ļ�������������������Ĳ���", fuck)
If iss = "" Then Exit Sub
 SetAttr iss & "\desktop.ini", 0
 Kill iss & "\desktop.ini"
RmDir iss & "/VISTA..\"
SetAttr iss, 0
MsgBox "����ʽ��������ȫ�����", vbSystemModal
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
'If Error Then MsgBox "���ִ����������Ĳ����Ƿ�����Ȼ������һ�Ρ�", vbSystemModal + vbCritical
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
   '�ҵĵ���
 Case 2
   cd = "{450D8FBA-AD25-11D0-98A8-0800361B1103}"
   '�ҵ��ĵ�
 Case 3
   cd = "{645FF040-5081-101B-9F08-00AA002F954E}"
   '����վ
 Case 4
   cd = "{208D2C60-3AEA-1069-A2D7-08002B30309D}"
   '�����ھ�
 Case 5
   cd = �������
 Case 6
   cd = "{2227A280-3AEA-1069-A2DE-08002B30309D}"
   '��ӡ��
 Case 7
   cd = "{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"
   '�ƻ�����
 Case 8
   cd = "{E211B736-43FD-11D1-9EFB-0000F8757FCD}"
   'ɨ����
 Case 9
   cd = "{871C5380-42A0-1069-A2EA-08002B30309D}"
   'IE
 Case 0
   cod = InputBox("�������������ճ��CLSID��", "CLSID����", cd)
   If cod = "" Then cod = cd
    If Mid(cod, 10, 1) <> "-" Or Mid(cod, 15, 1) <> "-" Or Mid(cod, 20, 1) <> "-" Or Mid(cod, 25, 1) <> "-" Or Len(cod) <> 38 Then MsgBox "��ʽ��ƥ�� ^?^", vbSystemModal + vbCritical: CLSID_Click (10)
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
   'HTML�ĵ�
End Select
End Sub

Private Sub CLSIDp_Click(iNdex As Integer)
Select Case iNdex
 Case 1
   cd = "{2559a1f1-21d7-11d4-bdaf-00c04f60b9f0}"
   '������֧��
 Case 2
   cd = "{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}"
   'Windows ��ȫ��
 Case 3
   cd = "{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}"
   '����
 Case 4
   cd = "{2559a1f0-21d7-11d4-bdaf-00c04f60b9f0}"
   '����
 Case 5
   cd = "{D20EA4E1-3957-11d2-A40B-0C5020524153}"
   '������
 Case 6
   cd = "{7007ACC7-3202-11D1-AAD2-00805FC1270E}"
   '��������
 Case 7
   cd = "{D20EA4E1-3957-11d2-A40B-0C5020524152}"
   '����
End Select
End Sub

Private Sub delt_Click(iNdex As Integer)
Select Case iNdex
  Case 1
    Clipboard.Clear
    Clipboard.SetText cd
    MsgBox "��ǰCLSID�Ѹ��Ƶ�ϵͳ�����塣" & Chr(13) & Chr(13) & "���ļ������ƺ��һ�����С���㣬�ٽ�����Ϣճ�������ɴﵽ����Ч����", vbInformation + vbSystemModal, "��ǰCLSID��" & cd
  Case 2
    delt(iNdex).Checked = IIf(MsgBox("��Ҫ�Լ����ļ��н��з�ɾ������" & Chr(13) & Chr(13) & "���ú�һ���˾��޷�ɾ�����������ˡ�", vbYesNo + vbQuestion + vbSystemModal, "��ɾѡ��") = vbYes, True, False)
  Case 3
    If Dir("C:\Program Files\WinRAR\WinRAR.exe") = "" Then MsgBox "�Ҳ���WinRAR��", vbCritical + vbSystemModal, "����": Exit Sub
    Call Shell("C:\Program Files\WinRAR\WinRAR.exe", vbNormalFocus)
  Case 4
    On Error Resume Next
    Call Shell("NOTEPAD.exe " & sis & "\desktop.ini", vbNormalFocus)
End Select
End Sub

Private Sub g2_Click(iNdex As Integer)
Select Case iNdex
 Case 1
  MsgBox "�������� ����ʿ Ϊ ӥ������ ��д�ģ�����˽��ʹ�ã����Դ������ϳ�ȥ��", vbExclamation + vbSystemModal, "����"
 Case 2
  End
End Select
End Sub

Private Sub Form_DblClick()
g2_Click (1)
End Sub

Private Sub bel_dblClick()
 MsgBox "�������ˣ�������ʿ�ɣ���Ը��û��:)", vbSystemModal + vbQuestion, "������ʾ"
End Sub
