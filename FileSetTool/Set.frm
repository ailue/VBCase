VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Sett 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�ļ�����С����  ����ʿ��д"
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
      DialogTitle     =   "�ļ�ѡ��"
      Filter          =   "�����ļ� (*.*)"
   End
   Begin VB.CommandButton Co4 
      Caption         =   "������"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "���ָ���ļ��Ƿ���ڡ���ݼ� F4"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Co3 
      Caption         =   "���ƾ���·��"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      ToolTipText     =   "����ǰ�ļ��ľ���·�����Ƶ��������ϡ���ݼ� F3"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Co2 
      Caption         =   "��������"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "����ָ���ļ������ԡ���ݼ� F2"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Co1 
      Caption         =   "��׺ͳ��"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "��ָ��Ŀ¼��ͬһ��չ�����ļ�ͳͳ��Ϊ��һ��չ������ݼ� F1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Co0 
      Caption         =   "����(&G)"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "��ָ���ļ�����һ���ơ���ݼ� F5"
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
      ToolTipText     =   "����ı����밴ESC"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label La1 
      Caption         =   "��˫���˴����ļ��ϵ�����"
      Height          =   420
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   "����ֱ�������ļ������е�ǰĿ¼�µĲ���"
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
'������Ϣ
MsgBox "�ļ�����С����" & Chr(13) & "�汾: 1.0" & Chr(13) & "�����س�Ʒ" & Chr(13) & Chr(13) & "�����ûʲô������������ҪΪ�˷����ҵ��ļ����ò�����" & Chr(13) & "������ɫ��ѣ�С�����硣ϣ�����Ƽ�����������ǡ�", vbInformation + vbSystemModal, "����"
End Sub

Private Sub Form_Load()
SetFormTopmost Me
add = ""
If Not Command = "" Then Te1 = Command
End Sub

Private Sub La1_dblClick()
'ͨ�öԻ���
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
'OLE�϶�
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
'����
 If Te1 = "" Then Te1.SetFocus: Exit Sub
' Te1 = Dir(Te1 & "*")
  If Dir(Te1, vbHidden Or vbSystem) = "" Then MsgBox "û��ѡ���ļ���", vbSystemModal: Te1.SetFocus: Exit Sub
ttt = InputBox("��ѡ����ļ��� " & Dir(Te1, vbHidden Or vbSystem) & Chr(13) & Chr(13) & "���������ļ���:", "ǿ������", Te1, 6795, 4595)
  If ttt = "" Then Te1.SetFocus: Exit Sub
On Error Resume Next
Name Te1 As ttt
 If Not Err Then
   Te1 = ttt: MsgBox "�����ɹ���", vbSystemModal
  Else: MsgBox "����ʧ�ܣ������ļ���ʹ�������"
 End If
End Sub

Private Sub Co1_Click()
'��չ��
��׺.Show
End Sub

Private Sub Co2_Click()
'����
 If Te1 = "" Then Te1.SetFocus: Exit Sub
' Te1 = Dir(Te1 & "*")
����.Show
If Err = 53 Then MsgBox "�ļ�δ�Ҽ�����ȷ��������ԣ�Ȼ������һ�Ρ�", vbSystemModal: ����.Hide: Sett.Show
End Sub

Private Sub Co3_Click()
'���ƾ���·��
 Clipboard.Clear
    If Len(Te1) <> 0 Then
      cud = IIf(Not Right(CurDir, 1) = "\", CurDir & "\", CurDir)
      Clipboard.SetText cud & Te1
    Else: Te1.SetFocus: Exit Sub
    End If
If Mid(Te1, 2, 1) = ":" Then Clipboard.Clear: Clipboard.SetText Te1
'If add <> "" Then Clipboard.SetText add
MsgBox "��ַ�Ѹ��Ƶ��������С�", vbSystemModal
End Sub

Private Sub Co4_Click()
'������
 If Te1 = "" Then Te1.SetFocus: Exit Sub
On Error Resume Next
MkDir Te1
If Err Then
MsgBox "���ڸ��ļ���", vbSystemModal
Else: RmDir Te1: MsgBox "ָ���ļ������ڣ�", vbSystemModal
End If
End Sub

Private Sub Te1_KeyUp(KeyCode As Integer, Shift As Integer)
'��ݼ�
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
