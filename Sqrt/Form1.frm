VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���ģ���ֹ�����"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5010
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Text            =   "2231341"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "�����������Ա�"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "������"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�����������"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SOS = 15           '������Чλ��
Dim a(SOS) As Integer  '����ֶ�����
Dim poin As Integer  '����С����λ��
 
Private Sub Command1_Click()

On Error Resume Next
'��ֲ���
 inpint = Int(Text1)
 inpflt = Val(Text1) - inpint   '��ֵת��ȷ��������С��
 
 If Err Then MsgBox "�벻Ҫ��������ֵĶ���": Exit Sub  '������
 
 For i = 0 To SOS - 1
     a(i) = inpint Mod 100
     inpint = inpint \ 100
 Next i                '��ȡ�����ֶ�ֵ

 For zero = SOS - 1 To 0 Step -1
     If a(zero) <> 0 Then Exit For
 Next zero              '��ȡ�����
 
 y = 0
 For k = zero To zero / 2 Step -1
     qq = a(k): a(k) = a(y): a(y) = qq
     y = y + 1
 Next k             '���黻λ���±������ʾ��λ����λ

 For i = zero + 1 To SOS
     a(i) = Int(inpflt * 100)
     inpflt = inpflt * 100 - a(i)
 Next i             'С����λ
 
 poin = SOS - 1 - zero     'С����λ��

'���㲿��
 nextbchs = a(0)     '������ʱ��������
 shang = 0           '������ֵ
    
 For i = 0 To SOS - 1    '�������ѭ��
     For smp = 1 To nextbchs / 2 + 1  'Ѱ�Һ��ʵĿɿ�����
       If (shang * 20 + smp) * smp > nextbchs Then Exit For
     Next smp
     smp = smp - 1          'Խ����һ
     yushu = nextbchs - (shang * 20 + smp) * smp    '���μ��������
     shang = shang * 10 + smp         '���ν��ֵ
     nextbchs = yushu * 100 + a(i + 1)      '��һ�εı�������
 Next i

 Text2 = shang / 10 ^ poin      'С�����������������
 Text3 = Val(Text1) ^ 0.5      '��Ϊ��������ĶԱ�
End Sub
