VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "编程模拟手工开方"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5010
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "给我算"
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
      Caption         =   "机器计算结果对比"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "计算结果"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "请输入计算数"
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
Const SOS = 15           '定义有效位数
Dim a(SOS) As Integer  '定义分段数组
Dim poin As Integer  '定义小数点位置
 
Private Sub Command1_Click()

On Error Resume Next
'拆分部分
 inpint = Int(Text1)
 inpflt = Val(Text1) - inpint   '数值转换确定整数和小数
 
 If Err Then MsgBox "请不要输入非数字的东东": Exit Sub  '错误处理
 
 For i = 0 To SOS - 1
     a(i) = inpint Mod 100
     inpint = inpint \ 100
 Next i                '提取整数分段值

 For zero = SOS - 1 To 0 Step -1
     If a(zero) <> 0 Then Exit For
 Next zero              '获取非零点
 
 y = 0
 For k = zero To zero / 2 Step -1
     qq = a(k): a(k) = a(y): a(y) = qq
     y = y + 1
 Next k             '数组换位，下标递增表示高位到低位

 For i = zero + 1 To SOS
     a(i) = Int(inpflt * 100)
     inpflt = inpflt * 100 - a(i)
 Next i             '小数补位
 
 poin = SOS - 1 - zero     '小数点位置

'计算部分
 nextbchs = a(0)     '定义临时被开方数
 shang = 0           '定义结果值
    
 For i = 0 To SOS - 1    '计算次数循环
     For smp = 1 To nextbchs / 2 + 1  '寻找合适的可开方数
       If (shang * 20 + smp) * smp > nextbchs Then Exit For
     Next smp
     smp = smp - 1          '越界后减一
     yushu = nextbchs - (shang * 20 + smp) * smp    '本次计算的余数
     shang = shang * 10 + smp         '本次结果值
     nextbchs = yushu * 100 + a(i + 1)      '下一次的被开方数
 Next i

 Text2 = shang / 10 ^ poin      '小数点调整，输出最后结果
 Text3 = Val(Text1) ^ 0.5      '作为机器运算的对比
End Sub
