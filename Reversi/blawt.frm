VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ߵ��ڰ� 1.2 �ڰ���С���� ����ʿԭ���㷨���԰�"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8445
   Icon            =   "blawt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "blawt.frx":628A
   ScaleHeight     =   6465
   ScaleWidth      =   8445
   StartUpPosition =   1  '����������
   Begin VB.OptionButton fg 
      Caption         =   "�ڷ�����"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   110
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton fg 
      Caption         =   "�׷�����"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   109
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox M_M 
      AutoSize        =   -1  'True
      Height          =   420
      Index           =   1
      Left            =   6840
      Picture         =   "blawt.frx":6594
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   106
      ToolTipText     =   "�����Ļ���������뵥�������ػ�����"
      Top             =   3840
      Width           =   420
   End
   Begin VB.PictureBox M_M 
      AutoSize        =   -1  'True
      Height          =   420
      Index           =   0
      Left            =   6840
      Picture         =   "blawt.frx":6C8E
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   105
      ToolTipText     =   "�����Ļ���������뵥�������ػ�����"
      Top             =   3360
      Width           =   420
   End
   Begin VB.CommandButton Restat 
      Caption         =   "�ؿ�һ��(&O)"
      Height          =   495
      Left            =   6720
      TabIndex        =   104
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton rules 
      Appearance      =   0  'Flat
      Caption         =   "����(&R)"
      Height          =   495
      Left            =   6720
      TabIndex        =   103
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton ed1 
      Cancel          =   -1  'True
      Caption         =   "����(&E)"
      Height          =   495
      Left            =   6720
      TabIndex        =   102
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Comst 
      Caption         =   "��ʼ(&B)"
      Default         =   -1  'True
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame FM 
      Caption         =   "���ڰ���˫�˶�ս"
      Enabled         =   0   'False
      Height          =   5655
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   5775
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   101
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   100
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   99
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   98
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   97
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   96
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   95
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   94
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   93
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   92
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   10
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   91
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   90
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   12
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   89
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   13
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   88
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   14
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   87
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   86
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   85
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   84
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   18
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   83
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   19
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   82
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   20
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   81
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   21
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   80
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   22
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   79
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   23
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   78
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   24
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   77
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   25
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   76
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   26
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   75
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   27
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   74
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   28
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   73
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   29
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   72
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   30
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   71
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   31
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   70
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   32
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   69
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   33
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   68
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   34
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   67
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   35
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   66
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   36
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   65
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   37
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   64
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   38
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   63
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   39
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   62
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   40
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   61
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   41
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   60
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   42
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   59
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   43
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   58
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   44
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   57
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   45
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   56
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   46
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   55
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   47
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   54
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   48
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   53
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   49
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   52
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   50
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   51
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   51
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   50
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   52
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   49
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   53
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   48
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   54
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   47
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   55
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   46
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   56
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   45
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   57
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   44
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   58
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   43
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   59
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   42
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   60
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   41
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   61
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   40
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   62
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   39
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   63
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   38
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   64
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   37
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   65
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   36
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   66
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   35
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   67
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   34
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   68
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   33
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   69
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   32
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   70
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   31
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   71
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   30
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   72
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   29
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   73
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   28
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   74
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   27
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   75
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   26
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   76
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   25
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   77
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   24
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   78
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   23
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   79
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   22
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   80
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   21
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   81
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   20
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   82
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   19
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   83
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   18
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   84
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   17
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   85
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   16
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   86
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   15
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   87
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   14
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   88
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   13
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   89
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   12
         Top             =   4320
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   90
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   11
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   91
         Left            =   960
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   10
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   92
         Left            =   1440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   9
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   93
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   8
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   94
         Left            =   2400
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   7
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   95
         Left            =   2880
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   6
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   96
         Left            =   3360
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   97
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   4
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   98
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   3
         Top             =   4800
         Width           =   495
      End
      Begin VB.PictureBox POi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   99
         Left            =   4800
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   2
         Top             =   4800
         Width           =   495
      End
   End
   Begin VB.CommandButton SL 
      Caption         =   "�������(&S)"
      Height          =   495
      Index           =   0
      Left            =   6720
      TabIndex        =   111
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton SL 
      Caption         =   "�������(&L)"
      Height          =   495
      Index           =   1
      Left            =   6720
      TabIndex        =   112
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      Height          =   960
      Left            =   6810
      Top             =   3330
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   6600
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label sc0re�� 
      Caption         =   "0"
      Height          =   255
      Left            =   7440
      TabIndex        =   108
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label sc0re�� 
      Caption         =   "0"
      Height          =   255
      Left            =   7440
      TabIndex        =   107
      Top             =   3480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const �� = 1, �� = -1  '��������
Dim blk(10, 10) As Integer '���ӱ������
Dim movea(2, 8) As Integer  '�˸�����
Dim flag As Integer  'һ�����ӱ��

Private Sub Comst_Click()   '�����Ԥ����
Comst.Visible = False
SL(0).Visible = True: SL(1).Visible = True   '���ؿ�ʼ��ť��¶��SL��ť
rules.Top = 1440 + 220 'λ���ص�
 FM.Enabled = True    '��ܽ���
  blk(4, 4) = ��
  POi(44).FontSize = 24
  POi(44).Print "��"
  blk(4, 5) = ��
  POi(45).FontSize = 24
  POi(45).Print "��"
  blk(5, 4) = ��
  POi(54).FontSize = 24
  POi(54).Print "��"
  blk(5, 5) = ��
  POi(55).FontSize = 24
  POi(55).Print "��"        '�ڰ�������������д�λ����
    sc0re�� = 2: sc0re�� = 2    '������Ŀ��ʶ
  If flag = �� Then fg(1).Value = True Else fg(0).Value = True  'ִ��ͬ��
End Sub

Private Sub Form_DblClick()
' Dialog.Show
End Sub

Private Sub Form_Load()
 flag = ��   '�׷�����
 movea(0, 0) = 0
 movea(1, 0) = 1
 movea(0, 1) = 1
 movea(1, 1) = 1
 movea(0, 2) = 1
 movea(1, 2) = 0
 movea(0, 3) = 1
 movea(1, 3) = -1
 movea(0, 4) = 0
 movea(1, 4) = -1
 movea(0, 5) = -1
 movea(1, 5) = -1
 movea(0, 6) = -1
 movea(1, 6) = 0
 movea(0, 7) = -1
 movea(1, 7) = 1
 '���������ʼ��
   For p = 0 To 99
     blk(p \ 10, p Mod 10) = 0
     POi(p).Cls
   Next p
 '��������ʼ��
End Sub

Private Sub POi_Click(Index As Integer)
  X% = Index \ 10: Y% = Index Mod 10 '�ӿؼ�������ȡ�����ö�ά����
  �� = ����(X, Y)  '�����ú���һ��
    '��������
    '3  ��ǰλ������
    '4  û�������������,ͬʱѯ���û��Ƿ�����
 Select Case ��    '��׽����ֵ��������ʾ����������
  Case 3
    MsgBox "������3: ��ǰλ�����ӣ����������ӡ�", _
            vbExclamation, "3�Ŵ��� �򵥵��ʧ��": Exit Sub
  Case 4
    ���з� = MsgBox("������4: û������������������ڱ����ӡ�" _
            & Chr(13) & Chr(13) & "������������'����'" & Chr(13) _
            & "����������������'ȡ��'", _
            vbCritical + vbRetryCancel, "4�Ŵ��� ����Ԥ����")
    If ���з� = vbCancel Then flag = -flag: opTn (flag)  '���д���
    Exit Sub
 End Select
 '��������ȷ���ݴ�����
    blk(X, Y) = ��
      '���Ӻ������ص�ǰλ���Ǻ��ӣ����ӻ��ǿ�
      sc0re�� = 0: sc0re�� = 0  '������Ŀ��ʼ��
    For k = 0 To 99  '�����ӱ�������������ϻ���ʵ��
     POi(k).FontSize = 24   'ͼ����С
     POi(k).Cls             '�����ػ�
     Select Case blk(k \ 10, k Mod 10)
       Case 0:  POi(k).Print       '����0�����
       Case ��: POi(k).Print "��": sc0re�� = sc0re�� + 1 '����1������ӣ��ۼƺ�����Ŀ
       Case ��: POi(k).Print "��": sc0re�� = sc0re�� + 1 '������1������ӣ��ۼư�����Ŀ
     End Select
    Next k
If Val(sc0re��) + Val(sc0re��) = 100 Or Val(sc0re��) = 0 Or Val(sc0re��) = 0 Then _
         MsgBox "��ֽ���!" & Chr(13) & Chr(13) _
                 & "�׷� " & sc0re�� & " ��" & Chr(13) _
                 & "�ڷ� " & sc0re�� & " ��" & Chr(13) _
                 & Chr(13) & "��ʤ�븺һĿ��Ȼ ^_^", vbInformation, _
                 "�ڰ���Ƿ���"   'ʤ���ж�
End Sub

Private Sub M_M_Click(Index As Integer)  '�ػ�
   For k = 0 To 99
     POi(k).FontSize = 24
     POi(k).Cls
     Select Case blk(k \ 10, k Mod 10)
       Case ��: POi(k).Print "��"
       Case ��: POi(k).Print "��"
     End Select
    Next k
End Sub

Function ����(X As Integer, Y As Integer) As Integer '��Ҫ�ĺ���
  If blk(X, Y) <> 0 Then ���� = 3: Exit Function      '�����򷵻ش���ֵ
  
  e = False '�ж��Ƿ����ӵ��ۻ�����ֵ
   For i = 0 To 7  '���Ű˸��������ν��������жϣ�������ֵ
    e = e Or fang(X + movea(0, i), Y + movea(1, i), i)
   Next i
   
   If e Then ���� = flag: flag = -flag: opTn (flag) Else ���� = 4  'ȷ�����ӣ����ӿ����򷵻ش���ֵ
End Function
 Sub opTn(ByVal flag As Integer)   '��ǰִ�ӱ�ǵ�ѡ��ť
  Select Case flag
   Case ��: fg(0).Value = True
   Case ��: fg(1).Value = True
  End Select
 End Sub
 
Function fang(ByVal xx As Integer, ByVal yy As Integer, ByVal sign As Integer) As Boolean '�����жϺ���������Ҫ�ĺ���
x_x = xx: y_y = yy   '��ǳ�ʼֵ
  Do  '�ɽ���Զ��һ������������
    If xx < 0 Or xx > 9 Or yy < 0 Or yy > 9 Then fang = False: Exit Function  '�߽����
    If blk(xx, yy) = flag Then Exit Do   '���ֵ�һ���͵�ǰͬ�Ӿ��˳�ѭ������������������
    If blk(xx, yy) = 0 Then fang = False: Exit Function   '�ڷ��ֵ�һ��ͬ��֮ǰ����0�ͽ���������������һ������ļ��
    If blk(xx, yy) = -flag Then xx = xx + movea(0, sign): yy = yy + movea(1, sign)  '�������Ӽ�����������
  Loop
  '����ֹ������ѭ����ȫƾ�ж����ǿ������
      If xx = x_x And yy = y_y Then fang = False: Exit Function '�����һ���Ӿ���ͬ�����޷�ת���˳�
  Do  '��Զ������������һ��ת
    xx = xx - movea(0, sign): yy = yy - movea(1, sign)
    blk(xx, yy) = flag
  Loop Until xx = x_x And yy = y_y
   fang = True  '˳����ɵ�����ʱ������ֵ����������
End Function

Private Sub Restat_Click()   '�ؿ��ж�
  Form_Load    '��ʼ��
  flag = IIf(MsgBox("�׷����֣�", vbYesNo + vbQuestion, "�����ж�") = vbYes, ��, ��)
  Comst_Click  '�忪֮ǰ
End Sub

Private Sub SL_Click(Index As Integer)   'SL�ļ��������
On Error Resume Next
Dim ���ױ�ʶ As Integer   '���׵��ļ�������Ψһ�ģ���չ��lbw
 Select Case Index
   Case 0  '�������
    fnm = InputBox("�����뱣���ļ���:", "�ڰ���浵", "NewOne")
    If fnm = "" Then Exit Sub   '��ȡ���Ļ��˳�
     Open fnm & ".lbw" For Random As #1 Len = 2
     If Err Then MsgBox "���ִ���", vbCritical: Close #1: Exit Sub
       Put #1, 2, -425   '�������ļ�����
       Put #1, 4, flag   '4�Ŵ洢��ǰ���ӱ��
        For k = 0 To 99
          Put #1, k + 5, blk(k \ 10, k Mod 10)   '5���Ժ�洢����
        Next k
     Close #1
     If Not Err Then MsgBox "����ɹ���", vbInformation
   Case 1  '�������
     fnm = InputBox("����������ļ���:", "�ڰ���浵", "NewOne")
     If fnm = "" Then Exit Sub
     Open fnm & ".lbw" For Random As #1 Len = 2
       Get #1, 2, ���ױ�ʶ   '�����жϣ��������
       If ���ױ�ʶ <> -425 Then MsgBox "������ʧ�ܣ����ļ����������ļ�." _
          & Chr(13) & "�ļ�����ɾ��.", vbCritical: Close #1: Kill fnm & ".lbw": Exit Sub
       Get #1, 4, flag      'ȡ��ǰ���ӱ��
        For k = 0 To 99
          Get #1, k + 5, blk(k \ 10, k Mod 10)
        Next k
     Close #1
     MsgBox "����ɹ���", vbInformation
      opTn (flag)
      M_M_Click (Index)  '�ػ�
End Select
End Sub

Private Sub rules_Click()  '�����������д�ģ�����ժ���ı�
msg$ = msg$ + "    �ڰ����Ǵ��й��Ŵ��Ϳ�ʼ��Դ��������Ϸ������������" + Chr(13)
msg$ = msg$ + "����ǳ��鷳����ΪҪ��ͣ�ؽ��ڰ����廥�����������˽���" + Chr(13)
msg$ = msg$ + "��������ͿΪ��ɫ����Ϊ�ڰ���ľ���ģʽһֱ���������ڡ�" + Chr(13) + Chr(13)
msg$ = msg$ + "    ��Ϸ����ǳ��򵥣���һ���������ȷ�����ö���ӣ��ڰ�˫���������壬" + Chr(13)
msg$ = msg$ + "ÿһ�����������̵Ŀհ״����ܡ��Ե����Է�����һ�ӣ������Լ������¶�" + Chr(13)
msg$ = msg$ + "�ɶԷ������£����Ե����Է���������ָ�������Ϊ���ģ�����ȥ������" + Chr(13)
msg$ = msg$ + "���Է������Ӻ������Լ�������(�м䲻���пո�)�����⼸���Է������Ӿ�" + Chr(13)
msg$ = msg$ + "�������ҷ������ӣ������Ե��ˣ��Է����������ӱ��ǳ��˼��ӡ����֡���" + Chr(13)
msg$ = msg$ + "����ӦͬʱӦ�������ҡ����ϡ����µȰ˸�����" + Chr(13)
msg$ = msg$ + "  ������һ�����̵ľֲ�Ϊ����������Ϊ��λ������Ϊ���塢����Ϊ���壺" + Chr(13)
msg$ = msg$ + "   ��������d4���ӣ����̱�Ϊͼ��������Ե�����4�ӡ�" + Chr(13)
msg$ = msg$ + "" + Chr(13)
msg$ = msg$ + "ͼһ:������������ ͼ��:" + Chr(13)
msg$ = msg$ + "  1 2 3 4 5         1 2 3 4 5 " + Chr(13)
msg$ = msg$ + "a �𡡡��𡡡�����a �𡡡���" + Chr(13)
msg$ = msg$ + "b ����񡡡�������b ����񡡡�" + Chr(13)
msg$ = msg$ + "c ������񡡡�����c �������" + Chr(13)
msg$ = msg$ + "d ���񡡡�������d �����" + Chr(13)
msg$ = msg$ + "e ����������������e ����������" + Chr(13)
msg$ = msg$ + "�ж�ʤ��:" + Chr(13)
msg$ = msg$ + "    �����������ˣ������Ӷ��һ��ʤ��������;һ�����ӱ�ȫ�����꣬" + Chr(13)
msg$ = msg$ + "�Է�ʤ����" + Chr(13)
MsgBox msg$, , "�ڰ������"
End Sub

Private Sub ed1_Click()  '�������
  End
End Sub

