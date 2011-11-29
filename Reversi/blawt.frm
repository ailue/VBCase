VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "颠倒黑白 1.2 黑白棋小程序 璐绥居士原创算法测试版"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton fg 
      Caption         =   "黑方下子"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   110
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton fg 
      Caption         =   "白方下子"
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
      ToolTipText     =   "如果屏幕被擦除，请单击这里重绘棋子"
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
      ToolTipText     =   "如果屏幕被擦除，请单击这里重绘棋子"
      Top             =   3360
      Width           =   420
   End
   Begin VB.CommandButton Restat 
      Caption         =   "重开一局(&O)"
      Height          =   495
      Left            =   6720
      TabIndex        =   104
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton rules 
      Appearance      =   0  'Flat
      Caption         =   "规则(&R)"
      Height          =   495
      Left            =   6720
      TabIndex        =   103
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton ed1 
      Cancel          =   -1  'True
      Caption         =   "结束(&E)"
      Height          =   495
      Left            =   6720
      TabIndex        =   102
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Comst 
      Caption         =   "开始(&B)"
      Default         =   -1  'True
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame FM 
      Caption         =   "璐绥黑白棋双人对战"
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
      Caption         =   "保存棋局(&S)"
      Height          =   495
      Index           =   0
      Left            =   6720
      TabIndex        =   111
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton SL 
      Caption         =   "载入棋局(&L)"
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
   Begin VB.Label sc0re白 
      Caption         =   "0"
      Height          =   255
      Left            =   7440
      TabIndex        =   108
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label sc0re黑 
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
Private Const 黑 = 1, 白 = -1  '常量声明
Dim blk(10, 10) As Integer '棋子标记数组
Dim movea(2, 8) As Integer  '八个方向
Dim flag As Integer  '一方下子标记

Private Sub Comst_Click()   '新棋局预处理
Comst.Visible = False
SL(0).Visible = True: SL(1).Visible = True   '隐藏开始按钮，露出SL按钮
rules.Top = 1440 + 220 '位置重调
 FM.Enabled = True    '框架解锁
  blk(4, 4) = 黑
  POi(44).FontSize = 24
  POi(44).Print "●"
  blk(4, 5) = 白
  POi(45).FontSize = 24
  POi(45).Print "○"
  blk(5, 4) = 白
  POi(54).FontSize = 24
  POi(54).Print "○"
  blk(5, 5) = 黑
  POi(55).FontSize = 24
  POi(55).Print "●"        '黑白棋规则，棋盘正中错位四字
    sc0re黑 = 2: sc0re白 = 2    '棋子数目标识
  If flag = 黑 Then fg(1).Value = True Else fg(0).Value = True  '执子同步
End Sub

Private Sub Form_DblClick()
' Dialog.Show
End Sub

Private Sub Form_Load()
 flag = 白   '白方先手
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
 '方向数组初始化
   For p = 0 To 99
     blk(p \ 10, p Mod 10) = 0
     POi(p).Cls
   Next p
 '棋标数组初始化
End Sub

Private Sub POi_Click(Index As Integer)
  X% = Index \ 10: Y% = Index Mod 10 '从控件数组提取分析用二维数组
  を = 下子(X, Y)  '仅调用函数一次
    '错误常量表：
    '3  当前位置有子
    '4  没有满足吃子条件,同时询问用户是否让行
 Select Case を    '捕捉错误值，警告提示并结束运行
  Case 3
    MsgBox "错误编号3: 当前位置有子！请重新下子。", _
            vbExclamation, "3号错误 简单点击失误": Exit Sub
  Case 4
    让行否 = MsgBox("错误编号4: 没有满足吃子条件。请在别处下子。" _
            & Chr(13) & Chr(13) & "重新下子请点击'重试'" & Chr(13) _
            & "放弃下子让行请点击'取消'", _
            vbCritical + vbRetryCancel, "4号错误 规则预定义")
    If 让行否 = vbCancel Then flag = -flag: opTn (flag)  '让行处理
    Exit Sub
 End Select
 '以下是正确数据处理方法
    blk(X, Y) = を
      '下子函数返回当前位置是黑子，白子还是空
      sc0re黑 = 0: sc0re白 = 0  '棋子数目初始化
    For k = 0 To 99  '由棋子标记数组在棋盘上画出实体
     POi(k).FontSize = 24   '图案大小
     POi(k).Cls             '清屏重绘
     Select Case blk(k \ 10, k Mod 10)
       Case 0:  POi(k).Print       '常数0代表空
       Case 黑: POi(k).Print "●": sc0re黑 = sc0re黑 + 1 '常数1代表黑子，累计黑子数目
       Case 白: POi(k).Print "○": sc0re白 = sc0re白 + 1 '常数－1代表白子，累计白子数目
     End Select
    Next k
If Val(sc0re黑) + Val(sc0re白) = 100 Or Val(sc0re黑) = 0 Or Val(sc0re白) = 0 Then _
         MsgBox "棋局结束!" & Chr(13) & Chr(13) _
                 & "白方 " & sc0re白 & " 子" & Chr(13) _
                 & "黑方 " & sc0re黑 & " 子" & Chr(13) _
                 & Chr(13) & "孰胜孰负一目了然 ^_^", vbInformation, _
                 "黑白棋记分牌"   '胜负判断
End Sub

Private Sub M_M_Click(Index As Integer)  '重绘
   For k = 0 To 99
     POi(k).FontSize = 24
     POi(k).Cls
     Select Case blk(k \ 10, k Mod 10)
       Case 黑: POi(k).Print "●"
       Case 白: POi(k).Print "○"
     End Select
    Next k
End Sub

Function 下子(X As Integer, Y As Integer) As Integer '主要的函数
  If blk(X, Y) <> 0 Then 下子 = 3: Exit Function      '有子则返回错误值
  
  e = False '判断是否下子的累积布尔值
   For i = 0 To 7  '沿着八个方向依次进行下子判断，传递临值
    e = e Or fang(X + movea(0, i), Y + movea(1, i), i)
   Next i
   
   If e Then 下子 = flag: flag = -flag: opTn (flag) Else 下子 = 4  '确认下子，无子可下则返回错误值
End Function
 Sub opTn(ByVal flag As Integer)   '当前执子标记单选按钮
  Select Case flag
   Case 白: fg(0).Value = True
   Case 黑: fg(1).Value = True
  End Select
 End Sub
 
Function fang(ByVal xx As Integer, ByVal yy As Integer, ByVal sign As Integer) As Boolean '下子判断函数，最重要的函数
x_x = xx: y_y = yy   '标记初始值
  Do  '由近向远逐一检索分析各子
    If xx < 0 Or xx > 9 Or yy < 0 Or yy > 9 Then fang = False: Exit Function  '边界控制
    If blk(xx, yy) = flag Then Exit Do   '发现第一个和当前同子就退出循环，继续程序后续语句
    If blk(xx, yy) = 0 Then fang = False: Exit Function   '在发现第一个同子之前发现0就结束函数，进行下一个方向的检测
    If blk(xx, yy) = -flag Then xx = xx + movea(0, sign): yy = yy + movea(1, sign)  '发现异子继续检索分析
  Loop
  '无终止条件的循环，全凭判断语句强行跳出
      If xx = x_x And yy = y_y Then fang = False: Exit Function '如果第一个子就是同子则无翻转，退出
  Do  '由远到近将异子逐一翻转
    xx = xx - movea(0, sign): yy = yy - movea(1, sign)
    blk(xx, yy) = flag
  Loop Until xx = x_x And yy = y_y
   fang = True  '顺利完成到本条时返回真值，可以下子
End Function

Private Sub Restat_Click()   '重开判断
  Form_Load    '初始化
  flag = IIf(MsgBox("白方先手？", vbYesNo + vbQuestion, "先手判断") = vbYes, 白, 黑)
  Comst_Click  '棋开之前
End Sub

Private Sub SL_Click(Index As Integer)   'SL文件控制语句
On Error Resume Next
Dim 棋谱标识 As Integer   '棋谱的文件类型是唯一的，扩展名lbw
 Select Case Index
   Case 0  '保存棋局
    fnm = InputBox("请输入保存文件名:", "黑白棋存档", "NewOne")
    If fnm = "" Then Exit Sub   '点取消的话退出
     Open fnm & ".lbw" For Random As #1 Len = 2
     If Err Then MsgBox "出现错误！", vbCritical: Close #1: Exit Sub
       Put #1, 2, -425   '与其他文件区别
       Put #1, 4, flag   '4号存储当前下子标记
        For k = 0 To 99
          Put #1, k + 5, blk(k \ 10, k Mod 10)   '5号以后存储棋子
        Next k
     Close #1
     If Not Err Then MsgBox "保存成功！", vbInformation
   Case 1  '载入棋局
     fnm = InputBox("请输入棋局文件名:", "黑白棋存档", "NewOne")
     If fnm = "" Then Exit Sub
     Open fnm & ".lbw" For Random As #1 Len = 2
       Get #1, 2, 棋谱标识   '类型判断，避免出错
       If 棋谱标识 <> -425 Then MsgBox "档案打开失败！该文件不是棋谱文件." _
          & Chr(13) & "文件将被删除.", vbCritical: Close #1: Kill fnm & ".lbw": Exit Sub
       Get #1, 4, flag      '取当前下子标记
        For k = 0 To 99
          Get #1, k + 5, blk(k \ 10, k Mod 10)
        Next k
     Close #1
     MsgBox "载入成功！", vbInformation
      opTn (flag)
      M_M_Click (Index)  '重绘
End Select
End Sub

Private Sub rules_Click()  '给不会玩的人写的，网上摘抄改编
msg$ = msg$ + "    黑白棋是从中国古代就开始起源的益智游戏，但在棋盘上" + Chr(13)
msg$ = msg$ + "下棋非常麻烦，因为要不停地将黑白两棋互换。后来有人将棋" + Chr(13)
msg$ = msg$ + "子正反面涂为异色，作为黑白棋的经典模式一直流传到现在。" + Chr(13) + Chr(13)
msg$ = msg$ + "    游戏规则非常简单：在一个棋盘上先放上四枚棋子，黑白双方轮流下棋，" + Chr(13)
msg$ = msg$ + "每一子须下在棋盘的空白处且能“吃掉”对方至少一子，否则自己不能下而" + Chr(13)
msg$ = msg$ + "由对方继续下；“吃掉”对方的棋子是指：以落点为中心，向左看去经过几" + Chr(13)
msg$ = msg$ + "个对方的棋子后又有自己的棋子(中间不能有空格)，则这几个对方的棋子就" + Chr(13)
msg$ = msg$ + "被换成我方的棋子，即被吃掉了，对方被换掉几子便是吃了几子。此种“看" + Chr(13)
msg$ = msg$ + "法”应同时应用于向右、向上、向下等八个方向。" + Chr(13)
msg$ = msg$ + "  以下面一个棋盘的局部为例，“　”为空位、“○”为白棋、“●”为黑棋：" + Chr(13)
msg$ = msg$ + "   若白棋在d4下子，棋盘变为图二，白棋吃掉黑棋4子。" + Chr(13)
msg$ = msg$ + "" + Chr(13)
msg$ = msg$ + "图一:　　　　　　 图二:" + Chr(13)
msg$ = msg$ + "  1 2 3 4 5         1 2 3 4 5 " + Chr(13)
msg$ = msg$ + "a ○　　○　　　　a ○　　○　" + Chr(13)
msg$ = msg$ + "b 　●●　　　　　b 　○●　　" + Chr(13)
msg$ = msg$ + "c 　　●●　　　　c 　　○●　" + Chr(13)
msg$ = msg$ + "d ○●●　　　　　d ○○○○　" + Chr(13)
msg$ = msg$ + "e 　　　　　　　　e 　　　　　" + Chr(13)
msg$ = msg$ + "判断胜负:" + Chr(13)
msg$ = msg$ + "    若棋盘下满了，则棋子多的一方胜利；若中途一方棋子被全部吃完，" + Chr(13)
msg$ = msg$ + "对方胜利。" + Chr(13)
MsgBox msg$, , "黑白棋规则"
End Sub

Private Sub ed1_Click()  '走人语句
  End
End Sub

