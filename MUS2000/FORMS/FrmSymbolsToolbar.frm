VERSION 5.00
Begin VB.Form fSymbolsToolbar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Float mouse over buttons for symbol key codes"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSymbolsToolbar.frx":0000
   ScaleHeight     =   1215
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   165
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   164
      Left            =   285
      Picture         =   "FrmSymbolsToolbar.frx":57D6
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   163
      Left            =   570
      Picture         =   "FrmSymbolsToolbar.frx":5928
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   162
      Left            =   855
      Picture         =   "FrmSymbolsToolbar.frx":5A7A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   161
      Left            =   1140
      Picture         =   "FrmSymbolsToolbar.frx":5BCC
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   160
      Left            =   1425
      Picture         =   "FrmSymbolsToolbar.frx":5D1E
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   159
      Left            =   1710
      Picture         =   "FrmSymbolsToolbar.frx":5E70
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   158
      Left            =   1995
      Picture         =   "FrmSymbolsToolbar.frx":5FC2
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   157
      Left            =   2280
      Picture         =   "FrmSymbolsToolbar.frx":6114
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   156
      Left            =   2565
      Picture         =   "FrmSymbolsToolbar.frx":6266
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   155
      Left            =   2850
      Picture         =   "FrmSymbolsToolbar.frx":63B8
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   154
      Left            =   3135
      Picture         =   "FrmSymbolsToolbar.frx":650A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   153
      Left            =   3435
      Picture         =   "FrmSymbolsToolbar.frx":665C
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   152
      Left            =   3705
      Picture         =   "FrmSymbolsToolbar.frx":67AE
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   151
      Left            =   3990
      Picture         =   "FrmSymbolsToolbar.frx":6900
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   150
      Left            =   4275
      Picture         =   "FrmSymbolsToolbar.frx":6A52
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   149
      Left            =   4560
      Picture         =   "FrmSymbolsToolbar.frx":6BA4
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   148
      Left            =   4845
      Picture         =   "FrmSymbolsToolbar.frx":6CF6
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   147
      Left            =   5130
      Picture         =   "FrmSymbolsToolbar.frx":6E48
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   146
      Left            =   5415
      Picture         =   "FrmSymbolsToolbar.frx":6F9A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   145
      Left            =   5700
      Picture         =   "FrmSymbolsToolbar.frx":70EC
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   144
      Left            =   5985
      Picture         =   "FrmSymbolsToolbar.frx":723E
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   143
      Left            =   6270
      Picture         =   "FrmSymbolsToolbar.frx":7390
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   142
      Left            =   6555
      Picture         =   "FrmSymbolsToolbar.frx":74E2
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   141
      Left            =   6840
      Picture         =   "FrmSymbolsToolbar.frx":7634
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   140
      Left            =   7125
      Picture         =   "FrmSymbolsToolbar.frx":7786
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   139
      Left            =   7410
      Picture         =   "FrmSymbolsToolbar.frx":78D8
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   138
      Left            =   7695
      Picture         =   "FrmSymbolsToolbar.frx":7A2A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   137
      Left            =   7980
      Picture         =   "FrmSymbolsToolbar.frx":7B7C
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   136
      Left            =   7980
      Picture         =   "FrmSymbolsToolbar.frx":7CCE
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   135
      Left            =   7695
      Picture         =   "FrmSymbolsToolbar.frx":7E20
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   134
      Left            =   7410
      Picture         =   "FrmSymbolsToolbar.frx":7F72
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   133
      Left            =   7125
      Picture         =   "FrmSymbolsToolbar.frx":80C4
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   132
      Left            =   6840
      Picture         =   "FrmSymbolsToolbar.frx":8216
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   131
      Left            =   6555
      Picture         =   "FrmSymbolsToolbar.frx":8368
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   130
      Left            =   6270
      Picture         =   "FrmSymbolsToolbar.frx":84BA
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   129
      Left            =   5985
      Picture         =   "FrmSymbolsToolbar.frx":860C
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   128
      Left            =   5700
      Picture         =   "FrmSymbolsToolbar.frx":875E
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   127
      Left            =   5415
      Picture         =   "FrmSymbolsToolbar.frx":88B0
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   126
      Left            =   5130
      Picture         =   "FrmSymbolsToolbar.frx":8A02
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   125
      Left            =   4845
      Picture         =   "FrmSymbolsToolbar.frx":8B54
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   124
      Left            =   4560
      Picture         =   "FrmSymbolsToolbar.frx":8CA6
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   123
      Left            =   4275
      Picture         =   "FrmSymbolsToolbar.frx":8DF8
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   122
      Left            =   3990
      Picture         =   "FrmSymbolsToolbar.frx":8F4A
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   121
      Left            =   3705
      Picture         =   "FrmSymbolsToolbar.frx":909C
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   120
      Left            =   3435
      Picture         =   "FrmSymbolsToolbar.frx":91EE
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   119
      Left            =   3135
      Picture         =   "FrmSymbolsToolbar.frx":9340
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   118
      Left            =   2850
      Picture         =   "FrmSymbolsToolbar.frx":9492
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   117
      Left            =   2565
      Picture         =   "FrmSymbolsToolbar.frx":95E4
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   116
      Left            =   2280
      Picture         =   "FrmSymbolsToolbar.frx":9736
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   115
      Left            =   1995
      Picture         =   "FrmSymbolsToolbar.frx":9888
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   114
      Left            =   1710
      Picture         =   "FrmSymbolsToolbar.frx":99DA
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   113
      Left            =   1425
      Picture         =   "FrmSymbolsToolbar.frx":9B2C
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   112
      Left            =   1140
      Picture         =   "FrmSymbolsToolbar.frx":9C7E
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   111
      Left            =   855
      Picture         =   "FrmSymbolsToolbar.frx":9DD0
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   110
      Left            =   570
      Picture         =   "FrmSymbolsToolbar.frx":9F22
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   109
      Left            =   285
      Picture         =   "FrmSymbolsToolbar.frx":A074
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   108
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   107
      Left            =   7125
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   1110
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   106
      Left            =   6840
      Picture         =   "FrmSymbolsToolbar.frx":A1C6
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   105
      Left            =   6555
      Picture         =   "FrmSymbolsToolbar.frx":A318
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   104
      Left            =   6270
      Picture         =   "FrmSymbolsToolbar.frx":A46A
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   103
      Left            =   5985
      Picture         =   "FrmSymbolsToolbar.frx":A5BC
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   102
      Left            =   5700
      Picture         =   "FrmSymbolsToolbar.frx":A70E
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   101
      Left            =   5415
      Picture         =   "FrmSymbolsToolbar.frx":A860
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   100
      Left            =   5130
      Picture         =   "FrmSymbolsToolbar.frx":A9B2
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   99
      Left            =   4845
      Picture         =   "FrmSymbolsToolbar.frx":AB04
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   98
      Left            =   4560
      Picture         =   "FrmSymbolsToolbar.frx":AC56
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   97
      Left            =   4275
      Picture         =   "FrmSymbolsToolbar.frx":ADA8
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   96
      Left            =   3990
      Picture         =   "FrmSymbolsToolbar.frx":AEFA
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   95
      Left            =   3705
      Picture         =   "FrmSymbolsToolbar.frx":B04C
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   94
      Left            =   3435
      Picture         =   "FrmSymbolsToolbar.frx":B19E
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   93
      Left            =   3135
      Picture         =   "FrmSymbolsToolbar.frx":B2F0
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   92
      Left            =   2850
      Picture         =   "FrmSymbolsToolbar.frx":B442
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   91
      Left            =   2565
      Picture         =   "FrmSymbolsToolbar.frx":B594
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   90
      Left            =   2280
      Picture         =   "FrmSymbolsToolbar.frx":B6E6
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   89
      Left            =   1995
      Picture         =   "FrmSymbolsToolbar.frx":B838
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   88
      Left            =   1710
      Picture         =   "FrmSymbolsToolbar.frx":B98A
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   87
      Left            =   1425
      Picture         =   "FrmSymbolsToolbar.frx":BADC
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   86
      Left            =   1140
      Picture         =   "FrmSymbolsToolbar.frx":BC2E
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   85
      Left            =   855
      Picture         =   "FrmSymbolsToolbar.frx":BD80
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   84
      Left            =   285
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   525
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   83
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2130
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   58
      Left            =   0
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   59
      Left            =   285
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   530
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   60
      Left            =   855
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   61
      Left            =   1140
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   62
      Left            =   1425
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   63
      Left            =   1710
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   64
      Left            =   1995
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   65
      Left            =   2280
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   66
      Left            =   2565
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   67
      Left            =   2850
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   68
      Left            =   3135
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   69
      Left            =   3435
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   70
      Left            =   3705
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   71
      Left            =   3990
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   72
      Left            =   4275
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   73
      Left            =   4560
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   74
      Left            =   4845
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   75
      Left            =   5130
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   76
      Left            =   5415
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   77
      Left            =   5700
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   78
      Left            =   5985
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   79
      Left            =   6270
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   80
      Left            =   6555
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   81
      Left            =   6840
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   82
      Left            =   7125
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   1110
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   57
      Left            =   0
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   56
      Left            =   285
      MousePointer    =   99  'Custom
      Top             =   405
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   55
      Left            =   570
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   54
      Left            =   855
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   53
      Left            =   1140
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   52
      Left            =   1425
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   51
      Left            =   1710
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   50
      Left            =   1995
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   49
      Left            =   2280
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   48
      Left            =   2565
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   47
      Left            =   2850
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   46
      Left            =   3135
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   45
      Left            =   3435
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   44
      Left            =   3705
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   43
      Left            =   3990
      MousePointer    =   99  'Custom
      Top             =   405
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   42
      Left            =   4275
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   41
      Left            =   4560
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   40
      Left            =   4845
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   39
      Left            =   5130
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   38
      Left            =   5415
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   37
      Left            =   5700
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   36
      Left            =   5985
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   35
      Left            =   6270
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   34
      Left            =   6555
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   33
      Left            =   6840
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   32
      Left            =   7125
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   31
      Left            =   7410
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   30
      Left            =   7695
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   29
      Left            =   7980
      MousePointer    =   99  'Custom
      Top             =   410
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   28
      Left            =   7980
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   27
      Left            =   7700
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   26
      Left            =   7410
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   25
      Left            =   7130
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   24
      Left            =   6840
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   23
      Left            =   6550
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   22
      Left            =   6270
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   21
      Left            =   5990
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   20
      Left            =   5700
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   19
      Left            =   5420
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   18
      Left            =   5130
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   17
      Left            =   4850
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   16
      Left            =   4560
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   15
      Left            =   4280
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   14
      Left            =   3995
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   13
      Left            =   3700
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   12
      Left            =   3430
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   11
      Left            =   3130
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   10
      Left            =   2855
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   9
      Left            =   2570
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   8
      Left            =   2280
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   7
      Left            =   2000
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   6
      Left            =   1710
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   5
      Left            =   1420
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   4
      Left            =   1140
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   3
      Left            =   850
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   2
      Left            =   575
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   1
      Left            =   280
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image SymbolKeyCode 
      Height          =   375
      Index           =   0
      Left            =   0
      MouseIcon       =   "FrmSymbolsToolbar.frx":BED2
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "fSymbolsToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' SET FORM ON TOP
    Call SetFormOnTop(Me)
End Sub

Private Sub SymbolKeyCode_Click(Index As Integer)
    ' DRAWS CURRENT SYMBOL KEY'S INDEX TO CURRENT MOUSE CURSOR AND
    ' WHEN CURRENT MUSIC BOARD'S HAVE BEEN CLICK IT WILL CREATE AN INDEX ARRAY
    ' TO ActiveSymbolKey ARRAY VARIABLE
    CurrentKey = "Symbols"
    Screen.MousePointer = vbCustom
    If Index = 0 Then
        Call PushKey(0, 0)
        Screen.MouseIcon = SymbolKeyCode(165)
        KeypressSymbol = "SymbolKeyCode(165)"
    ElseIf Index = 1 Then
        Call PushKey(0, 1)
        Screen.MouseIcon = SymbolKeyCode(164)
        KeypressSymbol = "SymbolKeyCode(164)"
    ElseIf Index = 2 Then
        Call PushKey(0, 2)
        Screen.MouseIcon = SymbolKeyCode(163)
        KeypressSymbol = "SymbolKeyCode(163)"
    ElseIf Index = 3 Then
        Call PushKey(0, 3)
        Screen.MouseIcon = SymbolKeyCode(162)
        KeypressSymbol = "SymbolKeyCode(162)"
    ElseIf Index = 4 Then
        Call PushKey(0, 4)
        Screen.MouseIcon = SymbolKeyCode(161)
        KeypressSymbol = "SymbolKeyCode(161)"
    ElseIf Index = 5 Then
        Call PushKey(0, 5)
        Screen.MouseIcon = SymbolKeyCode(160)
        KeypressSymbol = "SymbolKeyCode(160)"
    ElseIf Index = 6 Then
        Call PushKey(0, 6)
        Screen.MouseIcon = SymbolKeyCode(159)
        KeypressSymbol = "SymbolKeyCode(159)"
    ElseIf Index = 7 Then
        Call PushKey(0, 7)
        Screen.MouseIcon = SymbolKeyCode(158)
        KeypressSymbol = "SymbolKeyCode(158)"
    ElseIf Index = 8 Then
        Call PushKey(0, 8)
        Screen.MouseIcon = SymbolKeyCode(157)
        KeypressSymbol = "SymbolKeyCode(157)"
    ElseIf Index = 9 Then
        Call PushKey(0, 9)
        Screen.MouseIcon = SymbolKeyCode(156)
        KeypressSymbol = "SymbolKeyCode(156)"
    ElseIf Index = 10 Then
        Call PushKey(0, 10)
        Screen.MouseIcon = SymbolKeyCode(155)
        KeypressSymbol = "SymbolKeyCode(155)"
    ElseIf Index = 11 Then
        Call PushKey(0, 11)
        Screen.MouseIcon = SymbolKeyCode(154)
        KeypressSymbol = "SymbolKeyCode(154)"
    ElseIf Index = 12 Then
        Call PushKey(0, 12)
        Screen.MouseIcon = SymbolKeyCode(153)
        KeypressSymbol = "SymbolKeyCode(153)"
    ElseIf Index = 13 Then
        Call PushKey(0, 13)
        Screen.MouseIcon = SymbolKeyCode(152)
        KeypressSymbol = "SymbolKeyCode(152)"
    ElseIf Index = 14 Then
        Call PushKey(0, 14)
        Screen.MouseIcon = SymbolKeyCode(151)
        KeypressSymbol = "SymbolKeyCode(151)"
    ElseIf Index = 15 Then
        Call PushKey(0, 15)
        Screen.MouseIcon = SymbolKeyCode(150)
        KeypressSymbol = "SymbolKeyCode(150)"
    ElseIf Index = 16 Then
        Call PushKey(0, 16)
        Screen.MouseIcon = SymbolKeyCode(149)
        KeypressSymbol = "SymbolKeyCode(149)"
    ElseIf Index = 17 Then
        Call PushKey(0, 17)
        Screen.MouseIcon = SymbolKeyCode(148)
        KeypressSymbol = "SymbolKeyCode(148)"
    ElseIf Index = 18 Then
        Call PushKey(0, 18)
        Screen.MouseIcon = SymbolKeyCode(147)
        KeypressSymbol = "SymbolKeyCode(147)"
    ElseIf Index = 19 Then
        Call PushKey(0, 19)
        Screen.MouseIcon = SymbolKeyCode(146)
        KeypressSymbol = "SymbolKeyCode(146)"
    ElseIf Index = 20 Then
        Call PushKey(0, 20)
        Screen.MouseIcon = SymbolKeyCode(145)
        KeypressSymbol = "SymbolKeyCode(145)"
    ElseIf Index = 21 Then
        Call PushKey(0, 21)
        Screen.MouseIcon = SymbolKeyCode(144)
        KeypressSymbol = "SymbolKeyCode(144)"
    ElseIf Index = 22 Then
        Call PushKey(0, 22)
        Screen.MouseIcon = SymbolKeyCode(143)
        KeypressSymbol = "SymbolKeyCode(143)"
    ElseIf Index = 23 Then
        Call PushKey(0, 23)
        Screen.MouseIcon = SymbolKeyCode(142)
        KeypressSymbol = "SymbolKeyCode(142)"
    ElseIf Index = 24 Then
        Call PushKey(0, 24)
        Screen.MouseIcon = SymbolKeyCode(141)
        KeypressSymbol = "SymbolKeyCode(141)"
    ElseIf Index = 25 Then
        Call PushKey(0, 25)
        Screen.MouseIcon = SymbolKeyCode(140)
        KeypressSymbol = "SymbolKeyCode(140)"
    ElseIf Index = 26 Then
        Call PushKey(0, 26)
        Screen.MouseIcon = SymbolKeyCode(139)
        KeypressSymbol = "SymbolKeyCode(139)"
    ElseIf Index = 27 Then
        Call PushKey(0, 27)
        Screen.MouseIcon = SymbolKeyCode(138)
        KeypressSymbol = "SymbolKeyCode(138)"
    ElseIf Index = 28 Then
        Call PushKey(0, 28)
        Screen.MouseIcon = SymbolKeyCode(137)
        KeypressSymbol = "SymbolKeyCode(137)"
    ElseIf Index = 29 Then
        Call PushKey(0, 29)
        Screen.MouseIcon = SymbolKeyCode(136)
        KeypressSymbol = "SymbolKeyCode(136)"
    ElseIf Index = 30 Then
        Call PushKey(0, 30)
        Screen.MouseIcon = SymbolKeyCode(135)
        KeypressSymbol = "SymbolKeyCode(135)"
    ElseIf Index = 31 Then
        Call PushKey(0, 31)
        Screen.MouseIcon = SymbolKeyCode(134)
        KeypressSymbol = "SymbolKeyCode(134)"
    ElseIf Index = 32 Then
        Call PushKey(0, 32)
        Screen.MouseIcon = SymbolKeyCode(133)
        KeypressSymbol = "SymbolKeyCode(133)"
    ElseIf Index = 33 Then
        Call PushKey(0, 33)
        Screen.MouseIcon = SymbolKeyCode(132)
        KeypressSymbol = "SymbolKeyCode(132)"
    ElseIf Index = 34 Then
        Call PushKey(0, 34)
        Screen.MouseIcon = SymbolKeyCode(131)
        KeypressSymbol = "SymbolKeyCode(131)"
    ElseIf Index = 35 Then
        Call PushKey(0, 35)
        Screen.MouseIcon = SymbolKeyCode(130)
        KeypressSymbol = "SymbolKeyCode(130)"
    ElseIf Index = 36 Then
        Call PushKey(0, 36)
        Screen.MouseIcon = SymbolKeyCode(129)
        KeypressSymbol = "SymbolKeyCode(129)"
    ElseIf Index = 37 Then
        Call PushKey(0, 37)
        Screen.MouseIcon = SymbolKeyCode(128)
        KeypressSymbol = "SymbolKeyCode(128)"
    ElseIf Index = 38 Then
        Call PushKey(0, 38)
        Screen.MouseIcon = SymbolKeyCode(127)
        KeypressSymbol = "SymbolKeyCode(127)"
    ElseIf Index = 39 Then
        Call PushKey(0, 39)
        Screen.MouseIcon = SymbolKeyCode(126)
        KeypressSymbol = "SymbolKeyCode(126)"
    ElseIf Index = 40 Then
        Call PushKey(0, 40)
        Screen.MouseIcon = SymbolKeyCode(125)
        KeypressSymbol = "SymbolKeyCode(125)"
    ElseIf Index = 41 Then
        Call PushKey(0, 41)
        Screen.MouseIcon = SymbolKeyCode(124)
        KeypressSymbol = "SymbolKeyCode(124)"
    ElseIf Index = 42 Then
        Call PushKey(0, 42)
        Screen.MouseIcon = SymbolKeyCode(123)
        KeypressSymbol = "SymbolKeyCode(123)"
    ElseIf Index = 43 Then
        Call PushKey(0, 43)
        Screen.MouseIcon = SymbolKeyCode(122)
        KeypressSymbol = "SymbolKeyCode(122)"
    ElseIf Index = 44 Then
        Call PushKey(0, 44)
        Screen.MouseIcon = SymbolKeyCode(121)
        KeypressSymbol = "SymbolKeyCode(121)"
    ElseIf Index = 45 Then
        Call PushKey(0, 45)
        Screen.MouseIcon = SymbolKeyCode(120)
        KeypressSymbol = "SymbolKeyCode(120)"
    ElseIf Index = 46 Then
        Call PushKey(0, 46)
        Screen.MouseIcon = SymbolKeyCode(119)
        KeypressSymbol = "SymbolKeyCode(119)"
    ElseIf Index = 47 Then
        Call PushKey(0, 47)
        Screen.MouseIcon = SymbolKeyCode(118)
        KeypressSymbol = "SymbolKeyCode(118)"
    ElseIf Index = 48 Then
        Call PushKey(0, 48)
        Screen.MouseIcon = SymbolKeyCode(117)
        KeypressSymbol = "SymbolKeyCode(117)"
    ElseIf Index = 49 Then
        Call PushKey(0, 49)
        Screen.MouseIcon = SymbolKeyCode(116)
        KeypressSymbol = "SymbolKeyCode(116)"
    ElseIf Index = 50 Then
        Call PushKey(0, 50)
        Screen.MouseIcon = SymbolKeyCode(115)
        KeypressSymbol = "SymbolKeyCode(115)"
    ElseIf Index = 51 Then
        Call PushKey(0, 51)
        Screen.MouseIcon = SymbolKeyCode(114)
        KeypressSymbol = "SymbolKeyCode(114)"
    ElseIf Index = 52 Then
        Call PushKey(0, 52)
        Screen.MouseIcon = SymbolKeyCode(113)
        KeypressSymbol = "SymbolKeyCode(113)"
    ElseIf Index = 53 Then
        Call PushKey(0, 53)
        Screen.MouseIcon = SymbolKeyCode(112)
        KeypressSymbol = "SymbolKeyCode(112)"
    ElseIf Index = 54 Then
        Call PushKey(0, 54)
        Screen.MouseIcon = SymbolKeyCode(111)
        KeypressSymbol = "SymbolKeyCode(111)"
    ElseIf Index = 55 Then
        Call PushKey(0, 55)
        Screen.MouseIcon = SymbolKeyCode(110)
        KeypressSymbol = "SymbolKeyCode(110)"
    ElseIf Index = 56 Then
        Call PushKey(0, 56)
        Screen.MouseIcon = SymbolKeyCode(109)
        KeypressSymbol = "SymbolKeyCode(109)"
    ElseIf Index = 57 Then
        Call PushKey(0, 57)
        Screen.MouseIcon = SymbolKeyCode(108)
        KeypressSymbol = "SymbolKeyCode(108)"
    ElseIf Index = 58 Then
        Call PushKey(0, 58)
        Screen.MouseIcon = SymbolKeyCode(83)
        KeypressSymbol = "SymbolKeyCode(83)"
    ElseIf Index = 59 Then
        Call PushKey(0, 59)
        Screen.MouseIcon = SymbolKeyCode(84)
        KeypressSymbol = "SymbolKeyCode(84)"
    ElseIf Index = 60 Then
        Call PushKey(0, 60)
        Screen.MouseIcon = SymbolKeyCode(85)
        KeypressSymbol = "SymbolKeyCode(85)"
    ElseIf Index = 61 Then
        Call PushKey(0, 61)
        Screen.MouseIcon = SymbolKeyCode(86)
        KeypressSymbol = "SymbolKeyCode(86)"
    ElseIf Index = 62 Then
        Call PushKey(0, 62)
        Screen.MouseIcon = SymbolKeyCode(87)
        KeypressSymbol = "SymbolKeyCode(87)"
    ElseIf Index = 63 Then
        Call PushKey(0, 63)
        Screen.MouseIcon = SymbolKeyCode(88)
        KeypressSymbol = "SymbolKeyCode(88)"
    ElseIf Index = 64 Then
        Call PushKey(0, 64)
        Screen.MouseIcon = SymbolKeyCode(89)
        KeypressSymbol = "SymbolKeyCode(89)"
    ElseIf Index = 65 Then
        Call PushKey(0, 65)
        Screen.MouseIcon = SymbolKeyCode(90)
        KeypressSymbol = "SymbolKeyCode(90)"
    ElseIf Index = 66 Then
        Call PushKey(0, 66)
        Screen.MouseIcon = SymbolKeyCode(91)
        KeypressSymbol = "SymbolKeyCode(91)"
    ElseIf Index = 67 Then
        Call PushKey(0, 67)
        Screen.MouseIcon = SymbolKeyCode(92)
        KeypressSymbol = "SymbolKeyCode(92)"
    ElseIf Index = 68 Then
        Call PushKey(0, 68)
        Screen.MouseIcon = SymbolKeyCode(93)
        KeypressSymbol = "SymbolKeyCode(93)"
    ElseIf Index = 69 Then
        Call PushKey(0, 69)
        Screen.MouseIcon = SymbolKeyCode(94)
        KeypressSymbol = "SymbolKeyCode(94)"
    ElseIf Index = 70 Then
        Call PushKey(0, 70)
        Screen.MouseIcon = SymbolKeyCode(95)
        KeypressSymbol = "SymbolKeyCode(95)"
    ElseIf Index = 71 Then
        Call PushKey(0, 71)
        Screen.MouseIcon = SymbolKeyCode(96)
        KeypressSymbol = "SymbolKeyCode(96)"
    ElseIf Index = 72 Then
        Call PushKey(0, 72)
        Screen.MouseIcon = SymbolKeyCode(97)
        KeypressSymbol = "SymbolKeyCode(97)"
    ElseIf Index = 73 Then
        Call PushKey(0, 73)
        Screen.MouseIcon = SymbolKeyCode(98)
        KeypressSymbol = "SymbolKeyCode(98)"
    ElseIf Index = 74 Then
        Call PushKey(0, 74)
        Screen.MouseIcon = SymbolKeyCode(99)
        KeypressSymbol = "SymbolKeyCode(99)"
    ElseIf Index = 75 Then
        Call PushKey(0, 75)
        Screen.MouseIcon = SymbolKeyCode(100)
        KeypressSymbol = "SymbolKeyCode(100)"
    ElseIf Index = 76 Then
        Call PushKey(0, 76)
        Screen.MouseIcon = SymbolKeyCode(101)
        KeypressSymbol = "SymbolKeyCode(101)"
    ElseIf Index = 77 Then
        Call PushKey(0, 77)
        Screen.MouseIcon = SymbolKeyCode(102)
        KeypressSymbol = "SymbolKeyCode(102)"
    ElseIf Index = 78 Then
        Call PushKey(0, 78)
        Screen.MouseIcon = SymbolKeyCode(103)
        KeypressSymbol = "SymbolKeyCode(103)"
    ElseIf Index = 79 Then
        Call PushKey(0, 79)
        Screen.MouseIcon = SymbolKeyCode(104)
        KeypressSymbol = "SymbolKeyCode(104)"
    ElseIf Index = 80 Then
        Call PushKey(0, 80)
        Screen.MouseIcon = SymbolKeyCode(105)
        KeypressSymbol = "SymbolKeyCode(105)"
    ElseIf Index = 81 Then
        Call PushKey(0, 81)
        Screen.MouseIcon = SymbolKeyCode(106)
        KeypressSymbol = "SymbolKeyCode(106)"
    ElseIf Index = 82 Then
        Call PushKey(0, 82)
        Screen.MouseIcon = SymbolKeyCode(107)
        KeypressSymbol = "SymbolKeyCode(107)"
    End If
End Sub
