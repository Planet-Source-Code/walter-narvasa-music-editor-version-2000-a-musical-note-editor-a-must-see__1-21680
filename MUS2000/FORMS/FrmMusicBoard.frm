VERSION 5.00
Begin VB.Form fMusicBoard 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Untitled:"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   Icon            =   "FrmMusicBoard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Begin VB.TextBox ActiveTextLine 
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   15255
   End
   Begin VB.PictureBox StartProgress 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   11115
      Left            =   0
      ScaleHeight     =   11115
      ScaleWidth      =   15210
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15210
      Begin VB.Label ProgressCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   10515
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   11100
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image BassTabLine 
      Height          =   9450
      Index           =   5
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":0442
      Top             =   5040
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BassTabLine 
      Height          =   7875
      Index           =   4
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":9C2FC
      Top             =   5040
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BassTabLine 
      Height          =   6300
      Index           =   3
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":11E2A2
      Top             =   5040
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BassTabLine 
      Height          =   4725
      Index           =   2
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":186334
      Top             =   5040
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BassTabLine 
      Height          =   3150
      Index           =   1
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":1D44B2
      Top             =   4800
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BassTabLine 
      Height          =   1575
      Index           =   0
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":20871C
      Top             =   4560
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image TabLine 
      Height          =   9450
      Index           =   5
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":222A72
      Top             =   4560
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image TabLine 
      Height          =   7875
      Index           =   4
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":2BE92C
      Top             =   4440
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image TabLine 
      Height          =   6300
      Index           =   3
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":3408D2
      Top             =   4320
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image TabLine 
      Height          =   4725
      Index           =   2
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":3A8964
      Top             =   3960
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image TabLine 
      Height          =   3150
      Index           =   1
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":3F6AE2
      Top             =   3840
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image TabLine 
      Height          =   1575
      Index           =   0
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":42AD4C
      Top             =   3600
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image KeyboardTrebleBassLine 
      Height          =   2520
      Left            =   5760
      Picture         =   "FrmMusicBoard.frx":4450A2
      Top             =   1800
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image ActiveAlternateKey 
      Height          =   45
      Index           =   0
      Left            =   0
      ToolTipText     =   "Double Click Symbol to Cut.."
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Line PlainLine 
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   15070
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image PercLine 
      Height          =   1575
      Left            =   5040
      Picture         =   "FrmMusicBoard.frx":4471C4
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image DrumLine 
      Height          =   1575
      Left            =   4320
      Picture         =   "FrmMusicBoard.frx":44866E
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image EightvbTrebleLine 
      Height          =   1575
      Left            =   3600
      Picture         =   "FrmMusicBoard.frx":449B18
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image EightvaTrebleLine 
      Height          =   1575
      Left            =   2880
      Picture         =   "FrmMusicBoard.frx":44AFC2
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image BarLine 
      Height          =   9510
      Index           =   5
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":44C46C
      Top             =   3600
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BarLine 
      Height          =   7935
      Index           =   4
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":4E92F6
      Top             =   3600
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BarLine 
      Height          =   6360
      Index           =   3
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":56C26C
      Top             =   3600
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BarLine 
      Height          =   4785
      Index           =   2
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":5D52CE
      Top             =   3720
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BarLine 
      Height          =   3165
      Index           =   1
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":62441C
      Top             =   3720
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image BarLine 
      Height          =   1575
      Index           =   0
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":658A7A
      Top             =   3720
      Visible         =   0   'False
      Width           =   15180
   End
   Begin VB.Image TenorLine 
      Height          =   1575
      Left            =   2160
      Picture         =   "FrmMusicBoard.frx":672DD0
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image AltoLine 
      Height          =   1575
      Left            =   1440
      Picture         =   "FrmMusicBoard.frx":67427A
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image FrontEdgeLine 
      Height          =   1575
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image BassLine 
      Height          =   1575
      Left            =   720
      Picture         =   "FrmMusicBoard.frx":675724
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image TrebleLine 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmMusicBoard.frx":676BCE
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image ActiveSymbolKey 
      Height          =   45
      Index           =   0
      Left            =   0
      ToolTipText     =   "Double Click Symbol to Cut.."
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "fMusicBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nTop, nLeft, MaxIdx

Private Sub Form_Load()
    MaxIdx = 0
End Sub

Private Sub Form_Click()
    ' DRAW SYMBOL KEYS TO CURRENT MUSIC BOARD'S INDEX
    On Error Resume Next
    If MaxIdx = 0 Then MaxIdx = 1
        MaxIdx = MaxIdx + 1
    If CurrentKey = "Symbols" Then
        Load ActiveSymbolKey(MaxIdx)
        ActiveSymbolKey(MaxIdx).Visible = True
        ActiveSymbolKey(MaxIdx).Top = nTop
        ActiveSymbolKey(MaxIdx).Left = nLeft
        ActiveSymbolKey(MaxIdx).Picture = Screen.MouseIcon
        ActiveSymbolKeyVal_(MaxIdx) = KeypressSymbol
        ActiveSymbolKeyTopCoordinates_(MaxIdx) = nTop
        ActiveSymbolKeyLeftCoordinates_(MaxIdx) = nLeft
    Else
        Load ActiveAlternateKey(MaxIdx)
        ActiveAlternateKey(MaxIdx).Visible = True
        ActiveAlternateKey(MaxIdx).Top = nTop
        ActiveAlternateKey(MaxIdx).Left = nLeft
        ActiveAlternateKey(MaxIdx).Picture = Screen.MouseIcon
        ActiveAlternateKeyVal_(MaxIdx) = KeypressAlternate
        ActiveAlternateKeyTopCoordinates_(MaxIdx) = nTop
        ActiveAlternateKeyLeftCoordinates_(MaxIdx) = nLeft
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' DISPLAY MOUSE COORDINATES
    nTop = Y - 350
    nLeft = x - 15
    fMain.StatusBar.Panels(1).Text = "X-Coordinates = " & x & "    /   " & "Y-Coordinates = " & Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Save changes to " & Me.Caption, vbYesNo + vbDefaultButton2 + vbCritical, "Music Editor 2000-Information"
End Sub

Private Sub ActiveTextLine_Change(Index As Integer)
    ' PLACE CURRENT ActiveTextLine INDEX VALUES TO ActiveTextLineVal_ ARRAY VARIABLE
    ActiveTextLineVal_(Index) = MusicBoard(fIndex).ActiveTextLine(Index).Text
End Sub

Private Sub ActiveAlternateKey_Click(Index As Integer)
    ' CUT CURRENT ALTERNATE KEY'S INDEX
    ActiveAlternateKey(Index).Picture = LoadPicture()
    ActiveAlternateKeyVal_(Index) = ""
    ActiveAlternateKeyTopCoordinates_(Index) = 0
    ActiveAlternateKeyLeftCoordinates_(Index) = 0
End Sub

Private Sub ActiveSymbolKey_DblClick(Index As Integer)
    ' CUT CURRENT SYMBOL KEY'S INDEX
    ActiveSymbolKey(Index).Picture = LoadPicture()
    ActiveSymbolKeyVal_(Index) = ""
    ActiveSymbolKeyTopCoordinates_(Index) = 0
    ActiveSymbolKeyLeftCoordinates_(Index) = 0
End Sub

