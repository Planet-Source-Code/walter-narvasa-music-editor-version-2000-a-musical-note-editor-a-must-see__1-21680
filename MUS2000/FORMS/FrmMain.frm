VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm fMain 
   BackColor       =   &H8000000C&
   Caption         =   "Music Editor Version 2000 by Walter A. Narvasa"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox OpenSaveContainer 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   6135
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   6195
      Begin VB.TextBox txtSection 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   4680
         TabIndex        =   4
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtString 
         Height          =   285
         Left            =   9360
         TabIndex        =   3
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Section"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Key"
         Height          =   195
         Left            =   2400
         TabIndex        =   9
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "File"
         Height          =   195
         Left            =   4680
         TabIndex        =   8
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "String"
         Height          =   195
         Left            =   9360
         TabIndex        =   7
         Top             =   120
         Width           =   435
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1CFA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":214E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":25A2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":26FE
            Key             =   "TextLine"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":33DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":40B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":674A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar MainToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Music Editor File.."
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Music Editor File.."
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Music Editor File.."
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TextLine"
            Object.ToolTipText     =   "Inser Text Line.."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TrebleLine"
            Object.ToolTipText     =   "Insert Treble Line.."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BassLine"
            Object.ToolTipText     =   "Insert Bass Line.."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KeyboardTrebleBassLine"
            Object.ToolTipText     =   "Insert Keyboard (Treble+Bass) Line.."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TabLine"
            Object.ToolTipText     =   "Insert Tab Line.."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BassTabLine"
            Object.ToolTipText     =   "Insert Bass Tab Line.."
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3675
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "Coordinates"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "3/13/01"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12:57 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CMDialog 
      Left            =   120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New.."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open.."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save.."
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As.."
      End
      Begin VB.Menu mnuFBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print.."
      End
      Begin VB.Menu mnuFBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMusicFiles 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit.."
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuTextLine 
         Caption         =   "&Text Line"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuTreble 
         Caption         =   "T&reble.."
      End
      Begin VB.Menu mnuBass 
         Caption         =   "&Bass.."
      End
      Begin VB.Menu mnuIBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKeyboardStaff 
         Caption         =   "&Keyboard Staff.."
         Begin VB.Menu mnuAlto 
            Caption         =   "&Alto.."
         End
         Begin VB.Menu mnuTenor 
            Caption         =   "&Tenor.."
         End
         Begin VB.Menu mnu8vaTreble 
            Caption         =   "8v&a Treble.."
         End
         Begin VB.Menu mnu8vbTreble 
            Caption         =   "8v&bTreble.."
         End
         Begin VB.Menu mnuDrums 
            Caption         =   "&Drums.."
         End
         Begin VB.Menu mnuPercLine 
            Caption         =   "&Perc Line.."
         End
         Begin VB.Menu mnuNoClef 
            Caption         =   "&No Clef.."
         End
      End
      Begin VB.Menu mnuIBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKeyboardTrebleBassLine 
         Caption         =   "Keyboard(Treble+Bass) Staff"
      End
      Begin VB.Menu mnuTabLine 
         Caption         =   "T&ab Line.."
      End
      Begin VB.Menu mnuBassTabLine 
         Caption         =   "&Bass Tab Line.."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuMainToolbar 
         Caption         =   "&Main Toolbar.."
      End
      Begin VB.Menu mnuSymbolsToolbar 
         Caption         =   "&Symbols Toolbar.."
      End
      Begin VB.Menu mnuAlternateToolbar 
         Caption         =   "&Alternate Toolbar.."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuUseArrowCursor 
         Caption         =   "&Use Arrow Cursor.."
      End
      Begin VB.Menu mnuKeyboardSetCursorSymbol 
         Caption         =   "&Keyboard Set Cursor Symbol.."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuTile 
         Caption         =   "&Tile.."
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade.."
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "&Arrange Icons.."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics.."
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTop, MaxIdx, nTop2, MaxIdx2, nCount, xTop
Dim OmitLines As Boolean, OmitNext As Boolean

Private Sub MDIForm_Load()
    ' STARTUP SETTINGS
    Call ResetSettings(True)
    CurrentMusicBoardIndex = fIndex
    MainToolbar.Visible = True
    mnuMainToolbar.Checked = True
End Sub

Private Sub MainToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    ' MAIN TOOLBAR OPTIONS
    If Button.Key = "New" Then
        mnuNew_Click
    ElseIf Button.Key = "Open" Then
        mnuOpen_Click
    ElseIf Button.Key = "Save" Then
        mnuSaveAs_Click
    ElseIf Button.Key = "TextLine" Then
        mnuTextLine_Click
    ElseIf Button.Key = "TrebleLine" Then
        mnuTreble_Click
    ElseIf Button.Key = "BassLine" Then
        mnuBass_Click
    ElseIf Button.Key = "KeyboardTrebleBassLine" Then
        mnuKeyboardTrebleBassLine_Click
    ElseIf Button.Key = "TabLine" Then
        mnuTabLine_Click
    ElseIf Button.Key = "BassTabLine" Then
        mnuBassTabLine_Click
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ' END ALL LOAD FORMS
    End
End Sub

Private Sub mnuHelpTopics_Click()
    ' HELP TOPICS/ABOUT THE SYSTEM
    MsgBox "Music Editor" & vbCrLf & _
            "Version 2000" & vbCrLf & _
            "Copyright (c) 2000" & vbCrLf & _
            "All rights reserved.", vbOKOnly + vbInformation, "Music Editor 2000-Information"
End Sub

Private Sub mnuMainToolbar_Click()
    ' SHOW/HIDE MAIN TOOLBAR
    If mnuMainToolbar.Checked = False Then
        mnuMainToolbar.Checked = True
        MainToolbar.Visible = True
    Else
        mnuMainToolbar.Checked = False
        MainToolbar.Visible = False
    End If
End Sub

Private Sub mnuSymbolsToolbar_Click()
    ' LOAD fSymbolsToolbar FORM
    If mnuSymbolsToolbar.Checked = False Then
        fSymbolsToolbar.Show
        fSymbolsToolbar.SetFocus
        mnuSymbolsToolbar.Checked = True
    Else
        Unload fSymbolsToolbar
        mnuSymbolsToolbar.Checked = False
    End If
End Sub

Private Sub mnuAlternateToolbar_Click()
    ' LOAD fAltermnateToolbar FORM
    If mnuAlternateToolbar.Checked = False Then
        fAlternateToolbar.Show
        fAlternateToolbar.SetFocus
        mnuAlternateToolbar.Checked = True
    Else
        Unload fAlternateToolbar
        mnuAlternateToolbar.Checked = False
    End If
End Sub

Private Sub mnuNew_Click()
    ' CREATE NEW MUSIC BOARD FILE BY USING ARRAY COUNT
    fIndex = fIndex + 1
    MusicBoard(fIndex).Tag = fIndex
    MusicBoard(fIndex).Caption = "Untitled:" & fIndex
    MusicBoard(fIndex).Show
    MusicBoard(fIndex).Width = Me.ScaleWidth
    MusicBoard(fIndex).Height = Me.ScaleHeight
    mnuFBar3.Visible = True
    Load mnuMusicFiles(fIndex)
    mnuMusicFiles(fIndex).Visible = True
    mnuMusicFiles(fIndex).Caption = "Untitled:" & fIndex
    Call ResetSettings(False)
    CurrentMusicBoardIndex = fIndex
End Sub

Private Sub mnuOpen_Click()
    ' OPEN MUSIC EDITOR FILES "*.MUS" FORMAT BY GATHERING ALL ARRAY VARIABLES
    ' INCLUDING THE SYMBOLS, COORDINATES, LINES AND RESTORE IT TO NEW MUSIC EDITOR 2000 FILE
    ' INCLUDING ITS FILENAME
    'On Error Resume Next
    Dim Directory As String, TmpIdx
    Dim i As Integer, xVal As Integer, xWithPercLine As Boolean
    Dim xKeypressSymbol, xKeypressAlternate, KeypressValue
Start:
    CMDialog.filename = ""
    CMDialog.DialogTitle = "Open"
    CMDialog.InitDir = App.Path
    CMDialog.Filter = "Music Editor 2000 Key Files|*.MUS"
    CMDialog.ShowOpen
    If CMDialog.filename <> "" Then
        Directory$ = CMDialog.filename
    Else
        Exit Sub
    End If
    TmpIdx = fIndex + 1
    Load mnuMusicFiles(TmpIdx)
    mnuFBar3.Visible = True
    mnuMusicFiles(TmpIdx).Visible = True
    mnuMusicFiles(TmpIdx).Caption = Directory$
    Call ResetSettings(False)
    fIndex = fIndex + 1
    MusicBoard(fIndex).Tag = fIndex
    MusicBoard(fIndex).Caption = Directory$
    MusicBoard(fIndex).Show
    MusicBoard(fIndex).Width = Me.ScaleWidth
    MusicBoard(fIndex).Height = Me.ScaleHeight
    MusicBoard(fIndex).StartProgress.Visible = True
    MusicBoard(fIndex).ProgressCaption.Caption = MusicBoard(fIndex).ProgressCaption.Caption + Directory$
    txtFile = Directory$
    txtSection = "ActiveSymbolKeyValues"
    For xVal = 2 To 1000
        txtKey = ""
        txtKey = "ActiveSymbolKeyVal_(" & xVal & ")"
        txtString = OpenMusicFile(txtSection.Text, txtKey.Text, txtFile.Text)
        Load MusicBoard(fIndex).ActiveSymbolKey(xVal)
        MusicBoard(fIndex).ActiveSymbolKey(xVal).Visible = True
        xKeypressSymbol = ExtractArgument(2, txtString, "(")
        If xKeypressSymbol <> "" Then
            KeypressValue = Mid(xKeypressSymbol, 1, Len(xKeypressSymbol) - 1)
        Else
            MusicBoard(fIndex).ActiveSymbolKey(xVal).Visible = False
        End If
        MusicBoard(fIndex).ActiveSymbolKey(xVal).Picture = fSymbolsToolbar.SymbolKeyCode(KeypressValue).Picture
        txtKey = ""
        txtString = ""
        txtKey = "ActiveSymbolKeyTopCoordinates_(" & xVal & ")"
        txtString = OpenMusicFile(txtSection.Text, txtKey.Text, txtFile.Text)
        MusicBoard(fIndex).ActiveSymbolKey(xVal).Top = Val(txtString)
        txtKey = ""
        txtString = ""
        txtKey = "ActiveSymbolKeyLeftCoordinates_(" & xVal & ")"
        txtString = OpenMusicFile(txtSection.Text, txtKey.Text, txtFile.Text)
        MusicBoard(fIndex).ActiveSymbolKey(xVal).Left = Val(txtString)
        xVal = ((xVal + 1) - 1)
    Next xVal
    txtSection = "ActiveAlternateKeyValues"
    For xVal = 2 To 1000
        txtKey = ""
        txtKey = "ActiveAlternateKeyVal_(" & xVal & ")"
        txtString = OpenMusicFile(txtSection.Text, txtKey.Text, txtFile.Text)
        Load MusicBoard(fIndex).ActiveAlternateKey(xVal)
        MusicBoard(fIndex).ActiveAlternateKey(xVal).Visible = True
        xKeypressAlternate = ExtractArgument(2, txtString, "(")
        If xKeypressAlternate <> "" Then
            KeypressValue = Mid(xKeypressAlternate, 1, Len(xKeypressAlternate) - 1)
        Else
            MusicBoard(fIndex).ActiveAlternateKey(xVal).Visible = False
        End If
        MusicBoard(fIndex).ActiveAlternateKey(xVal).Picture = fAlternateToolbar.AlternateKeyCode(KeypressValue).Picture
        txtKey = ""
        txtString = ""
        txtKey = "ActiveAlternateKeyTopCoordinates_(" & xVal & ")"
        txtString = OpenMusicFile(txtSection.Text, txtKey.Text, txtFile.Text)
        MusicBoard(fIndex).ActiveAlternateKey(xVal).Top = Val(txtString)
        txtKey = ""
        txtString = ""
        txtKey = "ActiveAlternateKeyLeftCoordinates_(" & xVal & ")"
        txtString = OpenMusicFile(txtSection.Text, txtKey.Text, txtFile.Text)
        MusicBoard(fIndex).ActiveAlternateKey(xVal).Left = Val(txtString)
        xVal = ((xVal + 1) - 1)
    Next xVal
    txtSection = "InsertedLinesType"
    For xVal = 0 To 5
        txtKey = ""
        txtString = ""
        txtKey = "InsertedLinesType(" & xVal & ")"
        txtString = OpenMusicFile(txtSection.Text, txtKey.Text, txtFile.Text)
        If xVal <> 0 Then
            Load MusicBoard(fIndex).FrontEdgeLine(xVal)
        End If
        MusicBoard(fIndex).FrontEdgeLine(xVal).Visible = True
        If xVal = 0 Then
            nTop = 0
        Else
            If Trim(txtString) = "Keyboard (Treble+Bass) Line" Then
                nTop = nTop + 3180
            Else
                nTop = nTop + 1590
            End If
        End If
        If xVal = 4 Then
            If Trim(txtString) = "Keyboard (Treble+Bass) Line" Then
                nTop = 4340
            Else
                nTop = 6340
            End If
        ElseIf xVal = 5 Then
            nTop = 7920
        End If
        If Trim(txtString) = "Treble Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).TrebleLine.Picture
            nCount = nCount + 1
        ElseIf Trim(txtString) = "Bass Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).BassLine.Picture
            nCount = nCount + 1
        ElseIf Trim(txtString) = "Alto Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).AltoLine.Picture
            nCount = nCount + 1
        ElseIf Trim(txtString) = "Tenor Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).TenorLine.Picture
            nCount = nCount + 1
        ElseIf Trim(txtString) = "8va Treble Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).EightvaTrebleLine.Picture
            nCount = nCount + 1
        ElseIf Trim(txtString) = "8vb Treble Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).EightvbTrebleLine.Picture
            nCount = nCount + 1
        ElseIf Trim(txtString) = "Drum Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).DrumLine.Picture
            nCount = nCount + 1
        ElseIf Trim(txtString) = "Perc Line" Then
            Load MusicBoard(fIndex).PlainLine(xVal)
            MusicBoard(fIndex).PlainLine(xVal).Visible = True
            MusicBoard(fIndex).PlainLine(xVal).Y1 = nTop + 740
            MusicBoard(fIndex).PlainLine(xVal).Y2 = nTop + 740
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).PercLine.Picture
            xWithPercLine = True
            nCount = nCount + 1
        ElseIf Trim(txtString) = "No Clef" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Visible = False
            nCount = nCount + 1
        ElseIf Trim(txtString) = "Keyboard (Treble+Bass) Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Top = nTop
            MusicBoard(fIndex).FrontEdgeLine(xVal).Picture = MusicBoard(fIndex).KeyboardTrebleBassLine.Picture
            nCount = nCount + 2
        ElseIf Trim(txtString) = "Tab Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Visible = False
            nCount = nCount + 1
        ElseIf Trim(txtString) = "Bass Tab Line" Then
            MusicBoard(fIndex).FrontEdgeLine(xVal).Visible = False
            nCount = nCount + 1
        End If
        If Trim(txtString) <> "" Then
            If nCount = 1 Then
                If Trim(txtString) = "Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(0).Picture
                ElseIf Trim(txtString) = "Bass Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(0).Picture
                Else
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(0).Picture
                End If
                If xWithPercLine = True Then
                    MusicBoard(fIndex).Picture = LoadPicture("")
                End If
            ElseIf nCount = 2 Then
                If Trim(txtString) = "Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(1).Picture
                ElseIf Trim(txtString) = "Bass Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(1).Picture
                Else
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(1).Picture
                End If
                If xWithPercLine = True Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(0).Picture
                End If
            ElseIf nCount = 3 Then
                If Trim(txtString) = "Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(2).Picture
                ElseIf Trim(txtString) = "Bass Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(2).Picture
                Else
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(2).Picture
                End If
                If xWithPercLine = True Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(1).Picture
                End If
            ElseIf nCount = 4 Then
                If Trim(txtString) = "Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(3).Picture
                ElseIf Trim(txtString) = "Bass Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(3).Picture
                Else
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(3).Picture
                End If
                If xWithPercLine = True Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(2).Picture
                End If
            ElseIf nCount = 5 Then
                If Trim(txtString) = "Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(4).Picture
                ElseIf Trim(txtString) = "Bass Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(4).Picture
                Else
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(4).Picture
                End If
                If xWithPercLine = True Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(3).Picture
                End If
            ElseIf nCount = 6 Then
                If Trim(txtString) = "Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(5).Picture
                ElseIf Trim(txtString) = "Bass Tab Line" Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(5).Picture
                Else
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(5).Picture
                End If
                If xWithPercLine = True Then
                    MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(4).Picture
                End If
            End If
        End If
        xVal = ((xVal + 1) - 1)
    Next xVal
    txtSection = "ActiveTextLines"
    For xVal = 0 To 30
        txtKey = ""
        txtKey = "ActiveTextLineVal_(" & xVal & ")"
        txtString = OpenMusicFile(txtSection.Text, txtKey.Text, txtFile.Text)
        If txtString <> "" Then
            If xVal <> 0 Then
                Load MusicBoard(fIndex).ActiveTextLine(xVal)
            End If
            If xVal = 0 Then
                xTop = 0
            Else
                xTop = ((xTop + 310) - 1)
            End If
            MusicBoard(fIndex).ActiveTextLine(xVal).Visible = True
            MusicBoard(fIndex).ActiveTextLine(xVal).Top = xTop
            MusicBoard(fIndex).ActiveTextLine(xVal).Text = Trim(txtString.Text)
        End If
        xVal = ((xVal + 1) - 1)
    Next xVal
    MusicBoard(fIndex).StartProgress.Visible = False
End Sub

Private Sub mnuSave_Click()
    CMDialog.filename = Trim(mnuMusicFiles(fIndex).Caption)
    mnuSaveAs_Click
End Sub

Private Sub mnuSaveAs_Click()
    ' SAVE MUSIC EDITOR FILES INTO "*.MUS" FORMAT BY GATHERING ALL ARRAY VARIABLES
    ' INCLUDING THE SYMBOLS, COORDINATES, LINES AND SAVE IT TO CURRENT MUSIC EDITOR 2000 FILE
    On Error Resume Next
    Dim Directory As String
    Dim i As Integer, xVal As Integer
Start:
    CMDialog.filename = ""
    CMDialog.DialogTitle = "Save"
    CMDialog.InitDir = App.Path
    CMDialog.Filter = "Music Editor 2000 Key Files|*.MUS"
    CMDialog.ShowSave
    If CMDialog.filename <> "" Then
        Directory$ = CMDialog.filename
    Else
        Exit Sub
    End If
    If FileExists(Directory$) = True Then
        If MsgBox("Are you sure you want to" & vbCrLf & _
                "overwrite the previous file?", _
                48 + vbYesNo, "Music Editor 2000-Message") = vbYes Then
            GoTo SaveFile
        Else
            GoTo Start
        End If
    End If
SaveFile:
    txtFile = ""
    txtSection = ""
    txtFile = Directory$
    txtSection = "ActiveSymbolKeyValues"
    For xVal = 0 To 1000
        If ActiveSymbolKeyVal_(xVal) <> "" Then
            txtString = ""
            txtKey = ""
            txtString = ActiveSymbolKeyVal_(xVal)
            txtKey = "ActiveSymbolKeyVal_(" & xVal & ")"
        End If
        Call SaveMusicFile(txtSection.Text, txtKey.Text, txtString.Text, txtFile.Text)
        If ActiveSymbolKeyTopCoordinates_(xVal) <> 0 Then
            txtString = ""
            txtKey = ""
            txtString = ActiveSymbolKeyTopCoordinates_(xVal)
            txtKey = "ActiveSymbolKeyTopCoordinates_(" & xVal & ")"
        End If
        Call SaveMusicFile(txtSection.Text, txtKey.Text, txtString.Text, txtFile.Text)
        If ActiveSymbolKeyLeftCoordinates_(xVal) <> 0 Then
            txtString = ""
            txtKey = ""
            txtString = ActiveSymbolKeyLeftCoordinates_(xVal)
            txtKey = "ActiveSymbolKeyLeftCoordinates_(" & xVal & ")"
        End If
        Call SaveMusicFile(txtSection.Text, txtKey.Text, txtString.Text, txtFile.Text)
        xVal = ((xVal + 1) - 1)
    Next xVal
    txtSection = ""
    txtSection = "ActiveAlternateKeyValues"
    For xVal = 0 To 1000
        If ActiveAlternateKeyVal_(xVal) <> "" Then
            txtString = ""
            txtKey = ""
            txtString = ActiveAlternateKeyVal_(xVal)
            txtKey = "ActiveAlternateKeyVal_(" & xVal & ")"
        End If
        Call SaveMusicFile(txtSection.Text, txtKey.Text, txtString.Text, txtFile.Text)
        If ActiveAlternateKeyTopCoordinates_(xVal) <> 0 Then
            txtString = ""
            txtKey = ""
            txtString = ActiveAlternateKeyTopCoordinates_(xVal)
            txtKey = "ActiveAlternateKeyTopCoordinates_(" & xVal & ")"
        End If
        Call SaveMusicFile(txtSection.Text, txtKey.Text, txtString.Text, txtFile.Text)
        If ActiveAlternateKeyLeftCoordinates_(xVal) <> 0 Then
            txtString = ""
            txtKey = ""
            txtString = ActiveAlternateKeyLeftCoordinates_(xVal)
            txtKey = "ActiveAlternateKeyLeftCoordinates_(" & xVal & ")"
        End If
        Call SaveMusicFile(txtSection.Text, txtKey.Text, txtString.Text, txtFile.Text)
        xVal = ((xVal + 1) - 1)
    Next xVal
    txtSection = ""
    txtSection = "InsertedLinesType"
    For xVal = 0 To 5
        If InsertedLinesType(xVal) <> "" Then
            txtString = ""
            txtKey = ""
            txtString = InsertedLinesType(xVal)
            txtKey = "InsertedLinesType(" & xVal & ")"
        End If
        Call SaveMusicFile(txtSection.Text, txtKey.Text, txtString.Text, txtFile.Text)
        xVal = ((xVal + 1) - 1)
    Next xVal
    txtSection = ""
    txtSection = "ActiveTextLines"
    For xVal = 0 To 30
        If ActiveTextLineVal_(xVal) <> "" Then
            txtString = ""
            txtKey = ""
            txtString = ActiveTextLineVal_(xVal)
            txtKey = "ActiveTextLineVal_(" & xVal & ")"
        End If
        Call SaveMusicFile(txtSection.Text, txtKey.Text, txtString.Text, txtFile.Text)
        xVal = ((xVal + 1) - 1)
    Next xVal
    MusicBoard(fIndex).Caption = Directory$
    mnuMusicFiles(fIndex).Caption = Directory$
    For xVal = 0 To 1000
        If ActiveSymbolKeyVal_(xVal) <> "" Then
            ActiveSymbolKeyVal_(xVal) = ""
        End If
        If ActiveSymbolKeyTopCoordinates_(xVal) <> 0 Then
            ActiveSymbolKeyTopCoordinates_(xVal) = 0
        End If
        If ActiveSymbolKeyLeftCoordinates_(xVal) <> 0 Then
            ActiveSymbolKeyLeftCoordinates_(xVal) = 0
        End If
        xVal = ((xVal + 1) - 1)
    Next xVal
    For xVal = 0 To 5
        If InsertedLinesType(xVal) <> "" Then
            InsertedLinesType(xVal) = ""
        End If
    Next xVal
    Call ResetSettings(False)
End Sub

Private Sub mnuPrint_Click()
    ' PRINT CURRENT MUSIC BOARD INDEX
    If MsgBox("Are you sure you want to" & vbCrLf & _
                "print " & MusicBoard(CurrentMusicBoardIndex).Caption & "?", _
                48 + vbYesNo, "Music Editor 2000-Message") = vbYes Then
        MusicBoard(CurrentMusicBoardIndex).PrintForm
    End If
End Sub

Private Sub mnuTile_Click()
    ' ARRANGE ALL CHILD MUSIC BOARD'S IN TILE MODE
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuCascade_Click()
    ' ARRANGE ALL CHILD MUSIC BOARD'S IN CASCADE MODE
    Me.Arrange vbCascade
End Sub

Private Sub mnuArrangeIcons_Click()
    ' ARRANGE ALL CHILD MUSIC BOARD'S ICONS
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuExit_Click()
    ' END ALL LOAD FORMS
    End
End Sub

Private Sub mnuMusicFiles_Click(Index As Integer)
    ' SETFOCUS CURRENT INDEX CHILD MUSIC BOARD
    MusicBoard(Index).SetFocus
    CurrentMusicBoardIndex = Index
End Sub

Private Sub mnuTextLine_Click()
    ' INSERT TEXT LINE TO CURRENT MUSIC BOARD'S INDEX
    On Error Resume Next
    If MaxIdx2 = 0 Then
        MusicBoard(fIndex).ActiveTextLine(MaxIdx2).Visible = True
        MusicBoard(fIndex).ActiveTextLine(MaxIdx2).Top = nTop2
        MusicBoard(fIndex).ActiveTextLine(MaxIdx2).SetFocus
        ActiveTextLineVal_(MaxIdx2) = MusicBoard(fIndex).ActiveTextLine(MaxIdx2).Text
        MaxIdx2 = 1
    ElseIf MaxIdx2 >= 30 Then
        MsgBox "Sorry..You already exceeded the number of rows.", vbOKOnly + vbInformation, "Music Editor 2000-Information"
        Exit Sub
    Else
        Load MusicBoard(fIndex).ActiveTextLine(MaxIdx2)
        MusicBoard(fIndex).ActiveTextLine(MaxIdx2).Visible = True
        MusicBoard(fIndex).ActiveTextLine(MaxIdx2).Top = nTop2
        MusicBoard(fIndex).ActiveTextLine(MaxIdx2).SetFocus
        ActiveTextLineVal_(MaxIdx2) = MusicBoard(fIndex).ActiveTextLine(MaxIdx2).Text
    End If
    nTop2 = nTop2 + 310
    MaxIdx2 = MaxIdx2 + 1
End Sub

Private Sub mnuTreble_Click()
    ' INSERT TREBLE LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Treble Line", MusicBoard(fIndex).TrebleLine.Picture, False)
End Sub

Private Sub mnuBass_Click()
    ' INSERT BASS LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Bass Line", MusicBoard(fIndex).BassLine.Picture, False)
End Sub

Private Sub mnuAlto_Click()
    ' INSERT ALTO LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Alto Line", MusicBoard(fIndex).AltoLine.Picture, False)
End Sub

Private Sub mnuTenor_Click()
    ' INSERT TENOR LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Tenor Line", MusicBoard(fIndex).TenorLine.Picture, False)
End Sub

Private Sub mnu8vaTreble_Click()
    ' INSERT 8va TREBLE LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("8va Treble Line", MusicBoard(fIndex).EightvaTrebleLine.Picture, False)
End Sub

Private Sub mnu8vbTreble_Click()
    ' INSERT 8vb TREBLE LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("8vb Treble Line", MusicBoard(fIndex).EightvbTrebleLine.Picture, False)
End Sub

Private Sub mnuDrums_Click()
    ' INSERT DRUM LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Drum Line", MusicBoard(fIndex).DrumLine.Picture, False)
End Sub

Private Sub mnuPercLine_Click()
    ' INSERT PERC LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Perc Line", MusicBoard(fIndex).PercLine.Picture, True)
End Sub

Private Sub mnuNoClef_Click()
    ' INSERT NO CLEF TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("No Clef", LoadPicture(""), False)
End Sub

Private Sub mnuKeyboardTrebleBassLine_Click()
    ' INSERT KEYBOARD (TREBLE+BASS) LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Keyboard (Treble+Bass) Line", MusicBoard(fIndex).KeyboardTrebleBassLine.Picture, True)
End Sub

Private Sub mnuTabLine_Click()
    ' INSERT TAB LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Tab Line", LoadPicture(""), False)
End Sub

Private Sub mnuBassTabLine_Click()
    ' INSERT BASS TAB LINE TO CURRENT MUSIC BOARD INDEX
    Call DrawLines_Initialization("Bass Tab Line", LoadPicture(""), False)
End Sub

Function DrawLines_Initialization(xInsertedLinesType As String, xPicture As Object, xLogic As Boolean)
    ' DRAW LINES TO MUSIC BOARD CURRENT INDEX
    On Error GoTo ErrorDrawLines
    If MaxIdx = 0 Then
        MaxIdx = 1
        nTop = 0
    Else
        If xInsertedLinesType = "Keyboard (Treble+Bass) Line" Then
            nTop = nTop + 3180
            MaxIdx = MaxIdx + 2
        Else
            nTop = nTop + 1590
            MaxIdx = MaxIdx + 1
        End If
    End If
    If MaxIdx = 4 Then
        If xInsertedLinesType = "Keyboard (Treble+Bass) Line" Then
            nTop = 4340
        End If
    ElseIf MaxIdx = 5 Then
        nTop = 6340
    ElseIf MaxIdx = 6 Then
        nTop = 7920
    ElseIf MaxIdx > 6 Then
        MsgBox "Sorry..You already exceeded the number of rows.", vbOKOnly + vbInformation, "Music Editor 2000-Information"
        Exit Function
    End If
    If OmitLines = True Then
        Load MusicBoard(fIndex).FrontEdgeLine(MaxIdx)
        If xInsertedLinesType = "Perc Line" Then
            MusicBoard(fIndex).FrontEdgeLine(MaxIdx).Visible = True
        Else
            MusicBoard(fIndex).FrontEdgeLine(MaxIdx).Visible = False
        End If
        MusicBoard(fIndex).FrontEdgeLine(MaxIdx).Top = nTop
        MusicBoard(fIndex).FrontEdgeLine(MaxIdx).Picture = xPicture
    Else
        Load MusicBoard(fIndex).FrontEdgeLine(MaxIdx)
        MusicBoard(fIndex).FrontEdgeLine(MaxIdx).Visible = True
        MusicBoard(fIndex).FrontEdgeLine(MaxIdx).Top = nTop
        MusicBoard(fIndex).FrontEdgeLine(MaxIdx).Picture = xPicture
    End If
    If xInsertedLinesType = "Keyboard (Treble+Bass) Line" Then
        If OmitLines = False Then
            If nCount = 0 Then
                MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(1).Picture
                InsertedLinesType(0) = xInsertedLinesType
            ElseIf nCount = 1 Then
                MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(3).Picture
                InsertedLinesType(1) = xInsertedLinesType
            ElseIf nCount = 2 Then
                MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(5).Picture
                InsertedLinesType(2) = xInsertedLinesType
            End If
        End If
        OmitNext = True
    Else
        If OmitNext = True Then
            MsgBox "Sorry..You are only allowed to insert Keyboard" & vbCrLf & _
                    "(Treble+Bass) Lines at this time.", vbOKOnly + vbInformation, "Music Editor 2000-Information"
        Else
            If xLogic = True Then
                Load MusicBoard(fIndex).PlainLine(MaxIdx)
                MusicBoard(fIndex).PlainLine(MaxIdx).Visible = True
                MusicBoard(fIndex).PlainLine(MaxIdx).Y1 = nTop + 740
                MusicBoard(fIndex).PlainLine(MaxIdx).Y2 = nTop + 740
                If nCount = 0 Then
                    InsertedLinesType(0) = xInsertedLinesType
                ElseIf nCount = 1 Then
                    InsertedLinesType(1) = xInsertedLinesType
                ElseIf nCount = 2 Then
                    InsertedLinesType(2) = xInsertedLinesType
                ElseIf nCount = 3 Then
                    InsertedLinesType(3) = xInsertedLinesType
                ElseIf nCount = 4 Then
                    InsertedLinesType(4) = xInsertedLinesType
                ElseIf nCount = 5 Then
                    InsertedLinesType(5) = xInsertedLinesType
                End If
                OmitLines = True
            Else
                If OmitLines = False Then
                    If nCount = 0 Then
                        If xInsertedLinesType = "Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(0).Picture
                        ElseIf xInsertedLinesType = "Bass Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(0).Picture
                        Else
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(0).Picture
                        End If
                        InsertedLinesType(0) = xInsertedLinesType
                    ElseIf nCount = 1 Then
                        If xInsertedLinesType = "Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(1).Picture
                        ElseIf xInsertedLinesType = "Bass Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(1).Picture
                        Else
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(1).Picture
                        End If
                        InsertedLinesType(1) = xInsertedLinesType
                    ElseIf nCount = 2 Then
                        If xInsertedLinesType = "Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(2).Picture
                        ElseIf xInsertedLinesType = "Bass Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(2).Picture
                        Else
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(2).Picture
                        End If
                        InsertedLinesType(2) = xInsertedLinesType
                    ElseIf nCount = 3 Then
                        If xInsertedLinesType = "Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(3).Picture
                        ElseIf xInsertedLinesType = "Bass Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(3).Picture
                        Else
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(3).Picture
                        End If
                        InsertedLinesType(3) = xInsertedLinesType
                    ElseIf nCount = 4 Then
                        If xInsertedLinesType = "Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(4).Picture
                        ElseIf xInsertedLinesType = "Bass Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(4).Picture
                        Else
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(4).Picture
                            InsertedLinesType(4) = xInsertedLinesType
                        End If
                    ElseIf nCount = 5 Then
                        If xInsertedLinesType = "Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).TabLine(5).Picture
                        ElseIf xInsertedLinesType = "Bass Tab Line" Then
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BassTabLine(5).Picture
                        Else
                            MusicBoard(fIndex).Picture = MusicBoard(fIndex).BarLine(5).Picture
                        End If
                        InsertedLinesType(5) = xInsertedLinesType
                    End If
                End If
            End If
        End If
    End If
    nCount = nCount + 1
    Exit Function
ErrorDrawLines:
    MsgBox "Error " & Err & ": " & Error, vbOKOnly, "Music Editor 2000-Error"
    Resume Next
End Function

Function ResetSettings(lOmitfIndex As Boolean)
    ' RESET ALL VARIABLE SETTINGS TO 0
    nTop = 0
    MaxIdx = 0
    nTop2 = 0
    MaxIdx2 = 0
    nCount = 0
    xTop = 0
    If lOmitfIndex = True Then
        fIndex = 0
    End If
    OmitLines = False
    OmitNext = False
End Function

Private Sub mnuUseArrowCursor_Click()
    ' SET CURRENT MOUSE CURSOR TO DEFAULT
    Screen.MousePointer = 0
End Sub

