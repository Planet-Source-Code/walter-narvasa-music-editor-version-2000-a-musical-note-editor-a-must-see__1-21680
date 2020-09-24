Attribute VB_Name = "mInitialization"
Option Explicit

Global MusicBoard(1000) As New fMusicBoard ' ARRAY OF CHILD FORM OBJECTS
Global CurrentMusicBoardIndex As Integer
Global CurrentKey As String
Global fIndex As Integer
Global KeypressSymbol As String
Global ActiveSymbolKeyVal_(1000)
Global ActiveSymbolKeyTopCoordinates_(1000)
Global ActiveSymbolKeyLeftCoordinates_(1000)
Global KeypressAlternate As String
Global ActiveAlternateKeyVal_(1000)
Global ActiveAlternateKeyTopCoordinates_(1000)
Global ActiveAlternateKeyLeftCoordinates_(1000)
Global InsertedLinesType(5)
Global ActiveTextLineVal_(30)

' SET FORM ON TOP Declare our API functions
Declare Function SetWindowPos Lib "user32" ( _
                                ByVal hWnd As Long, _
                                ByVal hWndInsertAfter As Long, _
                                ByVal x As Long, ByVal Y As Long, _
                                ByVal cx As Long, ByVal cy As Long, _
                                ByVal wFlags As Long) As Long
'Declare our constants
'SWP stands for SetWindowPos
'SWP_NoSize tells SetWindowPos to ignore the cx and cy arguments
Private Const SWP_NOSIZE = &H1
'SWP_NoMove tells SetWindowPos to ignore the x and y arguments.
Private Const SWP_NOMOVE = &H2
'HWND_TOPMOST is passed to SetWindowPos to set the target window Always On Top.
Private Const HWND_TOPMOST = -1
'HWN_NOTOPMOST is passed to SetWindowPos to remove Always on Top
Private Const HWND_NOTOPMOST = -2
'Declare our variables

' FOR READING/OPENING AND WRITING/SAVING USE DECLARING API FUNCTIONS
#If Win16 Then
    Public Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
    Public Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
#Else
    Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If
   
Function OpenMusicFile(Section, KeyName, filename As String) As String
    ' OPEN SAVED MUSIC EDITOR 2000 FILE "*.MUS" EXTENSION
    Dim sRet As String
    sRet = String(255, Chr(0))
    OpenMusicFile = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Function SaveMusicFile(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    ' SAVE MUSIC EDITOR 2000 FILE TO "*.MUS" EXTENSION
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function


Public Sub SetFormOnTop(myForm As Object)
    ' SET FORM ON TOP functions
     SetWindowPos myForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Function PushKey(nKey As Integer, xArray As Integer)
' FOR SYMBOLS TOOLBAR USAGE (SCREEN LAYOUT)
    Dim i
    If nKey = 0 Then
        For i = 0 To 82
            fSymbolsToolbar.SymbolKeyCode(i).BorderStyle = 0
        Next i
        fSymbolsToolbar.SymbolKeyCode(xArray).BorderStyle = 1
    Else
        For i = 0 To 175
            fAlternateToolbar.AlternateKeyCode(i).BorderStyle = 0
        Next i
        fAlternateToolbar.AlternateKeyCode(xArray).BorderStyle = 1
    End If
End Function

Function FileExists(strFile As String) As String
    ' FILE DUPLICATE VALIDATION
    On Error Resume Next
    FileExists = Dir(strFile, vbHidden) <> ""
End Function

Function ExtractArgument(ArgNum As Integer, srchstr As String, Delim As String) As String
    'EXTRACT AN ARGUMENT OR TOKEN FROM A STRING BASED ON ITS POSITION AND A DELIMETER.
    On Error GoTo Err_ExtractArgument
    Dim ArgCount As Integer
    Dim LastPos As Integer
    Dim Pos As Integer
    Dim Arg As String
    Arg = ""
    LastPos = 1
    If ArgNum = 1 Then Arg = srchstr
        Do While InStr(srchstr, Delim) > 0
            Pos = InStr(LastPos, srchstr, Delim)
            If Pos = 0 Then
                If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
                Exit Do
            Else
                ArgCount = ArgCount + 1
                If ArgCount = ArgNum Then
                    Arg = Mid(srchstr, LastPos, Pos - LastPos)
                    Exit Do
                End If
            End If
        LastPos = Pos + 1
    Loop
    '---------
    ExtractArgument = Arg
    Exit Function
Err_ExtractArgument:
    MsgBox "Error " & Err & ": " & Error, vbOKOnly, "Music Editor 2000-Error"
    Resume Next
End Function

