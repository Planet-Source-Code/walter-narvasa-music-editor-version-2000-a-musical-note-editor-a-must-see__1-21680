VERSION 5.00
Begin VB.Form fGraphicsContainer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Graphics Container"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Symbols Toolbar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "fGraphicsContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Visible = False
End Sub
