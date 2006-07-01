VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Worcking..."
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prbPer 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnWork As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    blnWork = False
End Sub

Private Sub Form_Load()
    blnWork = True
End Sub
