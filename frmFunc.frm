VERSION 5.00
Begin VB.Form frmFunc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Введите аргумент"
   ClientHeight    =   2415
   ClientLeft      =   210
   ClientTop       =   765
   ClientWidth     =   4635
   Icon            =   "frmFunc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4635
   Begin VB.ComboBox cmbFunc 
      Height          =   315
      ItemData        =   "frmFunc.frx":0442
      Left            =   360
      List            =   "frmFunc.frx":0444
      TabIndex        =   5
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   840
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Выберите функцию:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblY 
      Caption         =   "Y="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblX 
      Caption         =   "x="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
End
Attribute VB_Name = "frmFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intcount As Integer
Dim objGetFunc As New GetFunc

Private Sub cmbFunc_Change()
txtY = ""
End Sub

Private Sub cmbFunc_DropDown()
txtY = ""
End Sub

Private Sub cmdOk_Click()
txtX = Replace(txtX, ",", ".")
txtX = Replace(txtX, "pi", 3.14159265358979)
For intcount = 1 To Len(txtX)
    If (Asc(Mid(txtX, intcount, 1)) + 2) \ 10 <> 5 And _
    Mid(txtX, intcount, 1) <> "." And _
    Mid(txtX, intcount, 1) <> "-" _
    Then MsgBox "Неверное число", vbCritical, "Error": Exit Sub
Next intcount
txtY = objGetFunc.GetFun(cmbFunc, Val(txtX))
If txtY = 1000000.3 Then txtY = "не существует"
End Sub

Private Sub txtX_Change()
txtY = ""
End Sub
