VERSION 5.00
Begin VB.Form frmCentr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Установки"
   ClientHeight    =   2310
   ClientLeft      =   345
   ClientTop       =   735
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   7035
   Begin VB.TextBox txtExact 
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtStep 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   5280
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtCY 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtCX 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Точность:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Шаг:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Mах функции:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Min функции:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Центр:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   135
   End
End
Attribute VB_Name = "frmCentr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intcount As Integer

Private Sub cmdOk_Click()
txtCX = Replace(txtCX, ",", ".")
txtCY = Replace(txtCY, ",", ".")
txtStep = Replace(txtStep, ",", ".")
txtMin = Replace(txtMin, ",", ".")
txtMax = Replace(txtMax, ",", ".")
txtExact = Replace(txtExact, ",", ".")
For intcount = 1 To Len(txtCX)
    If (Asc(Mid(txtCX, intcount, 1)) + 2) \ 10 <> 5 And Mid(txtCX, intcount, 1) <> "." And Mid(txtCX, intcount, 1) <> "-" Then MsgBox "Неверное число", vbCritical, "Error": Exit Sub
Next intcount
For intcount = 1 To Len(txtCY)
    If (Asc(Mid(txtCY, intcount, 1)) + 2) \ 10 <> 5 And Mid(txtCY, intcount, 1) <> "." And Mid(txtCY, intcount, 1) <> "-" Then MsgBox "Неверное число", vbCritical, "Error": Exit Sub
Next intcount
For intcount = 1 To Len(txtStep)
    If (Asc(Mid(txtStep, intcount, 1)) + 2) \ 10 <> 5 And Mid(txtStep, intcount, 1) <> "." Then MsgBox "Неверное число", vbCritical, "Error": Exit Sub
Next intcount
For intcount = 1 To Len(txtMin)
    If (Asc(Mid(txtMin, intcount, 1)) + 2) \ 10 <> 5 And Mid(txtMin, intcount, 1) <> "." And Mid(txtMin, intcount, 1) <> "-" Then MsgBox "Неверное число", vbCritical, "Error": Exit Sub
Next intcount
For intcount = 1 To Len(txtMax)
    If (Asc(Mid(txtMax, intcount, 1)) + 2) \ 10 <> 5 And Mid(txtMax, intcount, 1) <> "." And Mid(txtMax, intcount, 1) <> "-" Then MsgBox "Неверное число", vbCritical, "Error": Exit Sub
Next intcount
For intcount = 1 To Len(txtExact)
    If (Asc(Mid(txtExact, intcount, 1)) + 2) \ 10 <> 5 Then MsgBox "Неверное число", vbCritical, "Error": Exit Sub
Next intcount
Hide
End Sub

