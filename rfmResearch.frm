VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form rfmResearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Результаты иследования"
   ClientHeight    =   5175
   ClientLeft      =   3930
   ClientTop       =   3840
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7920
   Begin RichTextLib.RichTextBox rtbResearch 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9128
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"rfmResearch.frx":0000
   End
End
Attribute VB_Name = "rfmResearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objFunc As New GetFunc

Private Sub Form_Load()
rtbResearch.Text = "D(f):"

'пересекает оси
rtbResearch.Text = "пересекает оси:" & vbCrLf & "oX:  "

'критические точки

End Sub
