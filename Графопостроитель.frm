VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Constr 
   AutoRedraw      =   -1  'True
   Caption         =   "Graphic Constructor"
   ClientHeight    =   8190
   ClientLeft      =   1845
   ClientTop       =   675
   ClientWidth     =   11880
   Icon            =   "Графопостроитель.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10100
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFunc 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   11895
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   480
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   100
      DialogTitle     =   "Сохранить как"
      FileName        =   "График_01"
      Filter          =   ".bmp(точечный рисунок)"
      Flags           =   4100
      FontBold        =   -1  'True
      MaxFileSize     =   26000
      Orientation     =   2
   End
   Begin VB.TextBox txtOt 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   8040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtArg 
      Height          =   285
      Left            =   6600
      TabIndex        =   3
      Text            =   "1"
      Top             =   7560
      Width           =   585
   End
   Begin VB.TextBox txtMash 
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Text            =   "1"
      Top             =   7560
      Width           =   840
   End
   Begin VB.TextBox txtFunc 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   7560
      Width           =   3735
   End
   Begin VB.CommandButton cmdPG 
      Caption         =   "Построить график"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prbWork 
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.UpDown updMash 
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Top             =   7560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      OrigLeft        =   5640
      OrigTop         =   8280
      OrigRight       =   5895
      OrigBottom      =   8535
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updArg 
      Height          =   285
      Left            =   7185
      TabIndex        =   7
      Top             =   7560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtFunc"
      BuddyDispid     =   196613
      OrigLeft        =   7800
      OrigTop         =   8280
      OrigRight       =   8055
      OrigBottom      =   8565
      Enabled         =   -1  'True
   End
   Begin VB.Line LineX 
      BorderColor     =   &H0080FF80&
      X1              =   120.202
      X2              =   11900
      Y1              =   119.621
      Y2              =   119.621
   End
   Begin VB.Line LineY 
      BorderColor     =   &H0080FF80&
      X1              =   121.212
      X2              =   121.212
      Y1              =   119.621
      Y2              =   7800.061
   End
   Begin VB.Label Label4 
      Caption         =   "Введите функцию:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label lblOt 
      Caption         =   "Ответы:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblY 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   12
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblY 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   11
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label lblX 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   10
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "X:Y"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Укажите масштаб:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Menu Fail 
      Caption         =   "Файл"
      Begin VB.Menu SaveAs 
         Caption         =   "Сохранить как"
      End
      Begin VB.Menu a1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Правка"
      Begin VB.Menu Setk 
         Caption         =   "Сетка"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Constr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Double
Dim intcount As Integer
Dim intPrm(1) As Double
Dim txtfuncc As String
Dim objFunc As New GetFunc
Dim objFuncNe As New GetFuncNe
Dim blnFunc As Boolean
Dim btColFunc As Byte
Dim stroper As String
Dim dblCX As Double
Dim dblCY As Double
Dim dblMin As Double
Dim dblMax As Double
Dim dblStep As Double
Dim intExact As Integer
Dim blnSet As Byte 'сетка
Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal X As Long, ByVal у As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long

Private Sub cmdPG_Click()
Dim intOper As Integer
txtOt.Visible = False
lblOt.Visible = False
If InStr(txtFunc, ">=") <> 0 Then intOper = 1: stroper = ">=": GoTo EndSearch
If InStr(txtFunc, "<=") <> 0 Then intOper = 1: stroper = "<=": GoTo EndSearch
If InStr(txtFunc, "<>") <> 0 Then intOper = 1: stroper = "<>": GoTo EndSearch
If InStr(txtFunc, "<") <> 0 Then intOper = 1: stroper = "<": GoTo EndSearch
If InStr(txtFunc, ">") <> 0 Then intOper = 1: stroper = ">": GoTo EndSearch
If InStr(txtFunc, "=") <> 0 Then intOper = 1: stroper = "=": GoTo EndSearch
EndSearch:
If intOper = 1 Then
    'проверка функции
    If txtFunc = "" Then MsgBox "Неизвестная функция", vbCritical, "Error of function": Exit Sub
    If Not (objFuncNe.GetFunCheck(txtFunc)) Then MsgBox "Неизвестная функция", vbCritical, "Error of function": Exit Sub
    PointNe (stroper)
    btColFunc = 2
    lblY(1).Visible = True
    blnFunc = True
    Constr.Picture = Constr.Image
Else
    'проверка функции
    If txtFunc = "" Then MsgBox "Неизвестная функция", vbCritical, "Error of function": Exit Sub
    If Not (objFunc.GetFunCheck(txtFunc)) Then MsgBox "Неизвестная функция", vbCritical, "Error of function": Exit Sub
    PointGraph
    btColFunc = 1
    lblY(1).Visible = False
    blnFunc = True

End If
End Sub

Private Sub Cetca()
Cls
Picture = Nothing
Line (120, 4000)-(11860, 4000)
Line (6000, 120)-(6000, 7800)
txtOt.Visible = False
lblOt.Visible = False
lblY(1).Visible = False
Dim lgCount As Long
'центр (6000;4000)
For lgCount = 2 To 58
Line (lgCount * 200 + 10, 3990)-(lgCount * 200 + 10, 4040)
If lgCount - (lgCount \ 2) * 2 = 0 Then
If lgCount = 30 Then GoTo 2
PSet (lgCount * 200 - 100, 4040), &HFFFFFF
Print (lgCount - 30) * Val(txtMash) * Val(txtArg) + dblCX
End If
2:
Next lgCount
For lgCount = 2 To 38
Line (5980, lgCount * 200 + 10)-(6030, lgCount * 200 + 10)
If lgCount - (lgCount \ 2) * 2 = 0 Then
If lgCount = 20 Then GoTo 1
PSet (6040, lgCount * 200 - 100), &HFFFFFF
Print (20 - lgCount) * Val(txtMash) + dblCY
1:
End If
Next lgCount
blnFunc = False
End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyWord KeyCode, Shift
End Sub

Private Sub Form_Load()
Dim n As Long
Randomize Timer
' показываем эту форму
Show
'показываем экран - заставку
frmSplash.Show
DoEvents
For n = -200000 To 20000: Print "": Next n
' удаляем экран-застааку
Unload frmSplash
dblMin = -250
dblMax = 250
dblStep = 0.1
intExact = 8
blnSet = True
Setk.Checked = blnSet
Cetca
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dblX As Double
Dim dblY As Double
Dim intPar As Integer
Dim strFn() As String
If blnSet And Y < 7800 And Y > 100 And X > 100 And X < 11900 Then
    LineX.Visible = True
    LineY.Visible = True
    LineX.Y1 = Y
    LineX.Y2 = Y
    LineY.X1 = X
    LineY.X2 = X
Else
    LineX.Visible = False
    LineY.Visible = False
End If
dblX = ((X - 10) / 200 - 30) * Val(txtMash) * Val(txtArg) + dblCX
lblX = "X: " & Str(dblX)
If dblX < -28 * Val(txtMash) * Val(txtArg) + dblCX Or dblX > 28 * Val(txtMash) * Val(txtArg) + dblCX Or Not (blnFunc) Then lblY(0) = "Y:": lblY(1) = "Y:": lblX = "X:": Exit Sub
For intcount = 0 To btColFunc - 1
strFn = Split(txtFunc, stroper)
lblX = "X: " & Str(dblX)
dblY = objFunc.GetFun(strFn(intcount), dblX)
If dblY = 1000000.3 Then
   lblY(intcount) = "Y: несуществует"
Else
   lblY(intcount) = "Y: " & Str(dblY)
End If
Next intcount
End Sub

Private Sub Form_Resize()
    On Error GoTo Ext
    Scale (0, 0)-(12000, 10100)
    Label4.Top = 6960 + 1180
    Label3.Top = 6960 + 1180
    Label2.Top = 8140
    lblOt.Top = 7920 + 1180
    lblX.Top = 7080 + 1180
    lblY(0).Top = 7440 + 1180
    lblY(1).Top = 7800 + 1180
    prbWork.Top = 8280 + 1080
    txtArg.Top = 7440 + 1180
    txtFunc.Top = 7440 + 1180
    txtOt.Top = 7920 + 1180
    updArg.Top = 7440 + 1180
    updMash.Top = 7440 + 1180
    txtMash.Top = 7440 + 1180
    cmdPG.Top = 7560 + 1180
    LineX.X2 = 11900
    LineY.Y2 = 7800
    picFunc.Height = 8200
    picFunc.Width = 12000
    Cetca
Ext:
End Sub

Private Sub SaveAs_Click()
    LineX.Visible = False
    LineY.Visible = False
    cdSave.FileName = txtFunc
    cdSave.ShowSave
    If cdSave.FileName = "" Then GoTo ExitSub
    picFunc.Cls
    picFunc.Picture = Nothing
    BitBlt picFunc.hDC, 0, 0, picFunc.ScaleWidth, picFunc.ScaleHeight, Constr.hDC, 0, 0, vbSrcCopy
    SavePicture picFunc.Image, cdSave.FileName & ".bmp"
ExitSub:
    LineX.Visible = blnSet
    LineY.Visible = blnSet
End Sub

Private Sub Setk_Click()
    blnSet = Not (blnSet)
    Setk.Checked = blnSet
End Sub

Private Sub txtFunc_Change()
blnFunc = False
txtOt.Visible = False
lblOt.Visible = False
lblY(1).Visible = False
End Sub

Private Sub txtFunc_KeyDown(KeyCode As Integer, Shift As Integer)
KeyWord KeyCode, Shift
End Sub

Private Sub updArg_DownClick()
If Val(txtArg) <= 1 Then
    txtArg = Str(Val(txtArg) / 2)
Else
    txtArg = Val(txtArg) - 1
End If
Cetca
End Sub

Private Sub updArg_UpClick()
If Val(txtArg) <= 1 Then
    txtArg = Str(Val(txtArg) * 2)
Else
    txtArg = Val(txtArg) + 1
End If
Cetca
End Sub

Private Sub updMash_DownClick()
If Val(txtMash) <= 1 Then
    txtMash = Str(Val(txtMash) / 2)
Else
    txtMash = Val(txtMash) - 1
End If
Cetca
End Sub

Private Sub updMash_UpClick()
If Val(txtMash) <= 1 Then
    txtMash = Str(Val(txtMash) * 2)
Else
    txtMash = Val(txtMash) + 1
End If
Cetca
End Sub

Private Sub PointGraph()
'Чертим систему координат
Cetca
MousePointer = vbHourglass
prbWork.Visible = True
intPrm(1) = 1000000.3
For intcount = -5600 To 5600
Rem" блок вычислений и построения графика"
X = intcount * Val(txtMash) * Val(txtArg) / 200 + dblCX
prbWork.Value = (intcount + 5600) / 112
intPrm(0) = intPrm(1)
intPrm(1) = objFunc.GetFun(txtFunc, X) - dblCY
If intPrm(1) < -18 * Val(txtMash) Or intPrm(1) > 18 * Val(txtMash) Or intPrm(0) < -18 * Val(txtMash) Or intPrm(0) > 18 * Val(txtMash) Then
Else
    Line (6000 + ((X - dblCX) - (0.005 * Val(txtMash) * Val(txtArg))) / (0.005 * Val(txtMash) * Val(txtArg)), 4000 - intPrm(0) / (0.005 * Val(txtMash)))-(6000 + (X - dblCX) / (0.005 * Val(txtMash) * Val(txtArg)), 4000 - intPrm(1) / (0.005 * Val(txtMash)))
End If
Next intcount
prbWork.Value = 0
prbWork.Visible = False
MousePointer = vbDefault
DoEvents
End Sub

Private Sub PointNe(stroper As String)
Cetca
Dim strFn() As String
Dim strWork As String
Dim intCnt As Integer
If txtFunc = "" Then MsgBox "Неизвестная функция", vbCritical, "Error of function": Exit Sub
MousePointer = vbHourglass
prbWork.Visible = True
For intCnt = 0 To 1
intPrm(1) = 1000000.3
For intcount = -5600 To 5600
Rem" блок вычислений и построения графика"
strFn = Split(txtFunc, stroper)
X = intcount * Val(txtMash) * Val(txtArg) / 200 + dblCX
prbWork.Value = (intcount + 5600) / 224 + 50 * intCnt
intPrm(0) = intPrm(1)
intPrm(1) = objFunc.GetFun(strFn(intCnt), X) - dblCY
If intPrm(1) < -18 * Val(txtMash) Or intPrm(1) > 18 * Val(txtMash) Or intPrm(0) < -18 * Val(txtMash) Or intPrm(0) > 18 * Val(txtMash) Then
Else
    Line (6000 + ((X - dblCX) - (0.005 * Val(txtMash) * Val(txtArg))) / (0.005 * Val(txtMash) * Val(txtArg)), 4000 - intPrm(0) / (0.005 * Val(txtMash)))-(6000 + (X - dblCX) / (0.005 * Val(txtMash) * Val(txtArg)), 4000 - intPrm(1) / (0.005 * Val(txtMash)))
End If
Next intcount
Next intCnt
strWork = objFuncNe.GetFun(txtFunc, dblMin, dblMax, dblStep, intExact)
MsgBox strWork, vbOKOnly, "Ответы"
intcount = InStrRev(strWork, vbCrLf)
txtOt = Right(strWork, Len(strWork) - intcount - 1)
If Len(txtOt) = 0 Then txtOt = "нет корней"
txtOt.Visible = True
lblOt.Visible = True
prbWork.Value = 0
prbWork.Visible = False
MousePointer = vbDefault
DoEvents
End Sub

Private Sub KeyWord(KeyCode As Integer, Shift As Integer)
Dim strWork() As String
If KeyCode = 112 Then MsgBox "F2-значения функции" & vbCrLf & "F3-установки", vbInformation, "Information"
If KeyCode = 113 Then
    If stroper <> "" Then
        strWork = Split(txtFunc, stroper)
        For intcount = 0 To UBound(strWork)
        frmFunc.cmbFunc.List(intcount) = strWork(intcount)
        Next intcount
    Else
      frmFunc.cmbFunc.List(0) = txtFunc
    End If
frmFunc.Show
End If
If KeyCode = 114 Then
    frmCentr.txtCX = dblCX
    frmCentr.txtCY = dblCY
    frmCentr.txtStep = dblStep
    frmCentr.txtMin = dblMin
    frmCentr.txtMax = dblMax
    frmCentr.txtExact = intExact
    frmCentr.Show vbModal
    dblCX = Val(frmCentr.txtCX)
    dblCY = Val(frmCentr.txtCY)
    dblStep = Val(frmCentr.txtStep)
    dblMin = Val(frmCentr.txtMin)
    dblMax = Val(frmCentr.txtMax)
    intExact = Val(frmCentr.txtExact)
    Unload frmCentr
End If

End Sub
