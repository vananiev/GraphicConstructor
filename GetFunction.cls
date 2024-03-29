VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim intPos As Integer
Dim intSkb As Integer
Dim intCount As Integer
Dim FOper As Byte
Dim SOper As Byte

Public Function GetFun(strFunc As String, X As Double) As Double
X = Round(X, 9)
strFunc = Replace(strFunc, "x", X)
strFunc = Replace(strFunc, "--", "+")
If Mid(strFunc, 1, 1) = "+" Then strFunc = Right(strFunc, Len(strFunc) - 1)
'����� ������
FndSkb:
intPos = InStr(strFunc, "(")
If intPos = 0 Then
    GoTo FndMdl
Else
    intSkb = 1
End If
For intCount = intPos + 1 To Len(strFunc)
    If Mid(strFunc, intCount, 1) = "(" Then intSkb = intSkb + 1
    If Mid(strFunc, intCount, 1) = ")" Then intSkb = intSkb - 1
    If intSkb = 0 Then strFunc = Replace(strFunc, Mid(strFunc, intPos, intCount - intPos + 1), Round(GetFun(Mid(strFunc, intPos + 1, intCount - intPos - 1), X), 9)): GoTo FndSkb
Next intCount

'����� �������
FndMdl:
intPos = InStr(strFunc, "[")
If intPos = 0 Then
    GoTo FndTrg
Else
    intSkb = 1
End If
For intCount = intPos + 1 To Len(strFunc)
    If Mid(strFunc, intCount, 1) = "[" Then intSkb = intSkb + 1
    If Mid(strFunc, intCount, 1) = "]" Then intSkb = intSkb - 1
    If intSkb = 0 Then strFunc = Replace(strFunc, Mid(strFunc, intPos, intCount - intPos + 1), Abs(Round(GetFun(Mid(strFunc, intPos + 1, intCount - intPos - 1), X), 9))): GoTo FndMdl
Next intCount

'����� ������������������ �������
FndTrg:
If InStr(strFunc, "sin") = 0 And InStr(strFunc, "cos") = 0 And InStr(strFunc, "tg") = 0 And InStr(strFunc, "ctg") = 0 Then GoTo FndStp
If InStr(strFunc, "t") > InStr(strFunc, "s") Then
FndCosSin:
    intPos = InStr(strFunc, "s")
    If intPos = 0 Then GoTo FndTgCtg
    If InStr(strFunc, "cos") > InStr(strFunc, "sin") Then
FndSin:
        intPos = InStr(strFunc, "sin") + 2
        If intPos = 2 Then GoTo fndCos
        '������� ������ ��������� �� �����-����
        SOper = SOperat(strFunc, intPos)
        strFunc = Replace(strFunc, Mid(strFunc, intPos - 2, 3 + SOper), Round(Sin(Val(Mid(strFunc, intPos + 1, SOper))), 9))
    Else
fndCos:
        intPos = InStr(strFunc, "cos") + 2
        If intPos = 2 Then GoTo FndSin
        '������� ������ ��������� �� �����-����
        SOper = SOperat(strFunc, intPos)
        strFunc = Replace(strFunc, Mid(strFunc, intPos - 2, 3 + SOper), Round(Cos(Val(Mid(strFunc, intPos + 1, SOper))), 9))
    End If
    
Else
FndTgCtg:
    intPos = InStr(strFunc, "t")
    If intPos = 0 Then GoTo FndCosSin
    If InStr(strFunc, "ctg") > InStr(strFunc, "tg") Then
FndTg:
        intPos = InStr(strFunc, "tg") + 1
        If intPos = 1 Then GoTo FndCtg
        '������� ������ ��������� �� �����-����
        SOper = SOperat(strFunc, intPos)
        If Cos(Val(Mid(strFunc, intPos + 1, SOper))) = 0 Then GetFun = 10 ^ 6: Exit Function
        strFunc = Replace(strFunc, Mid(strFunc, intPos - 1, 2 + SOper), Round(Sin(Val(Mid(strFunc, intPos + 1, SOper))) / Cos(Val(Mid(strFunc, intPos + 1, SOper))), 9))
    Else
FndCtg:
        intPos = InStr(strFunc, "ctg") + 2
        If intPos = 2 Then GoTo FndTg
        '������� ������ ��������� �� �����-����
        SOper = SOperat(strFunc, intPos)
        If Sin(Val(Mid(strFunc, intPos + 1, SOper))) = 0 Then GetFun = 10 ^ 6: Exit Function
        strFunc = Replace(strFunc, Mid(strFunc, intPos - 2, 3 + SOper), Round(Cos(Val(Mid(strFunc, intPos + 1, SOper))) / Sin(Val(Mid(strFunc, intPos + 1, SOper))), 9))
    End If
End If
GoTo FndTrg

'����� �������
FndStp:
intPos = InStr(strFunc, "^")
If intPos = 0 Then GoTo FndMD
 
 '������� ������ ��������� �� �����-����
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
    If Val(Mid(strFunc, intPos + 1, SOper)) < 1 And Val(Mid(strFunc, intPos - FOper, FOper)) < 0 Then GetFun = 0: Exit Function
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) ^ Val(Mid(strFunc, intPos + 1, SOper)), 9))
GoTo FndStp

'����� � ������ ���/���
FndMD:
If InStr(strFunc, "/") = 0 And InStr(strFunc, "*") = 0 Then GoTo FndSS
If InStr(strFunc, "/") > InStr(strFunc, "*") Then
FndMul:
    intPos = InStr(strFunc, "*")
    If intPos = 0 Then GoTo FndDev
    '������� ������ ��������� �� �����-����
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) * Val(Mid(strFunc, intPos + 1, SOper)), 9))
Else
FndDev:
    intPos = InStr(strFunc, "/")
    If intPos = 0 Then GoTo FndMul
     '������� ������ ��������� �� �����-����
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
    If Val(Mid(strFunc, intPos + 1, SOper)) = 0 Then GetFun = 10 ^ 6: Exit Function
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) / Val(Mid(strFunc, intPos + 1, SOper)), 9))
End If
GoTo FndMD

'����� ����/���
FndSS:
If Mid(strFunc, 1, 1) = "-" Then strFunc = Right(strFunc, Len(strFunc) - 1): strFunc = "~" & strFunc
If InStr(strFunc, "+") = 0 And InStr(strFunc, "-") = 0 Then GoTo Fnd
If InStrRev(strFunc, "-") > InStrRev(strFunc, "+") Then
FndSumm:
    intPos = InStr(strFunc, "+")
    If intPos = 0 Then GoTo FndSub
    If Mid(strFunc, 1, 1) = "~" Then strFunc = Replace(strFunc, "~", "-")
     '������� ������ ��������� �� �����-����
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) + Val(Mid(strFunc, intPos + 1, SOper)), 9))
Else
FndSub:
    intPos = InStr(strFunc, "-")
    If intPos = 0 Then GoTo FndSumm
    If Mid(strFunc, 1, 1) = "~" Then strFunc = Replace(strFunc, "~", "-")
     '������� ������ ��������� �� �����-����
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
If FOper = 0 Then
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(0 - Val(Mid(strFunc, intPos + 1, SOper)), 9))
Else
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) - Val(Mid(strFunc, intPos + 1, SOper)), 9))
End If
End If
GoTo FndSS

Fnd:
If Mid(strFunc, 1, 1) = "~" Then strFunc = Replace(strFunc, "~", "-")
strFunc = Replace(strFunc, ",", ".")
GetFun = Val(strFunc)
End Function


Public Function FOperat(strFunc As String, intPos As Integer) As Byte
'������� ������ 1 �������a �� �����-����
Dim FOp As Byte
strFunc = Replace(strFunc, ",", ".")
If intPos = 1 Then FOperat = 0: Exit Function
 intCount = intPos - 1
    Do While (Asc(Mid(strFunc, intCount, 1)) + 2) \ 10 = 5 Or Mid(strFunc, intCount, 1) = "." Or Mid(strFunc, intCount, 1) = "-"
        If Mid(strFunc, intCount, 1) = "-" And Mid(strFunc, intCount + 1, 1) = "-" Then Exit Do
        If (Asc(Mid(strFunc, intCount, 1)) + 2) \ 10 = 5 And Mid(strFunc, intCount + 1, 1) = "-" And intCount + 1 <> intPos Then FOp = FOp - 1: Exit Do
        FOp = FOp + 1
        If intCount = 1 Then Exit Do
        intCount = intCount - 1
    Loop
FOperat = FOp
End Function

Public Function SOperat(strFunc As String, intPos As Integer) As Byte
'������� ������ 2 �������a �� �����-����
Dim SOp As Byte
strFunc = Replace(strFunc, ",", ".")
    intCount = intPos + 1
    Do While (Asc(Mid(strFunc, intCount, 1)) + 2) \ 10 = 5 Or Mid(strFunc, intCount, 1) = "." Or Mid(strFunc, intCount, 1) = "-"
        If intCount <> intPos + 1 And Mid(strFunc, intCount, 1) = "-" Then Exit Do
        SOp = SOp + 1
        If intCount = Len(strFunc) Then Exit Do
        intCount = intCount + 1
    Loop
SOperat = SOp
End Function

Public Function GetFunCheck(strFunc As String) As Boolean
Dim intMod As Integer
intMod = 0
intSkb = 0
strFunc = Replace(strFunc, "x", "")
strFunc = Replace(strFunc, "^", "")
strFunc = Replace(strFunc, "sin", "")
strFunc = Replace(strFunc, "cos", "")
strFunc = Replace(strFunc, "tg", "")
strFunc = Replace(strFunc, "ctg", "")
strFunc = Replace(strFunc, "*", "")
strFunc = Replace(strFunc, "/", "")
strFunc = Replace(strFunc, "+", "")
strFunc = Replace(strFunc, "-", "")
strFunc = Replace(strFunc, ",", "")
strFunc = Replace(strFunc, ".", "")
For intCount = 1 To Len(strFunc)
    If (Asc(Mid(strFunc, intCount, 1)) + 2) \ 10 = 5 Then strFunc = Replace(strFunc, Mid(strFunc, intCount, 1), "+")
    If Mid(strFunc, intCount, 1) = "(" Then intSkb = intSkb + 1
    If Mid(strFunc, intCount, 1) = ")" Then intSkb = intSkb - 1
    If Mid(strFunc, intCount, 1) = "[" Then intMod = intMod + 1
    If Mid(strFunc, intCount, 1) = "]" Then intMod = intMod - 1
Next intCount
If intSkb <> 0 Then MsgBox "�������� �������� ������", vbCritical, "Error of function": GetFunCheck = False: Exit Function
If intMod <> 0 Then MsgBox "�������� ������������� �������", vbCritical, "Error of function": GetFunCheck = False: Exit Function
strFunc = Replace(strFunc, "(", "")
strFunc = Replace(strFunc, ")", "")
strFunc = Replace(strFunc, "[", "")
strFunc = Replace(strFunc, "]", "")
strFunc = Replace(strFunc, "+", "")
If Len(strFunc) = 0 Then
    GetFunCheck = True
Else
    GetFunCheck = False
End If
End Function


