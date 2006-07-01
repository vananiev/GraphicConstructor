VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetFunc"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intPos As Integer
Dim intSkb As Integer
Dim intCount As Integer
Dim FOper As Byte
Dim SOper As Byte

Public Function GetFun(strFunc As String, X As Double) As Double
X = Round(X, 12)
strFunc = Replace(strFunc, " ", "")
strFunc = Replace(strFunc, "x", X)
strFunc = Replace(strFunc, "--", "+")
strFunc = Replace(strFunc, "pi", 3.14159265358979)
If Mid(strFunc, 1, 1) = "+" Then strFunc = Right(strFunc, Len(strFunc) - 1)
'поиск скобок
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
    If intSkb = 0 Then strFunc = Replace(strFunc, Mid(strFunc, intPos, intCount - intPos + 1), Round(GetFun(Mid(strFunc, intPos + 1, intCount - intPos - 1), X), 12)): GoTo FndSkb
Next intCount

'поиск модулей
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
    If intSkb = 0 Then strFunc = Replace(strFunc, Mid(strFunc, intPos, intCount - intPos + 1), Abs(Round(GetFun(Mid(strFunc, intPos + 1, intCount - intPos - 1), X), 12))): GoTo FndMdl
Next intCount

'поиск тригонометрических функций
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
        'рассчет длинны операндов по Аскил-коду
        SOper = SOperat(strFunc, intPos)
        strFunc = Replace(strFunc, Mid(strFunc, intPos - 2, 3 + SOper), Round(Sin(Val(Mid(strFunc, intPos + 1, SOper))), 12))
    Else
fndCos:
        intPos = InStr(strFunc, "cos") + 2
        If intPos = 2 Then GoTo FndSin
        'рассчет длинны операндов по Аскил-коду
        SOper = SOperat(strFunc, intPos)
        strFunc = Replace(strFunc, Mid(strFunc, intPos - 2, 3 + SOper), Round(Cos(Val(Mid(strFunc, intPos + 1, SOper))), 12))
    End If
    
Else
FndTgCtg:
    intPos = InStr(strFunc, "t")
    If intPos = 0 Then GoTo FndCosSin
    If InStr(strFunc, "ctg") > InStr(strFunc, "tg") Then
FndTg:
        intPos = InStr(strFunc, "tg") + 1
        If intPos = 1 Then GoTo FndCtg
        'рассчет длинны операндов по Аскил-коду
        SOper = SOperat(strFunc, intPos)
        If Cos(Val(Mid(strFunc, intPos + 1, SOper))) = 0 Then GetFun = 1000000.3: Exit Function
        strFunc = Replace(strFunc, Mid(strFunc, intPos - 1, 2 + SOper), Round(Sin(Val(Mid(strFunc, intPos + 1, SOper))) / Cos(Val(Mid(strFunc, intPos + 1, SOper))), 12))
    Else
FndCtg:
        intPos = InStr(strFunc, "ctg") + 2
        If intPos = 2 Then GoTo FndTg
        'рассчет длинны операндов по Аскил-коду
        SOper = SOperat(strFunc, intPos)
        If Sin(Val(Mid(strFunc, intPos + 1, SOper))) = 0 Then GetFun = 1000000.3: Exit Function
        strFunc = Replace(strFunc, Mid(strFunc, intPos - 2, 3 + SOper), Round(Cos(Val(Mid(strFunc, intPos + 1, SOper))) / Sin(Val(Mid(strFunc, intPos + 1, SOper))), 12))
    End If
End If
GoTo FndTrg

'поиск степени
FndStp:
intPos = InStr(strFunc, "^")
If intPos = 0 Then GoTo FndMD
 
 'рассчет длинны операндов по Аскил-коду
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
    If Val(Mid(strFunc, intPos + 1, SOper)) < 1 And Val(Mid(strFunc, intPos - FOper, FOper)) < 0 Then GetFun = 1000000.3: Exit Function
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) ^ Val(Mid(strFunc, intPos + 1, SOper)), 12))
GoTo FndStp

'поиск и замена умн/дел
FndMD:
If InStr(strFunc, "/") = 0 And InStr(strFunc, "*") = 0 Then GoTo FndSS
If InStr(strFunc, "/") > InStr(strFunc, "*") Then
FndMul:
    intPos = InStr(strFunc, "*")
    If intPos = 0 Then GoTo FndDev
    'рассчет длинны операндов по Аскил-коду
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) * Val(Mid(strFunc, intPos + 1, SOper)), 12))
Else
FndDev:
    intPos = InStr(strFunc, "/")
    If intPos = 0 Then GoTo FndMul
     'рассчет длинны операндов по Аскил-коду
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
    If Val(Mid(strFunc, intPos + 1, SOper)) = 0 Then GetFun = 1000000.3: Exit Function
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) / Val(Mid(strFunc, intPos + 1, SOper)), 12))
End If
GoTo FndMD

'поиск слож/выч
FndSS:
If Mid(strFunc, 1, 1) = "-" Then strFunc = Right(strFunc, Len(strFunc) - 1): strFunc = "~" & strFunc
If InStr(strFunc, "+") = 0 And InStr(strFunc, "-") = 0 Then GoTo Fnd
If InStrRev(strFunc, "-") > InStrRev(strFunc, "+") Then
FndSumm:
    intPos = InStr(strFunc, "+")
    If intPos = 0 Then GoTo FndSub
    If Mid(strFunc, 1, 1) = "~" Then strFunc = Replace(strFunc, "~", "-")
     'рассчет длинны операндов по Аскил-коду
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) + Val(Mid(strFunc, intPos + 1, SOper)), 12))
Else
FndSub:
    intPos = InStr(strFunc, "-")
    If intPos = 0 Then GoTo FndSumm
    If Mid(strFunc, 1, 1) = "~" Then strFunc = Replace(strFunc, "~", "-")
     'рассчет длинны операндов по Аскил-коду
    FOper = FOperat(strFunc, intPos)
    SOper = SOperat(strFunc, intPos)
If FOper = 0 Then
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(0 - Val(Mid(strFunc, intPos + 1, SOper)), 12))
Else
    strFunc = Replace(strFunc, Mid(strFunc, intPos - FOper, FOper + SOper + 1), Round(Val(Mid(strFunc, intPos - FOper, FOper)) - Val(Mid(strFunc, intPos + 1, SOper)), 12))
End If
End If
GoTo FndSS

Fnd:
If Mid(strFunc, 1, 1) = "~" Then strFunc = Replace(strFunc, "~", "-")
strFunc = Replace(strFunc, ",", ".")
GetFun = Round(Val(strFunc), 12)
End Function


Private Function FOperat(strFunc As String, intPos As Integer) As Byte
'рассчет длинны 1 операндa по Аскил-коду
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

Private Function SOperat(strFunc As String, intPos As Integer) As Byte
'рассчет длинны 2 операндa по Аскил-коду
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
strFunc = Replace(strFunc, "pi", "")
strFunc = Replace(strFunc, "x", "")
strFunc = Replace(strFunc, "^", "")
strFunc = Replace(strFunc, "sin", "")
strFunc = Replace(strFunc, "cos", "")
strFunc = Replace(strFunc, "ctg", "")
strFunc = Replace(strFunc, "tg", "")
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
If intSkb <> 0 Then MsgBox "Неверное вложение скобок", vbCritical, "Error of function": GetFunCheck = False: Exit Function
If intMod <> 0 Then MsgBox "Неверное использование модулей", vbCritical, "Error of function": GetFunCheck = False: Exit Function
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

